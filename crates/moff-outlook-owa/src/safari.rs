use std::time::Duration;

use cider::sources::safari::{
    self, BinaryDirection, BinaryTransferOptions, DocumentTarget, EvalOptions, OperationPhase,
    PageOperationError, TabTarget,
};
use serde_json::{Value, json};
use url::Url;

use crate::{
    OutlookWebError, OutlookWebSession, SessionTransport, decode_response, validate_read_action,
};

pub(super) struct SafariSession {
    target: DocumentTarget,
    runtime: tokio::runtime::Runtime,
    allow_mutation: bool,
}

fn outlook_url(value: &str) -> bool {
    Url::parse(value).is_ok_and(|url| {
        url.scheme() == "https"
            && url.username().is_empty()
            && url.password().is_none()
            && url.port().is_none()
            && matches!(
                url.host_str(),
                Some("outlook.office.com" | "outlook.cloud.microsoft" | "outlook.live.com")
            )
            && (url.path().starts_with("/mail/")
                || url.path().starts_with("/calendar/")
                || hosted_outlook(&url))
    })
}

fn hosted_outlook(url: &Url) -> bool {
    let segments: Vec<_> = url.path().split('/').collect();
    matches!(segments.as_slice(), ["", "host", id, "auth"]
        if id.len() == 36 && id.bytes().all(|b| b.is_ascii_hexdigit() || b == b'-'))
}

async fn verify_host_runtime(
    tab: &safari::OpenTab,
    target: &DocumentTarget,
) -> Result<(), OutlookWebError> {
    if Url::parse(&tab.url).is_ok_and(|url| hosted_outlook(&url)) {
        let result = safari::eval_document(
            "typeof window.webpackChunkOwa === 'object' && window.webpackChunkOwa !== null",
            target,
            &eval_options(10),
        )
        .await
        .map_err(|_| OutlookWebError::SafariBeforeDispatch)?;
        if result.truncated || result.value != true {
            return Err(OutlookWebError::SafariUnavailable);
        }
    }
    Ok(())
}

impl OutlookWebSession {
    /// Check browser authentication without refreshing or invoking interactive login.
    pub fn safari_auth_status(target: Option<(u32, u32)>) -> Result<Value, OutlookWebError> {
        let session = Self::from_safari(target)?;
        #[cfg(not(test))]
        let SessionTransport::Safari(safari) = &session.transport;
        #[cfg(test)]
        let SessionTransport::Safari(safari) = &session.transport else {
            return Err(OutlookWebError::SafariUnavailable);
        };
        let script = format!(
            "(async function(){{const result=await {};return {{status:result.status,failure:result.failure,recovery_marker_present:sessionStorage.getItem('__moff_auth_recovery')!==null}};}})()",
            page_script(
                include_str!("safari_read.js"),
                &json!({"action":"__MoffAuthStatus"})
            )
        );
        let result = safari
            .runtime
            .block_on(safari::eval_document(
                &script,
                &safari.target,
                &eval_options(30),
            ))
            .map_err(|_| OutlookWebError::SafariAuthRecoveryFailed("silent token acquisition"))?;
        if result.truncated {
            return Err(OutlookWebError::InvalidResponse);
        }
        if result.value["failure"] == "timeout" {
            return Err(OutlookWebError::ReadTimedOut {
                action: "authentication check",
            });
        }
        if result.value["failure"] == "network" {
            return Err(OutlookWebError::ReadNetworkFailed {
                action: "authentication check",
            });
        }
        match result.value["status"].as_u64() {
            Some(status @ (200 | 401 | 422 | 425)) => {
                let (state, message) = match status {
                    200 => (
                        "authenticated",
                        "Silent browser authentication succeeded; this does not prove that a mailbox read will succeed.",
                    ),
                    401 => (
                        "session_expired",
                        "Run moff outlook auth refresh --confirm and finish any sign-in steps yourself in Safari.",
                    ),
                    425 => (
                        "outlook_not_ready",
                        "Open the selected Outlook tab, finish sign-in if prompted, and wait for Mail.",
                    ),
                    _ => (
                        "auth_adapter_changed",
                        "Outlook authentication modules changed; Moff needs an adapter update, not repeated refreshes.",
                    ),
                };
                Ok(json!({"authenticated":result.value["status"] == 200,
                "state":state,"message":message,
                "status":result.value["status"],"recovery_marker_present":result.value["recovery_marker_present"],
                "credential_export":false}))
            }
            _ => Err(OutlookWebError::InvalidResponse),
        }
    }

    /// Explicitly reload Outlook and let browser SSO recover; never interact with login UI.
    pub fn refresh_safari_auth(
        target: Option<(u32, u32)>,
        wait_seconds: u32,
    ) -> Result<Value, OutlookWebError> {
        if !(1..=120).contains(&wait_seconds) {
            return Err(OutlookWebError::InvalidInput(
                "wait seconds must be between 1 and 120",
            ));
        }
        let runtime =
            tokio::runtime::Runtime::new().map_err(|_| OutlookWebError::SafariOperationFailed)?;
        runtime.block_on(async {
            let tabs = safari::tabs().await.map_err(|_| OutlookWebError::SafariOperationFailed)?;
            let tab = selected_tab(&tabs, target)?;
            let original = safari::document_target(tab_target(tab)?).await
                .map_err(|_| OutlookWebError::SafariBeforeDispatch)?;
            verify_host_runtime(tab, &original).await?;
            let marker = original.document_id.clone();
            let script = format!("({})({})", include_str!("safari_refresh.js"), json!(marker));
            let result = safari::eval_document(&script, &original, &eval_options(10)).await
                .map_err(|_| OutlookWebError::SafariRefreshUnknown)?;
            if result.truncated || result.value["scheduled"] != true {
                return Err(OutlookWebError::SafariRefreshUnknown);
            }
            let mut last_status = None;
            let wait = tokio::time::timeout(Duration::from_secs(u64::from(wait_seconds)), async {
                'poll: loop {
                    tokio::time::sleep(Duration::from_secs(2)).await;
                    let tabs = safari::tabs().await
                        .map_err(|_| OutlookWebError::SafariOperationFailed)?;
                    let mut candidates = Vec::new();
                    for tab in tabs.iter().filter(|tab| tab.window_id == original.window_id && outlook_url(&tab.url)) {
                        let document = safari::document_target(tab_target(tab)?).await
                            .map_err(|_| OutlookWebError::SafariBeforeDispatch)?;
                        if document.document_id == original.document_id { continue; }
                        let probe = format!(
                            "sessionStorage.getItem('__moff_auth_recovery') === {}", json!(marker));
                        let result = match safari::eval_document(&probe, &document, &eval_options(10)).await {
                            Ok(result) => result,
                            Err(error) if error.downcast_ref::<PageOperationError>().is_some_and(|error|
                                matches!(error.phase, OperationPhase::BeforeDispatch)) => continue 'poll,
                            Err(_) => return Err(OutlookWebError::SafariAuthRecoveryFailed("document marker probe")),
                        };
                        if result.truncated {
                            return Err(OutlookWebError::InvalidResponse);
                        }
                        if result.value == true { candidates.push((document, tab)); }
                    }
                    if candidates.len() > 1 { return Err(OutlookWebError::SafariUnavailable); }
                    let Some((document, tab)) = candidates.pop() else { continue; };
                    let script = page_script(include_str!("safari_read.js"),
                        &json!({"action":"__MoffAuthStatus","recovery_marker":marker}));
                    let result = match safari::eval_document(&script, &document, &eval_options(15)).await {
                        Ok(result) => result,
                        Err(error) if error.downcast_ref::<PageOperationError>().is_some_and(|error|
                            matches!(error.phase, OperationPhase::BeforeDispatch)) => continue 'poll,
                        Err(_) => return Err(OutlookWebError::SafariAuthRecoveryFailed("silent token acquisition")),
                    };
                    if result.truncated { return Err(OutlookWebError::InvalidResponse); }
                    last_status = result.value["status"].as_u64();
                    match result.value["status"].as_u64() {
                        Some(200) => return Ok(json!({"complete":true,"state":"authenticated",
                            "window":tab.window,"tab":tab.tab,"credential_export":false})),
                        Some(401 | 409 | 422 | 425) => {}
                        _ => return Err(OutlookWebError::InvalidResponse),
                    }
                }
            }).await;
            match wait {
                Ok(result) => result,
                Err(_) if last_status == Some(422) => Err(OutlookWebError::SafariAuthRecoveryFailed("Outlook authentication modules did not become available")),
                Err(_) => Ok(json!({"complete":false,"state":"needs_user_action",
                    "message":"Recovery was not verified within the wait period: sign-in may need your clicks, or Outlook may have changed origin/cleared the tab marker. Inspect Safari, finish sign-in if needed, then run moff auth status. Do not refresh while entering credentials.",
                    "credential_export":false})),
            }
        })
    }

    /// Use a loaded Outlook tab without exporting browser credentials.
    pub fn from_safari(target: Option<(u32, u32)>) -> Result<Self, OutlookWebError> {
        Self::from_safari_with_mutations(target, false)
    }
}

fn selected_tab(
    tabs: &[safari::OpenTab],
    target: Option<(u32, u32)>,
) -> Result<&safari::OpenTab, OutlookWebError> {
    let matches: Vec<_> = tabs
        .iter()
        .filter(|tab| {
            outlook_url(&tab.url)
                && target.is_none_or(|(window, position)| {
                    tab.window == i64::from(window) && tab.tab == i64::from(position)
                })
        })
        .collect();
    let [tab] = matches.as_slice() else {
        return Err(OutlookWebError::SafariUnavailable);
    };
    Ok(tab)
}

fn tab_target(tab: &safari::OpenTab) -> Result<TabTarget, OutlookWebError> {
    Ok(TabTarget {
        window: u32::try_from(tab.window).map_err(|_| OutlookWebError::SafariUnavailable)?,
        tab: u32::try_from(tab.tab).map_err(|_| OutlookWebError::SafariUnavailable)?,
    })
}

fn eval_options(timeout: u32) -> EvalOptions {
    EvalOptions {
        allow_eval: true,
        max_chars: 1000,
        timeout,
    }
}

fn page_script(source: &str, config: &Value) -> String {
    format!(
        "({source})({config}, ({}), ({}))",
        include_str!("safari_auth.js"),
        include_str!("safari_response.js")
    )
}

pub(crate) fn request_script_bytes(action: &str, body: &Value) -> usize {
    page_script(
        include_str!("safari_read.js"),
        &json!({
            "action":action,"body":body,"mutation":true,"allow_mutation":true,"timeout_ms":25000
        }),
    )
    .len()
}

fn transfer_options(total_bytes: u64) -> BinaryTransferOptions {
    BinaryTransferOptions {
        allow_eval: true,
        chunk_bytes: 128_000,
        total_bytes,
        timeout: 60,
        idle_timeout: 120,
    }
}

fn transfer_factory(config: &Value) -> String {
    format!(
        "(signal) => ({})({}, ({}), ({}), signal)",
        include_str!("safari_transfer.js"),
        config,
        include_str!("safari_auth.js"),
        include_str!("safari_response.js")
    )
}

impl OutlookWebSession {
    /// Explicitly opt into typed, non-sending writes; the raw read surface stays read-only.
    pub fn from_safari_with_mutations(
        target: Option<(u32, u32)>,
        allow_mutation: bool,
    ) -> Result<Self, OutlookWebError> {
        let runtime =
            tokio::runtime::Runtime::new().map_err(|_| OutlookWebError::SafariOperationFailed)?;
        let tabs = runtime
            .block_on(safari::tabs())
            .map_err(|_| OutlookWebError::SafariOperationFailed)?;
        let tab = selected_tab(&tabs, target)?;
        let target = tab_target(tab)?;
        let target = runtime
            .block_on(safari::document_target(target))
            .map_err(|_| OutlookWebError::SafariBeforeDispatch)?;
        runtime.block_on(verify_host_runtime(tab, &target))?;
        Ok(Self {
            transport: SessionTransport::Safari(SafariSession {
                target,
                runtime,
                allow_mutation,
            }),
        })
    }
}

impl SafariSession {
    pub(super) fn set_pin(&self, id: &str, pinned: bool) -> Result<(), OutlookWebError> {
        if !self.allow_mutation {
            return Err(OutlookWebError::SafariReadOnly);
        }
        let script = format!(
            "({})({})",
            include_str!("safari_pin.js"),
            json!({"id":id,"pinned":pinned,"allow_mutation":true})
        );
        let result = self
            .runtime
            .block_on(safari::eval_document(
                &script,
                &self.target,
                &eval_options(25),
            ))
            .map_err(|_| OutlookWebError::OutcomeUnknown { action: "PinItem" })?;
        if result.value["dispatched"] == false {
            return Err(if result.value["reason"] == "outlook_not_ready" {
                OutlookWebError::OutlookNotReady
            } else {
                OutlookWebError::AuthenticationAdapterChanged
            });
        }
        if result.truncated || result.value["success"] != true {
            return Err(OutlookWebError::OutcomeUnknown { action: "PinItem" });
        }
        Ok(())
    }
    pub(super) fn product_read(&self, url: &str) -> Result<Value, OutlookWebError> {
        self.product_request(&json!({"url":url}), false)
    }

    pub(super) fn product_download(
        &self,
        url: &str,
        output: &mut impl std::io::Write,
        max_bytes: u64,
    ) -> Result<u64, OutlookWebError> {
        let item = self.product_read(url)?;
        let size = item["size"]
            .as_u64()
            .filter(|_| item["file"].is_object())
            .ok_or(OutlookWebError::InvalidInput(
                "download requires a file with a known size",
            ))?;
        crate::validate_file_size(size, max_bytes)?;
        let factory = transfer_factory(&json!({"url":url,"size":size,"direction":"download"}));
        self.runtime.block_on(async {
            let mut transfer = safari::open_binary_transfer(
                &factory,
                &self.target,
                BinaryDirection::Download,
                &transfer_options(max_bytes),
            )
            .await
            .map_err(|_| OutlookWebError::SafariOperationFailed)?;
            let result = async {
                let mut count = 0;
                loop {
                    let chunk = safari::read_binary_chunk(&mut transfer)
                        .await
                        .map_err(|_| OutlookWebError::SafariOperationFailed)?;
                    if chunk.offset != count {
                        return Err(OutlookWebError::InvalidResponse);
                    }
                    output
                        .write_all(&chunk.bytes)
                        .map_err(OutlookWebError::FileTransferIo)?;
                    count += u64::try_from(chunk.bytes.len())
                        .map_err(|_| OutlookWebError::InvalidResponse)?;
                    if chunk.done {
                        break;
                    }
                }
                if count != size {
                    return Err(OutlookWebError::InvalidInput(
                        "download length differs from item size",
                    ));
                }
                safari::finish_binary_transfer(&mut transfer)
                    .await
                    .map_err(|_| OutlookWebError::SafariOperationFailed)?;
                Ok(count)
            }
            .await;
            if result.is_err() && !transfer.closed {
                safari::abort_binary_transfer(&mut transfer)
                    .await
                    .map_err(|_| OutlookWebError::SafariOperationFailed)?;
            }
            result
        })
    }

    pub(super) fn product_upload(
        &self,
        request: &crate::ProductWrite,
        input: &mut impl std::io::Read,
        max_bytes: u64,
    ) -> Result<Value, OutlookWebError> {
        static SEQUENCE: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
        if !self.allow_mutation {
            return Err(OutlookWebError::SafariReadOnly);
        }
        let size = request.upload_size.ok_or(OutlookWebError::InvalidAction)?;
        crate::validate_file_size(size, max_bytes)?;
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map_err(|_| OutlookWebError::SafariOperationFailed)?
            .as_nanos();
        let key = format!(
            "__moff_upload_{}_{}_{}_{}",
            self.target.document_id,
            std::process::id(),
            now,
            SEQUENCE.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
        );
        let factory = transfer_factory(&json!({"url":request.url,"size":size,
            "direction":"upload","etag":request.etag,"allow_mutation":true,"result_key":key}));
        let unknown = || OutlookWebError::OutcomeUnknown {
            action: "file upload; inspect OneDrive before retrying",
        };
        let result = self.runtime.block_on(async {
            let mut transfer = safari::open_binary_transfer(
                &factory, &self.target, BinaryDirection::Upload, &transfer_options(max_bytes),
            ).await.map_err(|error| {
                if error.downcast_ref::<PageOperationError>().is_some_and(|e| matches!(e.phase, OperationPhase::BeforeDispatch)) {
                    OutlookWebError::SafariBeforeDispatch
                } else { unknown() }
            })?;
            let result = async {
                let mut buffer = vec![0; 128_000];
                let mut count = 0;
                loop {
                    let length = input.read(&mut buffer).map_err(|_| unknown())?;
                    if length == 0 { break; }
                    count += u64::try_from(length).map_err(|_| unknown())?;
                    if count > size { return Err(unknown()); }
                    safari::write_binary_chunk(&mut transfer, &buffer[..length]).await
                        .map_err(|_| unknown())?;
                }
                if count != size { return Err(unknown()); }
                safari::finish_binary_transfer(&mut transfer).await.map_err(|_| unknown())?;
                let script = format!("(()=>{{const key={};const value=window[key];delete window[key];return value;}})()", json!(key));
                let result = safari::eval_document(&script, &self.target, &EvalOptions {
                    allow_eval:true, max_chars:10_000, timeout:30,
                }).await.map_err(|_| unknown())?;
                if result.truncated || result.value["size"].as_u64() != Some(size)
                    || result.value["id"].as_str().is_none_or(str::is_empty) { return Err(unknown()); }
                Ok(result.value)
            }.await;
            if result.is_err() && !transfer.closed {
                safari::abort_binary_transfer(&mut transfer).await.map_err(|_| unknown())?;
            }
            result
        });
        if result.is_err() {
            let script = format!(
                "(()=>{{const key={};const value=window[key];delete window[key];return value ?? null;}})()",
                json!(format!("{key}_failure"))
            );
            if let Ok(diagnostic) = self.runtime.block_on(safari::eval_document(
                &script,
                &self.target,
                &eval_options(10),
            )) {
                let stage = match diagnostic.value["stage"].as_str() {
                    Some("authentication") => Some("authentication"),
                    Some("empty upload") => Some("empty upload"),
                    Some("upload session creation") => Some("upload session creation"),
                    Some("upload range") => Some("upload range"),
                    Some("upload finalization") => Some("upload finalization"),
                    _ => None,
                };
                if let Some(stage) = stage {
                    return Err(OutlookWebError::FileUploadOutcomeUnknown {
                        stage,
                        status: diagnostic.value["status"]
                            .as_u64()
                            .and_then(|n| u16::try_from(n).ok()),
                    });
                }
            }
        }
        result
    }

    pub(super) fn product_write(
        &self,
        request: &crate::ProductWrite,
    ) -> Result<Value, OutlookWebError> {
        if !self.allow_mutation {
            return Err(OutlookWebError::SafariReadOnly);
        }
        if request.upload_size.is_some() {
            return Err(OutlookWebError::InvalidInput(
                "use the streamed file upload API",
            ));
        }
        if let Some(path) = request
            .url
            .strip_prefix("https://graph.microsoft.com/v1.0/groups/")
        {
            let id = path
                .split('/')
                .next()
                .ok_or(OutlookWebError::InvalidResponse)?;
            let group =
                self.product_read(&format!("https://graph.microsoft.com/v1.0/groups/{id}"))?;
            if group["groupTypes"]
                .as_array()
                .is_none_or(|types| types.len() != 1 || types[0] != "Unified")
                || group["isAssignableToRole"] == true
            {
                return Err(OutlookWebError::InvalidInput(
                    "only Microsoft 365 groups can be managed",
                ));
            }
            let me = self.product_read("https://graph.microsoft.com/v1.0/me")?;
            let user_id = me["id"]
                .as_str()
                .filter(|s| !s.is_empty())
                .ok_or(OutlookWebError::InvalidResponse)?;
            let owners = self.product_read(&format!(
                "https://graph.microsoft.com/v1.0/groups/{id}/owners?$top=100"
            ))?;
            let owners = owners["value"]
                .as_array()
                .ok_or(OutlookWebError::InvalidResponse)?;
            if !owners
                .iter()
                .any(|owner| owner["id"].as_str() == Some(user_id))
            {
                return Err(OutlookWebError::InvalidInput(
                    "group mutations require verified ownership; the bounded owner page did not establish it",
                ));
            }
        }
        self.product_request(
            &json!({"url":request.url,"method":request.method,"body":request.body,
            "etag":request.etag,"allow_mutation":true}),
            true,
        )
    }

    fn product_request(&self, config: &Value, mutation: bool) -> Result<Value, OutlookWebError> {
        self.page_request(include_str!("safari_products.js"), config, mutation)
    }

    pub(super) fn page_request(
        &self,
        source: &str,
        config: &Value,
        mutation: bool,
    ) -> Result<Value, OutlookWebError> {
        if mutation && !self.allow_mutation {
            return Err(OutlookWebError::SafariReadOnly);
        }
        let unknown = || OutlookWebError::OutcomeUnknown {
            action: "product write",
        };
        let script = page_script(source, config);
        let result = self
            .runtime
            .block_on(safari::eval_document(
                &script,
                &self.target,
                &EvalOptions {
                    allow_eval: true,
                    max_chars: 1_000_000,
                    timeout: 60,
                },
            ))
            .map_err(|error| {
                if error
                    .downcast_ref::<PageOperationError>()
                    .is_some_and(|e| matches!(e.phase, OperationPhase::BeforeDispatch))
                {
                    OutlookWebError::SafariBeforeDispatch
                } else if mutation {
                    unknown()
                } else {
                    OutlookWebError::SafariOperationFailed
                }
            })?;
        if result.truncated {
            return Err(if mutation {
                unknown()
            } else {
                OutlookWebError::TruncatedSearch
            });
        }
        let status = result.value["status"]
            .as_u64()
            .and_then(|n| u16::try_from(n).ok())
            .ok_or_else(|| {
                if mutation {
                    unknown()
                } else {
                    OutlookWebError::InvalidResponse
                }
            })?;
        if status == 401 {
            return Err(OutlookWebError::ProductAuthenticationUnavailable);
        }
        if !(200..300).contains(&status) {
            if mutation && status >= 500 {
                return Err(unknown());
            }
            return Err(OutlookWebError::ProductRequestRejected(status));
        }
        if mutation && status == 204 {
            return Ok(json!({"status":204,"acknowledged":true,"verified":false}));
        }
        let value = decode_response(result.value["body"].as_str().ok_or_else(|| {
            if mutation {
                unknown()
            } else {
                OutlookWebError::InvalidResponse
            }
        })?)
        .map_err(|_| {
            if mutation {
                unknown()
            } else {
                OutlookWebError::InvalidResponse
            }
        })?;
        if !value.is_object() || value.get("error").is_some() {
            return Err(if mutation {
                unknown()
            } else {
                OutlookWebError::InvalidResponse
            });
        }
        Ok(value)
    }

    pub(super) fn call(
        &self,
        action: &'static str,
        body: &Value,
        mutation: bool,
    ) -> Result<Value, OutlookWebError> {
        self.call_with_timeout(action, body, mutation, Duration::from_secs(25))
    }

    pub(super) fn call_with_timeout(
        &self,
        action: &'static str,
        body: &Value,
        mutation: bool,
        timeout: Duration,
    ) -> Result<Value, OutlookWebError> {
        if mutation {
            if !self.allow_mutation
                || ![
                    "CreateItem",
                    "UpdateItem",
                    "UpdateCalendarEvent",
                    "DeleteItem",
                    "MoveItem",
                ]
                .contains(&action)
            {
                return Err(OutlookWebError::SafariReadOnly);
            }
        } else {
            validate_read_action(action)?;
        }
        let script = page_script(
            include_str!("safari_read.js"),
            &json!({ "action": action, "body": body, "mutation": mutation,
                "allow_mutation": self.allow_mutation, "timeout_ms": timeout.as_millis() }),
        );
        let result = self
            .runtime
            .block_on(safari::eval_document(
                &script,
                &self.target,
                &EvalOptions {
                    allow_eval: true,
                    max_chars: 1_000_000,
                    timeout: u32::try_from(timeout.as_secs().saturating_add(10)).unwrap_or(60),
                },
            ))
            .map_err(|error| {
                if let Some(operation) = error.downcast_ref::<PageOperationError>() {
                    if matches!(operation.phase, OperationPhase::BeforeDispatch) {
                        return OutlookWebError::SafariBeforeDispatch;
                    }
                    if mutation {
                        return OutlookWebError::SafariOutcomeUnknown {
                            action,
                            phase: operation.phase,
                            cancellation: operation.cancellation,
                        };
                    }
                }
                if mutation {
                    uncertain_write(
                        action,
                        "Safari evaluation failed without phase evidence",
                        None,
                    )
                } else {
                    OutlookWebError::SafariOperationFailed
                }
            })?;
        if result.truncated {
            return Err(if mutation {
                uncertain_write(action, "Safari result was truncated", None)
            } else {
                OutlookWebError::TruncatedSearch
            });
        }
        let diagnostic_action =
            if action == "UpdateItem" && body["Body"]["MessageDisposition"] == "SendAndSaveCopy" {
                "UpdateItem mail submission"
            } else {
                action
            };
        decode_call_result(diagnostic_action, &result.value, mutation)
    }
}

fn uncertain_write(
    action: &'static str,
    reason: &'static str,
    http_status: Option<u16>,
) -> OutlookWebError {
    if matches!(action, "SendItem" | "UpdateItem mail submission") {
        OutlookWebError::SendOutcomeUnknown {
            reason,
            http_status,
        }
    } else {
        OutlookWebError::OutcomeUnknown { action }
    }
}

fn decode_call_result(
    action: &'static str,
    value: &Value,
    mutation: bool,
) -> Result<Value, OutlookWebError> {
    if let Some(failure) = value["failure"].as_str() {
        return Err(if mutation && value["dispatched"] != false {
            uncertain_write(
                action,
                match failure {
                    "timeout" => "browser request timed out after dispatch",
                    "network" => "browser network failed after dispatch",
                    "response_too_large" => "browser response exceeded size limit after dispatch",
                    _ => "unrecognized browser failure after dispatch",
                },
                None,
            )
        } else {
            match failure {
                "timeout" => OutlookWebError::ReadTimedOut { action },
                "network" => OutlookWebError::ReadNetworkFailed { action },
                "response_too_large" => OutlookWebError::ResponseTooLarge,
                _ => OutlookWebError::InvalidResponse,
            }
        });
    }
    let status = value["status"]
        .as_u64()
        .and_then(|n| u16::try_from(n).ok())
        .ok_or(if mutation {
            uncertain_write(action, "browser result lacks a valid HTTP status", None)
        } else {
            OutlookWebError::InvalidResponse
        })?;
    if status == 401 {
        return Err(OutlookWebError::SessionExpired);
    }
    if status == 425 && value["stage"] == "authentication" {
        return Err(OutlookWebError::OutlookNotReady);
    }
    if status == 422 && value["stage"] == "authentication" {
        return Err(OutlookWebError::AuthenticationAdapterChanged);
    }
    if !(200..300).contains(&status) {
        if mutation && status >= 500 {
            return Err(uncertain_write(
                action,
                "server returned an error after dispatch",
                Some(status),
            ));
        }
        return Err(OutlookWebError::RequestRejected(status));
    }
    decode_response(value["body"].as_str().ok_or(if mutation {
        uncertain_write(
            action,
            "successful HTTP response lacks a text body",
            Some(status),
        )
    } else {
        OutlookWebError::InvalidResponse
    })?)
    .map_err(|_| {
        if mutation {
            uncertain_write(
                action,
                "successful HTTP response body is not valid JSON",
                Some(status),
            )
        } else {
            OutlookWebError::InvalidResponse
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn send_uncertainty_preserves_safe_phase_reason_and_http_status_without_response_content() {
        for (value, reason, status) in [
            (
                json!({"status":0,"failure":"network","dispatched":true}),
                "network",
                None,
            ),
            (
                json!({"status":0,"failure":"timeout","dispatched":true}),
                "timed out",
                None,
            ),
            (
                json!({"status":0,"failure":"response_too_large","dispatched":true}),
                "size limit",
                None,
            ),
            (json!({}), "HTTP status", None),
            (
                json!({"status":500,"body":"private body/token"}),
                "server returned",
                Some(500),
            ),
            (json!({"status":503}), "server returned", Some(503)),
            (json!({"status":200}), "lacks a text body", Some(200)),
            (
                json!({"status":200,"body":"private body/token"}),
                "not valid JSON",
                Some(200),
            ),
            (json!({"status":204,"body":""}), "not valid JSON", Some(204)),
        ] {
            let error = decode_call_result("SendItem", &value, true).unwrap_err();
            assert!(
                matches!(&error, OutlookWebError::SendOutcomeUnknown { http_status, .. } if *http_status == status)
            );
            let text = error.to_string();
            assert!(text.contains(reason), "{text}");
            assert!(text.contains("do not retry"), "{text}");
            assert!(!text.contains("private body/token"));
        }
        for status in [400, 403, 404, 409, 429] {
            assert!(
                matches!(decode_call_result("SendItem",&json!({"status":status}),true),
                Err(OutlookWebError::RequestRejected(code)) if code == status)
            );
        }
        assert!(matches!(
            decode_call_result(
                "SendItem",
                &json!({"failure":"timeout","dispatched":false}),
                true
            ),
            Err(OutlookWebError::ReadTimedOut { .. })
        ));
    }

    #[test]
    fn browser_diagnostics_preserve_auth_timeout_and_write_uncertainty() {
        for (value, expected) in [
            (json!({"status":401}), "expired"),
            (json!({"status":425,"stage":"authentication"}), "not loaded"),
            (json!({"status":422,"stage":"authentication"}), "adapter"),
            (
                json!({"status":0,"failure":"timeout","dispatched":true}),
                "deadline",
            ),
            (
                json!({"status":0,"failure":"network","dispatched":true}),
                "network",
            ),
            (
                json!({"status":0,"failure":"response_too_large","dispatched":true}),
                "size limit",
            ),
            (json!({"status":200,"body":"not JSON"}), "invalid data"),
        ] {
            let error = decode_call_result("GetItem", &value, false).unwrap_err();
            assert!(error.to_string().contains(expected), "{error}");
        }
        assert!(matches!(
            decode_call_result("GetItem", &json!({"status":422}), false),
            Err(OutlookWebError::RequestRejected(422))
        ));
        assert!(matches!(
            decode_call_result(
                "MoveItem",
                &json!({"status":0,"failure":"timeout","dispatched":true}),
                true
            ),
            Err(OutlookWebError::OutcomeUnknown { action: "MoveItem" })
        ));
        assert!(matches!(
            decode_call_result(
                "MoveItem",
                &json!({"status":0,"failure":"timeout","dispatched":false}),
                true
            ),
            Err(OutlookWebError::ReadTimedOut { .. })
        ));
    }

    #[test]
    fn transfer_options_keep_transport_bounded_and_total_is_a_ceiling() {
        let options = transfer_options(crate::DEFAULT_MAX_FILE_BYTES);
        assert!(options.validate().is_ok());
        assert!(options.allow_eval);
        assert_eq!(options.chunk_bytes, 128_000);
        assert_eq!(options.total_bytes, 250_000_000);
        assert_eq!(options.timeout, 60);
        assert_eq!(options.idle_timeout, 120);
        assert!(transfer_options(0).validate().is_ok());
    }

    #[test]
    fn accepts_only_loaded_outlook_product_tabs() {
        for url in [
            "https://outlook.office.com/mail/",
            "https://outlook.cloud.microsoft/calendar/view/week",
            "https://outlook.cloud.microsoft/host/4a6520d6-94a8-4730-9184-24997b9b3401/auth",
            "https://outlook.live.com/mail/",
            "https://outlook.live.com/calendar/0/view/week",
        ] {
            assert!(outlook_url(url));
        }
        for url in [
            "http://outlook.office.com/mail/",
            "https://outlook.office.com.evil/mail/",
            "https://outlook.live.com.evil/mail/",
            "https://outlook.live.com/owa/",
            "https://outlook.office.com/",
            "https://user@outlook.office.com/mail/",
            "https://outlook.cloud.microsoft/host/invalid/auth",
            "https://outlook.cloud.microsoft/host/4a6520d6-94a8-4730-9184-24997b9b3401/auth/other",
        ] {
            assert!(!outlook_url(url));
        }
    }

    #[test]
    fn refresh_requires_an_unambiguous_outlook_document_and_bounded_wait() {
        let tabs = vec![
            safari::OpenTab {
                title: String::new(),
                url: "https://outlook.office.com/mail/".into(),
                window: 1,
                tab: 1,
                window_id: 42,
            },
            safari::OpenTab {
                title: String::new(),
                url: "https://outlook.office.com/calendar/".into(),
                window: 1,
                tab: 2,
                window_id: 42,
            },
        ];
        assert!(selected_tab(&tabs, None).is_err());
        assert_eq!(selected_tab(&tabs, Some((1, 2))).unwrap().tab, 2);
        assert!(selected_tab(&tabs, Some((2, 1))).is_err());
        for wait in [0, 121] {
            assert!(matches!(
                OutlookWebSession::refresh_safari_auth(None, wait),
                Err(OutlookWebError::InvalidInput(_))
            ));
        }
    }
}
