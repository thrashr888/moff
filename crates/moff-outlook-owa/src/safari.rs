use std::time::Duration;

use cider::sources::safari::{
    self, DocumentTarget, EvalOptions, OperationPhase, PageOperationError, TabTarget,
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
                Some("outlook.office.com" | "outlook.cloud.microsoft")
            )
            && (url.path().starts_with("/mail/") || url.path().starts_with("/calendar/"))
    })
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
            "(async function(){{const result=await ({})({{\"action\":\"__MoffAuthStatus\"}});return {{status:result.status,recovery_marker_present:sessionStorage.getItem('__moff_auth_recovery')!==null}};}})()",
            include_str!("safari_read.js")
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
        match result.value["status"].as_u64() {
            Some(200 | 401 | 422 | 425) => {
                Ok(json!({"authenticated":result.value["status"] == 200,
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
                    let script = format!("({})({})", include_str!("safari_read.js"),
                        json!({"action":"__MoffAuthStatus","recovery_marker":marker}));
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
    pub(super) fn call(
        &self,
        action: &'static str,
        body: &Value,
        mutation: bool,
    ) -> Result<Value, OutlookWebError> {
        if mutation {
            if !self.allow_mutation
                || !["CreateItem", "UpdateItem", "UpdateCalendarEvent"].contains(&action)
            {
                return Err(OutlookWebError::SafariReadOnly);
            }
        } else {
            validate_read_action(action)?;
        }
        let script = format!(
            "({})({})",
            include_str!("safari_read.js"),
            json!({ "action": action, "body": body, "mutation": mutation, "allow_mutation": self.allow_mutation })
        );
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
                    OutlookWebError::OutcomeUnknown { action }
                } else {
                    OutlookWebError::SafariOperationFailed
                }
            })?;
        if result.truncated {
            return Err(if mutation {
                OutlookWebError::OutcomeUnknown { action }
            } else {
                OutlookWebError::TruncatedSearch
            });
        }
        let value = result.value;
        let status = value["status"]
            .as_u64()
            .and_then(|n| u16::try_from(n).ok())
            .ok_or(if mutation {
                OutlookWebError::OutcomeUnknown { action }
            } else {
                OutlookWebError::InvalidResponse
            })?;
        if status == 401 {
            return Err(OutlookWebError::SessionExpired);
        }
        if !(200..300).contains(&status) {
            if mutation && status >= 500 {
                return Err(OutlookWebError::OutcomeUnknown { action });
            }
            return Err(OutlookWebError::RequestRejected(status));
        }
        decode_response(value["body"].as_str().ok_or(if mutation {
            OutlookWebError::OutcomeUnknown { action }
        } else {
            OutlookWebError::InvalidResponse
        })?)
        .map_err(|_| {
            if mutation {
                OutlookWebError::OutcomeUnknown { action }
            } else {
                OutlookWebError::InvalidResponse
            }
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_only_loaded_outlook_product_tabs() {
        for url in [
            "https://outlook.office.com/mail/",
            "https://outlook.cloud.microsoft/calendar/view/week",
        ] {
            assert!(outlook_url(url));
        }
        for url in [
            "http://outlook.office.com/mail/",
            "https://outlook.office.com.evil/mail/",
            "https://outlook.office.com/",
            "https://user@outlook.office.com/mail/",
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
