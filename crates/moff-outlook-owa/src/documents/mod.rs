mod agenda;
mod package;
#[cfg(test)]
mod tests;

use base64::{Engine as _, engine::general_purpose::URL_SAFE_NO_PAD};
use cider::sources::safari::{
    self, BinaryDirection, BinaryTransferOptions, DocumentTarget, EvalOptions, TabTarget,
};
use serde_json::{Value, json};
use url::Url;

use crate::{OutlookWebError, OutlookWebSession};
use package::MAX_DOCUMENT_BYTES;
pub use package::{DocumentEdit, DocumentParagraph, WordPackage};

pub struct DocumentSession {
    location: DocumentLocation,
    target: DocumentTarget,
    runtime: tokio::runtime::Runtime,
    allow_mutation: bool,
}

#[derive(Debug)]
struct DocumentLocation {
    base: String,
    id: String,
    url: String,
}

fn sharepoint_url(input: &str) -> Result<Url, OutlookWebError> {
    let url =
        Url::parse(input).map_err(|_| OutlookWebError::InvalidInput("invalid document URL"))?;
    if url.scheme() != "https"
        || !url
            .host_str()
            .is_some_and(|host| host.ends_with(".sharepoint.com"))
        || !url.username().is_empty()
        || url.password().is_some()
        || url.port().is_some()
        || url.path().contains('\\')
        || ["%2f", "%5c", "%25"]
            .iter()
            .any(|value| url.path().to_ascii_lowercase().contains(value))
    {
        return Err(OutlookWebError::InvalidInput(
            "documents require an HTTPS SharePoint or Word document link without credentials",
        ));
    }
    Ok(url)
}

fn guid(value: &str) -> Option<String> {
    let value = value.trim_matches(['{', '}']);
    let value = if value.starts_with('w') && value.len() == 33 {
        &value[1..]
    } else {
        value
    };
    let compact = value.replace('-', "");
    if compact.len() != 32 || !compact.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    Some(
        format!(
            "{}-{}-{}-{}-{}",
            &compact[..8],
            &compact[8..12],
            &compact[12..16],
            &compact[16..20],
            &compact[20..]
        )
        .to_ascii_lowercase(),
    )
}

impl DocumentLocation {
    fn parse(input: &str) -> Result<Option<Self>, OutlookWebError> {
        let url = sharepoint_url(input)?;
        let path = url.path().strip_prefix("/:w:/r").unwrap_or(url.path());
        let parts = path
            .split('/')
            .filter(|part| !part.is_empty())
            .collect::<Vec<_>>();
        let ids = url
            .query_pairs()
            .filter(|(key, _)| key.eq_ignore_ascii_case("sourcedoc") || key == "d")
            .map(|(_, value)| {
                guid(&value).ok_or(OutlookWebError::InvalidInput(
                    "invalid Word document identity",
                ))
            })
            .collect::<Result<Vec<_>, _>>()?;
        if ids.is_empty() {
            return Ok(None);
        }
        if ids.iter().any(|id| id != &ids[0])
            || parts.len() < 3
            || !matches!(parts[0], "sites" | "teams" | "personal")
        {
            return Err(OutlookWebError::InvalidInput(
                "ambiguous Word identity or unsupported site URL",
            ));
        }
        Ok(Some(Self {
            base: format!(
                "{}/{}/{}",
                url.origin().ascii_serialization(),
                parts[0],
                parts[1]
            ),
            id: ids[0].clone(),
            url: input.to_owned(),
        }))
    }

    fn resolve(input: &str) -> Result<Self, OutlookWebError> {
        if let Some(location) = Self::parse(input)? {
            return Ok(location);
        }
        let input_url = sharepoint_url(input)?;
        let path = format!(
            "/v1.0/shares/u!{}/driveItem",
            URL_SAFE_NO_PAD.encode(input.as_bytes())
        );
        let item = OutlookWebSession::from_safari(None)?.product_read(&path)?;
        let resolved = item["webUrl"]
            .as_str()
            .ok_or(OutlookWebError::InvalidResponse)?;
        if sharepoint_url(resolved)?.origin() != input_url.origin() {
            return Err(OutlookWebError::InvalidInput(
                "resolved document changed SharePoint origin; use its canonical URL explicitly",
            ));
        }
        Self::parse(resolved)?.ok_or(OutlookWebError::InvalidInput(
            "could not resolve Word identity; use the document's Doc.aspx?sourcedoc= link",
        ))
    }
}

fn options() -> BinaryTransferOptions {
    BinaryTransferOptions {
        allow_eval: true,
        chunk_bytes: 128_000,
        total_bytes: MAX_DOCUMENT_BYTES,
        timeout: 60,
        idle_timeout: 120,
    }
}

impl DocumentSession {
    pub fn open(
        url: &str,
        target: Option<(u32, u32)>,
        allow_mutation: bool,
    ) -> Result<Self, OutlookWebError> {
        let location = DocumentLocation::resolve(url)?;
        let origin = sharepoint_url(&location.base)?.origin();
        let runtime =
            tokio::runtime::Runtime::new().map_err(|_| OutlookWebError::SafariOperationFailed)?;
        let tabs = runtime
            .block_on(safari::tabs())
            .map_err(|_| OutlookWebError::SafariUnavailable)?;
        let matches = tabs
            .iter()
            .filter(|tab| {
                Url::parse(&tab.url).is_ok_and(|url| url.origin() == origin)
                    && target.is_none_or(|(window, position)| {
                        tab.window == i64::from(window) && tab.tab == i64::from(position)
                    })
            })
            .collect::<Vec<_>>();
        let tab = match matches.as_slice() {
            [] => {
                return Err(OutlookWebError::InvalidInput(
                    "open this SharePoint document in Safari and complete sign-in; no matching tab is available",
                ));
            }
            [tab] => *tab,
            _ => {
                return Err(OutlookWebError::InvalidInput(
                    "multiple SharePoint tabs match; select --safari-window and --safari-tab",
                ));
            }
        };
        let target = runtime
            .block_on(safari::document_target(TabTarget {
                window: u32::try_from(tab.window)
                    .map_err(|_| OutlookWebError::SafariUnavailable)?,
                tab: u32::try_from(tab.tab).map_err(|_| OutlookWebError::SafariUnavailable)?,
            }))
            .map_err(|_| OutlookWebError::SafariBeforeDispatch)?;
        Ok(Self {
            location,
            target,
            runtime,
            allow_mutation,
        })
    }

    fn config(&self, direction: &str) -> Value {
        json!({"base":self.location.base,"id":self.location.id,"direction":direction})
    }

    fn evaluate(&self, script: &str) -> Result<Value, OutlookWebError> {
        let result = self
            .runtime
            .block_on(safari::eval_document(
                script,
                &self.target,
                &EvalOptions {
                    allow_eval: true,
                    max_chars: 12_000,
                    timeout: 30,
                },
            ))
            .map_err(|_| OutlookWebError::SafariOperationFailed)?;
        if result.truncated {
            return Err(OutlookWebError::ResponseTooLarge);
        }
        Ok(result.value)
    }

    fn metadata(&self) -> Result<Value, OutlookWebError> {
        let result = self.evaluate(&format!(
            "(async()=>{{const c=new AbortController();const t=setTimeout(()=>c.abort(),25000);try{{return await ({})({},c.signal)}}finally{{clearTimeout(t)}}}})()",
            include_str!("safari_document.js"), self.config("metadata"),
        ))?;
        match result["status"].as_u64() {
            Some(200) => Ok(result),
            Some(401 | 403) => Err(OutlookWebError::InvalidInput(
                "this SharePoint tab cannot read the document; check its account and document access in Safari",
            )),
            Some(status) => Err(OutlookWebError::RequestRejected(
                u16::try_from(status).map_err(|_| OutlookWebError::InvalidResponse)?,
            )),
            None => Err(OutlookWebError::InvalidResponse),
        }
    }

    fn factory(config: &Value) -> String {
        format!(
            "signal => ({})({},signal)",
            include_str!("safari_document.js"),
            config
        )
    }

    fn download(&self, metadata: &Value) -> Result<Vec<u8>, OutlookWebError> {
        let mut config = self.config("download");
        config["size"] = metadata["size"].clone();
        config["etag"] = metadata["etag"].clone();
        self.runtime.block_on(async {
            let mut transfer = safari::open_binary_transfer(
                &Self::factory(&config),
                &self.target,
                BinaryDirection::Download,
                &options(),
            )
            .await
            .map_err(|_| OutlookWebError::SafariOperationFailed)?;
            let result = async {
                let mut bytes = Vec::new();
                loop {
                    let chunk = safari::read_binary_chunk(&mut transfer)
                        .await
                        .map_err(|_| OutlookWebError::SafariOperationFailed)?;
                    if chunk.offset != bytes.len() as u64
                        || bytes.len() as u64 + chunk.bytes.len() as u64 > MAX_DOCUMENT_BYTES
                    {
                        return Err(OutlookWebError::InvalidResponse);
                    }
                    bytes.extend_from_slice(&chunk.bytes);
                    if chunk.done {
                        break;
                    }
                }
                if Some(bytes.len() as u64) != metadata["size"].as_u64() {
                    return Err(OutlookWebError::InvalidResponse);
                }
                safari::finish_binary_transfer(&mut transfer)
                    .await
                    .map_err(|_| OutlookWebError::SafariOperationFailed)?;
                Ok(bytes)
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

    pub fn read(&self) -> Result<Value, OutlookWebError> {
        let metadata = self.metadata()?;
        let package = WordPackage::open(self.download(&metadata)?)?;
        Ok(
            json!({"document":metadata,"url":self.location.url,"blocks":package.read()?,
            "scope":"Word main body, including tables and content-control paragraphs; headers/footers/comments are retained but not included"}),
        )
    }

    pub fn edit(
        &self,
        section: &str,
        edit: &DocumentEdit,
        expected_etag: Option<&str>,
    ) -> Result<Value, OutlookWebError> {
        if expected_etag.is_some() && !self.allow_mutation {
            return Err(OutlookWebError::SafariReadOnly);
        }
        let before = self.metadata()?;
        if expected_etag.is_some_and(|version| before["etag"] != version) {
            return Err(OutlookWebError::PlanChanged);
        }
        if expected_etag.is_some() && before["locked"] == true {
            return Err(OutlookWebError::InvalidInput(
                "SharePoint reports that this document is locked for editing; close its open editors and wait for the lock to clear, then preview again; no upload was attempted",
            ));
        }
        let package = WordPackage::open(self.download(&before)?)?;
        let changed = package.append_section(section, edit)?;
        if expected_etag.is_none() {
            return Ok(
                json!({"preview":true,"document":before,"url":self.location.url,
                "section":section,"create_before":edit.create_before,"append":edit.paragraphs,
                "result_bytes":changed.len(),"preserves":"all existing XML and package parts; only inserts new paragraphs and hyperlink relationships"}),
            );
        }
        self.upload(&changed, &before)?;
        let after = verify_saved_with(
            &before,
            &changed,
            || {
                let after = self.metadata()?;
                if after["etag"] == before["etag"]
                    || after["size"].as_u64() != Some(changed.len() as u64)
                {
                    return Err(OutlookWebError::VerificationFailed);
                }
                let saved = self.download(&after)?;
                Ok((after, saved))
            },
            || std::thread::sleep(std::time::Duration::from_secs(2)),
        )?;
        Ok(
            json!({"verified":true,"document":after,"section":section,"appended":edit.paragraphs.len(),
            "url":self.location.url,"automatic_retry":false}),
        )
    }

    fn upload(&self, bytes: &[u8], metadata: &Value) -> Result<(), OutlookWebError> {
        if !self.allow_mutation {
            return Err(OutlookWebError::SafariReadOnly);
        }
        let key = format!(
            "__moff_doc_{}_{}",
            self.target.document_id,
            std::process::id()
        );
        let mut config = self.config("upload");
        config["size"] = json!(bytes.len());
        config["etag"] = metadata["etag"].clone();
        config["allow_mutation"] = json!(true);
        config["result_key"] = json!(key);
        let result = self.runtime.block_on(async {
            let mut transfer = safari::open_binary_transfer(
                &Self::factory(&config),
                &self.target,
                BinaryDirection::Upload,
                &options(),
            )
            .await
            .map_err(|_| OutlookWebError::SafariBeforeDispatch)?;
            let result = async {
                for chunk in bytes.chunks(128_000) {
                    safari::write_binary_chunk(&mut transfer, chunk)
                        .await
                        .map_err(|_| OutlookWebError::SafariBeforeDispatch)?;
                }
                safari::finish_binary_transfer(&mut transfer)
                    .await
                    .map_err(|_| OutlookWebError::OutcomeUnknown {
                        action: "Word document update; read the document before retrying",
                    })?;
                Ok(())
            }
            .await;
            if result.is_err() && !transfer.closed {
                safari::abort_binary_transfer(&mut transfer).await.map_err(|_| {
                    OutlookWebError::OutcomeUnknown {
                        action: "Word transfer cleanup failed; inspect the document before retrying",
                    }
                })?;
            }
            result
        });
        let marker = self.evaluate(&format!(
            "(()=>{{const key={};const value=window[key];delete window[key];return value??null}})()",
            json!(key),
        )).map_err(|_| OutlookWebError::OutcomeUnknown {
            action: "Word update acknowledgement unavailable; read the document before retrying",
        })?;
        upload_acknowledgement(result, &marker)
    }
}

fn verify_saved_with(
    before: &Value,
    expected: &[u8],
    mut read: impl FnMut() -> Result<(Value, Vec<u8>), OutlookWebError>,
    mut pause: impl FnMut(),
) -> Result<Value, OutlookWebError> {
    for attempt in 0..3 {
        if let Ok((after, saved)) = read()
            && saved == expected
            && after["etag"] != before["etag"]
        {
            return Ok(after);
        }
        if attempt < 2 {
            pause();
        }
    }
    Err(OutlookWebError::VerificationFailed)
}

fn upload_acknowledgement(
    result: Result<(), OutlookWebError>,
    marker: &Value,
) -> Result<(), OutlookWebError> {
    if marker["status"] == 412 {
        return Err(OutlookWebError::PlanChanged);
    }
    if marker["stage"] == "digest" {
        return Err(OutlookWebError::InvalidInput(
            "SharePoint update preflight failed before the file write",
        ));
    }
    if marker["stage"] == "write"
        && let Some(status) = marker["status"].as_u64().filter(|s| (400..500).contains(s))
    {
        return Err(OutlookWebError::ProductRequestRejected(
            u16::try_from(status).map_err(|_| OutlookWebError::InvalidResponse)?,
        ));
    }
    result?;
    if !marker["status"]
        .as_u64()
        .is_some_and(|s| (200..300).contains(&s))
    {
        return Err(OutlookWebError::OutcomeUnknown {
            action: "Word document update; read the document before retrying",
        });
    }
    Ok(())
}
