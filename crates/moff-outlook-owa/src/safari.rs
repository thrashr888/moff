use std::collections::BTreeMap;

use cider::sources::safari::{self, EvalOptions, TabTarget};
use serde_json::{Value, json};
use url::Url;

use crate::{
    OutlookWebError, OutlookWebSession, decode_response, http_client, validate_read_action,
};

pub(super) struct SafariSession {
    target: TabTarget,
    runtime: tokio::runtime::Runtime,
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
    /// Use a loaded Outlook tab without exporting browser credentials.
    pub fn from_safari(target: Option<(u32, u32)>) -> Result<Self, OutlookWebError> {
        let runtime =
            tokio::runtime::Runtime::new().map_err(|_| OutlookWebError::SafariOperationFailed)?;
        let tabs = runtime
            .block_on(safari::tabs())
            .map_err(|_| OutlookWebError::SafariOperationFailed)?;
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
        let target = TabTarget {
            window: u32::try_from(tab.window).map_err(|_| OutlookWebError::SafariUnavailable)?,
            tab: u32::try_from(tab.tab).map_err(|_| OutlookWebError::SafariUnavailable)?,
        };
        Ok(Self {
            origin: Url::parse(&tab.url).map_err(|_| OutlookWebError::SafariUnavailable)?,
            headers: BTreeMap::new(),
            client: http_client()?,
            safari: Some(SafariSession { target, runtime }),
        })
    }
}

impl SafariSession {
    pub(super) fn call(&self, action: &str, body: &Value) -> Result<Value, OutlookWebError> {
        validate_read_action(action)?;
        let script = format!(
            "({})({})",
            include_str!("safari_read.js"),
            json!({ "action": action, "body": body })
        );
        let result = self
            .runtime
            .block_on(safari::eval(
                &script,
                self.target,
                &EvalOptions {
                    allow_eval: true,
                    max_chars: 1_000_000,
                    timeout: 60,
                },
            ))
            .map_err(|_| OutlookWebError::SafariOperationFailed)?;
        if result.truncated {
            return Err(OutlookWebError::TruncatedSearch);
        }
        let value = result.value;
        let status = value["status"]
            .as_u64()
            .and_then(|n| u16::try_from(n).ok())
            .ok_or(OutlookWebError::InvalidResponse)?;
        if status == 401 {
            return Err(OutlookWebError::SessionExpired);
        }
        if !(200..300).contains(&status) {
            return Err(OutlookWebError::RequestRejected(status));
        }
        decode_response(
            value["body"]
                .as_str()
                .ok_or(OutlookWebError::InvalidResponse)?,
        )
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
}
