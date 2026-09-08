use std::{collections::BTreeMap, fmt, path::Path};

use reqwest::{
    blocking::Client,
    header::{HeaderMap, HeaderName, HeaderValue},
};
use serde_json::Value;
use url::Url;

use crate::{OutlookWebError, RequestParts, http_client, read_session_file};

/// Microsoft Graph adapter backed by a browser-exported Graph cURL request.
pub struct GraphWebSession {
    origin: Url,
    headers: BTreeMap<String, String>,
    client: Client,
}
impl fmt::Debug for GraphWebSession {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("GraphWebSession")
            .field("origin", &self.origin)
            .field("header_names", &self.headers.keys())
            .finish_non_exhaustive()
    }
}

impl GraphWebSession {
    pub fn from_curl_file(path: &Path) -> Result<Self, OutlookWebError> {
        let text = read_session_file(path)?;
        Self::from_curl_text(&text)
    }

    pub fn from_curl_text(text: &str) -> Result<Self, OutlookWebError> {
        let args = shlex::split(text).ok_or(OutlookWebError::InvalidSessionExport)?;
        if args.first().is_none_or(|arg| !arg.ends_with("curl")) {
            return Err(OutlookWebError::InvalidSessionExport);
        }
        let mut url_text = None;
        let mut headers = BTreeMap::new();
        let mut index = 1;
        while index < args.len() {
            match args[index].as_str() {
                "--url" => {
                    index += 1;
                    url_text = args.get(index).cloned();
                }
                "-H" | "--header" => {
                    index += 1;
                    let header = args
                        .get(index)
                        .ok_or(OutlookWebError::InvalidSessionExport)?;
                    if let Some((name, value)) = header.split_once(':') {
                        let name = name.trim().to_ascii_lowercase();
                        if name
                            .bytes()
                            .all(|byte| byte.is_ascii_alphanumeric() || byte == b'-')
                        {
                            headers.insert(name, value.trim().to_owned());
                        }
                    }
                }
                value if value.starts_with("https://") || value.starts_with("http://") => {
                    url_text = Some(value.to_owned());
                }
                _ => {}
            }
            index += 1;
        }
        let mut origin = Url::parse(&url_text.ok_or(OutlookWebError::InvalidSessionExport)?)
            .map_err(|_| OutlookWebError::InvalidSessionOrigin)?;
        if origin.scheme() != "https"
            || origin.host_str() != Some("graph.microsoft.com")
            || !origin.username().is_empty()
            || origin.password().is_some()
            || origin.port().is_some()
        {
            return Err(OutlookWebError::InvalidSessionOrigin);
        }
        origin.set_path("/");
        origin.set_query(None);
        origin.set_fragment(None);
        headers.remove("content-length");
        headers.remove("content-type");
        Ok(Self {
            origin,
            headers,
            client: http_client()?,
        })
    }

    pub fn request_parts(&self, path: &str) -> Result<RequestParts, OutlookWebError> {
        if !path.starts_with("/v1.0/") {
            return Err(OutlookWebError::InvalidAction);
        }
        let url = self
            .origin
            .join(&path[1..])
            .map_err(|_| OutlookWebError::InvalidSessionOrigin)?;
        let segments: Vec<_> = url.path().split('/').filter(|s| !s.is_empty()).collect();
        let allowed = matches!(
            segments.as_slice(),
            ["v1.0", "me", "people" | "manager" | "directReports"]
                | ["v1.0", "groups"]
                | ["v1.0", "users", _]
                | ["v1.0", "users", _, "manager" | "directReports"]
                | ["v1.0", "me", "drive", ..]
        );
        if !allowed {
            return Err(OutlookWebError::InvalidAction);
        }
        Ok(RequestParts {
            method: "GET",
            url,
            headers: self.headers.clone(),
        })
    }

    pub fn get(&self, path: &str) -> Result<Value, OutlookWebError> {
        let parts = self.request_parts(path)?;
        let mut headers = HeaderMap::new();
        for (name, value) in parts.headers {
            let name = HeaderName::from_bytes(name.as_bytes())
                .map_err(|_| OutlookWebError::InvalidHeader)?;
            let value =
                HeaderValue::from_str(&value).map_err(|_| OutlookWebError::InvalidHeader)?;
            headers.insert(name, value);
        }
        let response = self
            .client
            .get(parts.url)
            .headers(headers)
            .send()
            .map_err(|_| OutlookWebError::RequestFailed)?;
        if response.status() == reqwest::StatusCode::UNAUTHORIZED {
            return Err(OutlookWebError::SessionExpired);
        }
        if !response.status().is_success() {
            return Err(OutlookWebError::RequestRejected(response.status().as_u16()));
        }
        let value: Value = response
            .json()
            .map_err(|_| OutlookWebError::InvalidResponse)?;
        if value.get("error").is_some() {
            return Err(OutlookWebError::InvalidResponse);
        }
        Ok(value)
    }
}
