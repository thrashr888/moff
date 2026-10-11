//! Outlook Web operations backed only by a caller-provided browser-session export.
#![allow(clippy::missing_errors_doc, clippy::must_use_candidate)]
mod graph;
mod operations;
mod safari;
mod scheduling;
#[cfg(test)]
mod test_support;

pub use graph::GraphWebSession;
pub use operations::*;
pub use scheduling::*;

use std::{collections::BTreeMap, fmt, fs, path::Path, time::Duration};

use reqwest::{
    blocking::Client,
    header::{HeaderMap, HeaderName, HeaderValue},
};
use serde_json::{Value, json};
use url::Url;

const EXCHANGE_HEADER: &str = "JsonRequestHeaders:#Exchange";

#[derive(Debug)]
pub enum OutlookWebError {
    SessionFileUnavailable,
    InvalidSessionExport,
    InvalidSessionOrigin,
    InvalidAction,
    InvalidHeader,
    RequestFailed,
    RequestRejected(u16),
    InvalidResponse,
    SessionPermissions,
    SessionExpired,
    InvalidInput(&'static str),
    OperationRejected {
        action: &'static str,
        code: ServiceCode,
    },
    OutcomeUnknown {
        action: &'static str,
    },
    UnsupportedItem,
    TruncatedSearch,
    SafariUnavailable,
    SafariOperationFailed,
    SafariReadOnly,
    VerificationFailed,
    PlanChanged,
}

#[derive(Debug, PartialEq, Eq)]
pub enum ServiceCode {
    ItemNotFound,
    Conflict,
    AccessDenied,
    Other,
}
impl fmt::Display for OutlookWebError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::SessionFileUnavailable => "the Outlook Web session export could not be read",
            Self::InvalidSessionExport => {
                "the Outlook Web session export is not a supported cURL command"
            }
            Self::InvalidSessionOrigin => {
                "the Outlook Web session export does not target an HTTPS Outlook host"
            }
            Self::InvalidAction => "the requested Outlook action is not a supported read operation",
            Self::InvalidHeader => "the Outlook Web session export contains an invalid HTTP header",
            Self::RequestFailed => "Outlook Web did not complete the request",
            Self::RequestRejected(status) => {
                return write!(f, "Outlook Web rejected the request (HTTP {status})");
            }
            Self::InvalidResponse => "Outlook Web returned an invalid response",
            Self::SessionPermissions => "session exports require an owner-only file (0600) and directory (0700)",
            Self::SessionExpired => "the browser session expired (HTTP 401); sign in again in Safari or refresh the browser session export",
            Self::InvalidInput(message) => message,
            Self::UnsupportedItem => "this mutation requires an unsent mail draft, mail message, or single organizer-owned appointment without attendees, as applicable",
            Self::TruncatedSearch => "Outlook returned incomplete search data; narrow the folder or calendar range",
            Self::SafariUnavailable => "select exactly one loaded Outlook Safari tab, or provide both --safari-window and --safari-tab",
            Self::SafariOperationFailed => "Safari operation failed; check Safari Automation permission, JavaScript from Apple Events, and that Outlook is signed in; the private Outlook adapter may need an update",
            Self::SafariReadOnly => "Safari writes require --allow-mutation and --confirm; deletion is not supported through Safari",
            Self::VerificationFailed => "the write completed but verification failed; inspect the affected item before retrying",
            Self::PlanChanged => "the calendar or target appointment changed since preview; generate and review a new plan",
            Self::OperationRejected { action, code } => {
                return write!(f, "Outlook Web rejected {action} ({code:?})");
            }
            Self::OutcomeUnknown { action } => {
                return write!(f, "the outcome of {action} is unknown; do not retry it automatically; inspect the affected folder or calendar first");
            }
        })
    }
}
impl std::error::Error for OutlookWebError {}

pub struct RequestParts {
    pub method: &'static str,
    pub url: Url,
    pub headers: BTreeMap<String, String>,
}

pub struct OutlookWebSession {
    origin: Url,
    headers: BTreeMap<String, String>,
    client: Client,
    safari: Option<safari::SafariSession>,
}
impl fmt::Debug for RequestParts {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("RequestParts")
            .field("method", &self.method)
            .field("url", &self.url)
            .field("header_names", &self.headers.keys())
            .finish()
    }
}
impl fmt::Debug for OutlookWebSession {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("OutlookWebSession")
            .field("origin", &self.origin)
            .field("header_names", &self.headers.keys())
            .finish_non_exhaustive()
    }
}

pub(crate) fn read_session_file(path: &Path) -> Result<String, OutlookWebError> {
    let metadata =
        fs::symlink_metadata(path).map_err(|_| OutlookWebError::SessionFileUnavailable)?;
    if !metadata.is_file() {
        return Err(OutlookWebError::SessionPermissions);
    }
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        let parent = fs::metadata(path.parent().unwrap_or(Path::new(".")))
            .map_err(|_| OutlookWebError::SessionFileUnavailable)?;
        if metadata.permissions().mode() & 0o777 != 0o600
            || parent.permissions().mode() & 0o777 != 0o700
        {
            return Err(OutlookWebError::SessionPermissions);
        }
    }
    fs::read_to_string(path).map_err(|_| OutlookWebError::SessionFileUnavailable)
}

pub(crate) fn http_client() -> Result<Client, OutlookWebError> {
    Client::builder()
        .redirect(reqwest::redirect::Policy::none())
        .retry(reqwest::retry::never())
        .connect_timeout(Duration::from_secs(10))
        .timeout(Duration::from_secs(30))
        .build()
        .map_err(|_| OutlookWebError::RequestFailed)
}

impl OutlookWebSession {
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
                    let Some((name, value)) = header.split_once(':') else {
                        index += 1;
                        continue;
                    };
                    let name = name.trim().to_ascii_lowercase();
                    if name
                        .bytes()
                        .all(|byte| byte.is_ascii_alphanumeric() || byte == b'-')
                    {
                        headers.insert(name, value.trim().to_owned());
                    }
                }
                "-b" | "--cookie" => {
                    index += 1;
                    let cookie = args
                        .get(index)
                        .ok_or(OutlookWebError::InvalidSessionExport)?;
                    if !cookie.contains('=') {
                        return Err(OutlookWebError::InvalidSessionExport);
                    }
                    headers.insert("cookie".to_owned(), cookie.clone());
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
            || !origin.username().is_empty()
            || origin.password().is_some()
            || origin.port().is_some()
            || !origin.host_str().is_some_and(|host| {
                matches!(host, "outlook.office.com" | "outlook.cloud.microsoft")
            })
        {
            return Err(OutlookWebError::InvalidSessionOrigin);
        }
        origin.set_path("/");
        origin.set_query(None);
        origin.set_fragment(None);
        for stale in [
            "action",
            "content-length",
            "content-type",
            "x-owa-urlpostdata",
        ] {
            headers.remove(stale);
        }
        Ok(Self {
            origin,
            headers,
            client: http_client()?,
            safari: None,
        })
    }
    pub fn request_parts(
        &self,
        action: &str,
        _body: &Value,
    ) -> Result<RequestParts, OutlookWebError> {
        validate_read_action(action)?;
        self.parts(action)
    }
    fn parts(&self, action: &str) -> Result<RequestParts, OutlookWebError> {
        let url = self
            .origin
            .join(&format!("owa/service.svc?action={action}&app=Mail"))
            .map_err(|_| OutlookWebError::InvalidSessionOrigin)?;
        let mut headers = self.headers.clone();
        headers.insert("action".to_owned(), action.to_owned());
        headers.insert("content-type".to_owned(), "application/json".to_owned());
        Ok(RequestParts {
            method: "POST",
            url,
            headers,
        })
    }
    pub fn call(&self, action: &str, body: &Value) -> Result<Value, OutlookWebError> {
        let action = READ_ACTIONS
            .iter()
            .copied()
            .find(|name| *name == action)
            .ok_or(OutlookWebError::InvalidAction)?;
        self.execute(action, body, false)
    }
    fn execute(
        &self,
        action: &'static str,
        body: &Value,
        mutation: bool,
    ) -> Result<Value, OutlookWebError> {
        if let Some(session) = &self.safari {
            let value = session.call(action, body, mutation)?;
            check_response(action, &value).map_err(|error| match error {
                OutlookWebError::InvalidResponse if mutation => {
                    OutlookWebError::OutcomeUnknown { action }
                }
                other => other,
            })?;
            return Ok(value);
        }
        let parts = self.parts(action)?;
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
            .post(parts.url)
            .headers(headers)
            .json(body)
            .send()
            .map_err(|_| {
                if mutation {
                    OutlookWebError::OutcomeUnknown { action }
                } else {
                    OutlookWebError::RequestFailed
                }
            })?;
        if response.status() == reqwest::StatusCode::UNAUTHORIZED {
            return Err(OutlookWebError::SessionExpired);
        }
        if !response.status().is_success() {
            if mutation && response.status().is_server_error() {
                return Err(OutlookWebError::OutcomeUnknown { action });
            }
            return Err(OutlookWebError::RequestRejected(response.status().as_u16()));
        }
        let parsed = response
            .text()
            .map_err(|_| OutlookWebError::InvalidResponse)
            .and_then(|text| decode_response(&text));
        let value = parsed.map_err(|_| {
            if mutation {
                OutlookWebError::OutcomeUnknown { action }
            } else {
                OutlookWebError::InvalidResponse
            }
        })?;
        check_response(action, &value).map_err(|error| match error {
            OutlookWebError::InvalidResponse if mutation => {
                OutlookWebError::OutcomeUnknown { action }
            }
            other => other,
        })?;
        Ok(value)
    }
}

pub fn decode_response(text: &str) -> Result<Value, OutlookWebError> {
    let trimmed = text.trim_start_matches(['\u{feff}', '\n', '\r', ' ', '\t']);
    let payload = if trimmed.starts_with(['{', '[']) {
        trimmed
    } else {
        trimmed
            .get(trimmed.find('{').ok_or(OutlookWebError::InvalidResponse)?..)
            .ok_or(OutlookWebError::InvalidResponse)?
    };
    serde_json::from_str(payload).map_err(|_| OutlookWebError::InvalidResponse)
}
const READ_ACTIONS: &[&str] = &[
    "FindItem",
    "FindConversation",
    "GetItem",
    "GetConversationItems",
    "GetCalendarFolders",
    "GetCalendarFoldersV2",
    "GetCalendarView",
];
pub fn validate_read_action(action: &str) -> Result<(), OutlookWebError> {
    if READ_ACTIONS.contains(&action) {
        Ok(())
    } else {
        Err(OutlookWebError::InvalidAction)
    }
}

fn check_response(action: &'static str, value: &Value) -> Result<(), OutlookWebError> {
    let rejected = |code| OutlookWebError::OperationRejected { action, code };
    if value.get("FaultMessage").is_some()
        || value.get("Fault").is_some()
        || value.pointer("/Body/FaultMessage").is_some()
        || value.pointer("/Body/Fault").is_some()
    {
        return Err(rejected(ServiceCode::Other));
    }
    if let Some(success) = value.get("WasSuccessful").and_then(Value::as_bool) {
        return if success && value.get("ErrorCode").and_then(Value::as_i64) == Some(0) {
            Ok(())
        } else {
            Err(rejected(ServiceCode::Other))
        };
    }
    let messages: Vec<&Value> = if let Some(items) = value
        .pointer("/Body/ResponseMessages/Items")
        .and_then(Value::as_array)
    {
        if items.is_empty() {
            return Err(OutlookWebError::InvalidResponse);
        }
        items.iter().collect()
    } else {
        vec![value.get("Body").ok_or(OutlookWebError::InvalidResponse)?]
    };
    for message in messages {
        let code = message
            .get("ResponseCode")
            .and_then(Value::as_str)
            .ok_or(OutlookWebError::InvalidResponse)?;
        if code != "NoError" {
            return Err(rejected(match code {
                "ErrorItemNotFound" => ServiceCode::ItemNotFound,
                "ErrorIrresolvableConflict"
                | "ErrorChangeKeyRequired"
                | "ErrorInvalidChangeKey"
                | "ErrorChangeKeyRequiredForWriteOperations" => ServiceCode::Conflict,
                "ErrorAccessDenied" => ServiceCode::AccessDenied,
                _ => ServiceCode::Other,
            }));
        }
        if message
            .get("ResponseClass")
            .is_some_and(|class| class != "Success")
        {
            return Err(rejected(ServiceCode::Other));
        }
    }
    Ok(())
}
fn header() -> Value {
    json!({"__type": EXCHANGE_HEADER, "RequestServerVersion":"V2018_01_08"})
}
pub fn find_item_request(limit: u16) -> Value {
    json!({"Header":header(),"Body":{"__type":"FindItemRequest:#Exchange","Traversal":"Shallow","ItemShape":{"__type":"ItemResponseShape:#Exchange","BaseShape":"Default"},"ParentFolderIds":[{"__type":"DistinguishedFolderId:#Exchange","Id":"inbox"}],"Paging":{"__type":"IndexedPageView:#Exchange","BasePoint":"Beginning","Offset":0,"MaxEntriesReturned":limit}}})
}
pub fn get_item_request(id: &str) -> Value {
    json!({"Header":header(),"Body":{"__type":"GetItemRequest:#Exchange","ItemShape":{"__type":"ItemResponseShape:#Exchange","BaseShape":"AllProperties","BodyType":"Text"},"ItemIds":[{"__type":"ItemId:#Exchange","Id":id}]}})
}
pub fn calendar_folders_request() -> Value {
    json!({"ShouldReturnLinkedCalendarEntryV2":false,"IncludeOrganizationCalendars":false,"Header":header()})
}
pub fn calendar_view_request(folder_id: &str, start: &str, end: &str) -> Value {
    json!({"Header":header(),"Body":{"__type":"GetCalendarViewRequest:#Exchange","CalendarId":{"__type":"TargetFolderId:#Exchange","BaseFolderId":{"__type":"FolderId:#Exchange","Id":folder_id}},"RangeStart":start,"RangeEnd":end}})
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Read, Write};
    use std::net::TcpListener;
    use std::thread;

    #[test]
    fn application_faults_and_unknown_response_shapes_are_not_success() {
        for value in [
            json!({"FaultMessage":"private-value"}),
            json!({"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Error","ResponseCode":"ErrorAccessDenied","MessageText":"private-value"}]}}}),
            json!({"WasSuccessful":false,"ErrorCode":1}),
        ] {
            let error = check_response("CreateItem", &value).unwrap_err();
            assert!(!error.to_string().contains("private-value"));
        }
        assert!(check_response("CreateItem", &json!({})).is_err());
        assert!(
            check_response(
                "CreateItem",
                &json!({"Body":{"ResponseMessages":{"Items":[]}}})
            )
            .is_err()
        );
        assert!(
            check_response(
                "GetCalendarFoldersV2",
                &json!({"WasSuccessful":true,"ErrorCode":0})
            )
            .is_ok()
        );
    }

    #[test]
    fn debug_formatting_does_not_expose_browser_headers() {
        let session = OutlookWebSession::from_curl_text(
            "curl 'https://outlook.cloud.microsoft/owa/service.svc' -H 'Authorization: synthetic-private-value' -b 'session=synthetic-private-value'"
        ).unwrap();
        assert!(!format!("{session:?}").contains("synthetic-private-value"));
        let request = session.request_parts("FindItem", &json!({})).unwrap();
        assert!(!format!("{request:?}").contains("synthetic-private-value"));
        assert!(request.headers.contains_key("cookie"));
        assert!(
            OutlookWebSession::from_curl_text(
                "curl 'https://notoutlook.office.com/owa/service.svc'"
            )
            .is_err()
        );
    }

    #[test]
    fn ambiguous_write_is_not_retried_and_does_not_follow_redirects() {
        for status in ["500 Internal Server Error", "302 Found"] {
            let listener = TcpListener::bind("127.0.0.1:0").unwrap();
            let address = listener.local_addr().unwrap();
            let worker = thread::spawn(move || {
                let (mut socket, _) = listener.accept().unwrap();
                socket
                    .set_read_timeout(Some(Duration::from_secs(5)))
                    .unwrap();
                let mut request = Vec::new();
                loop {
                    let mut bytes = [0u8; 2048];
                    let count = socket.read(&mut bytes).unwrap();
                    assert!(count > 0);
                    request.extend_from_slice(&bytes[..count]);
                    if let Some(end) = request.windows(4).position(|v| v == b"\r\n\r\n") {
                        let headers = std::str::from_utf8(&request[..end]).unwrap();
                        let length = headers
                            .lines()
                            .find_map(|line| {
                                let (name, value) = line.split_once(':')?;
                                name.eq_ignore_ascii_case("content-length")
                                    .then(|| value.trim().parse::<usize>().unwrap())
                            })
                            .unwrap();
                        if request.len() >= end + 4 + length {
                            break;
                        }
                    }
                }
                let response = format!(
                    "HTTP/1.1 {status}\r\nContent-Length: 0\r\nLocation: http://{address}/unexpected\r\nConnection: close\r\n\r\n"
                );
                socket.write_all(response.as_bytes()).unwrap();
                drop(socket);
                listener.set_nonblocking(true).unwrap();
                thread::sleep(Duration::from_millis(100));
                assert!(listener.accept().is_err());
                assert!(
                    String::from_utf8(request)
                        .unwrap()
                        .starts_with("POST /owa/service.svc?action=CreateItem&app=Mail ")
                );
            });
            let session = OutlookWebSession {
                origin: Url::parse(&format!("http://{address}/")).unwrap(),
                headers: BTreeMap::new(),
                client: http_client().unwrap(),
                safari: None,
            };
            let error = session.create_draft(&DraftInput::default()).unwrap_err();
            if status.starts_with("500") {
                assert!(matches!(
                    error,
                    OutlookWebError::OutcomeUnknown {
                        action: "CreateItem"
                    }
                ));
            } else {
                assert!(matches!(error, OutlookWebError::RequestRejected(302)));
            }
            worker.join().unwrap();
        }
    }
}
