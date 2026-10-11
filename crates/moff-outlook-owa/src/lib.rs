//! Outlook Mail and Calendar operations through Safari and the Cider library.
#![allow(clippy::missing_errors_doc, clippy::must_use_candidate)]
mod contacts;
mod operations;
pub use contacts::validate_contact_input;
mod products;
pub use products::{ProductWrite, product_segment};
mod safari;
mod scheduling;
#[cfg(test)]
mod test_support;

pub use operations::*;
pub use scheduling::*;

use serde_json::{Value, json};
use std::fmt;

const EXCHANGE_HEADER: &str = "JsonRequestHeaders:#Exchange";

#[derive(Debug)]
pub enum OutlookWebError {
    InvalidAction,
    RequestRejected(u16),
    InvalidResponse,
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
    SafariRefreshUnknown,
    SafariAuthRecoveryFailed(&'static str),
    ProductAuthenticationUnavailable,
    ProductRequestRejected(u16),
    SafariReadOnly,
    VerificationFailed,
    ItemVerificationFailed {
        id: String,
    },
    DraftVerificationFailed {
        draft_id: String,
        source_id: String,
    },
    PlanChanged,
    SafariBeforeDispatch,
    SafariOutcomeUnknown {
        action: &'static str,
        phase: cider::sources::safari::OperationPhase,
        cancellation: cider::sources::safari::CancellationStatus,
    },
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
            Self::InvalidAction => "the requested Outlook action is not a supported read operation",
            Self::RequestRejected(status) => {
                return write!(f, "Outlook Web rejected the request (HTTP {status})");
            }
            Self::InvalidResponse => "Outlook Web returned an invalid response",
            Self::SessionExpired => "the browser session expired (HTTP 401); run moff outlook auth refresh --confirm, then complete any sign-in steps yourself in Safari",
            Self::InvalidInput(message) => message,
            Self::UnsupportedItem => "this mutation requires an unsent mail draft, mail message, or single organizer-owned appointment without attendees, as applicable",
            Self::TruncatedSearch => "Outlook returned incomplete search data; narrow the folder or calendar range",
            Self::SafariUnavailable => "select exactly one loaded Outlook Safari tab, or provide both --safari-window and --safari-tab",
            Self::SafariOperationFailed => "Safari operation failed; check Safari Automation permission, JavaScript from Apple Events, and that Outlook is signed in; the private Outlook adapter may need an update",
            Self::SafariRefreshUnknown => "Safari did not confirm the scheduled refresh; it may still happen. Inspect the selected tab before retrying; no mailbox write was attempted",
            Self::SafariAuthRecoveryFailed(stage) => {
                return write!(f, "Safari authentication recovery failed during {stage}; no mailbox write was attempted; inspect Safari or update the private Outlook adapter");
            }
            Self::ProductAuthenticationUnavailable => "browser product authentication is unavailable; sign into Outlook and check the product's permitted browser access; Moff will not request consent or export credentials",
            Self::ProductRequestRejected(status) => {
                return write!(f, "the product API rejected the browser request (HTTP {status}); access may require permissions not held by this session");
            }
            Self::SafariReadOnly => "Safari writes require --allow-mutation and --confirm",
            Self::VerificationFailed => "the write completed but verification failed; inspect the affected item before retrying",
            Self::ItemVerificationFailed { id } => {
                return write!(f,"item {id} was written but read-back verification failed; inspect it before retrying");
            }
            Self::DraftVerificationFailed { draft_id, source_id } => {
                return write!(f, "draft {draft_id} was created but verification or read-state restoration failed for source {source_id}; inspect both items; do not recreate the draft automatically");
            }
            Self::PlanChanged => "the calendar or target appointment changed since preview; generate and review a new plan",
            Self::SafariBeforeDispatch => "Safari target failed before dispatch; no caller operation was started; sign in and select the intended document again",
            Self::SafariOutcomeUnknown { action, phase, cancellation } => {
                return write!(f,"the outcome of {action} is unknown (phase: {phase:?}, cancellation: {cancellation:?}); cancellation does not undo server writes; inspect the affected item before retrying");
            }
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

pub struct OutlookWebSession {
    transport: SessionTransport,
}
enum SessionTransport {
    Safari(safari::SafariSession),
    #[cfg(test)]
    Mock(test_support::MockSession),
}
impl fmt::Debug for OutlookWebSession {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("OutlookWebSession")
            .field("transport", &"safari")
            .finish_non_exhaustive()
    }
}

impl OutlookWebSession {
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
        let value = match &self.transport {
            SessionTransport::Safari(session) => session.call(action, body, mutation)?,
            #[cfg(test)]
            SessionTransport::Mock(session) => session.call(action, body),
        };
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
    fn debug_formatting_does_not_expose_message_data() {
        let (session, _, completion) = test_support::mock(vec![]);
        assert_eq!(
            format!("{session:?}"),
            "OutlookWebSession { transport: \"safari\", .. }"
        );
        completion.join().unwrap();
    }
}
