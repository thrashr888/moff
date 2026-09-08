use std::path::Path;

use moff_outlook_owa::{
    GraphWebSession, OutlookWebSession, calendar_view_request, find_item_request, get_item_request,
};
use serde_json::json;

#[test]
fn session_replaces_request_specific_headers_without_retaining_credentials_in_errors() {
    let session = OutlookWebSession::from_curl_text(
        "curl --url 'https://outlook.office.com/owa/service.svc?action=GetOwaUserConfiguration&app=Mail' -X POST -H 'Authorization: Bearer pretend-token' -H 'Action: GetOwaUserConfiguration' -H 'Content-Length: 0' -H 'X-OWA-URLPOSTDATA: stale' -H 'X-OWA-CANARY: canary-value'",
    )
    .unwrap();

    let request = session
        .request_parts("FindItem", &find_item_request(5))
        .unwrap();
    assert_eq!(request.method, "POST");
    assert_eq!(request.url.path(), "/owa/service.svc");
    assert_eq!(request.url.query(), Some("action=FindItem&app=Mail"));
    assert_eq!(request.headers.get("action").unwrap(), "FindItem");
    assert_eq!(
        request.headers.get("content-type").unwrap(),
        "application/json"
    );
    assert!(!request.headers.contains_key("content-length"));
    assert!(!request.headers.contains_key("x-owa-urlpostdata"));
    assert!(request.headers.contains_key("authorization"));
    assert!(request.headers.contains_key("x-owa-canary"));
}

#[test]
fn mail_and_calendar_requests_have_exchange_contracts() {
    let find = find_item_request(5);
    assert_eq!(find["Body"]["__type"], "FindItemRequest:#Exchange");
    assert_eq!(find["Body"]["ParentFolderIds"][0]["Id"], "inbox");
    assert_eq!(find["Body"]["Paging"]["MaxEntriesReturned"], 5);

    let get = get_item_request("opaque-id");
    assert_eq!(get["Body"]["__type"], "GetItemRequest:#Exchange");
    assert_eq!(get["Body"]["ItemShape"]["BodyType"], "Text");

    let calendar = calendar_view_request("folder-id", "2026-09-01T00:00:00", "2026-09-02T00:00:00");
    assert_eq!(
        calendar["Body"]["__type"],
        "GetCalendarViewRequest:#Exchange"
    );
    assert_eq!(
        calendar["Body"]["CalendarId"]["BaseFolderId"]["Id"],
        "folder-id"
    );
}

#[test]
fn framed_owa_response_is_decoded() {
    let parsed =
        moff_outlook_owa::decode_response("\n000\n000\n{\"Body\":{\"ResponseCode\":\"NoError\"}}")
            .unwrap();
    assert_eq!(parsed, json!({"Body":{"ResponseCode":"NoError"}}));
}

#[test]
fn session_file_must_exist() {
    let missing =
        OutlookWebSession::from_curl_file(Path::new("/definitely/missing/moff-session-curl"));
    assert!(missing.is_err());
    assert!(!missing.unwrap_err().to_string().contains("pretend-token"));
}

#[test]
fn graph_session_builds_get_without_forwarding_request_body_headers() {
    let session = GraphWebSession::from_curl_text(
        "curl --url 'https://graph.microsoft.com/v1.0/users/example' -X GET -H 'Authorization: Bearer ***' -H 'Content-Type: application/json' -H 'Content-Length: 99'",
    )
    .unwrap();
    let request = session
        .request_parts("/v1.0/me/drive/root/children?$top=25")
        .unwrap();
    assert_eq!(request.method, "GET");
    assert_eq!(
        request.url.as_str(),
        "https://graph.microsoft.com/v1.0/me/drive/root/children?$top=25"
    );
    assert!(request.headers.contains_key("authorization"));
    assert!(!request.headers.contains_key("content-type"));
    assert!(!request.headers.contains_key("content-length"));
}

#[test]
fn graph_session_rejects_outlook_mail_and_calendar_paths() {
    let session = GraphWebSession::from_curl_text(
        "curl --url 'https://graph.microsoft.com/v1.0/me/messages' -H 'Authorization: Bearer ***'",
    )
    .unwrap();
    assert!(
        session
            .request_parts("/v1.0/me/messages/message-id")
            .is_err()
    );
    assert!(session.request_parts("/v1.0/me/events/event-id").is_err());
}
