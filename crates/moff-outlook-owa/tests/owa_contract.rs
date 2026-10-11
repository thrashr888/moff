use moff_outlook_owa::{calendar_view_request, find_item_request, get_item_request};
use serde_json::json;

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
