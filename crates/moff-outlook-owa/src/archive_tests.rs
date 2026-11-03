use super::*;
use crate::test_support::{message, mock};

fn folder(id: &str, name: &str) -> Value {
    json!({"Body":{"ResponseMessages":{"Items":[{"ResponseCode":"NoError","Folders":[{
        "FolderId":{"Id":id},"DisplayName":name,"FolderClass":"IPF.Note"
    }]}]}}})
}

fn item(id: &str, key: &str, parent: &str) -> Value {
    message(json!({
        "__type":"Message:#Exchange","ItemId":{"Id":id,"ChangeKey":key},
        "Subject":"OCM synthetic fixture","ParentFolderId":{"Id":parent},"IsDraft":false
    }))
}

fn plan() -> ArchivePlan {
    ArchivePlan {
        archive_folder_id: "archive-id".into(),
        inbox_folder_id: "inbox-id".into(),
        messages: vec![ArchiveMessage {
            id: "message-id".into(),
            change_key: "key".into(),
            subject: "OCM synthetic fixture".into(),
        }],
    }
}

fn preview_replies() -> Vec<(&'static str, Value)> {
    vec![
        ("GetFolder", folder("archive-id", "Archive")),
        ("GetFolder", folder("inbox-id", "Inbox")),
        ("GetItem", item("message-id", "key", "inbox-id")),
    ]
}

#[test]
fn preview_reads_only_and_move_verifies_new_id_in_archive() {
    let (session, requests, completion) = mock(preview_replies());
    let actual = session
        .preview_mail_archive(&["message-id".into()], "archive-id")
        .unwrap();
    assert_eq!(actual.messages[0].change_key, "key");
    assert_eq!(requests.lock().unwrap().len(), 3);
    completion.join().unwrap();

    let mut replies = preview_replies();
    replies.extend([
        ("MoveItem", message(json!({"ItemId":{"Id":"moved-id"}}))),
        ("GetItem", item("moved-id", "new-key", "archive-id")),
    ]);
    let (session, requests, completion) = mock(replies);
    let result = session.apply_mail_archive(&plan()).unwrap();
    assert_eq!(result["ok"], true);
    assert_eq!(result["completed"][0]["archived_id"], "moved-id");
    let requests = requests.lock().unwrap();
    assert_eq!(
        requests[3]["Body"]["ToFolderId"]["BaseFolderId"]["Id"],
        "archive-id"
    );
    assert_eq!(requests[3]["Body"]["ItemIds"][0]["ChangeKey"], "key");
    assert_eq!(requests[4]["Body"]["ItemIds"][0]["Id"], "moved-id");
    completion.join().unwrap();
}

#[test]
fn meeting_notices_are_mail_but_calendar_events_are_not() {
    for kind in [
        "MeetingCancellationMessage:#Exchange",
        "MeetingResponseMessage:#Exchange",
        "MeetingRequestMessage:#Exchange",
        "CalendarItem:#Exchange",
    ] {
        let mut replies = preview_replies();
        replies[2].1["Body"]["ResponseMessages"]["Items"][0]["Items"][0]["__type"] = json!(kind);
        let (session, _, completion) = mock(replies);
        let result = session.preview_mail_archive(&["message-id".into()], "archive-id");
        assert_eq!(result.is_ok(), kind != "CalendarItem:#Exchange");
        completion.join().unwrap();
    }
}

#[test]
fn rejects_bad_ids_destinations_non_inbox_messages_and_stale_plans_before_writes() {
    let (session, requests, completion) = mock(vec![]);
    for ids in [
        vec![],
        vec!["a".into(), "a".into()],
        vec![String::new()],
        vec!["a".into(); 21],
    ] {
        assert!(session.preview_mail_archive(&ids, "archive-id").is_err());
    }
    assert!(requests.lock().unwrap().is_empty());
    completion.join().unwrap();
    let (session, _, completion) = mock(vec![("GetFolder", folder("archive-id", "Deleted Items"))]);
    assert!(session.apply_mail_archive(&plan()).is_err());
    completion.join().unwrap();
    for (key, parent) in [("key", "sent-id"), ("changed", "inbox-id")] {
        let mut replies = preview_replies();
        replies[2] = ("GetItem", item("message-id", key, parent));
        let (session, requests, completion) = mock(replies);
        assert!(session.apply_mail_archive(&plan()).is_err());
        assert_eq!(requests.lock().unwrap().len(), 3);
        completion.join().unwrap();
    }
}

#[test]
fn partial_failure_reports_completed_and_uncertain_ids_without_retry() {
    let mut plan = plan();
    plan.messages.push(ArchiveMessage {
        id: "second".into(),
        change_key: "key2".into(),
        subject: "OCM synthetic fixture".into(),
    });
    let mut replies = preview_replies();
    replies.push(("GetItem", item("second", "key2", "inbox-id")));
    replies.extend([
        ("MoveItem", message(json!({"ItemId":{"Id":"moved-id"}}))),
        ("GetItem", item("moved-id", "new-key", "archive-id")),
        ("MoveItem", json!({})),
    ]);
    let (session, requests, completion) = mock(replies);
    let result = session.apply_mail_archive(&plan).unwrap();
    assert_eq!(result["ok"], false);
    assert_eq!(result["completed"].as_array().unwrap().len(), 1);
    assert_eq!(result["failed_or_uncertain_id"], "second");
    assert!(result["error"].as_str().unwrap().contains("unknown"));
    assert_eq!(requests.lock().unwrap().len(), 7);
    completion.join().unwrap();
}

#[test]
fn failed_verification_is_not_reported_as_success() {
    let mut replies = preview_replies();
    replies.extend([
        ("MoveItem", message(json!({"ItemId":{"Id":"moved-id"}}))),
        ("GetItem", item("moved-id", "new-key", "inbox-id")),
    ]);
    let (session, _, completion) = mock(replies);
    let result = session.apply_mail_archive(&plan()).unwrap();
    assert_eq!(result["ok"], false);
    assert_eq!(result["completed"], json!([]));
    assert_eq!(result["failed_or_uncertain_id"], "message-id");
    completion.join().unwrap();
}
