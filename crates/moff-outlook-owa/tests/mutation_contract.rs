use moff_outlook_owa::{
    DraftInput, EventInput, EventMoveInput, MeetingResponse, MeetingResponseInput, ReplyInput,
    batch_item_request, cancel_meeting_request, categories_request, create_draft_request,
    create_event_request, delete_event_request, delete_message_request, move_event_request,
    reply_draft_request, respond_event_request, search_item_request, update_draft_request,
    update_event_request, validate_read_action,
};
use serde_json::json;

#[test]
fn draft_writes_use_observed_exchange_envelopes_and_never_send() {
    let input: DraftInput =
        serde_json::from_value(json!({"subject":"synthetic", "body":"<text>"})).unwrap();
    let create = create_draft_request(&input).unwrap();
    assert_eq!(create["__type"], "CreateItemJsonRequest:#Exchange");
    assert_eq!(create["Body"]["MessageDisposition"], "SaveOnly");
    assert_eq!(create["Body"]["Items"][0]["__type"], "Message:#Exchange");
    assert!(create["Body"].get("SendOnNotFoundError").is_none());
    let update = update_draft_request("synthetic-id", "synthetic-key", &input).unwrap();
    assert_eq!(
        update["Body"]["ItemChanges"][0]["ItemId"]["ChangeKey"],
        "synthetic-key"
    );
    assert_eq!(
        update["Body"]["ItemChanges"][0]["Updates"][0]["Path"]["FieldURI"],
        "Subject"
    );
    assert_eq!(update["Body"]["MessageDisposition"], "SaveOnly");
    let delete = delete_message_request("synthetic-id", "synthetic-key").unwrap();
    assert_eq!(delete["Body"]["DeleteType"], "MoveToDeletedItems");
    assert_eq!(delete["Body"]["SendMeetingCancellations"], "SendToNone");
    assert_eq!(delete["Body"]["SuppressReadReceipts"], true);
}

#[test]
fn search_uses_minimal_properties_instead_of_all_properties() {
    let request = search_item_request("synthetic-id");
    let shape = &request["Body"]["ItemShape"];
    assert_eq!(shape["BaseShape"], "IdOnly");
    assert_eq!(shape["BodyType"], "Text");
    assert_eq!(
        shape["AdditionalProperties"],
        json!([
            {"__type":"PropertyUri:#Exchange","FieldURI":"Subject"},
            {"__type":"PropertyUri:#Exchange","FieldURI":"Body"}
        ])
    );
}

#[test]
fn calendar_writes_use_distinct_observed_actions_and_single_event_scope() {
    let input: EventInput = serde_json::from_value(json!({
        "subject":"synthetic", "body":"test",
        "start":"2026-09-20T12:00:00Z", "end":"2026-09-20T13:00:00Z"
    }))
    .unwrap();
    let create = create_event_request("synthetic-folder", &input).unwrap();
    assert_eq!(create["__type"], "CreateItemJsonRequest:#Exchange");
    assert_eq!(
        create["Body"]["Items"][0]["__type"],
        "CalendarItem:#Exchange"
    );
    assert_eq!(create["Body"]["SendMeetingInvitations"], "SendToNone");
    assert_eq!(create["Body"]["Items"][0]["ReminderIsSet"], false);
    for property in [
        "StartTimeZoneId",
        "EndTimeZoneId",
        "ParentFolderId",
        "IsDraft",
        "IsOrganizer",
    ] {
        assert!(create["Body"]["Items"][0].get(property).is_none());
    }
    assert_eq!(create["Body"]["Items"][0]["LegacyFreeBusyStatus"], "Busy");
    let update = update_event_request("synthetic-id", &input).unwrap();
    assert_eq!(update["__type"], "UpdateCalendarEventJsonRequest:#Exchange");
    assert_eq!(update["Body"]["EventScope"], 0);
    assert_eq!(update["Body"]["ShouldSendUpdateToAttendees"], false);
    let delete = delete_event_request("synthetic-id", "synthetic-key").unwrap();
    assert_eq!(delete["Body"]["__type"], "DeleteItemRequest:#Exchange");
    assert_eq!(delete["Body"]["DeleteType"], "MoveToDeletedItems");
    assert_eq!(delete["Body"]["ItemIds"][0]["ChangeKey"], "synthetic-key");
}

#[test]
fn meeting_responses_send_only_typed_rsvps_with_optional_message() {
    let meeting = json!({
        "__type":"CalendarItem:#Exchange",
        "ItemId":{"__type":"ItemId:#Exchange","Id":"synthetic-id","ChangeKey":"synthetic-key"},
        "IsMeeting":true,
        "IsOrganizer":false,
        "CalendarItemType":"Occurrence"
    });
    let input = MeetingResponseInput {
        response: MeetingResponse::Decline,
        message: Some("I will be at HashiConf.".into()),
        proposed_start: Some("2026-10-28T16:00:00Z".into()),
        proposed_end: Some("2026-10-28T17:00:00Z".into()),
    };
    let request = respond_event_request(&meeting, &input).unwrap();
    assert_eq!(request["__type"], "CreateItemJsonRequest:#Exchange");
    assert_eq!(request["Body"]["MessageDisposition"], "SendAndSaveCopy");
    assert_eq!(
        request["Body"]["Items"][0]["__type"],
        "DeclineItem:#Exchange"
    );
    assert_eq!(
        request["Body"]["Items"][0]["ReferenceItemId"]["ChangeKey"],
        "synthetic-key"
    );
    assert_eq!(
        request["Body"]["Items"][0]["Body"]["Value"],
        "I will be at HashiConf."
    );
    assert_eq!(request["Body"]["Items"][0]["Body"]["BodyType"], "Text");
    assert_eq!(
        request["Body"]["Items"][0]["ProposedStart"],
        "2026-10-28T16:00:00Z"
    );
    for (meeting, response) in [
        (
            json!({"__type":"CalendarItem:#Exchange","IsMeeting":false,"IsOrganizer":false,"CalendarItemType":"Single"}),
            MeetingResponse::Accept,
        ),
        (
            json!({"__type":"CalendarItem:#Exchange","IsMeeting":true,"IsOrganizer":true,"CalendarItemType":"Single"}),
            MeetingResponse::Decline,
        ),
        (
            json!({"__type":"CalendarItem:#Exchange","IsMeeting":true,"IsOrganizer":false,"CalendarItemType":"RecurringMaster"}),
            MeetingResponse::Tentative,
        ),
    ] {
        assert!(
            respond_event_request(
                &meeting,
                &MeetingResponseInput {
                    response,
                    message: None,
                    proposed_start: None,
                    proposed_end: None,
                }
            )
            .is_err()
        );
        assert!(
            respond_event_request(
                &meeting,
                &MeetingResponseInput {
                    response: MeetingResponse::Tentative,
                    message: None,
                    proposed_start: Some("2026-10-28T17:00:00Z".into()),
                    proposed_end: Some("2026-10-28T16:00:00Z".into()),
                }
            )
            .is_err()
        );
    }
    assert!(
        serde_json::from_value::<MeetingResponseInput>(json!({
            "response":"decline","message":"Away","recipients":["someone@example.com"]
        }))
        .is_err()
    );
    assert!(
        respond_event_request(
            &meeting,
            &MeetingResponseInput {
                response: MeetingResponse::Accept,
                message: None,
                proposed_start: Some("2026-10-28T16:00:00Z".into()),
                proposed_end: Some("2026-10-28T17:00:00Z".into()),
            }
        )
        .is_err()
    );
}

#[test]
fn moving_meetings_requires_organizer_and_notifies_attendees() {
    let meeting = json!({
        "__type":"CalendarItem:#Exchange",
        "ItemId":{"__type":"ItemId:#Exchange","Id":"synthetic-id","ChangeKey":"synthetic-key"},
        "IsMeeting":true,
        "IsOrganizer":true,
        "CalendarItemType":"Occurrence"
    });
    let input = EventMoveInput {
        start: "2026-10-28T16:00:00Z".into(),
        end: "2026-10-28T17:00:00Z".into(),
    };
    let request = move_event_request(&meeting, &input).unwrap();
    assert_eq!(
        request["Body"]["__type"],
        "UpdateCalendarEventRequest:#Exchange"
    );
    assert_eq!(request["Body"]["ShouldSendUpdateToAttendees"], true);
    assert_eq!(
        request["Body"]["ItemChange"]["ItemId"]["ChangeKey"],
        "synthetic-key"
    );
    assert_eq!(
        request["Body"]["ItemChange"]["Updates"][0]["Path"]["FieldURI"],
        "Start"
    );
    assert_eq!(
        request["Body"]["ItemChange"]["Updates"][1]["Path"]["FieldURI"],
        "End"
    );
    let attendee_meeting = json!({
        "__type":"CalendarItem:#Exchange",
        "ItemId":{"Id":"id","ChangeKey":"key"},
        "IsMeeting":true,
        "IsOrganizer":false,
        "CalendarItemType":"Single"
    });
    assert!(move_event_request(&attendee_meeting, &input).is_err());
}

#[test]
fn cancelling_meetings_notifies_attendees_and_requires_the_organizer() {
    let meeting = json!({
        "__type":"CalendarItem:#Exchange",
        "ItemId":{"__type":"ItemId:#Exchange","Id":"synthetic-id","ChangeKey":"synthetic-key"},
        "IsMeeting":true,
        "IsOrganizer":true,
        "CalendarItemType":"Single"
    });
    let request = cancel_meeting_request(&meeting).unwrap();
    assert_eq!(request["Body"]["__type"], "DeleteItemRequest:#Exchange");
    assert_eq!(request["Body"]["DeleteType"], "MoveToDeletedItems");
    assert_eq!(
        request["Body"]["SendMeetingCancellations"],
        "SendToAllAndSaveCopy"
    );
    assert_eq!(request["Body"]["ItemIds"][0]["ChangeKey"], "synthetic-key");
    let attendee_meeting = json!({
        "__type":"CalendarItem:#Exchange",
        "ItemId":{"Id":"id","ChangeKey":"key"},
        "IsMeeting":true,
        "IsOrganizer":false,
        "CalendarItemType":"Single"
    });
    assert!(cancel_meeting_request(&attendee_meeting).is_err());
}

#[test]
fn triage_builders_do_not_mark_read_send_or_replace_unrelated_fields() {
    let update =
        categories_request("id", "key", &["Existing".into(), "Needs Response".into()]).unwrap();
    assert_eq!(update["Body"]["ConflictResolution"], "NeverOverwrite");
    let changes = update["Body"]["ItemChanges"][0]["Updates"]
        .as_array()
        .unwrap();
    assert_eq!(changes.len(), 1);
    assert_eq!(changes[0]["Path"]["FieldURI"], "Categories");
    let reply = reply_draft_request(
        "id",
        "key",
        &ReplyInput {
            body: "<draft>".into(),
            reply_all: false,
        },
    )
    .unwrap();
    assert_eq!(reply["Body"]["MessageDisposition"], "SaveOnly");
    assert_eq!(reply["Body"]["Items"][0]["__type"], "ReplyToItem:#Exchange");
    assert_eq!(reply["Body"]["Items"][0]["ReferenceItemId"]["Id"], "id");
    assert_eq!(
        reply["Body"]["Items"][0]["NewBodyContent"]["Value"],
        "&lt;draft&gt;"
    );
    assert!(batch_item_request(&[], false).is_err());
    assert!(batch_item_request(&["id".into(), "id".into()], false).is_err());
    let batch = batch_item_request(&["a".into(), "b".into()], false).unwrap();
    assert_eq!(batch["Body"]["ItemIds"].as_array().unwrap().len(), 2);
    assert!(batch.to_string().find("\"IsRead\":").is_none());
}

#[test]
fn typed_inputs_reject_arbitrary_payloads_and_invalid_intervals() {
    assert!(serde_json::from_value::<DraftInput>(json!({"toRecipients":[]})).is_err());
    assert!(serde_json::from_value::<EventInput>(json!({"attendees":[]})).is_err());
    let empty: DraftInput = serde_json::from_value(json!({})).unwrap();
    assert!(create_draft_request(&empty).is_ok());
    assert!(update_draft_request("id", "key", &empty).is_err());
    for (start, end) in [
        ("not-a-date", "2026-09-20T13:00:00Z"),
        ("2026-02-30T12:00:00Z", "2026-03-01T13:00:00Z"),
        ("2026-09-20T13:00:00Z", "2026-09-20T12:00:00Z"),
    ] {
        let input: EventInput =
            serde_json::from_value(json!({"subject":"synthetic","start":start,"end":end})).unwrap();
        assert!(create_event_request("folder", &input).is_err());
    }
}

#[test]
fn calendar_times_preserve_fractional_seconds_when_normalized() {
    let input: EventInput = serde_json::from_value(json!({
        "subject":"synthetic", "start":"2026-09-20T14:00:00.125+02:00",
        "end":"2026-09-20T14:00:00.875+02:00"
    }))
    .unwrap();
    let request = create_event_request("synthetic-folder", &input).unwrap();
    assert_eq!(
        request["Body"]["Items"][0]["Start"],
        "2026-09-20T12:00:00.125Z"
    );
    assert_eq!(
        request["Body"]["Items"][0]["End"],
        "2026-09-20T12:00:00.875Z"
    );
}

#[test]
fn read_escape_hatch_stays_read_only() {
    for action in [
        "CreateItem",
        "UpdateItem",
        "DeleteItem",
        "CreateCalendarEvent",
        "UpdateCalendarEvent",
        "DeleteCalendarEvent",
        "SendItem",
    ] {
        assert!(validate_read_action(action).is_err());
    }
}
