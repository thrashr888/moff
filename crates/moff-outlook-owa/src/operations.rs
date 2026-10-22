use std::collections::HashSet;
use std::time::{Duration, Instant};

use chrono::{DateTime, FixedOffset, SecondsFormat};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

use crate::{
    OutlookWebError, OutlookWebSession, ServiceCode, calendar_view_request, find_item_request,
    get_attachment_request, get_item_request, header,
};

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct DraftInput {
    pub subject: Option<String>,
    pub body: Option<String>,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct EventInput {
    pub subject: Option<String>,
    pub body: Option<String>,
    pub start: Option<String>,
    pub end: Option<String>,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
pub struct ReplyInput {
    pub body: String,
    #[serde(default)]
    pub reply_all: bool,
}

#[derive(Clone, Copy, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum MeetingResponse {
    Accept,
    Tentative,
    Decline,
}

impl MeetingResponse {
    fn item_type(self) -> &'static str {
        match self {
            Self::Accept => "AcceptItem:#Exchange",
            Self::Tentative => "TentativelyAcceptItem:#Exchange",
            Self::Decline => "DeclineItem:#Exchange",
        }
    }

    fn response_type(self) -> &'static str {
        match self {
            Self::Accept => "Accept",
            Self::Tentative => "Tentative",
            Self::Decline => "Decline",
        }
    }
}

#[derive(Clone, Debug, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct MeetingResponseInput {
    pub response: MeetingResponse,
    pub message: Option<String>,
    pub proposed_start: Option<String>,
    pub proposed_end: Option<String>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct EventMoveInput {
    pub start: String,
    pub end: String,
}

pub fn respond_event_request(
    item: &Value,
    input: &MeetingResponseInput,
) -> Result<Value, OutlookWebError> {
    type_is(item, "CalendarItem:#Exchange")?;
    if item.get("IsMeeting").and_then(Value::as_bool) != Some(true)
        || item.get("IsOrganizer").and_then(Value::as_bool) != Some(false)
        || !matches!(
            item.get("CalendarItemType").and_then(Value::as_str),
            Some("Single" | "Occurrence" | "Exception")
        )
    {
        return Err(OutlookWebError::UnsupportedItem);
    }
    let item_id = item.get("ItemId").ok_or(OutlookWebError::InvalidResponse)?;
    nonempty(field(item_id, "/Id")?)?;
    nonempty(field(item_id, "/ChangeKey")?)?;
    if input
        .message
        .as_ref()
        .is_some_and(|message| message.trim().is_empty())
    {
        return Err(OutlookWebError::InvalidInput(
            "meeting response message must not be empty",
        ));
    }
    let proposed_start = input.proposed_start.as_deref();
    let proposed_end = input.proposed_end.as_deref();
    if proposed_start.is_some() != proposed_end.is_some() {
        return Err(OutlookWebError::InvalidInput(
            "proposed_start and proposed_end must be provided together",
        ));
    }
    if let (Some(start), Some(end)) = (proposed_start, proposed_end) {
        if input.response == MeetingResponse::Accept {
            return Err(OutlookWebError::InvalidInput(
                "a proposed new time is only supported for tentative or decline responses",
            ));
        }
        interval(start, end)?;
    }

    let mut response = json!({
        "__type": input.response.item_type(),
        "ReferenceItemId": item_id,
    });
    if let Some(message) = &input.message {
        response["Body"] = json!({
            "__type":"BodyContentType:#Exchange",
            "BodyType":"Text",
            "Value":message,
        });
    }
    if let (Some(start), Some(end)) = (proposed_start, proposed_end) {
        response["ProposedStart"] = json!(utc(start)?);
        response["ProposedEnd"] = json!(utc(end)?);
    }
    Ok(envelope(
        "CreateItem",
        json!({
            "__type":"CreateItemRequest:#Exchange",
            "MessageDisposition":"SendAndSaveCopy",
            "Items":[response]
        }),
    ))
}

fn meeting_organizer(item: &Value) -> Result<(), OutlookWebError> {
    type_is(item, "CalendarItem:#Exchange")?;
    if item.get("IsMeeting").and_then(Value::as_bool) != Some(true)
        || item.get("IsOrganizer").and_then(Value::as_bool) != Some(true)
        || !matches!(
            item.get("CalendarItemType").and_then(Value::as_str),
            Some("Single" | "Occurrence" | "Exception")
        )
    {
        return Err(OutlookWebError::UnsupportedItem);
    }
    Ok(())
}

pub fn move_event_request(item: &Value, input: &EventMoveInput) -> Result<Value, OutlookWebError> {
    meeting_organizer(item)?;
    interval(&input.start, &input.end)?;
    let mut request = update_event_request(
        field(item, "/ItemId/Id")?,
        &EventInput {
            start: Some(input.start.clone()),
            end: Some(input.end.clone()),
            ..EventInput::default()
        },
    )?;
    request["Body"]["ShouldSendUpdateToAttendees"] = json!(true);
    request["Body"]["EventId"] = item["ItemId"].clone();
    request["Body"]["ItemChange"]["ItemId"] = item["ItemId"].clone();
    Ok(request)
}

pub fn batch_item_request(ids: &[String], calendar: bool) -> Result<Value, OutlookWebError> {
    if ids.is_empty() || ids.len() > 20 || ids.iter().collect::<HashSet<_>>().len() != ids.len() {
        return Err(OutlookWebError::InvalidInput(
            "provide 1-20 unique item IDs",
        ));
    }
    for id in ids {
        nonempty(id)?;
    }
    let mut request = if calendar {
        calendar_item_request(&ids[0])
    } else {
        get_item_request(&ids[0])
    };
    request["Body"]["ItemIds"] = json!(ids.iter().map(|id| item_id(id)).collect::<Vec<_>>());
    Ok(request)
}

pub fn reply_draft_request(
    id: &str,
    key: &str,
    input: &ReplyInput,
) -> Result<Value, OutlookWebError> {
    nonempty(id)?;
    nonempty(key)?;
    nonempty(&input.body)?;
    Ok(envelope(
        "CreateItem",
        json!({
            "__type":"CreateItemRequest:#Exchange","MessageDisposition":"SaveOnly",
            "ItemShape":{"__type":"ItemResponseShape:#Exchange","BaseShape":"IdOnly"},
            "Items":[{"__type":if input.reply_all { "ReplyAllToItem:#Exchange" } else { "ReplyToItem:#Exchange" },
                "ReferenceItemId":{"__type":"ItemId:#Exchange","Id":id,"ChangeKey":key},
                "NewBodyContent":text_body(&input.body)}]
        }),
    ))
}

pub fn mark_read_request(id: &str, key: &str, read: bool) -> Result<Value, OutlookWebError> {
    nonempty(id)?;
    nonempty(key)?;
    Ok(envelope(
        "UpdateItem",
        json!({
            "__type":"UpdateItemRequest:#Exchange","MessageDisposition":"SaveOnly",
            "ConflictResolution":"NeverOverwrite","SuppressReadReceipts":true,
            "ItemShape":{"__type":"ItemResponseShape:#Exchange","BaseShape":"IdOnly"},
            "ItemChanges":[{"__type":"ItemChange:#Exchange",
                "ItemId":{"__type":"ItemId:#Exchange","Id":id,"ChangeKey":key},
                "Updates":[change("IsRead",json!(read),"Message:#Exchange")]}]
        }),
    ))
}

pub fn categories_request(
    id: &str,
    key: &str,
    categories: &[String],
) -> Result<Value, OutlookWebError> {
    nonempty(id)?;
    nonempty(key)?;
    for category in categories {
        nonempty(category)?;
    }
    Ok(envelope(
        "UpdateItem",
        json!({
            "__type":"UpdateItemRequest:#Exchange","MessageDisposition":"SaveOnly",
            "ConflictResolution":"NeverOverwrite","SuppressReadReceipts":true,
            "ItemShape":{"__type":"ItemResponseShape:#Exchange","BaseShape":"IdOnly"},
            "ItemChanges":[{"__type":"ItemChange:#Exchange",
                "ItemId":{"__type":"ItemId:#Exchange","Id":id,"ChangeKey":key},
                "Updates":[change("Categories",json!(categories),"Message:#Exchange")]}]
        }),
    ))
}

fn categories(item: &Value) -> Result<Vec<String>, OutlookWebError> {
    match item.get("Categories") {
        None => Ok(Vec::new()),
        Some(value) => {
            serde_json::from_value(value.clone()).map_err(|_| OutlookWebError::InvalidResponse)
        }
    }
}

fn nonempty(value: &str) -> Result<(), OutlookWebError> {
    if value.trim().is_empty() {
        Err(OutlookWebError::InvalidInput(
            "identifiers and search queries must not be empty",
        ))
    } else {
        Ok(())
    }
}
fn limit_ok(limit: u16) -> Result<(), OutlookWebError> {
    if (1..=100).contains(&limit) {
        Ok(())
    } else {
        Err(OutlookWebError::InvalidInput(
            "limit must be between 1 and 100",
        ))
    }
}
fn folder_ok(folder: &str) -> Result<(), OutlookWebError> {
    if [
        "inbox",
        "drafts",
        "sentitems",
        "deleteditems",
        "archive",
        "junkemail",
        "outbox",
    ]
    .contains(&folder)
    {
        Ok(())
    } else {
        Err(OutlookWebError::InvalidInput(
            "unsupported mail folder; use inbox, drafts, sentitems, deleteditems, archive, junkemail, or outbox",
        ))
    }
}
fn datetime(value: &str) -> Result<DateTime<FixedOffset>, OutlookWebError> {
    DateTime::parse_from_rfc3339(value).map_err(|_| {
        OutlookWebError::InvalidInput(
            "start and end must be valid RFC3339 timestamps with an offset",
        )
    })
}
fn interval(start: &str, end: &str) -> Result<(), OutlookWebError> {
    if datetime(end)? <= datetime(start)? {
        return Err(OutlookWebError::InvalidInput("end must be after start"));
    }
    Ok(())
}
fn utc(value: &str) -> Result<String, OutlookWebError> {
    Ok(datetime(value)?
        .to_utc()
        .to_rfc3339_opts(SecondsFormat::AutoSi, true))
}
fn text_body(value: &str) -> Value {
    let html = value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&#39;")
        .replace('\n', "<br>");
    json!({"BodyType":"HTML","Value":html})
}
fn envelope(kind: &str, body: Value) -> Value {
    let mut request = json!({"__type":format!("{kind}JsonRequest:#Exchange"),"Header":header()});
    request["Body"] = body;
    request
}
fn item_id(id: &str) -> Value {
    json!({"__type":"ItemId:#Exchange","Id":id})
}
fn change(field: &str, value: Value, item_type: &str) -> Value {
    let mut item = json!({"__type":item_type});
    item[field] = value;
    let mut update = json!({"__type":"SetItemField:#Exchange","Path":{"__type":"PropertyUri:#Exchange","FieldURI":field}});
    update["Item"] = item;
    update
}
fn draft_updates(input: &DraftInput) -> Result<Vec<Value>, OutlookWebError> {
    let mut updates = Vec::new();
    if let Some(subject) = &input.subject {
        updates.push(change("Subject", json!(subject), "Message:#Exchange"));
    }
    if let Some(body) = &input.body {
        updates.push(change("Body", text_body(body), "Message:#Exchange"));
    }
    if updates.is_empty() {
        return Err(OutlookWebError::InvalidInput(
            "draft update requires subject or body",
        ));
    }
    Ok(updates)
}
pub fn create_draft_request(input: &DraftInput) -> Result<Value, OutlookWebError> {
    Ok(envelope(
        "CreateItem",
        json!({"__type":"CreateItemRequest:#Exchange","MessageDisposition":"SaveOnly",
        "ItemShape":{"__type":"ItemResponseShape:#Exchange","BaseShape":"IdOnly"},
        "Items":[{"__type":"Message:#Exchange","Subject":input.subject.as_deref().unwrap_or(""),
            "Body":text_body(input.body.as_deref().unwrap_or("")),"ToRecipients":[],"CcRecipients":[],"BccRecipients":[],"MessageDisposition":"SaveOnly"}]}),
    ))
}
pub fn update_draft_request(
    id: &str,
    key: &str,
    input: &DraftInput,
) -> Result<Value, OutlookWebError> {
    nonempty(id)?;
    nonempty(key)?;
    let updates = draft_updates(input)?;
    Ok(envelope(
        "UpdateItem",
        json!({"__type":"UpdateItemRequest:#Exchange","MessageDisposition":"SaveOnly",
        "ConflictResolution":"NeverOverwrite","SuppressReadReceipts":true,
        "ItemShape":{"__type":"ItemResponseShape:#Exchange","BaseShape":"IdOnly"},
        "ItemChanges":[{"__type":"ItemChange:#Exchange","ItemId":{"__type":"ItemId:#Exchange","Id":id,"ChangeKey":key},"Updates":updates}]}),
    ))
}
pub fn delete_message_request(id: &str, key: &str) -> Result<Value, OutlookWebError> {
    nonempty(id)?;
    nonempty(key)?;
    Ok(envelope(
        "DeleteItem",
        json!({"__type":"DeleteItemRequest:#Exchange","DeleteType":"MoveToDeletedItems",
        "SendMeetingCancellations":"SendToNone",
        "AffectedTaskOccurrences":"AllOccurrences","ReturnMovedItemIds":true,"SuppressReadReceipts":true,
        "ItemIds":[{"__type":"ItemId:#Exchange","Id":id,"ChangeKey":key}]}),
    ))
}
pub fn create_event_request(folder: &str, input: &EventInput) -> Result<Value, OutlookWebError> {
    nonempty(folder)?;
    let subject = input
        .subject
        .as_deref()
        .filter(|s| !s.trim().is_empty())
        .ok_or(OutlookWebError::InvalidInput(
            "event creation requires a nonempty subject",
        ))?;
    let start = input.start.as_deref().ok_or(OutlookWebError::InvalidInput(
        "event creation requires start",
    ))?;
    let end = input
        .end
        .as_deref()
        .ok_or(OutlookWebError::InvalidInput("event creation requires end"))?;
    interval(start, end)?;
    Ok(envelope(
        "CreateItem",
        json!({"__type":"CreateItemRequest:#Exchange","SendMeetingInvitations":"SendToNone",
        "SavedItemFolderId":{"__type":"TargetFolderId:#Exchange","BaseFolderId":{"__type":"FolderId:#Exchange","Id":folder}},
        "ItemShape":{"__type":"ItemResponseShape:#Exchange","BaseShape":"IdOnly"},
        "Items":[{"__type":"CalendarItem:#Exchange","Subject":subject,"Body":text_body(input.body.as_deref().unwrap_or("")),
            "Start":utc(start)?,"End":utc(end)?,"IsAllDayEvent":false,
            "LegacyFreeBusyStatus":"Busy","ReminderIsSet":false}]}),
    ))
}
pub fn update_event_request(id: &str, input: &EventInput) -> Result<Value, OutlookWebError> {
    nonempty(id)?;
    let mut updates = Vec::new();
    if let Some(subject) = &input.subject {
        if subject.trim().is_empty() {
            return Err(OutlookWebError::InvalidInput(
                "event subject must not be empty",
            ));
        }
        updates.push(change("Subject", json!(subject), "CalendarItem:#Exchange"));
    }
    if let Some(body) = &input.body {
        updates.push(change("Body", text_body(body), "CalendarItem:#Exchange"));
    }
    if let Some(start) = &input.start {
        updates.push(change(
            "Start",
            json!(utc(start)?),
            "CalendarItem:#Exchange",
        ));
    }
    if let Some(end) = &input.end {
        updates.push(change("End", json!(utc(end)?), "CalendarItem:#Exchange"));
    }
    if let (Some(start), Some(end)) = (&input.start, &input.end) {
        interval(start, end)?;
    }
    if updates.is_empty() {
        return Err(OutlookWebError::InvalidInput(
            "event update requires at least one field",
        ));
    }
    Ok(envelope(
        "UpdateCalendarEvent",
        json!({"__type":"UpdateCalendarEventRequest:#Exchange",
        "EventId":item_id(id),"EventScope":0,"TargetAudience":0,"ShouldSendUpdateToAttendees":false,
        "ItemChange":{"__type":"ItemChange:#Exchange","ItemId":item_id(id),"Updates":updates}}),
    ))
}
pub fn delete_event_request(id: &str, key: &str) -> Result<Value, OutlookWebError> {
    delete_message_request(id, key)
}

pub fn cancel_meeting_request(item: &Value) -> Result<Value, OutlookWebError> {
    meeting_organizer(item)?;
    let item_id = item.get("ItemId").ok_or(OutlookWebError::InvalidResponse)?;
    nonempty(field(item_id, "/Id")?)?;
    nonempty(field(item_id, "/ChangeKey")?)?;
    Ok(envelope(
        "DeleteItem",
        json!({
            "__type":"DeleteItemRequest:#Exchange",
            "DeleteType":"MoveToDeletedItems",
            "SendMeetingCancellations":"SendToAllAndSaveCopy",
            "SuppressReadReceipts":true,
            "ItemIds":[item_id]
        }),
    ))
}

enum Mutation {
    CreateDraft,
    UpdateDraft,
    DeleteMessage,
    CreateEvent,
    UpdateEvent,
    DeleteEvent,
    RespondEvent,
    MoveEvent,
    CancelMeeting,
}
impl Mutation {
    fn action(&self) -> &'static str {
        match self {
            Self::CreateDraft | Self::CreateEvent | Self::RespondEvent => "CreateItem",
            Self::UpdateDraft => "UpdateItem",
            Self::DeleteMessage | Self::DeleteEvent | Self::CancelMeeting => "DeleteItem",
            Self::UpdateEvent | Self::MoveEvent => "UpdateCalendarEvent",
        }
    }
}

pub(crate) fn single_item(value: &Value) -> Result<&Value, OutlookWebError> {
    let responses = value
        .pointer("/Body/ResponseMessages/Items")
        .and_then(Value::as_array)
        .filter(|items| items.len() == 1)
        .ok_or(OutlookWebError::InvalidResponse)?;
    let items = responses[0]
        .get("Items")
        .and_then(Value::as_array)
        .filter(|items| items.len() == 1)
        .ok_or(OutlookWebError::InvalidResponse)?;
    Ok(&items[0])
}
pub(crate) fn is_mail_message(item: &Value) -> bool {
    matches!(
        item["__type"].as_str(),
        Some(
            "Message:#Exchange"
                | "MeetingMessage:#Exchange"
                | "MeetingRequestMessage:#Exchange"
                | "MeetingResponseMessage:#Exchange"
                | "MeetingCancellationMessage:#Exchange"
        )
    )
}
fn type_is(item: &Value, name: &str) -> Result<(), OutlookWebError> {
    if item.get("__type").and_then(Value::as_str) == Some(name) {
        Ok(())
    } else {
        Err(OutlookWebError::UnsupportedItem)
    }
}
pub(crate) fn field<'a>(item: &'a Value, pointer: &str) -> Result<&'a str, OutlookWebError> {
    item.pointer(pointer)
        .and_then(Value::as_str)
        .filter(|s| !s.is_empty())
        .ok_or(OutlookWebError::InvalidResponse)
}
pub(crate) fn appointment(item: &Value) -> Result<(), OutlookWebError> {
    type_is(item, "CalendarItem:#Exchange")?;
    if item.get("IsMeeting").and_then(Value::as_bool) != Some(false)
        || item.get("IsOrganizer").and_then(Value::as_bool) != Some(true)
        || item.get("CalendarItemType").and_then(Value::as_str) != Some("Single")
        || item.get("IsRecurring").and_then(Value::as_bool) != Some(false)
        || item.get("Recurrence").is_some_and(|v| !v.is_null())
        || ["RequiredAttendees", "OptionalAttendees", "Resources"]
            .iter()
            .any(|key| {
                item.get(key)
                    .is_some_and(|v| !v.is_null() && !v.as_array().is_some_and(Vec::is_empty))
            })
    {
        return Err(OutlookWebError::UnsupportedItem);
    }
    Ok(())
}
fn contains_text(item: &Value, query: &str) -> Result<bool, OutlookWebError> {
    if item
        .get("Subject")
        .and_then(Value::as_str)
        .is_some_and(|v| v.to_lowercase().contains(query))
    {
        return Ok(true);
    }
    if item.pointer("/Body/IsTruncated").and_then(Value::as_bool) == Some(true) {
        return Err(OutlookWebError::TruncatedSearch);
    }
    let body = item
        .pointer("/Body/Value")
        .and_then(Value::as_str)
        .ok_or(OutlookWebError::InvalidResponse)?;
    Ok(body.to_lowercase().contains(query))
}
fn search_result(mut items: Vec<Value>, limit: u16) -> Value {
    let has_more = items.len() > usize::from(limit);
    items.truncate(usize::from(limit));
    json!({"count":items.len(),"items":items,"has_more":has_more})
}

fn item_properties_request(id: &str, fields: &[&str]) -> Value {
    let mut request = get_item_request(id);
    request["Body"]["ItemShape"]["BaseShape"] = json!("IdOnly");
    request["Body"]["ItemShape"]["AdditionalProperties"] = Value::Array(
        fields
            .iter()
            .map(|field| json!({"__type":"PropertyUri:#Exchange","FieldURI":field}))
            .collect(),
    );
    request
}

pub fn search_item_request(id: &str) -> Value {
    item_properties_request(id, &["Subject", "Body"])
}

pub fn calendar_item_request(id: &str) -> Value {
    item_properties_request(
        id,
        &[
            "Subject",
            "Body",
            "Start",
            "End",
            "ParentFolderId",
            "IsMeeting",
            "IsOrganizer",
            "IsRecurring",
            "CalendarItemType",
            "RequiredAttendees",
            "OptionalAttendees",
            "Resources",
            "Location",
            "IsAllDayEvent",
            "LegacyFreeBusyStatus",
            "MyResponseType",
        ],
    )
}

impl OutlookWebSession {
    pub fn get_batch(&self, ids: &[String], calendar: bool) -> Result<Value, OutlookWebError> {
        let response = self.call("GetItem", &batch_item_request(ids, calendar)?)?;
        let messages = response
            .pointer("/Body/ResponseMessages/Items")
            .and_then(Value::as_array)
            .ok_or(OutlookWebError::InvalidResponse)?;
        let mut items = Vec::new();
        for message in messages {
            for item in message
                .get("Items")
                .and_then(Value::as_array)
                .ok_or(OutlookWebError::InvalidResponse)?
            {
                if calendar {
                    type_is(item, "CalendarItem:#Exchange")?;
                }
                items.push(item.clone());
            }
        }
        let returned: HashSet<_> = items
            .iter()
            .map(|item| field(item, "/ItemId/Id"))
            .collect::<Result<_, _>>()?;
        if items.len() != ids.len()
            || returned.len() != ids.len()
            || !ids.iter().all(|id| returned.contains(id.as_str()))
        {
            return Err(OutlookWebError::InvalidResponse);
        }
        Ok(json!({"items":items,"count":items.len(),"has_more":false}))
    }

    pub fn mail_categories(
        &self,
        id: &str,
        add: &[String],
        remove: &[String],
    ) -> Result<Value, OutlookWebError> {
        if add.is_empty() && remove.is_empty() {
            return Err(OutlookWebError::InvalidInput(
                "provide at least one category to add or remove",
            ));
        }
        for category in add.iter().chain(remove) {
            nonempty(category)?;
        }
        if add.iter().any(|c| remove.contains(c)) {
            return Err(OutlookWebError::InvalidInput(
                "cannot add and remove the same category",
            ));
        }
        let shape =
            |id| item_properties_request(id, &["Subject", "Categories", "IsRead", "IsDraft"]);
        let before_response = self.call("GetItem", &shape(id))?;
        let before = single_item(&before_response)?;
        type_is(before, "Message:#Exchange")?;
        let read = before
            .get("IsRead")
            .and_then(Value::as_bool)
            .ok_or(OutlookWebError::InvalidResponse)?;
        let mut updated = categories(before)?;
        updated.retain(|c| !remove.contains(c));
        for category in add {
            if !updated.contains(category) {
                updated.push(category.clone());
            }
        }
        if updated != categories(before)? {
            self.execute(
                "UpdateItem",
                &categories_request(id, field(before, "/ItemId/ChangeKey")?, &updated)?,
                true,
            )?;
        }
        let response = self
            .call("GetItem", &shape(id))
            .map_err(|_| OutlookWebError::VerificationFailed)?;
        let after = single_item(&response).map_err(|_| OutlookWebError::VerificationFailed)?;
        let actual = categories(after).map_err(|_| OutlookWebError::VerificationFailed)?;
        if actual.iter().collect::<HashSet<_>>() != updated.iter().collect::<HashSet<_>>()
            || after.get("IsRead").and_then(Value::as_bool) != Some(read)
        {
            return Err(OutlookWebError::VerificationFailed);
        }
        Ok(json!({"id":id,"categories":actual,"read_state_preserved":true,"verified":true}))
    }

    pub fn reply_draft(&self, id: &str, input: &ReplyInput) -> Result<Value, OutlookWebError> {
        nonempty(&input.body)?;
        let response = self.mail_get(id)?;
        let original = single_item(&response)?;
        type_is(original, "Message:#Exchange")?;
        let read = original
            .get("IsRead")
            .and_then(Value::as_bool)
            .ok_or(OutlookWebError::InvalidResponse)?;
        if original.get("IsDraft").and_then(Value::as_bool) != Some(false) {
            return Err(OutlookWebError::UnsupportedItem);
        }
        let created = self.execute(
            "CreateItem",
            &reply_draft_request(id, field(original, "/ItemId/ChangeKey")?, input)?,
            true,
        )?;
        let draft_id = field(
            single_item(&created).map_err(|_| OutlookWebError::OutcomeUnknown {
                action: "CreateItem",
            })?,
            "/ItemId/Id",
        )
        .map_err(|_| OutlookWebError::OutcomeUnknown {
            action: "CreateItem",
        })?;
        let verification_error = || OutlookWebError::DraftVerificationFailed {
            draft_id: draft_id.to_owned(),
            source_id: id.to_owned(),
        };
        let original_after = self.mail_get(id).map_err(|_| verification_error())?;
        let after = single_item(&original_after).map_err(|_| verification_error())?;
        let after_read = after.get("IsRead").and_then(Value::as_bool);
        let restored = !read && after_read == Some(true);
        if restored {
            // Outlook's native reply creation marks the source read, even with SaveOnly.
            self.mail_mark(id, false)
                .map_err(|_| verification_error())?;
        } else if after_read != Some(read) {
            return Err(verification_error());
        }
        let verified = self.mail_get(draft_id).map_err(|_| verification_error())?;
        if single_item(&verified)
            .map_err(|_| verification_error())?
            .get("IsDraft")
            .and_then(Value::as_bool)
            != Some(true)
        {
            return Err(verification_error());
        }
        Ok(
            json!({"source_id":id,"draft_id":draft_id,"sent":false,"verified":true,
            "reply_all":input.reply_all,"source_read_restored":restored}),
        )
    }
    pub fn mail_mark(&self, id: &str, read: bool) -> Result<Value, OutlookWebError> {
        let response = self.mail_get(id)?;
        let original = single_item(&response)?;
        type_is(original, "Message:#Exchange")?;
        let current = original
            .get("IsRead")
            .and_then(Value::as_bool)
            .ok_or(OutlookWebError::InvalidResponse)?;
        if current != read {
            self.execute(
                "UpdateItem",
                &mark_read_request(id, field(original, "/ItemId/ChangeKey")?, read)?,
                true,
            )?;
        }
        let after = self.mail_get(id)?;
        if single_item(&after)?.get("IsRead").and_then(Value::as_bool) != Some(read) {
            return Err(OutlookWebError::VerificationFailed);
        }
        Ok(json!({"id":id,"is_read":read,"verified":true,"changed":current != read}))
    }
    pub fn mail_list(
        &self,
        folder: &str,
        limit: u16,
        offset: u32,
    ) -> Result<Value, OutlookWebError> {
        folder_ok(folder)?;
        limit_ok(limit)?;
        let mut body = find_item_request(limit);
        body["Body"]["ParentFolderIds"][0]["Id"] = json!(folder);
        body["Body"]["Paging"]["Offset"] = json!(offset);
        self.call("FindItem", &body)
    }
    pub fn mail_get(&self, id: &str) -> Result<Value, OutlookWebError> {
        nonempty(id)?;
        self.call("GetItem", &get_item_request(id))
    }
    pub fn mail_attachment_get(&self, id: &str) -> Result<Value, OutlookWebError> {
        nonempty(id)?;
        self.call("GetAttachment", &get_attachment_request(id))
    }

    fn search_batch(
        &self,
        ids: &[String],
        deadline: Instant,
        scanned: usize,
    ) -> Result<Vec<Value>, OutlookWebError> {
        if ids.is_empty() {
            return Ok(Vec::new());
        }
        let mut request = batch_item_request(ids, false)?;
        request["Body"]["ItemShape"] = search_item_request(&ids[0])["Body"]["ItemShape"].clone();
        let response = self.search_call("GetItem", &request, deadline, scanned)?;
        let messages = response
            .pointer("/Body/ResponseMessages/Items")
            .and_then(Value::as_array)
            .ok_or(OutlookWebError::InvalidResponse)?;
        let mut found = std::collections::HashMap::new();
        for message in messages {
            for full in message["Items"]
                .as_array()
                .ok_or(OutlookWebError::InvalidResponse)?
            {
                let id = field(full, "/ItemId/Id")?;
                if !ids.iter().any(|requested| requested == id) || found.insert(id, full).is_some()
                {
                    return Err(OutlookWebError::InvalidResponse);
                }
            }
        }
        if found.len() != ids.len() {
            return Err(OutlookWebError::InvalidResponse);
        }
        Ok(ids.iter().map(|id| found[id.as_str()].clone()).collect())
    }

    /// Scan one bounded page. An empty page of matches is not an exhausted folder.
    pub fn mail_search_page(
        &self,
        folder: &str,
        query: &str,
        limit: u16,
        offset: u32,
    ) -> Result<Value, OutlookWebError> {
        folder_ok(folder)?;
        nonempty(query)?;
        limit_ok(limit)?;
        let deadline = Instant::now() + Duration::from_secs(40);
        let page_size = limit.min(20);
        let mut request = find_item_request(page_size);
        request["Body"]["ParentFolderIds"][0]["Id"] = json!(folder);
        request["Body"]["Paging"]["Offset"] = json!(offset);
        let page = self.search_call("FindItem", &request, deadline, 0)?;
        let root = page
            .pointer("/Body/ResponseMessages/Items/0/RootFolder")
            .ok_or(OutlookWebError::InvalidResponse)?;
        let items = root["Items"]
            .as_array()
            .ok_or(OutlookWebError::InvalidResponse)?;
        if items.len() > usize::from(page_size) {
            return Err(OutlookWebError::InvalidResponse);
        }
        let has_more = !root["IncludesLastItemInRange"]
            .as_bool()
            .ok_or(OutlookWebError::InvalidResponse)?;
        let next = if has_more {
            Some(
                root["IndexedPagingOffset"]
                    .as_u64()
                    .and_then(|n| u32::try_from(n).ok())
                    .filter(|n| *n > offset && !items.is_empty())
                    .ok_or(OutlookWebError::InvalidResponse)?,
            )
        } else {
            None
        };
        let mut ids = Vec::new();
        let mut seen = HashSet::new();
        for item in items {
            let id = field(item, "/ItemId/Id")?;
            if !seen.insert(id) {
                return Err(OutlookWebError::InvalidResponse);
            }
            if item["__type"]
                .as_str()
                .ok_or(OutlookWebError::InvalidResponse)?
                .ends_with("Message:#Exchange")
            {
                ids.push(id.to_owned());
            }
        }
        let query = query.to_lowercase();
        let mut matches = Vec::new();
        for full in self.search_batch(&ids, deadline, 0)? {
            if contains_text(&full, &query)? {
                let id = field(&full, "/ItemId/Id")?;
                let mut item = items
                    .iter()
                    .find(|item| item["ItemId"]["Id"] == id)
                    .ok_or(OutlookWebError::InvalidResponse)?
                    .clone();
                item.as_object_mut()
                    .ok_or(OutlookWebError::InvalidResponse)?
                    .extend(
                        full.as_object()
                            .ok_or(OutlookWebError::InvalidResponse)?
                            .clone(),
                    );
                matches.push(item);
            }
        }
        Ok(
            json!({"count":matches.len(),"items":matches,"offset":offset,"scanned":items.len(),
            "has_more":has_more,"next_offset":next}),
        )
    }

    pub fn mail_search(
        &self,
        folder: &str,
        query: &str,
        limit: u16,
    ) -> Result<Value, OutlookWebError> {
        folder_ok(folder)?;
        nonempty(query)?;
        limit_ok(limit)?;
        let query = query.to_lowercase();
        let mut matches = Vec::new();
        let mut offset = 0;
        let mut seen = HashSet::new();
        let deadline = Instant::now() + Duration::from_secs(40);
        let mut scanned = 0;
        loop {
            let mut request = find_item_request(100);
            request["Body"]["ParentFolderIds"][0]["Id"] = json!(folder);
            request["Body"]["Paging"]["Offset"] = json!(offset);
            let page = self.search_call("FindItem", &request, deadline, scanned)?;
            let root = page
                .pointer("/Body/ResponseMessages/Items/0/RootFolder")
                .ok_or(OutlookWebError::InvalidResponse)?;
            let items = root
                .get("Items")
                .and_then(Value::as_array)
                .ok_or(OutlookWebError::InvalidResponse)?;
            let mut ids = Vec::new();
            for item in items {
                let id = field(item, "/ItemId/Id")?;
                if !seen.insert(id.to_owned()) {
                    return Err(OutlookWebError::InvalidResponse);
                }
                if !item
                    .get("__type")
                    .and_then(Value::as_str)
                    .is_some_and(|kind| kind.ends_with("Message:#Exchange"))
                {
                    continue;
                }
                ids.push(id.to_owned());
            }
            for batch in ids.chunks(20) {
                for full in self.search_batch(batch, deadline, scanned)? {
                    scanned += 1;
                    if contains_text(&full, &query)? {
                        matches.push(full.clone());
                    }
                    if matches.len() > usize::from(limit) {
                        return Ok(search_result(matches, limit));
                    }
                }
            }
            match root.get("IncludesLastItemInRange").and_then(Value::as_bool) {
                Some(true) => return Ok(search_result(matches, limit)),
                Some(false) => {
                    let next = root
                        .get("IndexedPagingOffset")
                        .and_then(Value::as_u64)
                        .and_then(|v| u32::try_from(v).ok())
                        .filter(|v| *v > offset && !items.is_empty())
                        .ok_or(OutlookWebError::InvalidResponse)?;
                    offset = next;
                }
                None => return Err(OutlookWebError::InvalidResponse),
            }
        }
    }
    pub fn create_draft(&self, input: &DraftInput) -> Result<Value, OutlookWebError> {
        let result = self.execute(
            Mutation::CreateDraft.action(),
            &create_draft_request(input)?,
            true,
        )?;
        let id = field(
            single_item(&result).map_err(|_| OutlookWebError::OutcomeUnknown {
                action: "CreateItem",
            })?,
            "/ItemId/Id",
        )
        .map_err(|_| OutlookWebError::OutcomeUnknown {
            action: "CreateItem",
        })?;
        let fetched = self
            .mail_get(id)
            .map_err(|_| OutlookWebError::VerificationFailed)?;
        if single_item(&fetched)?
            .get("IsDraft")
            .and_then(Value::as_bool)
            != Some(true)
        {
            return Err(OutlookWebError::VerificationFailed);
        }
        Ok(result)
    }
    pub fn move_event(&self, id: &str, input: &EventMoveInput) -> Result<Value, OutlookWebError> {
        nonempty(id)?;
        let response = self.calendar_get(id)?;
        let item = single_item(&response)?;
        let request = move_event_request(item, input)?;
        let result = self.execute(Mutation::MoveEvent.action(), &request, true)?;
        let verified = self
            .calendar_get(id)
            .map_err(|_| OutlookWebError::VerificationFailed)?;
        let moved = single_item(&verified).map_err(|_| OutlookWebError::VerificationFailed)?;
        if datetime(&input.start).ok() != moved["Start"].as_str().and_then(|v| datetime(v).ok())
            || datetime(&input.end).ok() != moved["End"].as_str().and_then(|v| datetime(v).ok())
        {
            return Err(OutlookWebError::VerificationFailed);
        }
        Ok(result)
    }
    pub fn cancel_meeting(&self, id: &str) -> Result<Value, OutlookWebError> {
        nonempty(id)?;
        let response = self.calendar_get(id)?;
        let item = single_item(&response)?;
        let request = cancel_meeting_request(item)?;
        self.execute(Mutation::CancelMeeting.action(), &request, true)?;
        match self.calendar_get(id) {
            Err(OutlookWebError::OperationRejected {
                code: ServiceCode::ItemNotFound,
                ..
            }) => Ok(json!({"cancelled":true,"item_id":id,"attendee_notifications":"sent"})),
            Ok(_) => Err(OutlookWebError::VerificationFailed),
            Err(error) => Err(error),
        }
    }
    pub fn respond_event(
        &self,
        id: &str,
        input: &MeetingResponseInput,
    ) -> Result<Value, OutlookWebError> {
        nonempty(id)?;
        let response = self.calendar_get(id)?;
        let item = single_item(&response)?;
        let request = respond_event_request(item, input)?;
        let result = self.execute(Mutation::RespondEvent.action(), &request, true)?;
        let verified = self
            .calendar_get(id)
            .map_err(|_| OutlookWebError::VerificationFailed)?;
        let verified_item =
            single_item(&verified).map_err(|_| OutlookWebError::VerificationFailed)?;
        if verified_item.get("MyResponseType").and_then(Value::as_str)
            != Some(input.response.response_type())
        {
            return Err(OutlookWebError::VerificationFailed);
        }
        Ok(result)
    }
    pub fn update_draft(&self, id: &str, input: &DraftInput) -> Result<Value, OutlookWebError> {
        draft_updates(input)?;
        let response = self.mail_get(id)?;
        let item = single_item(&response)?;
        type_is(item, "Message:#Exchange")?;
        if item.get("IsDraft").and_then(Value::as_bool) != Some(true) {
            return Err(OutlookWebError::UnsupportedItem);
        }
        let body = update_draft_request(id, field(item, "/ItemId/ChangeKey")?, input)?;
        self.execute(Mutation::UpdateDraft.action(), &body, true)
    }
    pub fn delete_message(&self, id: &str) -> Result<Value, OutlookWebError> {
        let response = self.mail_get(id)?;
        type_is(single_item(&response)?, "Message:#Exchange")?;
        let body =
            delete_message_request(id, field(single_item(&response)?, "/ItemId/ChangeKey")?)?;
        self.execute(Mutation::DeleteMessage.action(), &body, true)
    }
    pub fn calendar_list(
        &self,
        folder_id: &str,
        start: &str,
        end: &str,
    ) -> Result<Value, OutlookWebError> {
        nonempty(folder_id)?;
        interval(start, end)?;
        let resolved;
        let folder_id = if folder_id == "calendar" {
            let response = self.call("GetFolder", &json!({
                "Header":header(),"Body":{
                    "__type":"GetFolderRequest:#Exchange",
                    "FolderShape":{"__type":"FolderResponseShape:#Exchange","BaseShape":"IdOnly"},
                    "FolderIds":[{"__type":"DistinguishedFolderId:#Exchange","Id":"calendar"}]
                }
            }))?;
            let responses = response
                .pointer("/Body/ResponseMessages/Items")
                .and_then(Value::as_array)
                .ok_or(OutlookWebError::InvalidResponse)?;
            if responses.len() != 1 {
                return Err(OutlookWebError::InvalidResponse);
            }
            let folders = responses[0]["Folders"]
                .as_array()
                .ok_or(OutlookWebError::InvalidResponse)?;
            if folders.len() != 1 {
                return Err(OutlookWebError::InvalidResponse);
            }
            resolved = field(&folders[0], "/FolderId/Id")?.to_owned();
            resolved.as_str()
        } else {
            folder_id
        };
        self.call(
            "GetCalendarView",
            &calendar_view_request(folder_id, start, end),
        )
    }
    pub fn calendar_get(&self, id: &str) -> Result<Value, OutlookWebError> {
        nonempty(id)?;
        let response = self.call("GetItem", &calendar_item_request(id))?;
        type_is(single_item(&response)?, "CalendarItem:#Exchange")?;
        Ok(response)
    }
    pub fn calendar_search(
        &self,
        folder: &str,
        start: &str,
        end: &str,
        query: &str,
        limit: u16,
    ) -> Result<Value, OutlookWebError> {
        nonempty(query)?;
        limit_ok(limit)?;
        let response = self.calendar_list(folder, start, end)?;
        let body = response
            .get("Body")
            .ok_or(OutlookWebError::InvalidResponse)?;
        if body.get("IncludesLastItemInRange").and_then(Value::as_bool) == Some(false)
            || body.get("IsTruncated").and_then(Value::as_bool) == Some(true)
        {
            return Err(OutlookWebError::TruncatedSearch);
        }
        let items = body
            .get("Items")
            .and_then(Value::as_array)
            .ok_or(OutlookWebError::InvalidResponse)?;
        let query = query.to_lowercase();
        let mut matches = Vec::new();
        for item in items {
            let response =
                self.call("GetItem", &search_item_request(field(item, "/ItemId/Id")?))?;
            let full = single_item(&response)?;
            if contains_text(full, &query)? {
                matches.push(full.clone());
            }
            if matches.len() > usize::from(limit) {
                break;
            }
        }
        Ok(search_result(matches, limit))
    }
    pub fn create_event(&self, folder: &str, input: &EventInput) -> Result<Value, OutlookWebError> {
        let result = self.execute(
            Mutation::CreateEvent.action(),
            &create_event_request(folder, input)?,
            true,
        )?;
        let id = field(
            single_item(&result).map_err(|_| OutlookWebError::OutcomeUnknown {
                action: "CreateItem",
            })?,
            "/ItemId/Id",
        )
        .map_err(|_| OutlookWebError::OutcomeUnknown {
            action: "CreateItem",
        })?;
        self.verify_event(id, input)?;
        Ok(result)
    }
    pub fn update_event(&self, id: &str, input: &EventInput) -> Result<Value, OutlookWebError> {
        let body = update_event_request(id, input)?;
        let response = self.calendar_get(id)?;
        let item = single_item(&response)?;
        appointment(item)?;
        interval(
            input.start.as_deref().unwrap_or(field(item, "/Start")?),
            input.end.as_deref().unwrap_or(field(item, "/End")?),
        )?;
        let mut body = body;
        body["Body"]["EventId"] = item["ItemId"].clone();
        body["Body"]["ItemChange"]["ItemId"] = item["ItemId"].clone();
        let result = self.execute(Mutation::UpdateEvent.action(), &body, true)?;
        let verified = self.verify_event(id, input)?;
        if input.body.is_none() && verified.pointer("/Body/Value") != item.pointer("/Body/Value") {
            return Err(OutlookWebError::VerificationFailed);
        }
        Ok(result)
    }
    fn verify_event(&self, id: &str, input: &EventInput) -> Result<Value, OutlookWebError> {
        let response = self
            .calendar_get(id)
            .map_err(|_| OutlookWebError::VerificationFailed)?;
        let item = single_item(&response).map_err(|_| OutlookWebError::VerificationFailed)?;
        appointment(item).map_err(|_| OutlookWebError::VerificationFailed)?;
        if input
            .subject
            .as_ref()
            .is_some_and(|subject| item["Subject"].as_str() != Some(subject))
            || input.start.as_deref().is_some_and(|start| {
                datetime(start).ok() != item["Start"].as_str().and_then(|v| datetime(v).ok())
            })
            || input.end.as_deref().is_some_and(|end| {
                datetime(end).ok() != item["End"].as_str().and_then(|v| datetime(v).ok())
            })
        {
            return Err(OutlookWebError::VerificationFailed);
        }
        Ok(item.clone())
    }
    pub fn delete_event(&self, id: &str) -> Result<Value, OutlookWebError> {
        let response = self.calendar_get(id)?;
        appointment(single_item(&response)?)?;
        self.execute(
            Mutation::DeleteEvent.action(),
            &delete_event_request(id, field(single_item(&response)?, "/ItemId/ChangeKey")?)?,
            true,
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_support::{message, mock};

    #[test]
    fn default_calendar_is_resolved_before_view_and_explicit_ids_are_preserved() {
        let range = ("2026-10-23T00:00:00-04:00", "2026-10-24T00:00:00-04:00");
        let folder = json!({"Body":{"ResponseMessages":{"Items":[{
            "ResponseCode":"NoError","Folders":[{"FolderId":{"Id":"real-calendar-id"}}]
        }]}}});
        let (session, requests, completion) = mock(vec![
            ("GetFolder", folder),
            (
                "GetCalendarView",
                json!({"Body":{"ResponseCode":"NoError","Items":[]}}),
            ),
        ]);
        session.calendar_list("calendar", range.0, range.1).unwrap();
        completion.join().unwrap();
        let requests = requests.lock().unwrap();
        assert_eq!(
            requests[0]["Body"]["FolderIds"][0]["__type"],
            "DistinguishedFolderId:#Exchange"
        );
        assert_eq!(requests[0]["Body"]["FolderIds"][0]["Id"], "calendar");
        assert_eq!(
            requests[1]["Body"]["CalendarId"]["BaseFolderId"]["Id"],
            "real-calendar-id"
        );

        let (session, requests, completion) = mock(vec![(
            "GetCalendarView",
            json!({"Body":{"ResponseCode":"NoError","Items":[]}}),
        )]);
        session
            .calendar_list("explicit-id", range.0, range.1)
            .unwrap();
        completion.join().unwrap();
        assert_eq!(
            requests.lock().unwrap()[0]["Body"]["CalendarId"]["BaseFolderId"]["Id"],
            "explicit-id"
        );
    }

    #[test]
    fn unresolved_default_calendar_fails_before_view_without_empty_result_fallback() {
        let (session, requests, completion) = mock(vec![(
            "GetFolder",
            json!({
                "Body":{"ResponseMessages":{"Items":[{"ResponseCode":"NoError","Folders":[]}]}}
            }),
        )]);
        assert!(matches!(
            session.calendar_list("calendar", "2026-10-23T00:00:00Z", "2026-10-24T00:00:00Z"),
            Err(OutlookWebError::InvalidResponse)
        ));
        completion.join().unwrap();
        assert_eq!(requests.lock().unwrap().len(), 1);
    }

    #[test]
    fn deletion_recycles_messages_and_rejects_meeting_cancellations() {
        let (session, requests, completion) = mock(vec![
            (
                "GetItem",
                message(
                    json!({"__type":"Message:#Exchange","ItemId":{"Id":"id","ChangeKey":"key"},"IsRead":false}),
                ),
            ),
            (
                "DeleteItem",
                json!({"Body":{"ResponseMessages":{"Items":[{"ResponseCode":"NoError"}]}}}),
            ),
        ]);
        session.delete_message("id").unwrap();
        completion.join().unwrap();
        let requests = requests.lock().unwrap();
        assert_eq!(requests[1]["Body"]["DeleteType"], "MoveToDeletedItems");
        assert_eq!(requests[1]["Body"]["ItemIds"][0]["ChangeKey"], "key");
        assert_eq!(requests[1]["Body"]["SuppressReadReceipts"], true);
        let (session, _, completion) = mock(vec![(
            "GetItem",
            message(
                json!({"__type":"CalendarItem:#Exchange","ItemId":{"Id":"id","ChangeKey":"key"},"IsMeeting":true}),
            ),
        )]);
        assert!(session.delete_event("id").is_err());
        completion.join().unwrap();
    }

    #[test]
    fn category_update_preserves_existing_categories_and_unread_state() {
        let before = json!({"__type":"Message:#Exchange","ItemId":{"Id":"id","ChangeKey":"key"},
            "Categories":["Existing"],"IsRead":false,"Subject":"Test"});
        let after = json!({"__type":"Message:#Exchange","ItemId":{"Id":"id","ChangeKey":"new"},
            "Categories":["Existing","Needs Response"],"IsRead":false,"Subject":"Test"});
        let (session, requests, worker) = mock(vec![
            ("GetItem", message(before)),
            ("UpdateItem", message(json!({"ItemId":{"Id":"id"}}))),
            ("GetItem", message(after)),
        ]);
        let result = session
            .mail_categories("id", &["Needs Response".into()], &[])
            .unwrap();
        assert_eq!(result["read_state_preserved"], true);
        worker.join().unwrap();
        let requests = requests.lock().unwrap();
        assert_eq!(
            requests[1]["Body"]["ItemChanges"][0]["Updates"][0]["Item"]["Categories"],
            json!(["Existing", "Needs Response"])
        );
        assert_eq!(requests[1]["Body"]["ConflictResolution"], "NeverOverwrite");
    }

    #[test]
    fn reply_draft_is_threaded_and_verified_without_marking_source_read() {
        let original = json!({"__type":"Message:#Exchange","ItemId":{"Id":"source","ChangeKey":"key"},
            "IsRead":false,"IsDraft":false});
        let (session, requests, worker) = mock(vec![
            ("GetItem", message(original.clone())),
            ("CreateItem", message(json!({"ItemId":{"Id":"draft"}}))),
            ("GetItem", message(original)),
            (
                "GetItem",
                message(json!({"ItemId":{"Id":"draft"},"IsDraft":true})),
            ),
        ]);
        assert_eq!(
            session
                .reply_draft(
                    "source",
                    &ReplyInput {
                        body: "Hello".into(),
                        reply_all: false,
                    }
                )
                .unwrap()["sent"],
            false
        );
        worker.join().unwrap();
        assert_eq!(
            requests.lock().unwrap()[1]["Body"]["Items"][0]["ReferenceItemId"]["Id"],
            "source"
        );
    }

    #[test]
    fn reply_creation_restores_unread_without_recreating_the_draft() {
        let unread = json!({"__type":"Message:#Exchange","ItemId":{"Id":"source","ChangeKey":"key"},
            "IsRead":false,"IsDraft":false});
        let mut read = unread.clone();
        read["IsRead"] = json!(true);
        let (session, requests, worker) = mock(vec![
            ("GetItem", message(unread.clone())),
            ("CreateItem", message(json!({"ItemId":{"Id":"draft"}}))),
            ("GetItem", message(read.clone())),
            ("GetItem", message(read)),
            ("UpdateItem", message(json!({"ItemId":{"Id":"source"}}))),
            ("GetItem", message(unread)),
            (
                "GetItem",
                message(json!({"ItemId":{"Id":"draft"},"IsDraft":true})),
            ),
        ]);
        let result = session
            .reply_draft(
                "source",
                &ReplyInput {
                    body: "Hello".into(),
                    reply_all: true,
                },
            )
            .unwrap();
        assert_eq!(result["source_read_restored"], true);
        assert_eq!(result["draft_id"], "draft");
        worker.join().unwrap();
        let requests = requests.lock().unwrap();
        assert_eq!(
            requests[1]["Body"]["Items"][0]["__type"],
            "ReplyAllToItem:#Exchange"
        );
        assert_eq!(requests[4]["Body"]["SuppressReadReceipts"], true);
        assert_eq!(
            requests[4]["Body"]["ItemChanges"][0]["Updates"][0]["Path"]["FieldURI"],
            "IsRead"
        );
        assert_eq!(
            requests[4]["Body"]["ItemChanges"][0]["Updates"][0]["Item"]["IsRead"],
            false
        );
    }

    #[test]
    fn failed_draft_verification_preserves_created_id() {
        let original = json!({"__type":"Message:#Exchange","ItemId":{"Id":"source","ChangeKey":"key"},
            "IsRead":true,"IsDraft":false});
        let (session, _, worker) = mock(vec![
            ("GetItem", message(original.clone())),
            ("CreateItem", message(json!({"ItemId":{"Id":"draft"}}))),
            ("GetItem", message(original)),
            (
                "GetItem",
                message(json!({"ItemId":{"Id":"draft"},"IsDraft":false})),
            ),
        ]);
        let error = session
            .reply_draft(
                "source",
                &ReplyInput {
                    body: "Hello".into(),
                    reply_all: false,
                },
            )
            .unwrap_err();
        assert!(matches!(
            error,
            OutlookWebError::DraftVerificationFailed { .. }
        ));
        assert!(error.to_string().contains("draft"));
        worker.join().unwrap();
    }

    #[test]
    fn category_noop_does_not_write_and_verification_failure_is_explicit() {
        let item = json!({"__type":"Message:#Exchange","ItemId":{"Id":"id","ChangeKey":"key"},
            "Categories":["Needs Response"],"IsRead":false});
        let mut changed = item.clone();
        changed["IsRead"] = json!(true);
        let (session, requests, worker) = mock(vec![
            ("GetItem", message(item)),
            ("GetItem", message(changed)),
        ]);
        assert!(matches!(
            session.mail_categories("id", &["Needs Response".into()], &[]),
            Err(OutlookWebError::VerificationFailed)
        ));
        worker.join().unwrap();
        assert_eq!(requests.lock().unwrap().len(), 2);
    }

    #[test]
    fn appointments_with_attendees_or_recurrence_cannot_be_mutated() {
        let base = json!({"__type":"CalendarItem:#Exchange","IsMeeting":false,"IsOrganizer":true,
            "IsRecurring":false,"CalendarItemType":"Single"});
        assert!(appointment(&base).is_ok());
        for (key, value) in [
            ("IsMeeting", json!(true)),
            ("IsOrganizer", json!(false)),
            ("IsRecurring", json!(true)),
            ("CalendarItemType", json!("Occurrence")),
            ("RequiredAttendees", json!([{}])),
            ("OptionalAttendees", json!([{}])),
            ("Resources", json!([{}])),
            ("Recurrence", json!({})),
        ] {
            let mut item = base.clone();
            item[key] = value;
            assert!(appointment(&item).is_err());
        }
        assert!(appointment(&json!({"__type":"CalendarItem:#Exchange"})).is_err());
    }

    #[test]
    fn body_search_is_literal_and_does_not_silently_accept_truncation() {
        assert!(
            contains_text(
                &json!({"Subject":"SYNTHETIC","Body":{"Value":""}}),
                "synthetic"
            )
            .unwrap()
        );
        assert!(
            contains_text(
                &json!({"Subject":"","Body":{"Value":"MiXeD text"}}),
                "mixed"
            )
            .unwrap()
        );
        assert!(
            contains_text(
                &json!({"Subject":"","Body":{"Value":"","IsTruncated":true}}),
                "absent"
            )
            .is_err()
        );
        assert!(contains_text(&json!({"Subject":""}), "absent").is_err());
    }

    fn search_page(ids: &[String], last: bool) -> Value {
        json!({"Body":{"ResponseMessages":{"Items":[{"ResponseCode":"NoError","RootFolder":{
            "Items":ids.iter().map(|id| json!({"__type":"Message:#Exchange","ItemId":{"Id":id}})).collect::<Vec<_>>(),
            "IncludesLastItemInRange":last,"IndexedPagingOffset":ids.len()
        }}]}}})
    }

    fn search_items(ids: &[String]) -> Value {
        json!({"Body":{"ResponseMessages":{"Items":[{"ResponseCode":"NoError",
            "Items":ids.iter().rev().map(|id| json!({"ItemId":{"Id":id},"Subject":id,
                "Body":{"Value":"match in body"}})).collect::<Vec<_>>()
        }]}}})
    }

    #[test]
    fn mail_search_batches_twenty_and_preserves_folder_order() {
        let ids: Vec<_> = (0..21).map(|i| format!("message-{i}")).collect();
        let (session, requests, worker) = mock(vec![
            ("FindItem", search_page(&ids, true)),
            ("GetItem", search_items(&ids[..20])),
            ("GetItem", search_items(&ids[20..])),
        ]);
        let result = session.mail_search("inbox", "MATCH", 25).unwrap();
        worker.join().unwrap();
        assert_eq!(result["count"], 21);
        assert_eq!(result["has_more"], false);
        assert_eq!(result["items"][0]["Subject"], "message-0");
        let requests = requests.lock().unwrap();
        assert_eq!(requests.len(), 3);
        assert_eq!(requests[1]["Body"]["ItemIds"].as_array().unwrap().len(), 20);
        assert_eq!(requests[2]["Body"]["ItemIds"].as_array().unwrap().len(), 1);
        assert_eq!(requests[1]["Body"]["ItemShape"]["BaseShape"], "IdOnly");
    }

    #[test]
    fn mail_search_rejects_missing_duplicate_and_unrequested_batch_items() {
        let ids = vec!["a".into(), "b".into()];
        for returned in [
            vec!["a".into()],
            vec!["a".into(), "a".into()],
            vec!["a".into(), "other".into()],
        ] {
            let (session, _, worker) = mock(vec![
                ("FindItem", search_page(&ids, true)),
                ("GetItem", search_items(&returned)),
            ]);
            assert!(matches!(
                session.mail_search("inbox", "match", 25),
                Err(OutlookWebError::InvalidResponse)
            ));
            worker.join().unwrap();
        }
    }

    #[test]
    fn mail_search_paginates_and_reports_more_matches_without_claiming_completeness() {
        let ids = vec!["a".into()];
        let next = vec!["b".into(), "c".into()];
        let (session, requests, worker) = mock(vec![
            ("FindItem", search_page(&ids, false)),
            ("GetItem", search_items(&ids)),
            ("FindItem", search_page(&next, true)),
            ("GetItem", search_items(&next)),
        ]);
        let result = session.mail_search("inbox", "match", 1).unwrap();
        worker.join().unwrap();
        assert_eq!(result["count"], 1);
        assert_eq!(result["has_more"], true);
        assert_eq!(requests.lock().unwrap()[2]["Body"]["Paging"]["Offset"], 1);
    }

    #[test]
    fn exhausted_search_budget_is_an_error_not_an_empty_result() {
        let (session, requests, worker) = mock(vec![]);
        let error = session
            .search_call("FindItem", &json!({}), Instant::now(), 120)
            .unwrap_err();
        assert!(matches!(
            error,
            OutlookWebError::SearchTimedOut { scanned: 120 }
        ));
        assert!(error.to_string().contains("incomplete, not empty"));
        assert!(requests.lock().unwrap().is_empty());
        worker.join().unwrap();
    }

    #[test]
    fn search_page_without_matches_returns_continuation_not_a_whole_folder_negative() {
        let ids = vec!["a".into(), "b".into()];
        let (session, requests, worker) = mock(vec![
            ("FindItem", search_page(&ids, false)),
            ("GetItem", search_items(&ids)),
        ]);
        let result = session.mail_search_page("inbox", "absent", 100, 0).unwrap();
        worker.join().unwrap();
        assert_eq!(result["count"], 0);
        assert_eq!(result["scanned"], 2);
        assert_eq!(result["has_more"], true);
        assert_eq!(result["next_offset"], 2);
        assert_eq!(requests.lock().unwrap().len(), 2);
        assert_eq!(
            requests.lock().unwrap()[0]["Body"]["Paging"]["MaxEntriesReturned"],
            20
        );
    }

    #[test]
    fn search_continuation_reaches_end_and_keeps_sender_metadata() {
        let ids = vec!["a".into()];
        let mut page = search_page(&ids, true);
        page["Body"]["ResponseMessages"]["Items"][0]["RootFolder"]["Items"][0]["From"] =
            json!({"Mailbox":{"Name":"Synthetic sender"}});
        let (session, requests, worker) =
            mock(vec![("FindItem", page), ("GetItem", search_items(&ids))]);
        let result = session.mail_search_page("inbox", "match", 10, 20).unwrap();
        worker.join().unwrap();
        assert_eq!(result["offset"], 20);
        assert_eq!(result["has_more"], false);
        assert!(result["next_offset"].is_null());
        assert_eq!(
            result["items"][0]["From"]["Mailbox"]["Name"],
            "Synthetic sender"
        );
        assert_eq!(requests.lock().unwrap()[0]["Body"]["Paging"]["Offset"], 20);
    }

    #[test]
    fn search_pages_reject_invalid_paging_metadata_before_reading_bodies() {
        let ids = vec!["a".into()];
        for (key, value) in [
            ("IndexedPagingOffset", json!(0)),
            ("IncludesLastItemInRange", Value::Null),
            ("Items", json!([])),
        ] {
            let mut page = search_page(&ids, false);
            page["Body"]["ResponseMessages"]["Items"][0]["RootFolder"][key] = value;
            let (session, _, worker) = mock(vec![("FindItem", page)]);
            assert!(matches!(
                session.mail_search_page("inbox", "match", 20, 0),
                Err(OutlookWebError::InvalidResponse)
            ));
            worker.join().unwrap();
        }
    }
}
