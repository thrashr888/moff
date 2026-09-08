use std::collections::HashSet;

use chrono::{DateTime, FixedOffset, SecondsFormat};
use serde::Deserialize;
use serde_json::{Value, json};

use crate::{
    OutlookWebError, OutlookWebSession, calendar_view_request, find_item_request, get_item_request,
    header,
};

#[derive(Default, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct DraftInput {
    pub subject: Option<String>,
    pub body: Option<String>,
}

#[derive(Default, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct EventInput {
    pub subject: Option<String>,
    pub body: Option<String>,
    pub start: Option<String>,
    pub end: Option<String>,
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
        "ConflictResolution":"AlwaysOverwrite","SuppressReadReceipts":true,
        "ItemShape":{"__type":"ItemResponseShape:#Exchange","BaseShape":"IdOnly"},
        "ItemChanges":[{"__type":"ItemChange:#Exchange","ItemId":{"__type":"ItemId:#Exchange","Id":id,"ChangeKey":key},"Updates":updates}]}),
    ))
}
pub fn delete_message_request(id: &str, key: &str) -> Result<Value, OutlookWebError> {
    nonempty(id)?;
    nonempty(key)?;
    Ok(envelope(
        "DeleteItem",
        json!({"__type":"DeleteItemRequest:#Exchange","DeleteType":"HardDelete",
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
            "Start":utc(start)?,"End":utc(end)?,"StartTimeZoneId":"UTC","EndTimeZoneId":"UTC",
            "ParentFolderId":{"Id":folder},"IsAllDayEvent":false,"IsDraft":false,"IsOrganizer":true,
            "ReminderIsSet":false,"Resources":[]}]}),
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
pub fn delete_event_request(id: &str) -> Result<Value, OutlookWebError> {
    nonempty(id)?;
    Ok(envelope(
        "DeleteCalendarEvent",
        json!({"__type":"DeleteCalendarEventRequest:#Exchange","EventId":item_id(id),"EventScope":0}),
    ))
}

enum Mutation {
    CreateDraft,
    UpdateDraft,
    DeleteMessage,
    CreateEvent,
    UpdateEvent,
    DeleteEvent,
}
impl Mutation {
    fn action(&self) -> &'static str {
        match self {
            Self::CreateDraft => "CreateItem",
            Self::UpdateDraft => "UpdateItem",
            Self::DeleteMessage => "DeleteItem",
            Self::CreateEvent => "CreateCalendarEvent",
            Self::UpdateEvent => "UpdateCalendarEvent",
            Self::DeleteEvent => "DeleteCalendarEvent",
        }
    }
}

fn single_item(value: &Value) -> Result<&Value, OutlookWebError> {
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
fn type_is(item: &Value, name: &str) -> Result<(), OutlookWebError> {
    if item.get("__type").and_then(Value::as_str) == Some(name) {
        Ok(())
    } else {
        Err(OutlookWebError::UnsupportedItem)
    }
}
fn field<'a>(item: &'a Value, pointer: &str) -> Result<&'a str, OutlookWebError> {
    item.pointer(pointer)
        .and_then(Value::as_str)
        .filter(|s| !s.is_empty())
        .ok_or(OutlookWebError::InvalidResponse)
}
fn appointment(item: &Value) -> Result<(), OutlookWebError> {
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
        ],
    )
}

impl OutlookWebSession {
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
        loop {
            let page = self.mail_list(folder, 100, offset)?;
            let root = page
                .pointer("/Body/ResponseMessages/Items/0/RootFolder")
                .ok_or(OutlookWebError::InvalidResponse)?;
            let items = root
                .get("Items")
                .and_then(Value::as_array)
                .ok_or(OutlookWebError::InvalidResponse)?;
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
                let response = self.call("GetItem", &search_item_request(id))?;
                let full = single_item(&response)?;
                if contains_text(full, &query)? {
                    matches.push(full.clone());
                }
                if matches.len() > usize::from(limit) {
                    return Ok(search_result(matches, limit));
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
        self.execute(
            Mutation::CreateDraft.action(),
            &create_draft_request(input)?,
            true,
        )
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
        self.execute(
            Mutation::CreateEvent.action(),
            &create_event_request(folder, input)?,
            true,
        )
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
        self.execute(Mutation::UpdateEvent.action(), &body, true)
    }
    pub fn delete_event(&self, id: &str) -> Result<Value, OutlookWebError> {
        let response = self.calendar_get(id)?;
        appointment(single_item(&response)?)?;
        self.execute(
            Mutation::DeleteEvent.action(),
            &delete_event_request(id)?,
            true,
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

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
}
