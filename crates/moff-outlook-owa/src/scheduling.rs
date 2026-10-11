use chrono::{DateTime, Duration, FixedOffset};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

use crate::operations::{appointment, field, single_item};
use crate::{EventInput, OutlookWebError, OutlookWebSession};

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct WorkWindow {
    pub start: String,
    pub end: String,
}

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct WorkTask {
    pub subject: String,
    pub body: Option<String>,
    pub minutes: u16,
    pub deadline: Option<String>,
    /// An existing personal appointment to move, rather than creating a duplicate.
    pub id: Option<String>,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
pub struct WorkInput {
    pub folder_id: String,
    pub windows: Vec<WorkWindow>,
    pub tasks: Vec<WorkTask>,
    #[serde(default)]
    pub allow_overlap_titles: Vec<String>,
}

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct PlannedBlock {
    pub id: Option<String>,
    pub expected: Option<Value>,
    pub event: EventInput,
    pub overlaps: Vec<String>,
}

#[derive(Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct WorkPlan {
    pub version: u8,
    pub folder_id: String,
    pub windows: Vec<WorkWindow>,
    pub allow_overlap_titles: Vec<String>,
    pub snapshot: Vec<Value>,
    pub blocks: Vec<PlannedBlock>,
}

#[derive(Clone, Copy)]
struct Span {
    start: DateTime<FixedOffset>,
    end: DateTime<FixedOffset>,
}

fn timestamp(value: &str) -> Result<DateTime<FixedOffset>, OutlookWebError> {
    DateTime::parse_from_rfc3339(value)
        .map_err(|_| OutlookWebError::InvalidInput("times require RFC3339 offsets"))
}
fn span(start: &str, end: &str) -> Result<Span, OutlookWebError> {
    let result = Span {
        start: timestamp(start)?,
        end: timestamp(end)?,
    };
    if result.end <= result.start {
        return Err(OutlookWebError::InvalidInput("end must be after start"));
    }
    Ok(result)
}
fn intersects(a: Span, b: Span) -> bool {
    a.start < b.end && b.start < a.end
}
fn is_free(item: &Value) -> bool {
    item["FreeBusyType"].as_str() == Some("Free")
        || item["LegacyFreeBusyStatus"].as_str() == Some("Free")
        || item["IsCancelled"].as_bool() == Some(true)
}
fn event_span(item: &Value) -> Result<Span, OutlookWebError> {
    span(field(item, "/Start")?, field(item, "/End")?)
}
fn windows(input: &[WorkWindow]) -> Result<Vec<Span>, OutlookWebError> {
    if input.is_empty() || input.len() > 14 {
        return Err(OutlookWebError::InvalidInput(
            "provide 1-14 non-overlapping work windows",
        ));
    }
    let mut spans = input
        .iter()
        .map(|w| span(&w.start, &w.end))
        .collect::<Result<Vec<_>, _>>()?;
    spans.sort_by_key(|w| w.start);
    if spans.iter().any(|w| w.end - w.start > Duration::hours(24))
        || spans.windows(2).any(|w| intersects(w[0], w[1]))
    {
        return Err(OutlookWebError::InvalidInput(
            "work windows must not overlap or exceed 24 hours",
        ));
    }
    Ok(spans)
}
fn gaps(window: Span, busy: &[Span]) -> Vec<Span> {
    let mut ranges: Vec<_> = busy
        .iter()
        .copied()
        .filter(|b| intersects(window, *b))
        .collect();
    ranges.sort_by_key(|b| b.start);
    let mut cursor = window.start;
    let mut result = Vec::new();
    for range in ranges {
        if range.start > cursor {
            result.push(Span {
                start: cursor,
                end: range.start.min(window.end),
            });
        }
        cursor = cursor.max(range.end);
    }
    if cursor < window.end {
        result.push(Span {
            start: cursor,
            end: window.end,
        });
    }
    for gap in &mut result {
        gap.start = gap.start.with_timezone(window.start.offset());
        gap.end = gap.end.with_timezone(window.start.offset());
    }
    result
}

fn busy(
    items: &[Value],
    ignored_ids: &[String],
    ignored_titles: &[String],
) -> Result<Vec<Span>, OutlookWebError> {
    items
        .iter()
        .filter(|item| {
            !item["ItemId"]["Id"]
                .as_str()
                .is_some_and(|id| ignored_ids.iter().any(|v| v == id))
                && !item["Subject"]
                    .as_str()
                    .is_some_and(|title| ignored_titles.iter().any(|v| v == title))
                && !is_free(item)
        })
        .map(event_span)
        .collect()
}
fn conflicts(
    items: &[Value],
    slot: Span,
    ignored_ids: &[String],
) -> Result<Vec<String>, OutlookWebError> {
    let mut result = Vec::new();
    for item in items {
        if item["ItemId"]["Id"]
            .as_str()
            .is_some_and(|id| ignored_ids.iter().any(|v| v == id))
            || is_free(item)
        {
            continue;
        }
        if intersects(slot, event_span(item)?) {
            result.push(field(item, "/Subject")?.to_owned());
        }
    }
    Ok(result)
}
fn minimum(minutes: u16) -> Result<Duration, OutlookWebError> {
    if !(1..=480).contains(&minutes) {
        return Err(OutlookWebError::InvalidInput(
            "minutes must be between 1 and 480",
        ));
    }
    Ok(Duration::minutes(i64::from(minutes)))
}
fn nonempty(value: &str) -> Result<(), OutlookWebError> {
    if value.trim().is_empty() {
        Err(OutlookWebError::InvalidInput(
            "subjects and IDs must not be empty",
        ))
    } else {
        Ok(())
    }
}

impl OutlookWebSession {
    fn in_calendar(&self, folder: &str, item: &Value) -> Result<bool, OutlookWebError> {
        let id = field(item, "/ItemId/Id")?;
        let response = self.calendar_list(folder, field(item, "/Start")?, field(item, "/End")?)?;
        let body = response
            .get("Body")
            .ok_or(OutlookWebError::InvalidResponse)?;
        if body["IncludesLastItemInRange"] == false || body["IsTruncated"] == true {
            return Err(OutlookWebError::TruncatedSearch);
        }
        Ok(body["Items"]
            .as_array()
            .ok_or(OutlookWebError::InvalidResponse)?
            .iter()
            .any(|candidate| candidate["ItemId"]["Id"].as_str() == Some(id)))
    }

    fn work_snapshot(
        &self,
        folder: &str,
        input: &[WorkWindow],
    ) -> Result<Vec<Value>, OutlookWebError> {
        windows(input)?;
        let mut items = Vec::new();
        for window in input {
            let response = self.calendar_list(folder, &window.start, &window.end)?;
            let body = response
                .get("Body")
                .ok_or(OutlookWebError::InvalidResponse)?;
            if body["IncludesLastItemInRange"] == false || body["IsTruncated"] == true {
                return Err(OutlookWebError::TruncatedSearch);
            }
            for item in body["Items"]
                .as_array()
                .ok_or(OutlookWebError::InvalidResponse)?
            {
                field(item, "/ItemId/Id")?;
                event_span(item)?;
                if !items
                    .iter()
                    .any(|old: &Value| old["ItemId"]["Id"] == item["ItemId"]["Id"])
                {
                    items.push(item.clone());
                }
            }
        }
        items.sort_by(|a, b| a["ItemId"]["Id"].as_str().cmp(&b["ItemId"]["Id"].as_str()));
        // Calendar view omits free/busy on this Outlook build. Read all-day metadata
        // in batches so free reminder banners do not consume the whole workday.
        let all_day_ids: Vec<String> = items
            .iter()
            .filter(|item| {
                item["IsAllDayEvent"].as_bool() == Some(true) && item.get("FreeBusyType").is_none()
            })
            .map(|item| field(item, "/ItemId/Id").map(str::to_owned))
            .collect::<Result<_, _>>()?;
        for ids in all_day_ids.chunks(20) {
            let response = self.get_batch(ids, false)?;
            for details in response["items"]
                .as_array()
                .ok_or(OutlookWebError::InvalidResponse)?
            {
                let item = items
                    .iter_mut()
                    .find(|item| item["ItemId"]["Id"] == details["ItemId"]["Id"])
                    .ok_or(OutlookWebError::InvalidResponse)?;
                if let Some(status) = details.get("FreeBusyType") {
                    item["FreeBusyType"] = status.clone();
                }
                if let Some(cancelled) = details.get("IsCancelled") {
                    item["IsCancelled"] = cancelled.clone();
                }
            }
        }
        Ok(items)
    }

    pub fn availability(
        &self,
        folder: &str,
        start: &str,
        end: &str,
        minutes: u16,
        overlap_titles: &[String],
    ) -> Result<Value, OutlookWebError> {
        let duration = minimum(minutes)?;
        let window = span(start, end)?;
        let items = self.work_snapshot(
            folder,
            &[WorkWindow {
                start: start.to_owned(),
                end: end.to_owned(),
            }],
        )?;
        let free = gaps(window, &busy(&items, &[], &[])?);
        let fallback = gaps(window, &busy(&items, &[], overlap_titles)?);
        let mut slots = Vec::new();
        for (ranges, overlap) in [(&free, false), (&fallback, true)] {
            for slot in ranges.iter().filter(|s| s.end - s.start >= duration) {
                if overlap && conflicts(&items, *slot, &[])?.is_empty() {
                    continue;
                }
                slots.push(
                    json!({"start":slot.start.to_rfc3339(),"end":slot.end.to_rfc3339(),
                    "minutes":(slot.end-slot.start).num_minutes(),"fallback":overlap,
                    "overlaps":conflicts(&items,*slot,&[])?}),
                );
            }
        }
        Ok(
            json!({"folder_id":folder,"scope":"selected calendar only","slots":slots,
            "policy":"empty slots first; exact-title overlaps only as explicit fallback"}),
        )
    }

    pub fn plan_work_blocks(&self, input: &WorkInput) -> Result<WorkPlan, OutlookWebError> {
        nonempty(&input.folder_id)?;
        let work_windows = windows(&input.windows)?;
        if input.tasks.is_empty() || input.tasks.len() > 20 {
            return Err(OutlookWebError::InvalidInput("provide 1-20 work tasks"));
        }
        for title in &input.allow_overlap_titles {
            nonempty(title)?;
        }
        let snapshot = self.work_snapshot(&input.folder_id, &input.windows)?;
        let ids: Vec<_> = input.tasks.iter().filter_map(|t| t.id.clone()).collect();
        if ids.iter().collect::<std::collections::HashSet<_>>().len() != ids.len() {
            return Err(OutlookWebError::InvalidInput(
                "each appointment may be moved only once",
            ));
        }
        let mut allocated = Vec::new();
        let mut blocks = Vec::new();
        for task in &input.tasks {
            nonempty(&task.subject)?;
            let duration = minimum(task.minutes)?;
            let deadline = task.deadline.as_deref().map(timestamp).transpose()?;
            let expected = if let Some(id) = &task.id {
                nonempty(id)?;
                let response = self.calendar_get(id)?;
                let item = single_item(&response)?;
                appointment(item)?;
                if !self.in_calendar(&input.folder_id, item)?
                    || item["Subject"].as_str() != Some(&task.subject)
                    || event_span(item)?.end - event_span(item)?.start != duration
                {
                    return Err(OutlookWebError::InvalidInput(
                        "moves must preserve subject, duration and calendar",
                    ));
                }
                if task.body.is_some() {
                    return Err(OutlookWebError::InvalidInput(
                        "moves preserve existing notes; omit body",
                    ));
                }
                Some(item.clone())
            } else {
                None
            };
            let mut chosen = None;
            for ignored in [&[][..], input.allow_overlap_titles.as_slice()] {
                let mut occupied = busy(&snapshot, &ids, ignored)?;
                occupied.extend(allocated.iter().copied());
                for window in &work_windows {
                    for gap in gaps(*window, &occupied) {
                        let end = gap.start + duration;
                        if end <= gap.end && deadline.is_none_or(|d| end <= d) {
                            chosen = Some(Span {
                                start: gap.start,
                                end,
                            });
                            break;
                        }
                    }
                    if chosen.is_some() {
                        break;
                    }
                }
                if chosen.is_some() {
                    break;
                }
            }
            let slot = chosen.ok_or(OutlookWebError::InvalidInput("a task does not fit before its deadline; expand windows, reduce duration or explicitly permit overlap"))?;
            allocated.push(slot);
            blocks.push(PlannedBlock {
                id: task.id.clone(),
                expected,
                overlaps: conflicts(&snapshot, slot, &ids)?,
                event: EventInput {
                    subject: Some(task.subject.clone()),
                    body: task.body.clone(),
                    start: Some(slot.start.to_rfc3339()),
                    end: Some(slot.end.to_rfc3339()),
                },
            });
        }
        Ok(WorkPlan {
            version: 1,
            folder_id: input.folder_id.clone(),
            windows: input.windows.clone(),
            allow_overlap_titles: input.allow_overlap_titles.clone(),
            snapshot,
            blocks,
        })
    }

    #[allow(clippy::too_many_lines)]
    pub fn apply_work_blocks(&self, plan: &WorkPlan) -> Result<Value, OutlookWebError> {
        let work_windows = windows(&plan.windows)?;
        if plan.version != 1 || plan.blocks.is_empty() || plan.blocks.len() > 20 {
            return Err(OutlookWebError::InvalidInput(
                "unsupported or empty work-block plan",
            ));
        }
        if self.work_snapshot(&plan.folder_id, &plan.windows)? != plan.snapshot {
            return Err(OutlookWebError::PlanChanged);
        }
        let ids: Vec<_> = plan.blocks.iter().filter_map(|b| b.id.clone()).collect();
        if ids.iter().collect::<std::collections::HashSet<_>>().len() != ids.len() {
            return Err(OutlookWebError::InvalidInput("duplicate move IDs"));
        }
        let mut slots = Vec::new();
        for block in &plan.blocks {
            let slot = span(
                block
                    .event
                    .start
                    .as_deref()
                    .ok_or(OutlookWebError::InvalidResponse)?,
                block
                    .event
                    .end
                    .as_deref()
                    .ok_or(OutlookWebError::InvalidResponse)?,
            )?;
            nonempty(
                block
                    .event
                    .subject
                    .as_deref()
                    .ok_or(OutlookWebError::InvalidResponse)?,
            )?;
            if !work_windows
                .iter()
                .any(|w| slot.start >= w.start && slot.end <= w.end)
                || slots.iter().any(|s| intersects(*s, slot))
                || conflicts(&plan.snapshot, slot, &ids)?
                    .iter()
                    .any(|title| !plan.allow_overlap_titles.contains(title))
            {
                return Err(OutlookWebError::InvalidInput(
                    "plan contains an out-of-window or conflicting block",
                ));
            }
            slots.push(slot);
            if let Some(id) = &block.id {
                let response = self.calendar_get(id)?;
                let item = single_item(&response)?;
                appointment(item)?;
                if block.expected.as_ref() != Some(item)
                    || !self.in_calendar(&plan.folder_id, item)?
                    || block.event.subject.as_deref() != item["Subject"].as_str()
                    || slot.end - slot.start != event_span(item)?.end - event_span(item)?.start
                    || block.event.body.is_some()
                {
                    return Err(OutlookWebError::PlanChanged);
                }
            } else if block.expected.is_some() {
                return Err(OutlookWebError::InvalidInput(
                    "new blocks cannot carry move preconditions",
                ));
            }
        }
        let mut results = Vec::new();
        let mut completed_ids = Vec::new();
        for block in &plan.blocks {
            // Recheck between writes; exclude only items successfully applied by this batch.
            let operation = (|| {
                let current = self.work_snapshot(&plan.folder_id, &plan.windows)?;
                let filter = |items: &[Value]| {
                    items
                        .iter()
                        .filter(|item| {
                            !item["ItemId"]["Id"]
                                .as_str()
                                .is_some_and(|id| completed_ids.iter().any(|v| v == id))
                        })
                        .cloned()
                        .collect::<Vec<_>>()
                };
                if filter(&current) != filter(&plan.snapshot) {
                    return Err(OutlookWebError::PlanChanged);
                }
                if let Some(id) = &block.id {
                    let response = self.calendar_get(id)?;
                    if block.expected.as_ref() != Some(single_item(&response)?) {
                        return Err(OutlookWebError::PlanChanged);
                    }
                    self.update_event(id, &block.event)?;
                    Ok(id.clone())
                } else {
                    let response = self.create_event(&plan.folder_id, &block.event)?;
                    Ok(field(single_item(&response)?, "/ItemId/Id")?.to_owned())
                }
            })();
            match operation {
                Ok(id) => {
                    completed_ids.push(id.clone());
                    results.push(json!({"id":id,"event":block.event,"verified":true}));
                }
                Err(error) => {
                    return Ok(
                        json!({"complete":false,"results":results,"failed_block":block,
                    "error":error.to_string(),"retry_automatically":false}),
                    );
                }
            }
        }
        Ok(json!({"complete":true,"results":results,"invitations_sent":false}))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_support::{message, mock};
    #[allow(clippy::needless_pass_by_value)]
    fn calendar(items: Vec<Value>) -> Value {
        json!({"Body":{"ResponseCode":"NoError","Items":items,"IsTruncated":false}})
    }
    fn input() -> WorkInput {
        serde_json::from_value(json!({"folder_id":"calendar",
            "windows":[{"start":"2026-10-12T09:00:00-07:00","end":"2026-10-12T17:00:00-07:00"}],
            "tasks":[{"subject":"Prepare","minutes":60}],
            "allow_overlap_titles":["Focus Time"]}))
        .unwrap()
    }

    #[test]
    fn free_all_day_reminders_are_not_treated_as_day_long_meetings() {
        let reminder = json!({"__type":"CalendarItem:#Exchange","ItemId":{"Id":"reminder"},
            "Subject":"Deadline","IsAllDayEvent":true,"Start":"2026-10-12T00:00:00Z","End":"2026-10-13T00:00:00Z"});
        let mut details = reminder.clone();
        details["FreeBusyType"] = json!("Free");
        let (session, _, worker) = mock(vec![
            ("GetCalendarView", calendar(vec![reminder])),
            ("GetItem", message(details)),
        ]);
        let slots = session
            .availability(
                "calendar",
                "2026-10-12T09:00:00-07:00",
                "2026-10-12T15:00:00-07:00",
                60,
                &[],
            )
            .unwrap();
        assert_eq!(slots["slots"][0]["minutes"], 360);
        worker.join().unwrap();
    }

    #[test]
    fn move_apply_preserves_duration_notes_and_change_key_without_inviting_anyone() {
        let old = json!({"__type":"CalendarItem:#Exchange","ItemId":{"Id":"move","ChangeKey":"old-key"},
            "Subject":"Prepare","Body":{"Value":"Keep these notes"},
            "Start":"2026-10-12T22:00:00Z","End":"2026-10-12T23:00:00Z",
            "IsMeeting":false,"IsOrganizer":true,"IsRecurring":false,"CalendarItemType":"Single"});
        let mut new = old.clone();
        new["Start"] = json!("2026-10-12T16:00:00Z");
        new["End"] = json!("2026-10-12T17:00:00Z");
        let snapshot = vec![old.clone()];
        let plan = WorkPlan {
            version: 1,
            folder_id: "calendar".into(),
            windows: input().windows,
            allow_overlap_titles: vec![],
            snapshot: snapshot.clone(),
            blocks: vec![PlannedBlock {
                id: Some("move".into()),
                expected: Some(old.clone()),
                overlaps: vec![],
                event: EventInput {
                    subject: Some("Prepare".into()),
                    body: None,
                    start: Some("2026-10-12T09:00:00-07:00".into()),
                    end: Some("2026-10-12T10:00:00-07:00".into()),
                },
            }],
        };
        let (session, requests, worker) = mock(vec![
            ("GetCalendarView", calendar(snapshot.clone())),
            ("GetItem", message(old.clone())),
            ("GetCalendarView", calendar(snapshot.clone())),
            ("GetCalendarView", calendar(snapshot)),
            ("GetItem", message(old.clone())),
            ("GetItem", message(old)),
            (
                "UpdateCalendarEvent",
                message(json!({"ItemId":{"Id":"move"}})),
            ),
            ("GetItem", message(new)),
        ]);
        assert_eq!(session.apply_work_blocks(&plan).unwrap()["complete"], true);
        worker.join().unwrap();
        let requests = requests.lock().unwrap();
        assert_eq!(requests[6]["Body"]["ShouldSendUpdateToAttendees"], false);
        assert_eq!(
            requests[6]["Body"]["ItemChange"]["ItemId"]["ChangeKey"],
            "old-key"
        );
        assert_eq!(
            requests[6]["Body"]["ItemChange"]["Updates"]
                .as_array()
                .unwrap()
                .len(),
            3
        );
    }

    #[test]
    fn planning_uses_earlier_empty_time_before_focus_and_honors_deadlines() {
        let events = vec![
            json!({"ItemId":{"Id":"meeting"},"Subject":"Meeting",
            "Start":"2026-10-12T16:00:00Z","End":"2026-10-12T17:00:00Z"}),
            json!({"ItemId":{"Id":"focus"},"Subject":"Focus Time",
            "Start":"2026-10-12T22:00:00Z","End":"2026-10-13T00:00:00Z"}),
        ];
        let (session, _, worker) = mock(vec![
            ("GetCalendarView", calendar(events.clone())),
            ("GetCalendarView", calendar(events)),
        ]);
        let mut task_input = input();
        let plan = session.plan_work_blocks(&task_input).unwrap();
        assert_eq!(
            plan.blocks[0].event.start.as_deref(),
            Some("2026-10-12T10:00:00-07:00")
        );
        assert!(plan.blocks[0].overlaps.is_empty());
        task_input.tasks[0].deadline = Some("2026-10-12T10:30:00-07:00".into());
        assert!(session.plan_work_blocks(&task_input).is_err());
        worker.join().unwrap();
    }

    #[test]
    fn apply_rejects_stale_snapshot_before_any_write() {
        let (session, requests, worker) = mock(vec![
            ("GetCalendarView", calendar(vec![])),
            (
                "GetCalendarView",
                calendar(vec![json!({"ItemId":{"Id":"new"},"Subject":"New meeting",
                "Start":"2026-10-12T16:00:00Z","End":"2026-10-12T17:00:00Z"})]),
            ),
        ]);
        let plan = session.plan_work_blocks(&input()).unwrap();
        assert!(matches!(
            session.apply_work_blocks(&plan),
            Err(OutlookWebError::PlanChanged)
        ));
        worker.join().unwrap();
        assert_eq!(requests.lock().unwrap().len(), 2);
    }

    #[test]
    fn batch_apply_stops_after_failure_and_reports_completed_writes() {
        let created = json!({"__type":"CalendarItem:#Exchange","ItemId":{"Id":"created","ChangeKey":"key"},
            "Subject":"Prepare","Start":"2026-10-12T16:00:00Z","End":"2026-10-12T17:00:00Z",
            "IsMeeting":false,"IsOrganizer":true,"IsRecurring":false,"CalendarItemType":"Single"});
        let (session, requests, worker) = mock(vec![
            ("GetCalendarView", calendar(vec![])),
            ("GetCalendarView", calendar(vec![])),
            ("GetCalendarView", calendar(vec![])),
            ("CreateItem", message(json!({"ItemId":{"Id":"created"}}))),
            ("GetItem", message(created.clone())),
            ("GetCalendarView", calendar(vec![created])),
            (
                "CreateItem",
                json!({"Body":{"ResponseMessages":{"Items":[{"ResponseCode":"ErrorAccessDenied","ResponseClass":"Error"}]}}}),
            ),
        ]);
        let mut request = input();
        request.tasks.push(WorkTask {
            subject: "Next".into(),
            body: None,
            minutes: 30,
            deadline: None,
            id: None,
        });
        let plan = session.plan_work_blocks(&request).unwrap();
        let result = session.apply_work_blocks(&plan).unwrap();
        assert_eq!(result["complete"], false);
        assert_eq!(result["results"].as_array().unwrap().len(), 1);
        assert_eq!(result["retry_automatically"], false);
        worker.join().unwrap();
        assert_eq!(requests.lock().unwrap().len(), 7);
    }
    #[test]
    fn merges_overlaps_and_keeps_boundary_gaps() {
        let w = span("2026-10-12T09:00:00-07:00", "2026-10-12T15:00:00-07:00").unwrap();
        let b = [
            span("2026-10-12T08:00:00-07:00", "2026-10-12T10:00:00-07:00").unwrap(),
            span("2026-10-12T09:30:00-07:00", "2026-10-12T11:00:00-07:00").unwrap(),
            span("2026-10-12T12:00:00-07:00", "2026-10-12T13:00:00-07:00").unwrap(),
        ];
        let g = gaps(w, &b);
        assert_eq!(g.len(), 2);
        assert_eq!((g[0].end - g[0].start).num_minutes(), 60);
        assert_eq!((g[1].end - g[1].start).num_minutes(), 120);
    }
    #[test]
    fn missing_free_busy_blocks_and_exact_title_override_is_explicit() {
        let items = vec![json!({"ItemId":{"Id":"x"},"Subject":"Focus Time",
            "Start":"2026-10-12T22:00:00Z","End":"2026-10-13T00:00:00Z"})];
        assert_eq!(busy(&items, &[], &[]).unwrap().len(), 1);
        assert!(
            busy(&items, &[], &["Focus Time".into()])
                .unwrap()
                .is_empty()
        );
        assert_eq!(busy(&items, &[], &["focus".into()]).unwrap().len(), 1);
    }
    #[test]
    fn rejects_ambiguous_windows_and_invalid_durations() {
        assert!(span("2026-10-12T09:00:00", "2026-10-12T10:00:00").is_err());
        assert!(minimum(0).is_err());
        assert!(minimum(481).is_err());
        assert!(windows(&[]).is_err());
    }
}
