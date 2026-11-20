use chrono::{Datelike, Duration, NaiveDate, NaiveDateTime};
use serde::Deserialize;
use serde_json::{Value, json};

use crate::{
    OutlookWebError, OutlookWebSession, get_item_request,
    operations::{change, single_item},
};

#[derive(Debug, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct SeriesScheduleInput {
    /// Local wall-clock timestamps in the series' existing time zone, without an offset.
    pub start: String,
    pub end: String,
    pub interval_weeks: u16,
}

fn local_time(value: &str) -> Result<NaiveDateTime, OutlookWebError> {
    NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S")
        .ok()
        .filter(|date| date.format("%Y-%m-%dT%H:%M:%S").to_string() == value)
        .ok_or({
            OutlookWebError::InvalidInput(
                "series start/end require local YYYY-MM-DDTHH:MM:SS timestamps without an offset",
            )
        })
}

pub fn series_item_request(id: &str, master: bool) -> Value {
    let mut request = get_item_request(id);
    request["Body"]["ItemShape"]["BodyType"] = json!("HTML");
    if master {
        request["Body"]["ItemIds"] =
            json!([{"__type":"RecurringMasterItemId:#Exchange","OccurrenceId":id}]);
    }
    request
}

fn required<'a>(item: &'a Value, path: &str) -> Result<&'a str, OutlookWebError> {
    item.pointer(path)
        .and_then(Value::as_str)
        .filter(|value| !value.trim().is_empty())
        .ok_or(OutlookWebError::InvalidResponse)
}

fn organizer(item: &Value, types: &[&str]) -> Result<(), OutlookWebError> {
    if item["__type"] != "CalendarItem:#Exchange"
        || item["IsOrganizer"] != true
        || item["IsMeeting"] != true
        || item["IsCancelled"] == true
        || item["IsAllDayEvent"] == true
        || !item["CalendarItemType"]
            .as_str()
            .is_some_and(|kind| types.contains(&kind))
    {
        return Err(OutlookWebError::InvalidInput(
            "series edits require an active organizer-owned timed meeting occurrence",
        ));
    }
    required(item, "/ItemId/Id")?;
    required(item, "/ItemId/ChangeKey")?;
    Ok(())
}

fn last_weekly_occurrence(
    master: &Value,
    boundary: NaiveDate,
) -> Result<Option<NaiveDate>, OutlookWebError> {
    let weekdays = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday",
    ];
    let pattern = &master["Recurrence"]["RecurrencePattern"];
    let first_day = required(pattern, "/FirstDayOfWeek")?;
    let first_day = weekdays
        .iter()
        .position(|day| *day == first_day)
        .ok_or(OutlookWebError::InvalidResponse)?;
    let days = required(pattern, "/DaysOfWeek")?
        .split_whitespace()
        .collect::<Vec<_>>();
    if days.is_empty()
        || days
            .iter()
            .any(|day| !weekdays.contains(day) && !["Day", "Weekday", "WeekendDay"].contains(day))
    {
        return Err(OutlookWebError::InvalidResponse);
    }
    let interval = pattern["Interval"]
        .as_i64()
        .filter(|interval| *interval > 0)
        .ok_or(OutlookWebError::InvalidResponse)?;
    let start = required(master, "/Recurrence/RecurrenceRange/StartDate")?
        .get(..10)
        .ok_or(OutlookWebError::InvalidResponse)?;
    let start = NaiveDate::parse_from_str(start, "%Y-%m-%d")
        .map_err(|_| OutlookWebError::InvalidResponse)?;
    let offset = (i64::from(start.weekday().num_days_from_monday())
        - i64::try_from(first_day).map_err(|_| OutlookWebError::InvalidResponse)?)
    .rem_euclid(7);
    let anchor = start - Duration::days(offset);
    let mut last = None;
    for offset in 0..7 {
        let candidate = anchor + Duration::days(offset);
        let weekday = candidate.weekday();
        let name = weekdays[weekday.num_days_from_monday() as usize];
        let weekend = weekday.num_days_from_monday() >= 5;
        if candidate >= boundary
            || !days.iter().any(|day| {
                *day == name
                    || *day == "Day"
                    || (*day == "Weekday" && !weekend)
                    || (*day == "WeekendDay" && weekend)
            })
        {
            continue;
        }
        let weeks = (boundary - Duration::days(1) - candidate).num_days() / 7;
        let candidate = candidate + Duration::weeks((weeks / interval) * interval);
        if candidate >= start && last.is_none_or(|previous| candidate > previous) {
            last = Some(candidate);
        }
    }
    Ok(last)
}

pub fn reschedule_series_request(
    occurrence: &Value,
    master: &Value,
    input: &SeriesScheduleInput,
) -> Result<Value, OutlookWebError> {
    organizer(occurrence, &["Occurrence", "Exception"])?;
    organizer(master, &["RecurringMaster"])?;
    if occurrence["CalendarItemType"] == "Exception"
        && occurrence["OriginalStart"] != occurrence["Start"]
    {
        return Err(OutlookWebError::InvalidInput(
            "select an occurrence whose start still matches its original recurrence slot",
        ));
    }
    let start = local_time(&input.start)?;
    let end = local_time(&input.end)?;
    let original = required(occurrence, "/StartWallClock")?
        .get(..19)
        .ok_or(OutlookWebError::InvalidResponse)
        .and_then(local_time)?;
    let zone = required(master, "/StartTimeZoneId")?;
    if zone != required(master, "/EndTimeZoneId")?
        || zone != required(occurrence, "/StartTimeZoneId")?
        || zone != required(occurrence, "/EndTimeZoneId")?
        || required(occurrence, "/UID")? != required(master, "/UID")?
    {
        return Err(OutlookWebError::InvalidInput(
            "the occurrence and series must share an identity and start/end time zone",
        ));
    }
    if end <= start
        || end.date() != start.date()
        || start.date() < original.date()
        || start.date() > original.date() + Duration::days(31)
        || !(1..=52).contains(&input.interval_weeks)
    {
        return Err(OutlookWebError::InvalidInput(
            "use a positive same-day duration, interval_weeks 1-52, and a start within 31 days on or after the selected occurrence",
        ));
    }
    let mut recurrence = master["Recurrence"].clone();
    if recurrence["RecurrencePattern"]["__type"] != "WeeklyRecurrence:#Exchange"
        || recurrence["RecurrenceRange"]["__type"] != "EndDateRecurrence:#Exchange"
    {
        return Err(OutlookWebError::InvalidInput(
            "only weekly series with an existing end date are supported; that end date is preserved",
        ));
    }
    let _previous_slot = last_weekly_occurrence(master, original.date())?;
    let until = required(&recurrence, "/RecurrenceRange/EndDate")?
        .get(..10)
        .ok_or(OutlookWebError::InvalidResponse)?
        .to_owned();
    let start_date = start.date().to_string();
    let until_date = NaiveDate::parse_from_str(&until, "%Y-%m-%d")
        .map_err(|_| OutlookWebError::InvalidResponse)?;
    if start.date() > until_date {
        return Err(OutlookWebError::InvalidInput(
            "the new start must not follow the existing series end date",
        ));
    }
    recurrence["RecurrencePattern"]["Interval"] = json!(input.interval_weeks);
    recurrence["RecurrencePattern"]["DaysOfWeek"] = json!(match start.weekday() {
        chrono::Weekday::Mon => "Monday",
        chrono::Weekday::Tue => "Tuesday",
        chrono::Weekday::Wed => "Wednesday",
        chrono::Weekday::Thu => "Thursday",
        chrono::Weekday::Fri => "Friday",
        chrono::Weekday::Sat => "Saturday",
        chrono::Weekday::Sun => "Sunday",
    });
    recurrence["RecurrenceRange"]["StartDate"] = json!(start_date);
    recurrence["RecurrenceRange"]["EndDate"] = json!(until);
    Ok(json!({
        "__type":"UpdateCalendarEventJsonRequest:#Exchange",
        "Header":{
            "__type":"JsonRequestHeaders:#Exchange","RequestServerVersion":"V2018_01_08",
            "TimeZoneContext":{"TimeZoneDefinition":{"Id":zone}}
        },
        "Body":{
            "__type":"UpdateCalendarEventRequest:#Exchange",
            "EventId":occurrence["ItemId"],"EventScope":2,"TargetAudience":0,
            "ShouldSendUpdateToAttendees":true,
            "ItemChange":{
                "__type":"ItemChange:#Exchange","ItemId":occurrence["ItemId"],
                "Updates":[
                    change("Start",json!(input.start),"CalendarItem:#Exchange"),
                    change("End",json!(input.end),"CalendarItem:#Exchange"),
                    change("Recurrence",recurrence,"CalendarItem:#Exchange")
                ]
            }
        }
    }))
}

impl OutlookWebSession {
    fn series_items(&self, id: &str) -> Result<(Value, Value), OutlookWebError> {
        let occurrence = self.call("GetItem", &series_item_request(id, false))?;
        let master = self.call("GetItem", &series_item_request(id, true))?;
        Ok((
            single_item(&occurrence)?.clone(),
            single_item(&master)?.clone(),
        ))
    }

    pub fn preview_series_schedule(
        &self,
        id: &str,
        input: &SeriesScheduleInput,
    ) -> Result<Value, OutlookWebError> {
        let (occurrence, master) = self.series_items(id)?;
        let request = reschedule_series_request(&occurrence, &master, input)?;
        let count = ["RequiredAttendees", "OptionalAttendees", "Resources"]
            .iter()
            .try_fold(0, |count, key| {
                attendees(&master[*key]).map(|list| count + list.len())
            })?;
        Ok(json!({
            "preview":true,"subject":master["Subject"],"occurrence_id":id,
            "change_key":occurrence["ItemId"]["ChangeKey"],
            "series_change_key":master["ItemId"]["ChangeKey"],
            "scope":"this_and_following","start":input.start,"end":input.end,
            "time_zone":master["StartTimeZoneId"],"attendee_count":count,
            "recurrence":request["Body"]["ItemChange"]["Updates"][2]["Item"]["Recurrence"],
            "notifies_attendees":true,"preserves":["past occurrences","attendees","body","online meeting","series end date"]
        }))
    }

    pub fn reschedule_series(
        &self,
        id: &str,
        input: &SeriesScheduleInput,
        change_key: &str,
        series_change_key: &str,
    ) -> Result<Value, OutlookWebError> {
        let (occurrence, master) = self.series_items(id)?;
        let request = reschedule_series_request(&occurrence, &master, input)?;
        if occurrence["ItemId"]["ChangeKey"] != change_key
            || master["ItemId"]["ChangeKey"] != series_change_key
        {
            return Err(OutlookWebError::PlanChanged);
        }
        self.execute("UpdateCalendarEvent", &request, true)?;
        self.verify_series_schedule(&occurrence, &master, input, &request)
            .map_err(|_| OutlookWebError::VerificationFailed)
    }

    fn verify_series_schedule(
        &self,
        occurrence: &Value,
        master: &Value,
        input: &SeriesScheduleInput,
        request: &Value,
    ) -> Result<Value, OutlookWebError> {
        let original_start = required(occurrence, "/Start")?;
        let start = local_time(&input.start)?;
        let range_start = original_start
            .get(..10)
            .ok_or(OutlookWebError::InvalidResponse)?;
        let range_end = (start.date() + Duration::days(2)).to_string();
        let view = self.calendar_list(
            required(master, "/ParentFolderId/Id")?,
            &format!("{range_start}T00:00:00Z"),
            &format!("{range_end}T00:00:00Z"),
        )?;
        if view["Body"]["IncludesLastItemInRange"] == false {
            return Err(OutlookWebError::TruncatedSearch);
        }
        let items = view["Body"]["Items"]
            .as_array()
            .ok_or(OutlookWebError::InvalidResponse)?;
        let mut matched = Vec::new();
        for item in items
            .iter()
            .filter(|item| item["Subject"] == master["Subject"])
        {
            let fetched = self.call(
                "GetItem",
                &series_item_request(required(item, "/ItemId/Id")?, false),
            )?;
            let event = single_item(&fetched)?;
            if event["IsOrganizer"] != true {
                continue;
            }
            if event["Start"] == original_start
                && occurrence["StartWallClock"]
                    .as_str()
                    .and_then(|value| value.get(..19))
                    != Some(input.start.as_str())
            {
                return Err(OutlookWebError::VerificationFailed);
            }
            if event["StartWallClock"]
                .as_str()
                .and_then(|value| value.get(..19))
                == Some(input.start.as_str())
                && event["EndWallClock"]
                    .as_str()
                    .and_then(|value| value.get(..19))
                    == Some(input.end.as_str())
            {
                matched.push(event.clone());
            }
        }
        if matched.len() != 1 {
            return Err(OutlookWebError::VerificationFailed);
        }
        let event = &matched[0];
        let series = self.call(
            "GetItem",
            &series_item_request(required(event, "/ItemId/Id")?, true),
        )?;
        let updated = single_item(&series)?;
        let expected = &request["Body"]["ItemChange"]["Updates"][2]["Item"]["Recurrence"];
        let actual = &updated["Recurrence"];
        if actual["RecurrencePattern"] != expected["RecurrencePattern"]
            || ["StartDate", "EndDate"].iter().any(|key| {
                actual["RecurrenceRange"][key]
                    .as_str()
                    .and_then(|value| value.get(..10))
                    != expected["RecurrenceRange"][key].as_str()
            })
            || ["StartTimeZoneId", "EndTimeZoneId", "Subject"]
                .iter()
                .any(|key| updated[key] != master[key])
            || updated["Body"] != master["Body"]
        {
            return Err(OutlookWebError::VerificationFailed);
        }
        for key in ["RequiredAttendees", "OptionalAttendees", "Resources"] {
            if attendees(&updated[key])? != attendees(&master[key])? {
                return Err(OutlookWebError::VerificationFailed);
            }
        }
        self.verify_past_series(occurrence, master, updated)?;
        Ok(json!({
            "verified":true,"notifications_submitted":true,"scope":"this_and_following",
            "occurrence_id":event["ItemId"]["Id"],"series_id":updated["ItemId"]["Id"],
            "start":event["Start"],"end":event["End"],
            "time_zone":updated["StartTimeZoneId"],"recurrence":actual
        }))
    }

    fn verify_past_series(
        &self,
        occurrence: &Value,
        master: &Value,
        updated: &Value,
    ) -> Result<(), OutlookWebError> {
        let boundary = local_time(
            required(occurrence, "/StartWallClock")?
                .get(..19)
                .ok_or(OutlookWebError::InvalidResponse)?,
        )?
        .date();
        let old_start = required(master, "/Recurrence/RecurrenceRange/StartDate")?
            .get(..10)
            .ok_or(OutlookWebError::InvalidResponse)?;
        if boundary.to_string().as_str() > old_start {
            if updated["ItemId"]["Id"] == master["ItemId"]["Id"] {
                return Err(OutlookWebError::VerificationFailed);
            }
            let previous = self.call(
                "GetItem",
                &series_item_request(required(master, "/ItemId/Id")?, false),
            )?;
            let previous = single_item(&previous)?;
            let last = last_weekly_occurrence(master, boundary)?
                .ok_or(OutlookWebError::VerificationFailed)?;
            let cutoff = required(previous, "/Recurrence/RecurrenceRange/EndDate")?
                .get(..10)
                .ok_or(OutlookWebError::InvalidResponse)?;
            let cutoff = NaiveDate::parse_from_str(cutoff, "%Y-%m-%d")
                .map_err(|_| OutlookWebError::InvalidResponse)?;
            if cutoff < last
                || cutoff >= boundary
                || previous["Recurrence"]["RecurrencePattern"]
                    != master["Recurrence"]["RecurrencePattern"]
                || [
                    "Start",
                    "End",
                    "Body",
                    "Subject",
                    "StartTimeZoneId",
                    "EndTimeZoneId",
                ]
                .iter()
                .any(|key| previous[key] != master[key])
            {
                return Err(OutlookWebError::VerificationFailed);
            }
        }

        Ok(())
    }
}

fn attendees(value: &Value) -> Result<Vec<String>, OutlookWebError> {
    if value.is_null() {
        return Ok(Vec::new());
    }
    let mut addresses = value
        .as_array()
        .ok_or(OutlookWebError::InvalidResponse)?
        .iter()
        .map(|attendee| required(attendee, "/Mailbox/EmailAddress").map(str::to_ascii_lowercase))
        .collect::<Result<Vec<_>, _>>()?;
    addresses.sort();
    Ok(addresses)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_support::{message, mock};

    fn occurrence() -> Value {
        json!({
            "__type":"CalendarItem:#Exchange","IsMeeting":true,"IsOrganizer":true,
            "IsCancelled":false,
            "CalendarItemType":"Exception","UID":"series-uid",
            "ItemId":{"__type":"ItemId:#Exchange","Id":"occurrence","ChangeKey":"key"},
            "Start":"2026-12-01T18:35:00Z","OriginalStart":"2026-12-01T18:35:00Z","End":"2026-12-01T19:00:00Z",
            "StartWallClock":"2026-12-01T10:35:00-08:00",
            "StartTimeZoneId":"Pacific Standard Time","EndTimeZoneId":"Pacific Standard Time"
        })
    }

    fn master() -> Value {
        let mut item = occurrence();
        item["CalendarItemType"] = json!("RecurringMaster");
        item["ItemId"]["Id"] = json!("master");
        item["Subject"] = json!("Synthetic weekly meeting");
        item["ParentFolderId"] = json!({"Id":"folder"});
        item["Body"] = json!({"BodyType":"HTML","Value":"<a href='https://example.com'>Join</a>"});
        item["RequiredAttendees"] =
            json!([{"Mailbox":{"EmailAddress":"a@example.com"},"ResponseType":"Accept"}]);
        item["Recurrence"] = json!({
            "RecurrencePattern":{"__type":"WeeklyRecurrence:#Exchange","Interval":4,
                "DaysOfWeek":"Tuesday","FirstDayOfWeek":"Monday"},
            "RecurrenceRange":{"__type":"EndDateRecurrence:#Exchange",
                "StartDate":"2026-07-14-07:00","EndDate":"2027-01-14-08:00"}
        });
        item
    }

    fn input() -> SeriesScheduleInput {
        SeriesScheduleInput {
            start: "2026-12-08T10:15:00".into(),
            end: "2026-12-08T11:00:00".into(),
            interval_weeks: 2,
        }
    }

    #[test]
    fn future_schedule_uses_occurrence_id_local_zone_and_preserves_end_date() {
        let request = reschedule_series_request(&occurrence(), &master(), &input()).unwrap();
        assert_eq!(request["Body"]["EventScope"], 2);
        assert_eq!(request["Body"]["EventId"]["Id"], "occurrence");
        assert_eq!(request["Body"]["EventId"]["ChangeKey"], "key");
        assert_eq!(
            request["Header"]["TimeZoneContext"]["TimeZoneDefinition"]["Id"],
            "Pacific Standard Time"
        );
        let updates = request["Body"]["ItemChange"]["Updates"].as_array().unwrap();
        assert_eq!(updates.len(), 3);
        assert_eq!(updates[0]["Item"]["Start"], "2026-12-08T10:15:00");
        assert_eq!(
            updates[2]["Item"]["Recurrence"]["RecurrencePattern"]["Interval"],
            2
        );
        assert_eq!(
            updates[2]["Item"]["Recurrence"]["RecurrenceRange"]["StartDate"],
            "2026-12-08"
        );
        assert_eq!(
            updates[2]["Item"]["Recurrence"]["RecurrenceRange"]["EndDate"],
            "2027-01-14"
        );
    }

    #[test]
    fn invalid_scope_timezone_dates_and_intervals_are_rejected() {
        for (path, value) in [
            ("/IsOrganizer", json!(false)),
            ("/CalendarItemType", json!("Single")),
            ("/IsCancelled", json!(true)),
            ("/UID", json!("other")),
            ("/StartTimeZoneId", json!("UTC")),
            ("/ItemId/ChangeKey", json!("")),
        ] {
            let mut item = occurrence();
            *item
                .pointer_mut(path)
                .unwrap_or_else(|| panic!("missing {path}")) = value;
            assert!(reschedule_series_request(&item, &master(), &input()).is_err());
        }
        for start in [
            "2026-12-08T18:15:00Z",
            "2026-12-08T10:15:00-08:00",
            "2026-11-30T10:15:00",
            "2027-01-15T10:15:00",
        ] {
            let mut schedule = input();
            schedule.start = start.into();
            assert!(reschedule_series_request(&occurrence(), &master(), &schedule).is_err());
        }
        for interval in [0, 53] {
            let mut schedule = input();
            schedule.interval_weeks = interval;
            assert!(reschedule_series_request(&occurrence(), &master(), &schedule).is_err());
        }
        let mut series = master();
        series["Recurrence"]["RecurrenceRange"]["__type"] = json!("NoEndRecurrence:#Exchange");
        assert!(reschedule_series_request(&occurrence(), &series, &input()).is_err());
    }

    #[test]
    fn previous_weekly_slot_respects_interval_week_start_and_multiple_days() {
        let boundary = NaiveDate::from_ymd_opt(2026, 12, 1).unwrap();
        assert_eq!(
            last_weekly_occurrence(&master(), boundary).unwrap(),
            NaiveDate::from_ymd_opt(2026, 11, 3)
        );
        let mut series = master();
        series["Recurrence"]["RecurrencePattern"]["DaysOfWeek"] = json!("Tuesday Thursday");
        assert_eq!(
            last_weekly_occurrence(&series, boundary).unwrap(),
            NaiveDate::from_ymd_opt(2026, 11, 5)
        );
        series["Recurrence"]["RecurrencePattern"]["DaysOfWeek"] = json!("Sunday Tuesday");
        series["Recurrence"]["RecurrencePattern"]["FirstDayOfWeek"] = json!("Sunday");
        assert_eq!(
            last_weekly_occurrence(&series, boundary).unwrap(),
            NaiveDate::from_ymd_opt(2026, 11, 29)
        );
        assert_eq!(
            last_weekly_occurrence(&master(), NaiveDate::from_ymd_opt(2026, 7, 14).unwrap())
                .unwrap(),
            None
        );
    }

    #[test]
    fn preview_is_read_only_and_stale_apply_never_dispatches_update() {
        let replies = vec![
            ("GetItem", message(occurrence())),
            ("GetItem", message(master())),
        ];
        let (session, requests, done) = mock(replies.clone());
        let preview = session
            .preview_series_schedule("occurrence", &input())
            .unwrap();
        assert_eq!(preview["attendee_count"], 1);
        assert_eq!(requests.lock().unwrap().len(), 2);
        done.join().unwrap();
        let (session, requests, done) = mock(replies);
        assert!(matches!(
            session.reschedule_series("occurrence", &input(), "stale", "key"),
            Err(OutlookWebError::PlanChanged)
        ));
        assert_eq!(requests.lock().unwrap().len(), 2);
        done.join().unwrap();
    }

    #[test]
    fn apply_verifies_new_recurrence_and_preservation_without_retrying() {
        let schedule = input();
        let original = master();
        let request = reschedule_series_request(&occurrence(), &original, &schedule).unwrap();
        let mut new_event = occurrence();
        new_event["Subject"] = original["Subject"].clone();
        new_event["ItemId"]["Id"] = json!("new-occurrence");
        new_event["Start"] = json!("2026-12-08T18:15:00Z");
        new_event["End"] = json!("2026-12-08T19:00:00Z");
        new_event["StartWallClock"] = json!("2026-12-08T10:15:00-08:00");
        new_event["EndWallClock"] = json!("2026-12-08T11:00:00-08:00");
        let mut updated = original.clone();
        updated["ItemId"]["Id"] = json!("new-master");
        updated["Recurrence"] =
            request["Body"]["ItemChange"]["Updates"][2]["Item"]["Recurrence"].clone();
        updated["RequiredAttendees"][0]["ResponseType"] = json!("Unknown");
        for (cutoff, corrupt_body, success) in [
            ("2026-11-03-08:00", false, true),
            ("2026-11-30-08:00", false, true),
            ("2026-11-02-08:00", false, false),
            ("2026-12-01-08:00", false, false),
            ("2026-11-03-08:00", true, false),
        ] {
            let mut verified = updated.clone();
            if corrupt_body {
                verified["Body"]["Value"] = json!("replaced");
            }
            let mut replies = vec![
                ("GetItem", message(occurrence())),
                ("GetItem", message(original.clone())),
                (
                    "UpdateCalendarEvent",
                    message(json!({"ItemId":{"Id":"new-master"}})),
                ),
                (
                    "GetCalendarView",
                    json!({"Body":{"ResponseCode":"NoError","Items":[new_event.clone()]}}),
                ),
                ("GetItem", message(new_event.clone())),
                ("GetItem", message(verified)),
            ];
            if !corrupt_body {
                let mut previous = original.clone();
                previous["Recurrence"]["RecurrenceRange"]["EndDate"] = json!(cutoff);
                replies.push(("GetItem", message(previous)));
            }
            let (session, requests, done) = mock(replies);
            let result = session.reschedule_series("occurrence", &schedule, "key", "key");
            assert_eq!(result.is_ok(), success);
            assert_eq!(
                requests
                    .lock()
                    .unwrap()
                    .iter()
                    .filter(|r| r["Body"]["__type"] == "UpdateCalendarEventRequest:#Exchange")
                    .count(),
                1
            );
            done.join().unwrap();
        }
    }
}
