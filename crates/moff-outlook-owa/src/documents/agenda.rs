use std::collections::HashSet;

use chrono::{DateTime, Duration, FixedOffset};
use scraper::{Html, Selector};
use serde_json::{Value, json};

use crate::{OutlookWebError, OutlookWebSession, get_item_request, operations::single_item};

use super::sharepoint_url;

fn links(body: &str) -> Result<Vec<Value>, OutlookWebError> {
    let html = Html::parse_fragment(body);
    let selector = Selector::parse("a[href]").map_err(|_| OutlookWebError::InvalidResponse)?;
    let mut seen = HashSet::new();
    Ok(html
        .select(&selector)
        .filter_map(|anchor| {
            let link = anchor.value().attr("href")?;
            let url = sharepoint_url(link).ok()?;
            let word = url.path().starts_with("/:w:/")
                || url.path().to_ascii_lowercase().ends_with(".docx")
                || (url.path().ends_with("/Doc.aspx")
                    && url.query_pairs().any(|(key, value)| {
                        key == "file" && value.to_ascii_lowercase().ends_with(".docx")
                    }));
            if !word || !seen.insert(link.to_owned()) {
                return None;
            }
            Some(json!({"text":anchor.text().collect::<String>().trim(),"url":link}))
        })
        .collect())
}

fn start(item: &Value) -> Result<DateTime<FixedOffset>, OutlookWebError> {
    item["Start"]
        .as_str()
        .and_then(|value| DateTime::parse_from_rfc3339(value).ok())
        .ok_or(OutlookWebError::InvalidResponse)
}

impl OutlookWebSession {
    fn agenda_item(&self, id: &str) -> Result<Value, OutlookWebError> {
        let mut request = get_item_request(id);
        request["Body"]["ItemShape"]["BodyType"] = json!("HTML");
        self.call("GetItem", &request)
    }

    pub fn calendar_agenda(
        &self,
        id: &str,
        now: DateTime<FixedOffset>,
    ) -> Result<Value, OutlookWebError> {
        let response = self.agenda_item(id)?;
        let original = single_item(&response)?;
        let item = if original["CalendarItemType"] == "Single" {
            if start(original)? < now {
                return Err(OutlookWebError::InvalidInput(
                    "selected meeting is in the past",
                ));
            }
            original.clone()
        } else {
            self.next_agenda_occurrence(original, now)?
        };
        let body = item["Body"]["Value"]
            .as_str()
            .ok_or(OutlookWebError::InvalidResponse)?;
        let candidates = links(body)?;
        let selected = if candidates.len() == 1 {
            Some(&candidates[0])
        } else {
            None
        };
        Ok(
            json!({"meeting":{"id":item["ItemId"]["Id"],"subject":item["Subject"],
            "start":item["Start"],"end":item["End"],"time_zone":item["StartTimeZoneId"]},
            "agenda":selected,"candidates":candidates,"candidate_scope":"linked Word documents; multiple candidates require explicit selection",
            "search":{"lookahead_days":90,"next_occurrence_verified":true}}),
        )
    }

    fn next_agenda_occurrence(
        &self,
        original: &Value,
        now: DateTime<FixedOffset>,
    ) -> Result<Value, OutlookWebError> {
        let uid = original["UID"]
            .as_str()
            .filter(|v| !v.is_empty())
            .ok_or(OutlookWebError::InvalidResponse)?;
        let folder = original["ParentFolderId"]["Id"]
            .as_str()
            .ok_or(OutlookWebError::InvalidResponse)?;
        let deadline = now + Duration::days(90);
        let mut cursor = now;
        while cursor < deadline {
            let end = (cursor + Duration::days(7)).min(deadline);
            let view = self.calendar_list(folder, &cursor.to_rfc3339(), &end.to_rfc3339())?;
            if view["Body"]["IncludesLastItemInRange"] == false {
                return Err(OutlookWebError::TruncatedSearch);
            }
            let items = view["Body"]["Items"]
                .as_array()
                .ok_or(OutlookWebError::InvalidResponse)?;
            let mut matching = Vec::new();
            for item in items
                .iter()
                .filter(|item| item["Subject"] == original["Subject"])
            {
                let id = item["ItemId"]["Id"]
                    .as_str()
                    .ok_or(OutlookWebError::InvalidResponse)?;
                let response = self.agenda_item(id)?;
                let event = single_item(&response)?;
                let event_start = start(event)?;
                if event["UID"] == uid
                    && event["IsCancelled"] != true
                    && event_start >= now
                    && event_start < end
                {
                    matching.push((event_start, event.clone()));
                }
            }
            matching.sort_by_key(|(start, _)| *start);
            if let Some((_, event)) = matching.into_iter().next() {
                return Ok(event);
            }
            cursor = end;
        }
        Err(OutlookWebError::InvalidInput(
            "no upcoming occurrence of this series was found within 90 days; select the active series",
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn agenda_links_are_word_only_deduplicated_and_never_teams_credentials() {
        let body = r#"<a href="https://tenant.sharepoint.com/:w:/s/Project/opaque">Agenda &amp; notes</a>
            <a href="https://tenant.sharepoint.com/:w:/s/Project/opaque">Duplicate</a>
            <a href="https://teams.microsoft.com/meet/secret">Join</a>
            <a href="https://tenant.sharepoint.com/sites/Project/_layouts/15/Doc.aspx?file=Notes.docx&amp;sourcedoc=some-id">Other doc</a>
            <a href="https://tenant.sharepoint.com.evil.example/:w:/s/Project/opaque">Foreign</a>"#;
        let links = links(body).unwrap();
        assert_eq!(links.len(), 2);
        assert_eq!(links[0]["text"], "Agenda & notes");
        assert!(!serde_json::to_string(&links).unwrap().contains("secret"));
    }
}
