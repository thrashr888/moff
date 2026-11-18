use moff_outlook_owa::{
    EventInput, MeetingResponse, MeetingResponseInput, OutlookWebError, OutlookWebSession,
    SendMailInput, create_event_request,
};
use rmcp::model::{Tool, ToolAnnotations};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;

use super::{Failure, MailGetArgs, decode, default_limit, tool};
use crate::{groups_list_path, onedrive_list_path, people_list_path, search_pages};

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct SearchArgs {
    query: String,
    #[serde(default = "default_limit")]
    #[schemars(range(min = 1, max = 100))]
    limit: u16,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct DriveListArgs {
    /// Omit to list the drive root.
    item_id: Option<String>,
    #[serde(default = "default_limit")]
    #[schemars(range(min = 1, max = 100))]
    limit: u16,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct ListArgs {
    #[serde(default = "default_limit")]
    #[schemars(range(min = 1, max = 100))]
    limit: u16,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct SendArgs {
    /// Exact SMTP email addresses, not inferred names or groups.
    to: Vec<String>,
    #[serde(default)]
    cc: Vec<String>,
    #[serde(default)]
    bcc: Vec<String>,
    subject: String,
    body: String,
    /// True only after the user approves these exact recipients and content.
    confirm: bool,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct CreateArgs {
    folder_id: String,
    subject: String,
    body: Option<String>,
    /// Start timestamp with a timezone offset (RFC 3339).
    start: String,
    /// End timestamp with a timezone offset (RFC 3339).
    end: String,
    /// True only after the user approves this appointment.
    confirm: bool,
}

#[derive(Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
enum Rsvp {
    Accept,
    Tentative,
    Decline,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct RespondArgs {
    /// ID of a received meeting or occurrence, not a series master.
    id: String,
    response: Rsvp,
    message: Option<String>,
    proposed_start: Option<String>,
    proposed_end: Option<String>,
    /// True only after the user approves this exact RSVP.
    confirm: bool,
}

pub(super) fn read_catalog() -> Vec<Tool> {
    vec![
        tool::<MailGetArgs>(
            "calendar_get",
            "Read one calendar item by ID, including its body and attendees.",
        ),
        tool::<SearchArgs>(
            "people_search",
            "Search Outlook's People list. Results may be bounded; inspect the returned search scope.",
        ),
        tool::<DriveListArgs>(
            "onedrive_list",
            "List OneDrive item metadata without downloading any files.",
        ),
        tool::<ListArgs>("groups_list", "List Microsoft 365 groups."),
    ]
}

pub(super) fn write_catalog() -> Vec<Tool> {
    let write = |mut t: Tool| {
        t.annotations = Some(
            ToolAnnotations::new()
                .read_only(false)
                .destructive(true)
                .idempotent(false)
                .open_world(true),
        );
        t
    };
    vec![
        write(tool::<SendArgs>(
            "mail_send",
            "Send one plain-text email without attachments. Requires host user approval. An acknowledgement is not proof of delivery; never retry automatically.",
        )),
        write(tool::<CreateArgs>(
            "calendar_create",
            "Create one personal appointment without attendees or invitations. Requires host user approval.",
        )),
        write(tool::<RespondArgs>(
            "calendar_respond",
            "Send an accept, tentative, or decline RSVP for one received invitation. Requires host user approval. Never retry automatically.",
        )),
    ]
}

pub(super) fn writes(name: &str) -> bool {
    matches!(name, "mail_send" | "calendar_create" | "calendar_respond")
}

fn confirm(confirmed: bool) -> Result<(), Failure> {
    if confirmed {
        Ok(())
    } else {
        Err(Failure::invalid(
            "User approval and confirm=true are required",
        ))
    }
}

fn nonempty(value: &str, field: &str) -> Result<(), Failure> {
    if value.trim().is_empty() {
        Err(Failure::invalid(format!("{field} must not be empty")))
    } else {
        Ok(())
    }
}

fn check_limit(limit: u16) -> Result<(), Failure> {
    if (1..=100).contains(&limit) {
        Ok(())
    } else {
        Err(Failure::invalid("limit must be from 1 to 100"))
    }
}

pub(super) enum Request {
    CalendarGet(String),
    PeopleSearch(String, u16),
    ProductRead(String),
    Send(SendMailInput),
    Create(String, EventInput),
    Respond(String, MeetingResponseInput),
}

impl Request {
    pub(super) fn parse(
        name: &str,
        args: &Value,
        write_actions: bool,
    ) -> Result<Option<Self>, Failure> {
        if writes(name) && !write_actions {
            return Err(Failure::invalid("Write actions are not enabled"));
        }
        Ok(Some(match name {
            "calendar_get" => {
                let a: MailGetArgs = decode(args.clone())?;
                nonempty(&a.id, "id")?;
                Self::CalendarGet(a.id)
            }
            "people_search" => {
                let a: SearchArgs = decode(args.clone())?;
                nonempty(&a.query, "query")?;
                check_limit(a.limit)?;
                Self::PeopleSearch(a.query, a.limit)
            }
            "onedrive_list" => {
                let a: DriveListArgs = decode(args.clone())?;
                check_limit(a.limit)?;
                if let Some(id) = &a.item_id {
                    nonempty(id, "item_id")?;
                }
                Self::ProductRead(onedrive_list_path(a.item_id.as_deref(), a.limit))
            }
            "groups_list" => {
                let a: ListArgs = decode(args.clone())?;
                check_limit(a.limit)?;
                Self::ProductRead(groups_list_path(a.limit))
            }
            "mail_send" => {
                let a: SendArgs = decode(args.clone())?;
                confirm(a.confirm)?;
                let input = SendMailInput {
                    to: a.to,
                    cc: a.cc,
                    bcc: a.bcc,
                    subject: a.subject,
                    body: a.body,
                };
                input.validate()?;
                Self::Send(input)
            }
            "calendar_create" => {
                let a: CreateArgs = decode(args.clone())?;
                confirm(a.confirm)?;
                let input = EventInput {
                    subject: Some(a.subject),
                    body: a.body,
                    start: Some(a.start),
                    end: Some(a.end),
                };
                create_event_request(&a.folder_id, &input)?;
                Self::Create(a.folder_id, input)
            }
            "calendar_respond" => {
                let a: RespondArgs = decode(args.clone())?;
                confirm(a.confirm)?;
                nonempty(&a.id, "id")?;
                Self::Respond(
                    a.id,
                    MeetingResponseInput {
                        response: match a.response {
                            Rsvp::Accept => MeetingResponse::Accept,
                            Rsvp::Tentative => MeetingResponse::Tentative,
                            Rsvp::Decline => MeetingResponse::Decline,
                        },
                        message: a.message,
                        proposed_start: a.proposed_start,
                        proposed_end: a.proposed_end,
                    },
                )
            }
            _ => return Ok(None),
        }))
    }

    pub(super) fn run(self, target: Option<(u32, u32)>) -> Result<Value, Failure> {
        let is_write = matches!(self, Self::Send(_) | Self::Create(..) | Self::Respond(..));
        let session = OutlookWebSession::from_safari_with_mutations(target, is_write)?;
        Ok(match self {
            Self::CalendarGet(id) => session
                .calendar_get(&id)?
                .pointer("/Body/ResponseMessages/Items/0/Items/0")
                .cloned()
                .ok_or(OutlookWebError::InvalidResponse)?,
            Self::PeopleSearch(query, limit) => {
                search_pages(&session, &people_list_path(limit), &query, limit, false)
                    .map_err(|error| Failure("outlook_error".into(), error))?
            }
            Self::ProductRead(path) => session.product_read(&path)?,
            Self::Send(input) => session.send_mail(&input)?,
            Self::Create(folder, input) => session.create_event(&folder, &input)?,
            Self::Respond(id, input) => session.respond_event(&id, &input)?,
        })
    }
}
