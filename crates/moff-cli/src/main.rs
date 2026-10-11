use std::{
    fs,
    io::Read,
    path::{Path, PathBuf},
    process::ExitCode,
};

use clap::{Args, Parser, Subcommand};
use moff_outlook_owa::{
    DraftInput, EventInput, OutlookWebError, OutlookWebSession, ReplyInput, WorkInput, WorkPlan,
    calendar_folders_request, calendar_view_request, validate_read_action,
};
use serde_json::{Value, json};

#[derive(Debug, Parser)]
#[command(
    name = "moff",
    version,
    about = "Personal Microsoft 365 browser-session CLI"
)]
struct Cli {
    #[command(subcommand)]
    command: Option<Command>,
}

#[derive(Debug, Subcommand)]
enum Command {
    Auth {
        #[command(flatten)]
        session: SessionArgs,
        #[command(subcommand)]
        command: AuthCommand,
    },
    Outlook {
        #[command(flatten)]
        session: SessionArgs,
        #[command(subcommand)]
        command: OutlookCommand,
    },
    Schema,
}

#[derive(Debug, Args, Clone, Copy)]
struct SessionArgs {
    #[arg(long = "safari", hide = true)]
    _safari: bool,
    #[arg(
        long,
        help = "Permit typed non-sending Safari writes; each write also requires --confirm"
    )]
    allow_mutation: bool,
    #[arg(long, requires = "safari_tab", value_parser = clap::value_parser!(u32).range(1..))]
    safari_window: Option<u32>,
    #[arg(long, requires = "safari_window", value_parser = clap::value_parser!(u32).range(1..))]
    safari_tab: Option<u32>,
}

#[derive(Debug, Subcommand)]
enum OutlookCommand {
    Auth {
        #[command(subcommand)]
        command: AuthCommand,
    },
    Mail {
        #[command(subcommand)]
        command: MailCommand,
    },
    Calendar {
        #[command(subcommand)]
        command: CalendarCommand,
    },
    People {
        #[command(subcommand)]
        command: PeopleCommand,
    },
    Org {
        #[command(subcommand)]
        command: OrgCommand,
    },
    Onedrive {
        #[command(subcommand)]
        command: OneDriveCommand,
    },
    Groups {
        #[command(subcommand)]
        command: GroupsCommand,
    },
    Loop {
        #[command(subcommand)]
        command: LoopCommand,
    },
    Service {
        #[command(subcommand)]
        command: ServiceCommand,
    },
}

#[derive(Debug, Subcommand)]
enum AuthCommand {
    /// Check silent authentication without a refresh or interactive sign-in.
    Status,
    /// Reload the selected Outlook tab for browser SSO; leave account/2FA clicks to you.
    Refresh {
        #[arg(long, default_value_t = 60, value_parser = clap::value_parser!(u32).range(1..=120))]
        wait_seconds: u32,
        #[arg(long, required = true)]
        confirm: bool,
    },
}

#[derive(Debug, Subcommand)]
enum MailCommand {
    /// Set read state explicitly; reply drafts restore Outlook's automatic read side effect.
    Mark {
        #[arg(long)]
        id: String,
        #[arg(long, required_unless_present = "unread", conflicts_with = "unread")]
        read: bool,
        #[arg(long, required_unless_present = "read", conflicts_with = "read")]
        unread: bool,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Read 1-20 messages in one non-marking request.
    GetBatch {
        #[arg(long, required = true, num_args = 1..=20)]
        id: Vec<String>,
    },
    /// Add/remove categories without replacing unrelated categories or read state.
    Categories {
        #[arg(long)]
        id: String,
        #[arg(long)]
        add: Vec<String>,
        #[arg(long)]
        remove: Vec<String>,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Save a threaded reply draft; optional `reply_all` in JSON keeps thread recipients.
    ReplyDraft {
        #[arg(long)]
        id: String,
        #[arg(long)]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
    List {
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
        #[arg(long, default_value = "inbox")]
        folder: String,
        #[arg(long, default_value_t = 0)]
        offset: u32,
    },
    Get {
        #[arg(long)]
        id: String,
    },
    /// Search a folder's subjects and text bodies, using a case-insensitive literal query.
    Search {
        #[arg(long)]
        query: String,
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
        #[arg(long, default_value = "inbox")]
        folder: String,
    },
    /// Create an unsent draft from JSON containing optional subject and body strings.
    Create {
        #[arg(long, value_name = "PATH")]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Update only the specified subject/body fields of an unsent draft.
    Update {
        #[arg(long)]
        id: String,
        #[arg(long, value_name = "PATH")]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Unavailable through Safari; retained to return an explicit unsupported error.
    Delete {
        #[arg(long)]
        id: String,
        #[arg(long, required = true)]
        confirm: bool,
    },
}

#[derive(Debug, Subcommand)]
enum CalendarCommand {
    /// Read 1-20 calendar items in one request.
    GetBatch {
        #[arg(long, required = true, num_args = 1..=20)]
        id: Vec<String>,
    },
    /// Find empty gaps; explicitly permitted title overlaps are ranked after empty gaps.
    Availability {
        #[arg(long)]
        folder_id: String,
        #[arg(long)]
        start: String,
        #[arg(long)]
        end: String,
        #[arg(long, default_value_t = 30, value_parser = clap::value_parser!(u16).range(1..=480))]
        minutes: u16,
        #[arg(long)]
        allow_overlap_title: Vec<String>,
    },
    WorkBlocks {
        #[command(subcommand)]
        command: WorkBlockCommand,
    },
    Folders,
    View {
        #[arg(long)]
        folder_id: String,
        #[arg(long)]
        start: String,
        #[arg(long)]
        end: String,
    },
    List {
        #[arg(long)]
        folder_id: String,
        #[arg(long)]
        start: String,
        #[arg(long)]
        end: String,
    },
    /// Search subjects and bodies within an explicit calendar time range.
    Search {
        #[arg(long)]
        folder_id: String,
        #[arg(long)]
        start: String,
        #[arg(long)]
        end: String,
        #[arg(long)]
        query: String,
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    Get {
        #[arg(long)]
        id: String,
    },
    /// Create a single personal appointment without attendees or invitations.
    Create {
        #[arg(long)]
        folder_id: String,
        #[arg(long, value_name = "PATH")]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Update a single personal appointment without attendees.
    Update {
        #[arg(long)]
        id: String,
        #[arg(long, value_name = "PATH")]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Unavailable through Safari; retained to return an explicit unsupported error.
    Delete {
        #[arg(long)]
        id: String,
        #[arg(long, required = true)]
        confirm: bool,
    },
}

#[derive(Debug, Subcommand)]
enum WorkBlockCommand {
    /// Produce a read-only plan using explicit work windows and tasks.
    Plan {
        #[arg(long)]
        body_file: PathBuf,
    },
    /// Apply a reviewed plan, rejecting changed calendars and stopping at first failure.
    Apply {
        #[arg(long)]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
}

#[derive(Debug, Subcommand)]
enum PeopleCommand {
    Search {
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    Get {
        #[arg(long)]
        user: String,
    },
}

#[derive(Debug, Subcommand)]
enum OrgCommand {
    Manager {
        #[arg(long)]
        user: Option<String>,
    },
    Reports {
        #[arg(long)]
        user: Option<String>,
    },
}

#[derive(Debug, Subcommand)]
enum OneDriveCommand {
    List {
        #[arg(long)]
        item_id: Option<String>,
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    Get {
        #[arg(long)]
        id: String,
    },
}

#[derive(Debug, Subcommand)]
enum GroupsCommand {
    List {
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
}

#[derive(Debug, Subcommand)]
enum LoopCommand {
    List {
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    Get {
        #[arg(long)]
        id: String,
    },
}

#[derive(Debug, Subcommand)]
enum ServiceCommand {
    Read {
        #[arg(long)]
        action: String,
        #[arg(long, value_name = "PATH")]
        body_file: PathBuf,
    },
}

fn print_json(value: &Value) {
    println!(
        "{}",
        serde_json::to_string(value).expect("JSON values serialize")
    );
}

fn error(message: impl std::fmt::Display) -> ExitCode {
    eprintln!("[browser_session] {message}");
    ExitCode::from(2)
}

fn session(args: SessionArgs) -> Result<OutlookWebSession, String> {
    OutlookWebSession::from_safari_with_mutations(
        args.safari_window.zip(args.safari_tab),
        args.allow_mutation,
    )
    .map_err(|error| error.to_string())
}

fn encode_segment(segment: &str) -> String {
    url::form_urlencoded::byte_serialize(segment.as_bytes()).collect()
}

fn user_path(user: Option<&str>, relationship: &str) -> String {
    match user {
        Some(user) => format!("/v1.0/users/{}/{}", encode_segment(user), relationship),
        None => format!("/v1.0/me/{relationship}"),
    }
}

fn output(command: &str, value: &Value) {
    print_json(
        &json!({"schema":moff_core::SCHEMA_VERSION,"command":command,"provider":"browser-session","content_origin":"microsoft_browser_session_untrusted","value":value}),
    );
}

#[allow(clippy::needless_pass_by_value)]
fn graph_result(
    _args: SessionArgs,
    command: &str,
    _path: String,
    _only_loop_files: bool,
) -> Result<(&'static str, Value), String> {
    Err(format!(
        "{command} is not yet available through Safari/Cider; exported-cURL credentials are no longer supported"
    ))
}

fn confirmed_json(path: &Path, confirm: bool) -> Result<Value, String> {
    if !confirm {
        return Err("mutations require --confirm".to_owned());
    }
    read_json(path)
}

fn read_json(path: &Path) -> Result<Value, String> {
    let file = fs::File::open(path).map_err(|_| "--body-file could not be read".to_owned())?;
    let mut body = String::new();
    file.take(1_000_001)
        .read_to_string(&mut body)
        .map_err(|_| "--body-file must be UTF-8".to_owned())?;
    if body.len() > 1_000_000 {
        return Err("--body-file exceeds 1MB".to_owned());
    }
    serde_json::from_str(&body).map_err(|_| "--body-file must contain JSON".to_owned())
}

fn safari_write_guard(args: &SessionArgs, mutation: bool, delete: bool) -> Result<(), String> {
    if mutation && (!args.allow_mutation || delete) {
        return Err(OutlookWebError::SafariReadOnly.to_string());
    }
    Ok(())
}

fn with_owa(
    args: SessionArgs,
    operation: impl FnOnce(&OutlookWebSession) -> Result<Value, OutlookWebError>,
) -> Result<Value, String> {
    operation(&session(args)?).map_err(|error| error.to_string())
}

#[allow(clippy::too_many_lines)]
fn mail_result(args: SessionArgs, command: MailCommand) -> Result<(&'static str, Value), String> {
    safari_write_guard(
        &args,
        matches!(
            command,
            MailCommand::Create { .. }
                | MailCommand::Update { .. }
                | MailCommand::Delete { .. }
                | MailCommand::Categories { .. }
                | MailCommand::ReplyDraft { .. }
                | MailCommand::Mark { .. }
        ),
        matches!(command, MailCommand::Delete { .. }),
    )?;
    match command {
        MailCommand::Mark {
            id,
            read,
            unread: _,
            confirm,
        } => {
            if !confirm {
                return Err("mutations require --confirm".to_owned());
            }
            Ok((
                "outlook.mail.mark",
                with_owa(args, |s| s.mail_mark(&id, read))?,
            ))
        }
        MailCommand::GetBatch { id } => Ok((
            "outlook.mail.get-batch",
            with_owa(args, |s| s.get_batch(&id, false))?,
        )),
        MailCommand::Categories {
            id,
            add,
            remove,
            confirm,
        } => {
            if !confirm {
                return Err("mutations require --confirm".into());
            }
            if add.is_empty() && remove.is_empty() {
                return Err("provide --add or --remove".into());
            }
            Ok((
                "outlook.mail.categories",
                with_owa(args, |s| s.mail_categories(&id, &add, &remove))?,
            ))
        }
        MailCommand::ReplyDraft {
            id,
            body_file,
            confirm,
        } => {
            let input: ReplyInput = serde_json::from_value(confirmed_json(&body_file, confirm)?)
                .map_err(|_| "reply input accepts a body string and optional reply_all boolean")?;
            if input.body.trim().is_empty() {
                return Err("reply body must not be empty".into());
            }
            Ok((
                "outlook.mail.reply-draft",
                with_owa(args, |s| s.reply_draft(&id, &input))?,
            ))
        }
        MailCommand::List {
            folder,
            limit,
            offset,
        } => Ok((
            "outlook.mail.list",
            with_owa(args, |s| s.mail_list(&folder, limit, offset))?,
        )),
        MailCommand::Get { id } => Ok(("outlook.mail.get", with_owa(args, |s| s.mail_get(&id))?)),
        MailCommand::Search {
            folder,
            query,
            limit,
        } => Ok((
            "outlook.mail.search",
            with_owa(args, |s| s.mail_search(&folder, &query, limit))?,
        )),
        MailCommand::Create { body_file, confirm } => {
            let input: DraftInput = serde_json::from_value(confirmed_json(&body_file, confirm)?)
                .map_err(|_| "--body-file must contain only optional subject and body strings")?;
            Ok((
                "outlook.mail.create",
                with_owa(args, |s| s.create_draft(&input))?,
            ))
        }
        MailCommand::Update {
            id,
            body_file,
            confirm,
        } => {
            let input: DraftInput = serde_json::from_value(confirmed_json(&body_file, confirm)?)
                .map_err(|_| "--body-file must contain only optional subject and body strings")?;
            Ok((
                "outlook.mail.update",
                with_owa(args, |s| s.update_draft(&id, &input))?,
            ))
        }
        MailCommand::Delete { id, confirm } => {
            if !confirm {
                return Err("mutations require --confirm".to_owned());
            }
            Ok((
                "outlook.mail.delete",
                with_owa(args, |s| s.delete_message(&id))?,
            ))
        }
    }
}

#[allow(clippy::too_many_lines)]
fn calendar_result(
    args: SessionArgs,
    command: CalendarCommand,
) -> Result<(&'static str, Value), String> {
    safari_write_guard(
        &args,
        matches!(
            command,
            CalendarCommand::Create { .. }
                | CalendarCommand::Update { .. }
                | CalendarCommand::Delete { .. }
                | CalendarCommand::WorkBlocks {
                    command: WorkBlockCommand::Apply { .. }
                }
        ),
        matches!(command, CalendarCommand::Delete { .. }),
    )?;
    match command {
        CalendarCommand::GetBatch { id } => Ok((
            "outlook.calendar.get-batch",
            with_owa(args, |s| s.get_batch(&id, true))?,
        )),
        CalendarCommand::Availability {
            folder_id,
            start,
            end,
            minutes,
            allow_overlap_title,
        } => Ok((
            "outlook.calendar.availability",
            with_owa(args, |s| {
                s.availability(&folder_id, &start, &end, minutes, &allow_overlap_title)
            })?,
        )),
        CalendarCommand::WorkBlocks {
            command: WorkBlockCommand::Plan { body_file },
        } => {
            let input: WorkInput = serde_json::from_value(read_json(&body_file)?)
                .map_err(|_| "invalid work input: expected folder_id, windows, tasks and optional allow_overlap_titles")?;
            let s = session(args)?;
            let plan = s.plan_work_blocks(&input).map_err(|e| e.to_string())?;
            Ok((
                "outlook.calendar.work-blocks.plan",
                serde_json::to_value(plan).map_err(|e| e.to_string())?,
            ))
        }
        CalendarCommand::WorkBlocks {
            command: WorkBlockCommand::Apply { body_file, confirm },
        } => {
            let plan: WorkPlan = serde_json::from_value(confirmed_json(&body_file, confirm)?)
                .map_err(|_| "invalid work plan: use value from work-blocks plan output")?;
            Ok((
                "outlook.calendar.work-blocks.apply",
                with_owa(args, |s| s.apply_work_blocks(&plan))?,
            ))
        }
        CalendarCommand::Folders => Ok((
            "outlook.calendar.folders",
            with_owa(args, |s| {
                s.call("GetCalendarFoldersV2", &calendar_folders_request())
            })?,
        )),
        CalendarCommand::View {
            folder_id,
            start,
            end,
        } => Ok((
            "outlook.calendar.view",
            with_owa(args, |s| {
                s.call(
                    "GetCalendarView",
                    &calendar_view_request(&folder_id, &start, &end),
                )
            })?,
        )),
        CalendarCommand::List {
            folder_id,
            start,
            end,
        } => Ok((
            "outlook.calendar.list",
            with_owa(args, |s| s.calendar_list(&folder_id, &start, &end))?,
        )),
        CalendarCommand::Search {
            folder_id,
            start,
            end,
            query,
            limit,
        } => Ok((
            "outlook.calendar.search",
            with_owa(args, |s| {
                s.calendar_search(&folder_id, &start, &end, &query, limit)
            })?,
        )),
        CalendarCommand::Get { id } => Ok((
            "outlook.calendar.get",
            with_owa(args, |s| s.calendar_get(&id))?,
        )),
        CalendarCommand::Create {
            folder_id,
            body_file,
            confirm,
        } => {
            let input: EventInput = serde_json::from_value(confirmed_json(&body_file, confirm)?)
                .map_err(
                    |_| "--body-file must contain only subject, body, start, and end strings",
                )?;
            Ok((
                "outlook.calendar.create",
                with_owa(args, |s| s.create_event(&folder_id, &input))?,
            ))
        }
        CalendarCommand::Update {
            id,
            body_file,
            confirm,
        } => {
            let input: EventInput = serde_json::from_value(confirmed_json(&body_file, confirm)?)
                .map_err(
                    |_| "--body-file must contain only subject, body, start, and end strings",
                )?;
            Ok((
                "outlook.calendar.update",
                with_owa(args, |s| s.update_event(&id, &input))?,
            ))
        }
        CalendarCommand::Delete { id, confirm } => {
            if !confirm {
                return Err("mutations require --confirm".to_owned());
            }
            Ok((
                "outlook.calendar.delete",
                with_owa(args, |s| s.delete_event(&id))?,
            ))
        }
    }
}

#[allow(clippy::too_many_lines)]
fn main() -> ExitCode {
    let cli = Cli::parse();
    match cli.command {
        Some(Command::Schema) => {
            print_json(
                &json!({"schema":moff_core::SCHEMA_VERSION,"provider":"browser-session","default_transport":"safari","transports":{"safari":{"experimental":true,"mutations":true,"mutation_gate":"--allow-mutation and --confirm","delete":false,"products":["mail","calendar"],"requires":"loaded authenticated Outlook tab","credential_export":false,"auth_recovery":{"refresh":true,"status":true,"interactive_sign_in":false,"confirmation":"--confirm","wait_seconds_max":120}}},"commands":{"auth":{"refresh":true,"status":true},"mail":{"list":true,"search":true,"get":true,"get-batch":true,"categories":true,"reply-draft":true,"reply_all_draft":true,"mark":true,"create":true,"update":true,"delete":false,"send":false},"calendar":{"folders":true,"view":true,"list":true,"search":true,"get":true,"get-batch":true,"availability":true,"work-blocks":{"plan":true,"apply":true},"create":true,"update":true,"delete":false},"people":{"search":false,"get":false},"org":{"manager":false,"reports":false},"onedrive":{"list":false,"get":false},"groups":{"list":false},"loop":{"list":false,"get":false},"service":{"read":true}},"unimplemented_products":["people","org","onedrive","groups","loop"],"mutations":true,"mutation_confirmation":"--confirm","calendar_mutation_scope":"single appointments without attendees","scheduling_policy":"empty gaps first; explicit exact-title overlap fallback","batch_limit":20,"search_semantics":"case-insensitive literal subject/body search within the chosen folder or calendar range"}),
            );
            ExitCode::SUCCESS
        }
        Some(
            Command::Auth {
                session: args,
                command: AuthCommand::Status,
            }
            | Command::Outlook {
                session: args,
                command:
                    OutlookCommand::Auth {
                        command: AuthCommand::Status,
                    },
            },
        ) => match OutlookWebSession::safari_auth_status(args.safari_window.zip(args.safari_tab)) {
            Ok(value) => {
                output("outlook.auth.status", &value);
                if value["authenticated"] == true {
                    ExitCode::SUCCESS
                } else {
                    ExitCode::from(2)
                }
            }
            Err(e) => error(e),
        },
        Some(
            Command::Auth {
                session: args,
                command:
                    AuthCommand::Refresh {
                        wait_seconds,
                        confirm,
                    },
            }
            | Command::Outlook {
                session: args,
                command:
                    OutlookCommand::Auth {
                        command:
                            AuthCommand::Refresh {
                                wait_seconds,
                                confirm,
                            },
                    },
            },
        ) => {
            if !confirm {
                return error("auth refresh requires --confirm; no credentials are stored");
            }
            eprintln!(
                "Refreshing the selected Outlook tab; complete any account-selection or 2FA clicks yourself in Safari."
            );
            match OutlookWebSession::refresh_safari_auth(
                args.safari_window.zip(args.safari_tab),
                wait_seconds,
            ) {
                Ok(value) => {
                    output("outlook.auth.refresh", &value);
                    if value["complete"] == false {
                        ExitCode::from(2)
                    } else {
                        ExitCode::SUCCESS
                    }
                }
                Err(e) => error(e),
            }
        }
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Service {
                    command: ServiceCommand::Read { action, body_file },
                },
        }) => {
            if let Err(validation_error) = validate_read_action(&action) {
                return error(validation_error);
            }
            let Some(body) = fs::read_to_string(body_file)
                .ok()
                .and_then(|text| serde_json::from_str(&text).ok())
            else {
                return error("--body-file must contain a JSON request body");
            };
            match session(args).and_then(|s| s.call(&action, &body).map_err(|e| e.to_string())) {
                Ok(value) => {
                    output("outlook.service.read", &value);
                    ExitCode::SUCCESS
                }
                Err(e) => error(e),
            }
        }
        Some(Command::Outlook {
            session: args,
            command: OutlookCommand::Mail { command },
        }) => match mail_result(args, command) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command: OutlookCommand::Calendar { command },
        }) => match calendar_result(args, command) {
            Ok((command, value)) => {
                output(command, &value);
                if value["complete"] == false {
                    ExitCode::from(2)
                } else {
                    ExitCode::SUCCESS
                }
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::People {
                    command: PeopleCommand::Search { limit },
                },
        }) => match graph_result(
            args,
            "outlook.people.search",
            format!("/v1.0/me/people?$top={limit}"),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::People {
                    command: PeopleCommand::Get { user },
                },
        }) => match graph_result(
            args,
            "outlook.people.get",
            format!(
                "/v1.0/users/{}?$select=id,displayName,mail,userPrincipalName,userType",
                encode_segment(&user)
            ),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Org {
                    command: OrgCommand::Manager { user },
                },
        }) => match graph_result(
            args,
            "outlook.org.manager",
            user_path(user.as_deref(), "manager"),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Org {
                    command: OrgCommand::Reports { user },
                },
        }) => match graph_result(
            args,
            "outlook.org.reports",
            user_path(user.as_deref(), "directReports"),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Onedrive {
                    command: OneDriveCommand::List { item_id, limit },
                },
        }) => {
            let base = item_id.map_or_else(
                || "root".to_owned(),
                |id| format!("items/{}", encode_segment(&id)),
            );
            match graph_result(
                args,
                "outlook.onedrive.list",
                format!("/v1.0/me/drive/{base}/children?$top={limit}"),
                false,
            ) {
                Ok((command, value)) => {
                    output(command, &value);
                    ExitCode::SUCCESS
                }
                Err(e) => error(e),
            }
        }
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Onedrive {
                    command: OneDriveCommand::Get { id },
                },
        }) => match graph_result(
            args,
            "outlook.onedrive.get",
            format!("/v1.0/me/drive/items/{}", encode_segment(&id)),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Groups {
                    command: GroupsCommand::List { limit },
                },
        }) => match graph_result(
            args,
            "outlook.groups.list",
            format!(
                "/v1.0/groups?$filter=groupTypes%2Fany%28c%3Ac%20eq%20%27Unified%27%29&$top={limit}"
            ),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Loop {
                    command: LoopCommand::List { limit },
                },
        }) => match graph_result(
            args,
            "outlook.loop.list",
            format!("/v1.0/me/drive/root/search(q='.loop')?$top={limit}"),
            true,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Loop {
                    command: LoopCommand::Get { id },
                },
        }) => match graph_result(
            args,
            "outlook.loop.get",
            format!("/v1.0/me/drive/items/{}", encode_segment(&id)),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        None => {
            print_json(&json!({"command":"moff","provider":"browser-session","status":"ready"}));
            ExitCode::SUCCESS
        }
    }
}
