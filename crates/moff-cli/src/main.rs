use std::{
    fs,
    io::Read,
    path::{Path, PathBuf},
    process::ExitCode,
};

use clap::{Args, Parser, Subcommand};
use moff_outlook_owa::{
    DEFAULT_MAX_FILE_BYTES, DraftInput, EventInput, EventMoveInput, LoopWorkspaceRequest,
    MeetingResponseInput, OutlookWebError, OutlookWebSession, ProductWrite, ReplyInput,
    SendMailInput, WorkInput, WorkPlan, calendar_folders_request, calendar_view_request,
    load_mail_attachments, product_segment, validate_contact_input, validate_file_size,
    validate_read_action,
};
use serde_json::{Value, json};
mod archive_mcp;
mod mcp;
mod native_mcp;

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
    /// Serve CLI operations over MCP; read-only unless --full-access.
    Mcp {
        /// Add a separately approval-gated tool for every supported CLI operation.
        #[arg(long)]
        full_access: bool,
    },
    /// Serve native Outlook tools and MCP Apps cards; read-only unless --mail-actions.
    McpNative {
        #[command(flatten)]
        session: SessionArgs,
        /// Expose typed flag, native pin, and preview-gated archive actions to the host.
        #[arg(long)]
        mail_actions: bool,
    },
    /// Serve archive discovery, preview, and approval-gated archive tools (stdio).
    McpArchive {
        #[command(flatten)]
        session: SessionArgs,
    },
    /// Native Loop and drive-backed Loop file operations in the signed-in Outlook tab.
    Loop {
        #[command(flatten)]
        session: SessionArgs,
        #[command(subcommand)]
        command: LoopRootCommand,
    },
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
        help = "Permit typed Safari writes, including explicit mail sending; each write also requires --confirm"
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
    /// Submit new plain-text mail from the selected Outlook account; acknowledgement is not delivery.
    #[command(
        after_help = "Example: moff outlook --allow-mutation mail send --to recipient@example.com --subject 'Subject' --body 'Message' --attachment /absolute/path/avatar.png --confirm\nRequires host approval through moff_action for exact files and recipients. Repeat --to/--cc/--bcc for exact addresses (20 total max). Repeat --attachment for up to 5 regular non-symlink files, 500000 bytes combined, subject to a 1000000-byte serialized browser script limit. Attachments create/verify a draft then send it; failures may leave a draft or uncertain submission. No custom sender, HTML, arbitrary draft sending or automatic retries."
    )]
    Send {
        #[arg(long, required = true)]
        to: Vec<String>,
        #[arg(long)]
        cc: Vec<String>,
        #[arg(long)]
        bcc: Vec<String>,
        #[arg(long)]
        subject: String,
        #[arg(long)]
        body: String,
        /// Local regular file to attach; repeat for up to 5 files (500000 bytes combined).
        #[arg(long)]
        attachment: Vec<PathBuf>,
        #[arg(long, required = true)]
        confirm: bool,
    },
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
        /// Continue from `next_offset` returned by the previous search page.
        #[arg(long, default_value_t = 0)]
        offset: u32,
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
    /// Move one message to Deleted Items without sending a receipt.
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
    /// Recycle one personal appointment without sending cancellations.
    Delete {
        #[arg(long)]
        id: String,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Respond to one received meeting invitation and send the RSVP.
    Respond {
        #[arg(long)]
        id: String,
        #[arg(long, value_name = "PATH")]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Move one organizer-owned meeting and send an update to its attendees.
    Move {
        #[arg(long)]
        id: String,
        #[arg(long, value_name = "PATH")]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Cancel one organizer-owned meeting and notify all attendees.
    Cancel {
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
    /// Manage writable personal contacts; directory recommendations remain separate.
    Contacts {
        #[command(subcommand)]
        command: ContactCommand,
    },
    List {
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    Search {
        #[arg(long)]
        query: Option<String>,
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
    List {
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    Get {
        #[arg(long)]
        user: String,
    },
    Search {
        #[arg(long)]
        query: String,
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
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
    /// Grant one named recipient signed-in, view-only access to an existing file.
    Share {
        #[arg(long)]
        id: String,
        /// Owning drive ID from file metadata; omit for your own `OneDrive`.
        #[arg(long)]
        drive_id: Option<String>,
        /// Exact email address resolved from people/contacts; never guess.
        #[arg(long)]
        email: String,
        /// Ask Microsoft to send its sharing invitation (delivery is not verified).
        #[arg(long)]
        notify: bool,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Inspect existing file permissions without changing access.
    Permissions {
        #[arg(long)]
        id: String,
        #[arg(long)]
        drive_id: Option<String>,
    },
    #[command(flatten)]
    Manage(ManagementCommand),
    List {
        #[arg(long)]
        item_id: Option<String>,
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    Get {
        #[arg(long)]
        id: String,
        /// Owning drive ID for a `SharePoint` or another accessible drive item.
        #[arg(long, conflicts_with = "output")]
        drive_id: Option<String>,
        /// Stream a file to a new local path; never overwrite an existing file.
        #[arg(long)]
        output: Option<PathBuf>,
        #[arg(long, default_value_t = DEFAULT_MAX_FILE_BYTES, value_parser = clap::value_parser!(u64).range(1..=250_000_000))]
        max_bytes: u64,
    },
}

#[derive(Debug, Subcommand)]
enum GroupsCommand {
    AddMember {
        #[arg(long)]
        id: String,
        #[arg(long)]
        user: String,
        #[arg(long, required = true)]
        confirm: bool,
    },
    RemoveMember {
        #[arg(long)]
        id: String,
        #[arg(long)]
        user: String,
        #[arg(long, required = true)]
        confirm: bool,
    },
    #[command(flatten)]
    Manage(ManagementCommand),
    Get {
        #[arg(long)]
        id: String,
    },
    Members {
        #[arg(long)]
        id: String,
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    List {
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
}

#[derive(Debug, Subcommand)]
enum LoopCommand {
    /// Manage native Loop workspaces; deletion is permanent.
    Workspace {
        #[command(subcommand)]
        command: LoopWorkspaceCommand,
    },
    /// Search drive-backed Loop file metadata, not full workspaces.
    Search {
        #[arg(long)]
        query: String,
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    /// Rename a drive-backed Loop file; content editing requires a Loop-specific adapter.
    Update {
        #[arg(long)]
        id: String,
        #[arg(long)]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Recycle a drive-backed Loop file.
    Delete {
        #[arg(long)]
        id: String,
        #[arg(long, required = true)]
        confirm: bool,
    },
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
enum LoopRootCommand {
    Workspace {
        #[command(subcommand)]
        command: LoopWorkspaceCommand,
    },
}

#[derive(Debug, Subcommand)]
enum LoopWorkspaceCommand {
    /// List recent workspaces, not a complete tenant inventory.
    List {
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    Get {
        /// Base64 `PodId`, not an SPO_ recent-item ID or Graph drive ID.
        #[arg(long)]
        id: String,
    },
    /// Create an unnamed tenant-owned container; naming/pages are not supported.
    Create {
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Permanently delete a workspace owned by the current user; no recycle bin.
    Delete {
        #[arg(long)]
        id: String,
        #[arg(long)]
        mru_id: Option<String>,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Filter recent workspace titles with a case-insensitive literal query.
    Search {
        query: String,
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
}

#[derive(Debug, Subcommand)]
enum ManagementCommand {
    /// Create a personal contact, drive folder or Microsoft 365 group.
    Create {
        #[arg(long, required_unless_present = "file", conflicts_with = "file")]
        body_file: Option<PathBuf>,
        #[arg(long)]
        file: Option<PathBuf>,
        #[arg(long, requires = "file", default_value_t = DEFAULT_MAX_FILE_BYTES, value_parser = clap::value_parser!(u64).range(1..=250_000_000))]
        max_bytes: u64,
        #[arg(long, requires = "file")]
        name: Option<String>,
        #[arg(long)]
        parent_id: Option<String>,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Update only the supplied writable fields.
    Update {
        #[arg(long)]
        id: String,
        #[arg(long, required_unless_present = "file", conflicts_with = "file")]
        body_file: Option<PathBuf>,
        #[arg(long)]
        file: Option<PathBuf>,
        #[arg(long, requires = "file", default_value_t = DEFAULT_MAX_FILE_BYTES, value_parser = clap::value_parser!(u64).range(1..=250_000_000))]
        max_bytes: u64,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Delete the selected object; drive items go to the recycle bin.
    Delete {
        #[arg(long)]
        id: String,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Literal metadata search, following at most 100 pages; reports incomplete results.
    Search {
        #[arg(long)]
        query: String,
        /// Scope `OneDrive` search to this `SharePoint` library root or shared-library URL.
        #[arg(long)]
        library_url: Option<String>,
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
}

#[derive(Debug, Subcommand)]
enum ContactCommand {
    #[command(flatten)]
    Manage(ManagementCommand),
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

fn schema() -> Value {
    json!({
        "schema":moff_core::SCHEMA_VERSION,"provider":"browser-session","default_transport":"safari",
        "mcp":{"command":"moff mcp","transport":"stdio","read_only":true,"tool":"moff"},
        "full_access_mcp":{"command":"moff mcp --full-access","transport":"stdio",
            "tools":["moff","moff_action"],"requires_host_approval":true,
            "scope":"all supported CLI commands; no additional backend capabilities implied"},
        "archive_mcp":{"command":"moff mcp-archive","transport":"stdio","requires_host_approval":true,
            "tools":["mail_archive_folders","mail_archive_preview","mail_archive"],
            "batch_max":20,"delete":false,"send":false},
        "command_prefixes":{"auth":["auth"],"default":["outlook"]},
        "transports":{"safari":{"experimental":true,"mutations":true,"delete":true,
            "mutation_gate":"--allow-mutation and --confirm",
            "products":["mail","calendar","people","org","onedrive","groups","loop"],
            "requires":"loaded authenticated Outlook tab; Graph operations require corresponding delegated permissions",
            "credential_export":false,
            "auth_recovery":{"refresh":true,"status":true,"interactive_sign_in":false,"confirmation":"--confirm","wait_seconds_max":120}}},
        "commands":{
            "auth":{"refresh":true,"status":true},
            "mail":{"list":true,"search":true,"get":true,"get-batch":true,"categories":true,"reply-draft":true,"reply_all_draft":true,"mark":true,"create":true,"update":true,"delete":true,"send":true},
            "calendar":{"folders":true,"view":true,"list":true,"search":true,"get":true,"get-batch":true,"availability":true,"work-blocks":{"plan":true,"apply":true},"create":true,"update":true,"delete":true,"respond":true,"move":true,"cancel":true},
            "people":{"list":true,"search":true,"get":true,"contacts":{"create":true,"get":true,"update":true,"delete":true,"list":true,"search":true}},
            "org":{"list":true,"get":true,"search":true,"manager":true,"reports":true,"create":false,"update":false,"delete":false},
            "onedrive":{"create":true,"get":true,"update":true,"delete":true,"list":true,"search":true,"upload":true,"download":true,"share":true,"permissions":true},
            "groups":{"create":true,"get":true,"update":true,"delete":true,"list":true,"search":true,"members":true,"add-member":true,"remove-member":true},
            "loop":{"create":false,"get":true,"update":true,"delete":true,"list":true,"search":true,"content_editing":false,
                "workspace":{"list":true,"get":true,"create":true,"delete":true,"search":true,"update":false,"title":false,"id":"base64 ODSP PodId","delete_policy":"permanent; verified current-user Owner required","list_scope":"recent workspaces only"}},
            "service":{"read":true}},
        "mail_send":{"command":"moff outlook --allow-mutation mail send --to recipient@example.com --subject Subject --body Message --confirm",
            "transport":"selected authenticated Outlook tab; text-only CreateItem SendAndSaveCopy; attachments CreateItem SaveOnly/read-back/SendItem",
            "requires":["explicit SMTP recipients","nonempty subject","nonempty plain-text body","--allow-mutation","--confirm","host approval for moff_action"],
            "recipient_flags":["--to","--cc","--bcc"],"max_recipients":20,"max_subject_bytes":998,"max_body_bytes":200_000,
            "sender":"selected Outlook account; no override","attachments":true,"html":false,
            "attachment_flag":"--attachment","max_attachments":5,"max_attachment_bytes_combined":500_000,
            "attachment_source":"regular non-symlink local files; snapshotted before first write; duplicate names rejected",
            "attachment_transport":"CreateItem SaveOnly with FileAttachments; verify draft/bytes; versioned SendItem to Sent Items",
            "max_serialized_browser_script_bytes":1_000_000,"attachment_failure":"no retries/cleanup; draft ID and stage reported when known",
            "send_existing_draft":false,"reply_send":false,"automatic_retry":false,
            "verification":"single successful OWA response acknowledges submission only; delivery and sent copy are not verified",
            "live_send_validated":false},
        "unimplemented_products":[],
        "full_product_management":false,
        "product_reads":{"methods":["GET"],"max_page_size":100,"follow_next_links":"metadata search only, maximum 100 pages",
            "onedrive_search_scope":"personal drive by default; --library-url resolves a SharePoint library and searches its descendants; use returned owning drive_id for reads/sharing",
            "loop_scope":"native workspace metadata through Substrate; legacy drive-backed .loop file metadata, rename and recycle","interactive_consent":false},
        "product_writes":{"methods":["POST","PATCH","DELETE","PUT"],"requires":"--allow-mutation and --confirm plus browser delegated permissions",
            "contacts_scope":"personal default contacts folder through OWA; supplied fields are read-back verified",
            "onedrive_scope":"folder/file creation, bounded binary transfers, rename/description/move, recycle and named-recipient file sharing",
            "onedrive_sharing":{"roles":["read"],"require_sign_in":true,"recipients_per_call":1,"public_links":false,"notification":"opt-in; delivery not verified","verification":"read back the returned permission ID and read role","drive_id":"optional owning drive for SharePoint files"},
            "max_file_bytes":250_000_000,"default_max_file_bytes":DEFAULT_MAX_FILE_BYTES,
            "transfer_chunk_bytes":128_000,"upload_service_chunk_bytes":3_276_800,"transfer_limit_flag":"--max-bytes",
            "groups_scope":"Microsoft 365 groups; update/delete require verified ownership",
            "verification":"Graph metadata writes use API acknowledgement; uploads check acknowledged size; native Loop deletion verifies GET 404",
            "max_json_bytes":200_000},
        "remaining_management_gaps":["Loop workspace naming and page/content creation/editing","files above 250 MB","meeting/recurrence management"],
        "mutations":true,"mutation_confirmation":"--confirm","delete_policy":"mail/calendar move to Deleted Items; OneDrive recycle; contacts/groups API deletion; native Loop workspaces permanent",
        "calendar_mutation_scope":"single appointments without attendees","scheduling_policy":"empty gaps first; explicit exact-title overlap fallback","batch_limit":20,
        "search_semantics":"mail/calendar literal subject/body; contacts/groups literal metadata; OneDrive drive search then metadata filtering; native Loop workspaces literal recent titles"
    })
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
    url::form_urlencoded::byte_serialize(segment.as_bytes())
        .collect::<String>()
        .replace('+', "%20")
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
    args: SessionArgs,
    command: &str,
    path: String,
    only_loop_files: bool,
) -> Result<(&'static str, Value), String> {
    let mut value = session(args)?
        .product_read(&path)
        .map_err(|error| error.to_string())?;
    let is_loop = |item: &Value| {
        item["name"].as_str().is_some_and(|name| {
            Path::new(name)
                .extension()
                .is_some_and(|ext| ext.eq_ignore_ascii_case("loop"))
        })
    };
    if only_loop_files {
        let items = value
            .get_mut("value")
            .and_then(Value::as_array_mut)
            .ok_or("Loop search returned invalid product metadata")?;
        items.retain(is_loop);
    } else if command == "outlook.loop.get" && !is_loop(&value) {
        return Err("the requested drive item is not a .loop file".to_owned());
    }
    let command = match command {
        "outlook.people.search" => "outlook.people.search",
        "outlook.people.get" => "outlook.people.get",
        "outlook.org.manager" => "outlook.org.manager",
        "outlook.org.reports" => "outlook.org.reports",
        "outlook.onedrive.list" => "outlook.onedrive.list",
        "outlook.onedrive.get" => "outlook.onedrive.get",
        "outlook.groups.list" => "outlook.groups.list",
        "outlook.loop.list" => "outlook.loop.list",
        "outlook.loop.get" => "outlook.loop.get",
        _ => return Err("unsupported browser product command".to_owned()),
    };
    Ok((command, value))
}

fn metadata_search(value: &mut Value, query: &str) -> Result<(), String> {
    if query.trim().is_empty() {
        return Err("search query must not be empty".to_owned());
    }
    let query = query.to_lowercase();
    let has_more = value.get("@odata.nextLink").is_some();
    let items = value["value"]
        .as_array_mut()
        .ok_or("product list returned invalid data")?;
    let scanned = items.len();
    items.retain(|item| item.to_string().to_lowercase().contains(&query));
    let count = items.len();
    value["search"] = json!({"count":count,"scanned":scanned,"complete":!has_more,
        "scope":"one metadata page","has_more":has_more});
    Ok(())
}

fn search_pages(
    session: &OutlookWebSession,
    path: &str,
    query: &str,
    limit: u16,
    loop_files: bool,
) -> Result<Value, String> {
    search_pages_with(path, query, limit, loop_files, |path| {
        session.product_read(path).map_err(|e| e.to_string())
    })
}

fn search_pages_with(
    path: &str,
    query: &str,
    limit: u16,
    loop_files: bool,
    mut read: impl FnMut(&str) -> Result<Value, String>,
) -> Result<Value, String> {
    if query.trim().is_empty() || !(1..=100).contains(&limit) {
        return Err("search requires a nonempty query and a limit from 1 to 100".to_owned());
    }
    let original = url::Url::parse(&format!("https://graph.microsoft.com{path}"))
        .map_err(|e| e.to_string())?;
    let mut path = path.to_owned();
    let mut results = Vec::new();
    let mut visited = std::collections::HashSet::new();
    let mut scanned = 0;
    for _ in 0..100 {
        if !visited.insert(path.clone()) {
            return Err("product returned a repeated continuation link".to_owned());
        }
        let mut page = read(&path)?;
        let items = page["value"]
            .as_array()
            .ok_or("product search returned invalid data")?;
        if items.iter().any(|item| {
            item["id"].as_str().is_none_or(str::is_empty)
                || (loop_files && !item["name"].is_string())
        }) {
            return Err("product search returned incomplete item metadata".to_owned());
        }
        scanned += items.len();
        if loop_files {
            page["value"]
                .as_array_mut()
                .ok_or("Loop search returned invalid data")?
                .retain(|item| {
                    item["name"]
                        .as_str()
                        .is_some_and(|s| s.to_ascii_lowercase().ends_with(".loop"))
                });
        }
        if !original.path().contains("/search(q=") || loop_files {
            metadata_search(&mut page, query)?;
        }
        results.extend(
            page["value"]
                .as_array()
                .ok_or("product search returned invalid data")?
                .iter()
                .cloned(),
        );
        let next = page
            .get("@odata.nextLink")
            .map(|value| {
                value
                    .as_str()
                    .filter(|s| !s.is_empty())
                    .ok_or("invalid product continuation")
            })
            .transpose()?;
        if results.len() >= usize::from(limit) || next.is_none() {
            let has_more = results.len() > usize::from(limit) || next.is_some();
            results.truncate(usize::from(limit));
            return Ok(
                json!({"value":results,"search":{"count":results.len(),"scanned":scanned,
                "complete":!has_more,"has_more":has_more,"scope":"metadata across bounded pages"}}),
            );
        }
        let next = url::Url::parse(next.ok_or("missing continuation")?)
            .map_err(|_| "invalid product continuation")?;
        if next.origin() != original.origin()
            || next.path() != original.path()
            || next.username() != ""
            || next.password().is_some()
            || next.fragment().is_some()
        {
            return Err("unsafe product continuation; no request followed".to_owned());
        }
        path = format!(
            "{}{}",
            next.path(),
            next.query().map_or(String::new(), |q| format!("?{q}"))
        );
    }
    Err("product search exceeded 100 pages; narrow the query instead of treating incomplete data as empty".to_owned())
}

fn library_search_with(
    library: &str,
    query: &str,
    limit: u16,
    mut read: impl FnMut(&str) -> Result<Value, String>,
) -> Result<Value, String> {
    let (site_path, library_url) =
        moff_outlook_owa::sharepoint_library_location(library).map_err(|e| e.to_string())?;
    let site = read(&site_path)?;
    let site_id = product_segment(site["id"].as_str().ok_or("site returned no ID")?)
        .map_err(|e| e.to_string())?;
    let drives = search_pages_with(
        &format!("/v1.0/sites/{site_id}/drives?$top=100"),
        &library_url,
        100,
        false,
        &mut read,
    )?;
    let candidates: Vec<_> = drives["value"]
        .as_array()
        .ok_or("invalid library list")?
        .iter()
        .filter(|drive| {
            drive["webUrl"]
                .as_str()
                .is_some_and(|value| value.trim_end_matches('/') == library_url)
        })
        .collect();
    if drives["search"]["has_more"] == true || candidates.len() != 1 {
        return Err(
            "could not resolve exactly one accessible library; no search performed".to_owned(),
        );
    }
    let drive_id = candidates[0]["id"]
        .as_str()
        .ok_or("library returned no drive ID")?;
    let path = format!(
        "/v1.0/drives/{}/root/search(q='{}')?$top={limit}",
        product_segment(drive_id).map_err(|e| e.to_string())?,
        encode_segment(&query.replace('\'', "''"))
    );
    let mut value = search_pages_with(&path, query, limit, false, read)?;
    value["search"]["library_url"] = json!(library_url);
    value["search"]["drive_id"] = json!(drive_id);
    value["search"]["scope"] =
        json!("selected SharePoint library, including descendants; Graph filename/content search");
    Ok(value)
}

#[allow(clippy::too_many_lines, clippy::needless_pass_by_value)]
fn management_result(
    args: SessionArgs,
    product: &str,
    command: ManagementCommand,
) -> Result<(&'static str, Value), String> {
    let fail = |e: OutlookWebError| e.to_string();
    if let ManagementCommand::Search {
        query,
        limit,
        library_url,
    } = command
    {
        if query.trim().is_empty() {
            return Err("search query must not be empty".to_owned());
        }
        if let Some(library) = library_url {
            if product != "onedrive" {
                return Err("--library-url is only supported for OneDrive search".to_owned());
            }
            moff_outlook_owa::sharepoint_library_location(&library).map_err(|e| e.to_string())?;
            let session = session(args)?;
            let value = library_search_with(&library, &query, limit, |path| {
                session.product_read(path).map_err(|e| e.to_string())
            })?;
            return Ok(("search", value));
        }
        if product == "contacts" {
            return Ok((
                "search",
                with_owa(args, |session| session.contacts_search(&query, limit))?,
            ));
        }
        let path = match product {
            "people" => format!("/v1.0/me/people?$top={limit}"),
            "org" => format!("/v1.0/users?$top={limit}"),
            "onedrive" => format!(
                "/v1.0/me/drive/root/search(q='{}')?$top={limit}",
                encode_segment(&query.replace('\'', "''"))
            ),
            "groups" => format!(
                "/v1.0/groups?$filter=groupTypes%2Fany%28c%3Ac%20eq%20%27Unified%27%29&$top={limit}"
            ),
            "loop" => format!("/v1.0/me/drive/root/search(q='.loop')?$top={limit}"),
            _ => return Err("unsupported product".to_owned()),
        };
        let value = search_pages(&session(args)?, &path, &query, limit, product == "loop")?;
        return Ok(("search", value));
    }
    safari_write_guard(&args, true)?;
    match &command {
        ManagementCommand::Create {
            file: Some(file),
            name,
            parent_id,
            confirm,
            max_bytes,
            ..
        } => {
            if product != "onedrive" {
                return Err("--file is only supported for OneDrive".to_owned());
            }
            if !confirm {
                return Err("mutations require --confirm".to_owned());
            }
            return Ok((
                "create",
                file_mutation(
                    args,
                    None,
                    parent_id.as_deref(),
                    name.as_deref(),
                    file,
                    *max_bytes,
                )?,
            ));
        }
        ManagementCommand::Update {
            id,
            file: Some(file),
            confirm,
            max_bytes,
            ..
        } => {
            if product != "onedrive" {
                return Err("--file is only supported for OneDrive".to_owned());
            }
            if !confirm {
                return Err("mutations require --confirm".to_owned());
            }
            return Ok((
                "update",
                file_mutation(args, Some(id), None, None, file, *max_bytes)?,
            ));
        }
        _ => (),
    }
    let (action, id, body, parent) = match command {
        ManagementCommand::Create {
            body_file,
            parent_id,
            confirm,
            ..
        } => (
            "create",
            None,
            Some(confirmed_json(
                body_file.as_deref().ok_or("--body-file required")?,
                confirm,
            )?),
            parent_id,
        ),
        ManagementCommand::Update {
            id,
            body_file,
            confirm,
            ..
        } => (
            "update",
            Some(id),
            Some(confirmed_json(
                body_file.as_deref().ok_or("--body-file required")?,
                confirm,
            )?),
            None,
        ),
        ManagementCommand::Delete { id, confirm } => {
            if !confirm {
                return Err("mutations require --confirm".to_owned());
            }
            ("delete", Some(id), None, None)
        }
        ManagementCommand::Search { .. } => unreachable!(),
    };
    let method = match action {
        "create" => "POST",
        "update" => "PATCH",
        _ => "DELETE",
    };
    if product == "contacts" {
        if parent.is_some() {
            return Err("--parent-id is only valid for OneDrive folders".to_owned());
        }
        if let Some(body) = body.as_ref() {
            validate_contact_input(body).map_err(fail)?;
        }
        let value = with_owa(args, |session| match action {
            "create" => session.contact_create(
                body.as_ref()
                    .ok_or(OutlookWebError::InvalidInput("contact body required"))?,
            ),
            "update" => session.contact_update(
                id.as_deref()
                    .ok_or(OutlookWebError::InvalidInput("contact ID required"))?,
                body.as_ref()
                    .ok_or(OutlookWebError::InvalidInput("contact body required"))?,
            ),
            _ => session.contact_delete(
                id.as_deref()
                    .ok_or(OutlookWebError::InvalidInput("contact ID required"))?,
            ),
        })?;
        return Ok((action, value));
    }
    let request = match product {
        "groups" => {
            if parent.is_some() {
                return Err("--parent-id is only valid for OneDrive folders".to_owned());
            }
            ProductWrite::group(id.as_deref(), method, body.clone()).map_err(fail)?
        }
        "onedrive" if action == "create" => {
            let value = body.as_ref().ok_or("folder input required")?;
            let object = value.as_object().ok_or("folder input must be an object")?;
            if object.len() != 1 || !object.contains_key("name") {
                return Err("folder creation accepts only a name string".to_owned());
            }
            ProductWrite::drive_folder(
                parent.as_deref(),
                value["name"]
                    .as_str()
                    .ok_or("folder name must be a string")?,
            )
            .map_err(fail)?
        }
        "onedrive" | "loop" => ProductWrite::drive_item(
            id.as_deref().ok_or("item ID required")?,
            method,
            body.clone(),
            "preflight",
        )
        .map_err(fail)?,
        _ => return Err("unsupported product mutation".to_owned()),
    };
    let session = session(args)?;
    let request = if let Some(id) = id.as_deref() {
        match product {
            "onedrive" | "loop" => {
                let before = session
                    .product_read(&format!(
                        "/v1.0/me/drive/items/{}",
                        product_segment(id).map_err(fail)?
                    ))
                    .map_err(fail)?;
                if product == "loop"
                    && (!before["name"]
                        .as_str()
                        .is_some_and(|s| s.to_ascii_lowercase().ends_with(".loop"))
                        || body.as_ref().and_then(|v| v.get("name")).is_some_and(|v| {
                            !v.as_str()
                                .is_some_and(|s| s.to_ascii_lowercase().ends_with(".loop"))
                        }))
                {
                    return Err(
                        "Loop mutations require a .loop file and must preserve its extension"
                            .to_owned(),
                    );
                }
                let etag = before["eTag"]
                    .as_str()
                    .filter(|s| !s.is_empty())
                    .ok_or("drive item returned no entity tag; no write attempted")?;
                ProductWrite::drive_item(id, method, body, etag).map_err(fail)?
            }
            _ => request,
        }
    } else {
        request
    };
    let result = session.product_write(&request).map_err(fail)?;
    if method == "POST" && result["id"].as_str().is_none_or(str::is_empty) {
        return Err(
            "creation was acknowledged but returned no ID; inspect the product before retrying"
                .to_owned(),
        );
    }
    Ok((action, result))
}

fn file_mutation(
    args: SessionArgs,
    id: Option<&str>,
    parent: Option<&str>,
    name: Option<&str>,
    file: &Path,
    max_bytes: u64,
) -> Result<Value, String> {
    let name = if id.is_none() {
        Some(
            name.or_else(|| file.file_name()?.to_str())
                .ok_or("provide a UTF-8 --name")?,
        )
    } else {
        None
    };
    if let Some(name) = name {
        product_segment(name).map_err(|e| e.to_string())?;
    }
    let mut input = fs::File::open(file).map_err(|e| format!("cannot open upload file: {e}"))?;
    let metadata = input
        .metadata()
        .map_err(|e| format!("cannot inspect upload file: {e}"))?;
    if !metadata.is_file() {
        return Err("upload requires a regular file".to_owned());
    }
    let size = metadata.len();
    validate_file_size(size, max_bytes).map_err(|e| e.to_string())?;
    let request = ProductWrite::drive_upload(parent, id, name, size, id.map(|_| "preflight"))
        .map_err(|e| e.to_string())?;
    let session = session(args)?;
    let request = if let Some(id) = id {
        let before = session
            .product_read(&format!(
                "/v1.0/me/drive/items/{}",
                product_segment(id).map_err(|e| e.to_string())?
            ))
            .map_err(|e| e.to_string())?;
        if !before["file"].is_object() {
            return Err("content replacement requires an existing file, not a folder".to_owned());
        }
        let etag = before["eTag"]
            .as_str()
            .filter(|s| !s.is_empty())
            .ok_or("file returned no entity tag")?;
        ProductWrite::drive_upload(None, Some(id), None, size, Some(etag))
            .map_err(|e| e.to_string())?
    } else {
        request
    };
    let value = session
        .drive_upload(&request, &mut input, max_bytes)
        .map_err(|e| e.to_string())?;
    if value["id"].as_str().is_none_or(str::is_empty) {
        return Err(
            "upload was acknowledged without a file ID; inspect OneDrive before retrying"
                .to_owned(),
        );
    }
    Ok(value)
}

fn download_file(args: SessionArgs, id: &str, path: &Path, max_bytes: u64) -> ExitCode {
    let result = (|| -> Result<Value, String> {
        product_segment(id).map_err(|e| e.to_string())?;
        let session = session(args)?;
        let mut file = fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(path)
            .map_err(|e| {
                format!(
                    "cannot create download destination (existing files are not overwritten): {e}"
                )
            })?;
        let count = session
            .drive_download(id, &mut file, max_bytes)
            .map_err(|e| {
                format!(
                    "download write failed; a partial file may remain at {}: {e}",
                    path.display()
                )
            })?;
        Ok(json!({"path":path.display().to_string(),"bytes":count}))
    })();
    match result {
        Ok(value) => {
            output("outlook.onedrive.get", &value);
            ExitCode::SUCCESS
        }
        Err(e) => error(e),
    }
}

fn run_management(args: SessionArgs, product: &str, command: ManagementCommand) -> ExitCode {
    match management_result(args, product, command) {
        Ok((action, value)) => {
            let product = if product == "contacts" {
                "people.contacts"
            } else {
                product
            };
            output(&format!("outlook.{product}.{action}"), &value);
            ExitCode::SUCCESS
        }
        Err(e) => error(e),
    }
}

fn run_workspace(args: SessionArgs, command: &LoopWorkspaceCommand) -> ExitCode {
    let result = (|| -> Result<(&str, Value), String> {
        let (action, request) = match command {
            LoopWorkspaceCommand::List { limit } => (
                "list",
                LoopWorkspaceRequest::List {
                    query: None,
                    limit: *limit,
                },
            ),
            LoopWorkspaceCommand::Search { query, limit } => (
                "search",
                LoopWorkspaceRequest::List {
                    query: Some(query),
                    limit: *limit,
                },
            ),
            LoopWorkspaceCommand::Get { id } => ("get", LoopWorkspaceRequest::Get { id }),
            LoopWorkspaceCommand::Create { confirm } => {
                safari_write_guard(&args, *confirm)?;
                ("create", LoopWorkspaceRequest::Create)
            }
            LoopWorkspaceCommand::Delete {
                id,
                mru_id,
                confirm,
            } => {
                safari_write_guard(&args, *confirm)?;
                (
                    "delete",
                    LoopWorkspaceRequest::Delete {
                        id,
                        mru_id: mru_id.as_deref(),
                    },
                )
            }
        };
        Ok((
            action,
            session(args)?
                .loop_workspace(&request)
                .map_err(|e| e.to_string())?,
        ))
    })();
    match result {
        Ok((action, value)) => {
            output(&format!("outlook.loop.workspace.{action}"), &value);
            ExitCode::SUCCESS
        }
        Err(e) => error(e),
    }
}

fn run_product_read(args: SessionArgs, command: &str, path: impl AsRef<str>) -> ExitCode {
    match with_owa(args, |session| session.product_read(path.as_ref())) {
        Ok(value) => {
            output(command, &value);
            ExitCode::SUCCESS
        }
        Err(e) => error(e),
    }
}

fn group_member_result(
    args: SessionArgs,
    id: &str,
    user: &str,
    add: bool,
    confirm: bool,
) -> ExitCode {
    let result = (|| -> Result<Value, String> {
        safari_write_guard(&args, true)?;
        if !confirm {
            return Err("mutations require --confirm".to_owned());
        }
        let request = ProductWrite::group_member(id, user, add).map_err(|e| e.to_string())?;
        with_owa(args, |session| session.product_write(&request))
    })();
    match result {
        Ok(value) => {
            output(
                if add {
                    "outlook.groups.add-member"
                } else {
                    "outlook.groups.remove-member"
                },
                &value,
            );
            ExitCode::SUCCESS
        }
        Err(e) => error(e),
    }
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

fn safari_write_guard(args: &SessionArgs, mutation: bool) -> Result<(), String> {
    if mutation && !args.allow_mutation {
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
                | MailCommand::Send { .. }
                | MailCommand::Update { .. }
                | MailCommand::Delete { .. }
                | MailCommand::Categories { .. }
                | MailCommand::ReplyDraft { .. }
                | MailCommand::Mark { .. }
        ),
    )?;
    match command {
        MailCommand::Send {
            to,
            cc,
            bcc,
            subject,
            body,
            attachment,
            confirm,
        } => {
            if !confirm {
                return Err("mutations require --confirm".into());
            }
            let input = SendMailInput {
                to,
                cc,
                bcc,
                subject,
                body,
            };
            input.validate().map_err(|error| error.to_string())?;
            let files = if attachment.is_empty() {
                Vec::new()
            } else {
                load_mail_attachments(&attachment).map_err(|error| error.to_string())?
            };
            Ok((
                "outlook.mail.send",
                with_owa(args, |s| {
                    if files.is_empty() {
                        s.send_mail(&input)
                    } else {
                        s.send_mail_with_attachments(&input, &files)
                    }
                })?,
            ))
        }
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
            offset,
        } => Ok((
            "outlook.mail.search",
            with_owa(args, |s| s.mail_search_page(&folder, &query, limit, offset))?,
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
                | CalendarCommand::Respond { .. }
                | CalendarCommand::Move { .. }
                | CalendarCommand::Cancel { .. }
                | CalendarCommand::WorkBlocks {
                    command: WorkBlockCommand::Apply { .. }
                }
        ),
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
        CalendarCommand::Respond {
            id,
            body_file,
            confirm,
        } => {
            let input: MeetingResponseInput =
                serde_json::from_value(confirmed_json(&body_file, confirm)?)
                    .map_err(|_| "--body-file must contain response (accept, tentative, or decline), an optional message, and optional proposed_start/proposed_end RFC3339 timestamps")?;
            Ok((
                "outlook.calendar.respond",
                with_owa(args, |s| s.respond_event(&id, &input))?,
            ))
        }
        CalendarCommand::Move {
            id,
            body_file,
            confirm,
        } => {
            let input: EventMoveInput =
                serde_json::from_value(confirmed_json(&body_file, confirm)?)
                    .map_err(|_| "--body-file must contain start and end RFC3339 strings")?;
            Ok((
                "outlook.calendar.move",
                with_owa(args, |s| s.move_event(&id, &input))?,
            ))
        }
        CalendarCommand::Cancel { id, confirm } => {
            if !confirm {
                return Err("mutations require --confirm".to_owned());
            }
            Ok((
                "outlook.calendar.cancel",
                with_owa(args, |s| s.cancel_meeting(&id))?,
            ))
        }
    }
}

#[allow(clippy::too_many_lines)]
fn main() -> ExitCode {
    let cli = Cli::parse();
    match cli.command {
        Some(Command::Mcp { full_access }) => mcp::serve(full_access),
        Some(Command::McpNative {
            session,
            mail_actions,
        }) => {
            if session.allow_mutation {
                eprintln!(
                    "Native MCP rejects --allow-mutation; use the scoped --mail-actions opt-in"
                );
                return ExitCode::FAILURE;
            }
            match native_mcp::serve_stdio(
                session.safari_window.zip(session.safari_tab),
                mail_actions,
            ) {
                Ok(()) => ExitCode::SUCCESS,
                Err(error) => {
                    eprintln!("MCP server failed: {error}");
                    ExitCode::FAILURE
                }
            }
        }
        Some(Command::McpArchive { session }) => mcp::serve_archive(session),
        Some(
            Command::Loop {
                session: args,
                command: LoopRootCommand::Workspace { command },
            }
            | Command::Outlook {
                session: args,
                command:
                    OutlookCommand::Loop {
                        command: LoopCommand::Workspace { command },
                    },
            },
        ) => run_workspace(args, &command),
        Some(Command::Schema) => {
            print_json(&schema());
            ExitCode::SUCCESS
        }
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Groups {
                    command: GroupsCommand::AddMember { id, user, confirm },
                },
        }) => group_member_result(args, &id, &user, true, confirm),
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Groups {
                    command: GroupsCommand::RemoveMember { id, user, confirm },
                },
        }) => group_member_result(args, &id, &user, false, confirm),

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
                    command:
                        PeopleCommand::Contacts {
                            command: ContactCommand::Manage(command),
                        },
                },
        }) => run_management(args, "contacts", command),
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Onedrive {
                    command: OneDriveCommand::Manage(command),
                },
        }) => run_management(args, "onedrive", command),
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Onedrive {
                    command:
                        OneDriveCommand::Share {
                            id,
                            drive_id,
                            email,
                            notify,
                            confirm,
                        },
                },
        }) => {
            let result = (|| {
                safari_write_guard(&args, confirm)?;
                ProductWrite::drive_share(drive_id.as_deref(), &id, &email, notify)
                    .map_err(|e| e.to_string())?;
                session(args)?
                    .drive_share(drive_id.as_deref(), &id, &email, notify)
                    .map_err(|e| e.to_string())
            })();
            match result {
                Ok(value) => {
                    output("outlook.onedrive.share", &value);
                    ExitCode::SUCCESS
                }
                Err(e) => error(e),
            }
        }
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Onedrive {
                    command: OneDriveCommand::Permissions { id, drive_id },
                },
        }) => match moff_outlook_owa::drive_item_path(drive_id.as_deref(), &id) {
            Ok(path) => run_product_read(
                args,
                "outlook.onedrive.permissions",
                format!("{path}/permissions"),
            ),
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Groups {
                    command: GroupsCommand::Manage(command),
                },
        }) => run_management(args, "groups", command),
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Loop {
                    command: LoopCommand::Search { query, limit },
                },
        }) => run_management(
            args,
            "loop",
            ManagementCommand::Search {
                query,
                limit,
                library_url: None,
            },
        ),
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Loop {
                    command:
                        LoopCommand::Update {
                            id,
                            body_file,
                            confirm,
                        },
                },
        }) => run_management(
            args,
            "loop",
            ManagementCommand::Update {
                id,
                body_file: Some(body_file),
                file: None,
                max_bytes: DEFAULT_MAX_FILE_BYTES,
                confirm,
            },
        ),
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Loop {
                    command: LoopCommand::Delete { id, confirm },
                },
        }) => run_management(args, "loop", ManagementCommand::Delete { id, confirm }),
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::People {
                    command:
                        PeopleCommand::Contacts {
                            command: ContactCommand::List { limit },
                        },
                },
        }) => match with_owa(args, |session| session.contacts_list(limit, 0)) {
            Ok(value) => {
                output("outlook.people.contacts.list", &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::People {
                    command:
                        PeopleCommand::Contacts {
                            command: ContactCommand::Get { id },
                        },
                },
        }) => match with_owa(args, |session| session.contact_get(&id)) {
            Ok(value) => {
                output("outlook.people.contacts.get", &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Groups {
                    command: GroupsCommand::Get { id },
                },
        }) => match product_segment(&id) {
            Ok(id) => run_product_read(args, "outlook.groups.get", format!("/v1.0/groups/{id}")),
            Err(e) => error(e.to_string()),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Groups {
                    command: GroupsCommand::Members { id, limit },
                },
        }) => match product_segment(&id) {
            Ok(id) => run_product_read(
                args,
                "outlook.groups.members",
                format!("/v1.0/groups/{id}/members?$top={limit}"),
            ),
            Err(e) => error(e.to_string()),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::People {
                    command: PeopleCommand::List { limit },
                },
        }) => run_product_read(
            args,
            "outlook.people.list",
            format!("/v1.0/me/people?$top={limit}"),
        ),
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::People {
                    command:
                        PeopleCommand::Search {
                            limit,
                            query: Some(query),
                        },
                },
        }) => run_management(
            args,
            "people",
            ManagementCommand::Search {
                query,
                limit,
                library_url: None,
            },
        ),
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::People {
                    command: PeopleCommand::Search { limit, query: None },
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
                    command: OrgCommand::List { limit },
                },
        }) => run_product_read(
            args,
            "outlook.org.list",
            format!("/v1.0/users?$top={limit}"),
        ),
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Org {
                    command: OrgCommand::Get { user },
                },
        }) => match product_segment(&user) {
            Ok(user) => run_product_read(
                args,
                "outlook.org.get",
                format!(
                    "/v1.0/users/{user}?$select=id,displayName,mail,userPrincipalName,userType"
                ),
            ),
            Err(e) => error(e.to_string()),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Org {
                    command: OrgCommand::Search { query, limit },
                },
        }) => run_management(
            args,
            "org",
            ManagementCommand::Search {
                query,
                limit,
                library_url: None,
            },
        ),
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
                    command:
                        OneDriveCommand::Get {
                            id,
                            output: Some(path),
                            max_bytes,
                            ..
                        },
                },
        }) => download_file(args, &id, &path, max_bytes),
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Onedrive {
                    command:
                        OneDriveCommand::Get {
                            id,
                            drive_id,
                            output: None,
                            ..
                        },
                },
        }) => match graph_result(
            args,
            "outlook.onedrive.get",
            match moff_outlook_owa::drive_item_path(drive_id.as_deref(), &id) {
                Ok(path) => path,
                Err(e) => return error(e),
            },
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

#[cfg(test)]
mod management_tests {
    use super::*;

    #[test]
    fn library_search_uses_the_resolved_drive_and_preserves_scope_and_paging() {
        let library = "https://example.sharepoint.com/sites/PROJ-Example/Shared%20Documents";
        let mut paths = Vec::new();
        let result = library_search_with(library, "pipeline", 1, |path| {
            paths.push(path.to_owned());
            Ok(match paths.len() {
                1 => json!({"id":"site-id"}),
                2 => json!({"value":[{"id":"drive-id","webUrl":library}]}),
                _ => json!({"value":[{"id":"pdf","name":"pipeline.pdf","parentReference":{"driveId":"drive-id"}}],
                    "@odata.nextLink":"https://graph.microsoft.com/v1.0/drives/drive-id/root/search(q='pipeline')?$skip=1"}),
            })
        }).unwrap();
        assert_eq!(
            paths[0],
            "/v1.0/sites/example.sharepoint.com:/sites/PROJ-Example"
        );
        assert_eq!(paths[1], "/v1.0/sites/site-id/drives?$top=100");
        assert_eq!(
            paths[2],
            "/v1.0/drives/drive-id/root/search(q='pipeline')?$top=1"
        );
        assert_eq!(result["search"]["drive_id"], "drive-id");
        assert_eq!(result["search"]["library_url"], library);
        assert_eq!(result["search"]["has_more"], true);
        assert_eq!(result["value"][0]["id"], "pdf");
    }

    #[test]
    fn library_search_never_falls_back_when_library_resolution_fails() {
        let library = "https://example.sharepoint.com/sites/PROJ-Example/Shared%20Documents";
        for drives in [
            json!([]),
            json!([{"id":"a","webUrl":library},{"id":"b","webUrl":library}]),
        ] {
            let mut reads = 0;
            assert!(
                library_search_with(library, "pipeline", 5, |_| {
                    reads += 1;
                    assert!(reads <= 2, "must not search an unresolved drive");
                    Ok(if reads == 1 {
                        json!({"id":"site"})
                    } else {
                        json!({"value":drives})
                    })
                })
                .is_err()
            );
        }
        assert!(
            library_search_with(library, "pipeline", 5, |_| Err("access denied".to_owned()))
                .is_err()
        );
    }

    #[test]
    fn metadata_search_follows_sparse_pages_and_reports_truncation() {
        let mut paths = Vec::new();
        let result = search_pages_with("/v1.0/me/contacts?$top=1", "project", 1, false, |path| {
            paths.push(path.to_owned());
            Ok(if paths.len() == 1 {
                json!({"value":[],"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/contacts?$top=1&$skip=1"})
            } else {
                json!({"value":[{"id":"a","displayName":"Project Alpha"},{"id":"b","displayName":"Project Beta"}]})
            })
        }).unwrap();
        assert_eq!(paths.len(), 2);
        assert_eq!(result["value"].as_array().unwrap().len(), 1);
        assert_eq!(result["search"]["has_more"], true);
        assert_eq!(result["search"]["complete"], false);
        assert_eq!(result["search"]["scanned"], 2);
    }

    #[test]
    fn metadata_search_rejects_unsafe_continuations_and_invalid_shapes() {
        for next in [
            "https://evil.example/v1.0/me/contacts?$skip=1",
            "https://graph.microsoft.com/v1.0/groups?$skip=1",
            "https://graph.microsoft.com/v1.0/me/contacts#fragment",
        ] {
            let mut calls = 0;
            let result = search_pages_with("/v1.0/me/contacts?$top=1", "project", 1, false, |_| {
                calls += 1;
                Ok(json!({"value":[],"@odata.nextLink":next}))
            });
            assert!(result.is_err());
            assert_eq!(calls, 1);
        }
        for value in [
            json!({}),
            json!({"value":[{}]}),
            json!({"value":[],"@odata.nextLink":42}),
        ] {
            assert!(
                search_pages_with("/v1.0/me/contacts?$top=1", "x", 1, false, |_| Ok(
                    value.clone()
                ))
                .is_err()
            );
        }
    }

    #[test]
    fn loop_search_is_explicitly_limited_to_loop_files() {
        let result = search_pages_with("/v1.0/me/drive/root/search(q='.loop')?$top=5","project",5,true, |_| {
            Ok(json!({"value":[{"id":"a","name":"Project.LOOP"},{"id":"b","name":"Project.txt"}]}))
        }).unwrap();
        assert_eq!(result["value"], json!([{"id":"a","name":"Project.LOOP"}]));
        assert_eq!(result["search"]["complete"], true);
    }
}
