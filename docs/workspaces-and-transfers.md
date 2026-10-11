# Native Loop workspaces and bounded OneDrive transfers

All operations use a loaded, signed-in Outlook Safari document: Mail/Calendar,
or `/host/<guid>/auth` with `webpackChunkOwa` present. Hosted routes are probed
read-only before use; a same-origin URL alone does not establish a valid runtime.
Moff pins Cider's `DocumentTarget`: reordering tabs is safe, but reload, closure,
origin change or moving the document to another window invalidates the operation.
There is no automatic navigation, sign-in, consent, token export or write retry.

## Loop workspace contract

`moff loop workspace` and `moff outlook loop workspace` expose the same commands:

| Command | Request | Scope |
| --- | --- | --- |
| `list --limit N` | GET `/recommended/api/v1.1/loop/recent` | Recent workspace metadata; `is_complete` also reflects the local limit |
| `search <query> --limit N` | Same GET, literal case-insensitive title filter | Recent workspaces only; not service-wide search |
| `get --id <PodId>` | GET `/recommended/api/v1.1/loop/workspaces/{encodedPodId}` | Workspace metadata and permissions |
| `create --confirm` | POST `/recommended/api/v1.1/loop/workspaces`, `{"type":"tenantOwned"}` | Unnamed tenant-owned container |
| `delete --id <PodId> --confirm` | DELETE `/recommended/api/v1.1/loop/workspaces/{encodedPodId}`, `{}` or `{"mru_id":"<guid>"}` | Permanent deletion, not recycling |

Requests run on `https://substrate.office.com` with Outlook's silent token helper,
`Authorization`, `X-AnchorMailbox: Oid:<oid>@<tid>`, a unique `Client-Request-Id`,
and JSON content type when a body is present. Token claims and headers remain
inside the page. Both write commands additionally require `--allow-mutation`
before `workspace`.

Workspace lookup uses the canonical PodId:
`btoa("ODSP|<lowercase-site-domain>|<driveId b!...>|<rootItemId>")`.
List/get normalize `id` to `mfs_info.pod_id` and preserve the original identifier
as `source_id`. Raw `SPO_...` recent IDs and drive IDs are not lookup IDs.

Creation returns a drive ID and SharePoint site URL. Moff reads
`https://graph.microsoft.com/v1.0/drives/{driveId}/root?$select=id` with an
in-page Graph token and returns the derived PodId. If creation succeeds but
derivation fails, Moff reports verification failure **with the created drive ID**;
do not automatically create another container.

The container's initial title is a GUID and it does not necessarily appear in
`recent`. Workspace naming/renaming needs the pod/Fluid file that the Loop app
creates separately. Neither `update` nor `--title` is supported; `--title` is a
future follow-up, not silently ignored input. Pages and Fluid content are not
created, read or edited by these commands.

Deletion first GETs the exact PodId and requires an Owner permission whose UPN
matches the current token's user identity. Unknown/missing ownership refuses the
write. The acknowledgement must contain `response_code:10010`; a subsequent
workspace GET must return 404. Any uncertain acknowledgement or verification
failure reports unknown outcome and stops, without a legacy `?id=` fallback.
Existing `outlook loop delete` still recycles a drive-backed `.loop` file: it is
not workspace deletion.

## OneDrive transfer contract

The file-size ceiling is 250000000 bytes by default and at most 250000000.
`--max-bytes` can lower the ceiling for `create --file`, `update --file` and
`get --output`. The ceiling is not an expected length. Regular-file upload size
and Graph download item size are checked independently; final length must match.

Rust reads/writes at most 128000 bytes per Cider operation. Downloads resolve
Graph's content location inside the page, validate its Microsoft HTTPS host,
then request at most 3276800 bytes per HTTP range and incrementally consume each
`ReadableStream`. This keeps Safari's upstream buffering bounded even during
slow Apple Events transport. Retained upstream chunks are capped at 3276800
bytes and split to the transport size; invalid/ignored ranges and oversized chunks fail
rather than invoking `arrayBuffer()` or buffering the file. Destination creation
is exclusive. A failed transfer can leave a partial local file and reports its
path; it never overwrites an existing file.

Nonempty uploads create a Graph upload session. New-file conflict behavior is
`fail`; replacement requires preflight file metadata and the current entity tag
on session creation. The page accumulates at most 3276800 bytes (ten 320 KiB
units), sends PUT ranges with `Content-Range`, and accepts a single exact next
expected offset. Only the final range may be smaller than 320 KiB. Zero-byte
uploads use a bounded empty PUT. Final acknowledgement must identify the file
and match the local size; only sanitized metadata is returned.

Upload-session URLs and download URLs never leave the page. Preauthenticated
content/range requests do not carry Graph authorization headers. Cider options
use a 60-second per-operation timeout and 120-second idle timeout; no more than
four live transfers are permitted per pinned document. Finish/abort release
resources. After adapter error or timeout, Cider invalidates the handle and
requests cleanup; lost document identity can make cleanup unknown.

Upload errors can occur after a server write. Moff never retries a range, resumes
from a guessed offset, repeats session creation or silently reports success.
Inspect OneDrive to determine the outcome before retrying. An abandoned upload
session may remain until service expiry; no cleanup DELETE is sent after an
ambiguous write.

The underlying transport contract is
[Cider v0.12.0 binary transfers](https://github.com/thrashr888/cider/blob/v0.12.0/docs/safari-binary-transfers.md).
