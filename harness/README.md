# Moff Persistent Bot Harness

An experimental assistant running on [Pi Durable](https://earendil.com/posts/pi-durable/),
using **local Ollama** by default. Pi owns the agent loop, durable submissions,
tool tasks, checkpoints, compaction, and recovery. The model adapter retains
the tested Ollama streaming and repetition safeguards.

## Architecture

- **Model access**: Ollama's OpenAI-compatible API at `http://127.0.0.1:11434`. The experimental Copilot provider is blocked by HTTP 403 in this environment.
- **Default model**: `qwen3.5:9b`, selected after passing both real-mailbox
  regressions on October 19, 2026. Qwen-family thinking is disabled via
  `reasoning_effort: "none"` because the installed model returned reasoning-only
  responses without executing tool calls in thinking mode. Empty responses
  now fail explicitly. Ollama can also discard malformed model-generated tool
  arguments and return an empty response; the error identifies both possibilities.
  Llama 3.2 remains selectable but failed the live
  mailbox follow-up count check; it is not the recommended default.
- **Storage**: Pi Durable's Node SQLite backend persists committed state in
  `pi-bot.sqlite`. The CLI reopens the same root conversation each launch.
  Pending generations resume after a restart; committed tools are not run again.
  Read-only MCP tools are declared replay-safe: an interrupted, uncommitted read
  can run again and may see newer data.
- **Conversations**: ACP clients can create, list, and reload separate sessions.
  Request IDs on the harness API deduplicate retried submissions.
- **Tool execution**: Pi runs tools sequentially. Native Moff's stdio server is
  synchronous and browser-backed; parallel model calls must not queue overlapping
  MCP requests that can time out while waiting for another read. Tool failures
  remain explicitly labeled as failures in subsequent model context.
- **Protocol**: The official ACP SDK handles initialization, session creation,
  loading/history replay, listing, model configuration, text streaming, tool
  status, cancellation, and protocol errors. Unsupported content is rejected.
- **macOS**: The SwiftUI development app uses the same ACP agent; see
  [native app instructions](../apps/macos/README.md).

This is not yet an always-on service: the scheduler works while its process is
running, and resumes saved work when reopened. There is no launch agent,
proactive mail polling, reminder extension, or write approval workflow yet.
The Pi Durable API is experimental.

The old custom `bot-storage.sqlite` and JSON fallback files are **not migrated
or deleted**. Use a new Pi database path; do not point `BOT_STORAGE_PATH` at a
legacy database. The unused legacy storage module remains only for reference.

## Quickstart

```sh
# Run from the harness directory. Requires Ollama and an installed model.
# Rebuilds automatically and explicitly selects Ollama, even if BOT_PROVIDER=copilot.
npm run chat
npm run chat -- --model qwen3.5:9b

# Raw ACP stdio for editors or the macOS companion (no npm banners on stdout).
npm run build
node dist/cli.js acp

# Live CLI test using a temporary database (requires local qwen3.5:9b).
npm run test:live

# Opt-in real mailbox test: reads mail through MCP, checks source record counts
# and ISO timestamps, then asks a follow-up. Never logs message contents.
npm run test:mail:live

# Exact reported two-email/Monday-calendar regression.
npm run test:combined:live
```

`npm test` covers message compatibility, streaming/repetition regressions,
native MCP discovery, ACP lifecycle/cancellation, persisted context,
request deduplication, and a forced process crash/restart after a committed MCP
read, without reading the mailbox. It also exercises the local clock across
daylight-saving boundaries and the evidence hook across process restarts and
separate conversations.
The live mailbox test additionally checks that enumerated subjects come from
the returned emails and that a second turn reports the actual record count.
Its eight-minute test deadline accommodates local inference; it is not an
answer-length or runtime generation limit. Run model-dependent live tests
serially to avoid competing for Ollama resources.

Exit any previously running bot with `exit` or Ctrl+C before relaunching.
`OLLAMA_HOST` changes the API address; `BOT_MODEL` changes the default model;
`BOT_STORAGE_PATH` changes the database path.
`npm --prefix harness run start` from the repository root also starts chat.
From the harness directory, use `node dist/cli.js status --ollama` to list
persisted conversations.

## Extensions and clients

The bot chose **Cipher**, a **Ranger (Gloom Stalker)**, after reconsidering
party balance and its unique role as Paul's work-environment agent. Its avatar
is a shadow panther with amber eyes and midnight-blue (`#191970`) coloring; the original drawing is
[`assets/cipher.svg`](assets/cipher.svg). The `party-identity` Pi extension keeps
that identity in its context. Moff remains the name of the Microsoft 365 CLI.

The [pi-bundle Cipher profile](https://github.com/thrashr888/pi-bundle/tree/main/machines/cipher)
registers that identity for Tavern's standard host. Tavern uses Pi RPC rather
than this harness's ACP protocol; the profile starts a separate session and does
not migrate or share Companion's database. Registration is not a claim that
the bundle has been installed or that this ACP process is attached to Tavern.

Moff's MCP bridge is installed as a real Pi Durable extension, with discovered
schemas and system-prompt sections. Pi's automatic compaction is enabled.
The `local-context` extension exposes `current_time`: actual local time, timezone,
and next-occurrence weekday ranges with exclusive next-midnight boundaries.
For a weekday named today, these ranges select the following week; the assistant
is instructed to state the resolved date and clarify ambiguity.

The `moff-evidence-guard` Pi tool hook blocks opaque `--id` / `--*-id` values until
they have appeared in successful Moff JSON results in that conversation. This
prevents invented placeholders such as `primary` from reaching the CLI. Discover
folders/records before reading them by ID, including IDs supplied directly in a
prompt. Discovered IDs survive reopening; new conversations and forks must
rediscover them. This is a grounding aid, not an authorization boundary; native
Moff still enforces read-only access.

The next useful extensions are durable reminders/background tasks, explicit
approval hooks before any future writes, and deliberately scoped skills/context.
Do not install the broad coding-tools extension for a mailbox-only assistant:
it grants shell and filesystem capabilities this bot does not need.

Pi's TUI is a potential terminal client/toolkit, not a durability layer.
Coding-agent extensions should not be assumed compatible with Pi Durable's
extension registry. A richer TUI can consume the same committed events as ACP;
it is not installed or implemented here.

Local Ollama chat connects to `moff mcp` using the MCP client SDK. Tools,
input schemas, server instructions and `moff schema` are discovered on startup;
the harness contains no per-operation mail wrappers or copied pagination defaults.
Build the native server from the repository root with `cargo build -p moff-cli`.
The harness prefers the checkout's `target/debug/moff`, then Moff on PATH;
`MOFF_BIN` explicitly overrides either. Keep Safari signed into Outlook and ask:
`Can you read my latest emails from moff?`

Mail results are passed to the local model and saved in the local transcript.
Sending, marking, deleting, arbitrary shell commands, and other mailbox writes
are not available. Mail tools are disabled for remote Ollama URLs and Copilot
by default. If Moff reports an authentication error, sign into Outlook in
Safari and retry.

The discovered generic `moff` tool takes CLI arguments, for example
`{"args":["outlook","mail","list"]}`. Moff owns argument parsing, defaults,
validation and read-only policy; command `--help` supplies option details.
Startup fails explicitly if the selected binary lacks MCP support.
The server instructions also distinguish the returned page from
`TotalItemsInView`, the collection-wide total. That total must not be used as
the number of emails returned by one call.

After changing the server, rebuild it with Cargo; `npm run chat` rebuilds
the harness. MCP SDK discovery/defaults tests run with:
`npm run build && node --test tests/moff-mcp.test.mjs`.

The harness does not impose an answer-length or email-count cap. Six consecutive
copies of a substantial text block or identical list entry stop a stuck stream
with an explicit incomplete-answer error. Numbering is ignored, but dates and
IDs are preserved to avoid conflating similar records. Repeated entries whose
only differences are timestamps are also stopped when at least three of the
last six entries contain timestamps absent from the tool results. Similar
records with source-supported timestamps are allowed. This is a narrow loop
heuristic, not a general factuality guarantee; varied hallucinations can still
escape it. MCP results remain unchanged.

## Model reliability limitations

The reported request, `what's my latest 2 emails, and my calendar for monday?`,
now has an opt-in regression that checks successful reads, both source email
subjects, folder discovery, and the exact next-Monday midnight-to-midnight range.
Email and calendar records are checked separately; blocked tool attempts are
allowed only if the model subsequently completes the required reads.

After restoring Safari authentication, `qwen3:latest` passed the real-mailbox
scenario on October 19, 2026: both source email subjects, a discovered calendar
folder, and the exact upcoming-Monday range were verified. It took 346 seconds
and seven calls, recovering from three rejected calendar attempts. Earlier runs
failed or omitted subjects; one passing run does not establish general reliability.
The installed Granite 4.1 8B recovered from an invalid
ID and successfully read the calendar, but still failed the subject check and
took roughly six minutes. A later Qwen3 mail-only run failed the subject check.

Use `TEST_MODEL=<installed-model> npm run test:combined:live` for a comparison
without changing the app's selected model. The test does not log mailbox content.
With Safari restored and tool execution serialized, `qwen3.5:9b` passed the
mail summary/count follow-up in 84 seconds and the combined request in 207
seconds, using four successful calls and no rejected attempts. It is now the
default for new conversations. Existing saved conversations preserve their
selected model: use the app's model picker to select `qwen3.5:9b`, or start a
new conversation. These passes are evidence of improvement, not a guarantee
of correct answers on every request.

Earlier Qwen3.5 workflow runs timed out, including a synthetic-source run.
Enabling Qwen3
thinking or constraining intents to JSON also did not establish a passing fix;
the experimental JSON transport was removed.

Direct native mail reads also failed with `Safari operation failed` during the
investigation, independently of inference. Authentication was restored before
the passing run. If this recurs, check `moff auth status` before comparing models;
an answer describing that failure is not a successful mailbox test.

For a separate, explicitly synthetic-source check using real model inference,
run from the repository root:

```sh
MOFF_BIN="$PWD/harness/tests/fixtures/combined-mcp.mjs" \
  TEST_MODEL=qwen3.5:9b npm --prefix harness run test:combined:live
```

That fixture never reads the mailbox and does not substitute for live validation.
The live-test deadline aborts the conversation; it is not a runtime output cap.
