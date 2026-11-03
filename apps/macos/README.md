# Moff Companion

A SwiftUI development app for the Moff Pi Durable agent over ACP stdio.
It supports saved-session reload, streamed text, read-only tool activity,
model selection, cancellation, reconnect, and new conversations.

## Run

Requires macOS 14+, full Xcode with Swift 6, Node with `node:sqlite` support
(tested with Node 26), Rust, and local Ollama with `qwen3.5:9b`.
Keep Safari signed into Outlook for mailbox access.
macOS may ask to allow Moff Companion to automate Safari when its Moff
subprocess first reads mail. The packaged app has its own permission identity,
separate from Terminal or a test runner; native chat checks do not verify that
permission grant.

From the repository root:

```sh
# Restore JavaScript dependencies if this is a fresh checkout.
npm --prefix harness ci

# Build Moff, the harness, and a local .app bundle. Does not launch it.
bash apps/macos/build.sh

open "apps/macos/MoffCompanion/.build/Moff Companion.app"
```

This bundle is for local development, not distribution: it requires the
checkout's built harness, Moff binary, Node, and dependencies. It is not
notarized and does not install an always-on background service.

Alternatively, run directly:

```sh
DEVELOPER_DIR=/Applications/Xcode.app/Contents/Developer \
  swift run --package-path apps/macos/MoffCompanion MoffCompanion
```

The app uses `~/Library/Application Support/MoffCompanion/pi-bot.sqlite` and
remembers its session ID in UserDefaults. Quit/reopen to restore the session.
CLI chat uses a different database by default. To share a database, set
`BOT_STORAGE_PATH` when launching either client; run only one owner of that
database at a time.

`MOFF_HARNESS_DIR` overrides the harness directory; `MOFF_NODE` selects the Node
executable; `MOFF_BIN` selects the MCP-enabled Moff executable. The development
default resolves the harness relative to the Swift source location at build
time. When launching from Terminal, environment overrides are inherited.

Closing the window closes the agent's stdin; the agent closes Pi and MCP.
Work interrupted by shutdown stays durable and resumes on reopen, but the
agent does not keep running while the app is closed. **Stop** explicitly aborts
the current turn instead. Only text prompts and read-only Moff operations are
supported.

## Checks

```sh
npm --prefix harness test
DEVELOPER_DIR=/Applications/Xcode.app/Contents/Developer \
  swift test --package-path apps/macos/MoffCompanion

# Real Swift client -> ACP -> Pi Durable -> Ollama round trip.
MOFF_RUN_LIVE_TESTS=1 DEVELOPER_DIR=/Applications/Xcode.app/Contents/Developer \
  swift test --package-path apps/macos/MoffCompanion \
  --filter ACPClientTests/testLiveNativeClientStreamsQwenAnswer

# Separate opt-in real-mailbox grounding and follow-up check.
npm --prefix harness run test:mail:live
```

Tests use isolated temporary databases; live tests require Ollama and should
run serially. Transport tests exercise partial frames, multiple frames, ACP
errors, real subprocess initialization, and persisted session/model reload.
Native client tests do not constitute visual UI verification.

Qwen3.5 is the default for new conversations after passing both the real-mail
summary/count follow-up and the combined two-email/Monday-calendar regression.
The combined request took about 3.5 minutes with no rejected calls; see
[model reliability limitations](../../harness/README.md#model-reliability-limitations).
The app includes the local-clock and discovered-ID Pi extensions. After rebuilding, use **Reconnect** or quit
and reopen to restart the agent with updated tools and instructions.
Saved conversations preserve their previous model; select `qwen3.5:9b` in the
model picker or create a new conversation.
