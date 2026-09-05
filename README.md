# Moff

Moff is an **agent-first Microsoft 365 CLI**. This initial foundation is a safe
scaffold: it defines stable command and JSON contracts but deliberately performs
no authentication, Microsoft Graph, or other provider networking.

## Commands

The executable exposes `auth`, `doctor`, `schema`, and `outlook`. Use the CLI as
the authoritative command inventory:

```sh
cargo run -p moff-cli -- --help
cargo run -p moff-cli -- schema
cargo run -p moff-cli -- schema --source outlook
```

Command results are compact JSON on stdout by default. `--envelope` produces the
stable `{ok,data,warnings,request_id}` agent response object. Diagnostics go to
stderr, and invalid global mode combinations exit nonzero with a typed code.

## Build and test

Requires Rust 1.98 or newer (edition 2024).

```sh
cargo fmt --check
cargo clippy --workspace --all-targets -- -D warnings
cargo test --workspace
cargo audit
```

`cargo audit` requires the separately installed `cargo-audit` subcommand.

## License

MIT
