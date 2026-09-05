use assert_cmd::Command;
use predicates::prelude::PredicateBooleanExt;
use predicates::str::contains;
use serde_json::Value;

fn moff() -> Command {
    Command::cargo_bin("moff").expect("moff binary is built")
}
fn json_stdout(args: &[&str]) -> Value {
    let output = moff()
        .args(args)
        .assert()
        .success()
        .get_output()
        .stdout
        .clone();
    serde_json::from_slice(&output).expect("stdout is JSON")
}

#[test]
fn version_is_available() {
    moff()
        .arg("--version")
        .assert()
        .success()
        .stdout("moff 0.1.0\n");
}
#[test]
fn help_lists_the_root_commands_and_global_agent_flags() {
    moff()
        .arg("--help")
        .assert()
        .success()
        .stdout(contains("auth"))
        .stdout(contains("doctor"))
        .stdout(contains("schema"))
        .stdout(contains("outlook"))
        .stdout(contains("--pretty"))
        .stdout(contains("--envelope"))
        .stdout(contains("--json"))
        .stdout(contains("--profile"))
        .stdout(contains("--dry-run"))
        .stdout(contains("--yes"))
        .stdout(contains("--no-color"));
}
#[test]
fn no_color_is_accepted_as_a_global_agent_flag() {
    moff().args(["--no-color", "doctor"]).assert().success();
}
#[test]
fn default_command_output_is_compact_json() {
    let value = json_stdout(&["doctor"]);
    assert_eq!(value["command"], "doctor");
    assert_eq!(value["status"], "local_only");
}
#[test]
fn envelope_output_has_stable_success_shape() {
    let value = json_stdout(&["--envelope", "doctor"]);
    assert_eq!(value["ok"], true);
    assert_eq!(value["data"]["command"], "doctor");
    assert!(value["warnings"].is_array());
    assert!(value["request_id"].is_string());
}
#[test]
fn incompatible_json_and_pretty_modes_fail_with_a_typed_code() {
    moff()
        .args(["--json", "--pretty", "doctor"])
        .assert()
        .failure()
        .code(2)
        .stderr(contains("incompatible_output_modes"));
}
#[test]
fn unknown_outlook_subcommand_fails_at_the_cli_boundary() {
    moff()
        .args(["outlook", "unknown"])
        .assert()
        .failure()
        .code(2)
        .stderr(contains("unrecognized subcommand"));
}
#[test]
fn schema_filters_to_the_requested_source() {
    let value = json_stdout(&["schema", "--source", "outlook"]);
    assert_eq!(value["schema"], "moff.schema/v1");
    assert_eq!(value["commands"].as_object().expect("command map").len(), 1);
    assert!(value["commands"]["outlook"].is_object());
}

#[test]
fn outlook_mail_commands_expose_guarded_dry_run_and_send_confirmation() {
    let value = json_stdout(&["--dry-run", "outlook", "mail", "list", "--folder", "inbox"]);
    assert_eq!(value["command"], "outlook.mail.list");
    moff()
        .args([
            "outlook",
            "mail",
            "send",
            "--to",
            "a@example.test",
            "--body",
            "hello",
        ])
        .assert()
        .failure()
        .stderr(contains("confirm-send"));
    moff()
        .args(["outlook", "mail", "delete", "id"])
        .assert()
        .failure()
        .stderr(contains("--yes"));
}

#[test]
fn outlook_mail_read_grammar_validates_typed_inputs_before_any_networking() {
    let value = json_stdout(&[
        "--dry-run",
        "outlook",
        "mail",
        "list",
        "--unread",
        "--from",
        "alice@example.test",
        "--since",
        "2026-09-05T12:30:00Z",
    ]);
    assert_eq!(value["command"], "outlook.mail.list");
    moff()
        .args([
            "--dry-run",
            "outlook",
            "mail",
            "list",
            "--from",
            "not-an-address",
        ])
        .assert()
        .failure()
        .stderr(contains("invalid_email"));
    moff()
        .args([
            "--dry-run",
            "outlook",
            "mail",
            "list",
            "--since",
            "yesterday",
        ])
        .assert()
        .failure()
        .stderr(contains("invalid_rfc3339"));
    moff()
        .args([
            "--dry-run",
            "outlook",
            "mail",
            "get",
            "--id",
            "immutable-id",
            "--body",
            "none",
        ])
        .assert()
        .success();
    moff()
        .args([
            "--dry-run",
            "outlook",
            "mail",
            "attachments",
            "list",
            "--id",
            "immutable-id",
        ])
        .assert()
        .success();
}

#[test]
fn compose_requires_one_utf8_body_source_and_valid_recipients() {
    for args in [
        [
            "--dry-run",
            "outlook",
            "mail",
            "send",
            "--body",
            "hello",
            "--confirm-send",
        ]
        .as_slice(),
        [
            "--dry-run",
            "outlook",
            "mail",
            "forward",
            "--id",
            "immutable-id",
            "--body",
            "hello",
            "--confirm-send",
        ]
        .as_slice(),
    ] {
        moff()
            .args(args)
            .assert()
            .failure()
            .code(2)
            .stderr(contains("recipient_required"));
    }
    moff()
        .args([
            "--dry-run",
            "outlook",
            "mail",
            "send",
            "--to",
            "not-an-address",
            "--body",
            "hello",
            "--confirm-send",
        ])
        .assert()
        .failure()
        .stderr(contains("invalid_email"));
    moff()
        .args([
            "--dry-run",
            "outlook",
            "mail",
            "send",
            "--to",
            "a@example.test",
            "--body",
            "hello",
            "--body-stdin",
            "--confirm-send",
        ])
        .assert()
        .failure()
        .stderr(contains("body_source_conflict"));
    moff()
        .args([
            "--dry-run",
            "outlook",
            "mail",
            "reply",
            "--id",
            "immutable-id",
            "--comment",
            "thanks",
            "--body",
            "hello",
            "--confirm-send",
        ])
        .assert()
        .failure()
        .stderr(contains("invalid_reply"));
}

#[test]
#[allow(clippy::too_many_lines)]
fn compose_body_sources_are_required_exclusive_and_validated_before_authentication() {
    for args in [
        [
            "--dry-run",
            "outlook",
            "mail",
            "send",
            "--to",
            "a@example.test",
            "--confirm-send",
        ]
        .as_slice(),
        [
            "--dry-run",
            "outlook",
            "mail",
            "draft-create",
            "--to",
            "a@example.test",
        ]
        .as_slice(),
    ] {
        moff()
            .args(args)
            .assert()
            .failure()
            .code(2)
            .stderr(contains("body_source_required"));
    }

    let root = std::env::temp_dir().join(format!("moff-invalid-body-{}", std::process::id()));
    std::fs::create_dir_all(&root).unwrap();
    let invalid = root.join("invalid.txt");
    std::fs::write(&invalid, [0xff, 0xfe]).unwrap();
    let oversized = root.join("oversized.txt");
    std::fs::write(&oversized, vec![b'x'; 1_048_577]).unwrap();
    for (path, code) in [
        (&invalid, "invalid_body_utf8"),
        (&oversized, "body_too_large"),
    ] {
        moff()
            .args([
                "outlook",
                "mail",
                "send",
                "--to",
                "a@example.test",
                "--body-file",
                path.to_str().unwrap(),
                "--confirm-send",
            ])
            .assert()
            .failure()
            .code(2)
            .stderr(contains(code))
            .stderr(contains("authentication").not());
    }
    let valid = root.join("valid.txt");
    std::fs::write(&valid, "file body\n").unwrap();
    moff()
        .args([
            "--dry-run",
            "outlook",
            "mail",
            "send",
            "--to",
            "a@example.test",
            "--body-file",
            valid.to_str().unwrap(),
            "--confirm-send",
        ])
        .assert()
        .success();
    moff()
        .args([
            "--dry-run",
            "outlook",
            "mail",
            "reply",
            "--id",
            "immutable-id",
            "--body-stdin",
            "--body-type",
            "html",
            "--confirm-send",
        ])
        .write_stdin("<p>stdin body</p>")
        .assert()
        .success();
    std::fs::remove_dir_all(root).unwrap();

    moff()
        .args([
            "--dry-run",
            "outlook",
            "mail",
            "reply",
            "--id",
            "immutable-id",
            "--comment",
            "thanks",
            "--body-file",
            "missing.txt",
            "--confirm-send",
        ])
        .assert()
        .failure()
        .stderr(contains("invalid_reply"));
}

#[test]
fn draft_update_requires_a_change_and_rejects_invalid_rfc3339_dates() {
    moff()
        .args([
            "--dry-run",
            "outlook",
            "mail",
            "draft-update",
            "--id",
            "immutable-id",
        ])
        .assert()
        .failure()
        .code(2)
        .stderr(contains("draft_update_required"));
    for invalid in [
        "2026-99-99T12:30:00Z",
        "2026-02-30T12:30:00Z",
        "2026-09-05T25:30:00Z",
    ] {
        moff()
            .args(["--dry-run", "outlook", "mail", "list", "--since", invalid])
            .assert()
            .failure()
            .code(2)
            .stderr(contains("invalid_rfc3339"));
    }
}

#[test]
fn draft_update_and_forward_require_the_id_flag() {
    for args in [
        [
            "--dry-run",
            "outlook",
            "mail",
            "draft-update",
            "immutable-id",
            "--subject",
            "updated",
        ]
        .as_slice(),
        [
            "--dry-run",
            "outlook",
            "mail",
            "forward",
            "immutable-id",
            "--to",
            "a@example.test",
            "--body",
            "hello",
            "--confirm-send",
        ]
        .as_slice(),
    ] {
        moff()
            .args(args)
            .assert()
            .failure()
            .code(2)
            .stderr(contains("unexpected argument"));
    }

    for args in [
        [
            "--dry-run",
            "outlook",
            "mail",
            "draft-update",
            "--id",
            "immutable-id",
            "--subject",
            "updated",
        ]
        .as_slice(),
        [
            "--dry-run",
            "outlook",
            "mail",
            "forward",
            "--id",
            "immutable-id",
            "--to",
            "a@example.test",
            "--body",
            "hello",
            "--confirm-send",
        ]
        .as_slice(),
    ] {
        moff()
            .args(args)
            .assert()
            .success()
            .stdout(contains("network_mutation").and(contains("false")));
    }
}

#[test]
fn sync_reset_requires_explicit_local_state_confirmation_and_supports_limit() {
    moff()
        .args(["outlook", "mail", "sync", "--reset"])
        .assert()
        .failure()
        .stderr(contains("--yes"));
    json_stdout(&[
        "--dry-run",
        "--yes",
        "outlook",
        "mail",
        "sync",
        "--reset",
        "--limit",
        "1000",
    ]);
    moff()
        .args(["--dry-run", "outlook", "mail", "sync", "--limit", "1001"])
        .assert()
        .failure()
        .stderr(contains("invalid_limit"));
}

#[test]
fn auth_login_rejects_invalid_client_id_before_networking() {
    moff()
        .args(["auth", "login", "--client-id", "not-a-uuid"])
        .assert()
        .failure()
        .stderr(contains("invalid UUID"));
}
#[test]
fn auth_login_dry_run_prints_a_safe_plan_without_provider_calls() {
    moff()
        .args([
            "--dry-run",
            "auth",
            "login",
            "--client-id",
            "11111111-1111-1111-1111-111111111111",
            "--scope",
            "mail.read",
        ])
        .assert()
        .success()
        .stdout(contains("device_authorization"))
        .stdout(
            predicates::str::is_match("access_token|refresh_token")
                .unwrap()
                .not(),
        );
}
#[test]
fn auth_login_mail_write_dry_run_requests_mail_read_write() {
    let value = json_stdout(&[
        "--dry-run",
        "auth",
        "login",
        "--client-id",
        "11111111-1111-1111-1111-111111111111",
        "--scope",
        "mail.write",
    ]);
    assert_eq!(value["scopes"][0], "Mail.ReadWrite");
}
#[test]
fn invalid_profile_is_rejected_by_prompt_free_auth_status() {
    moff()
        .args(["--profile", "../bad", "auth", "status"])
        .assert()
        .failure()
        .stderr(contains("invalid profile name"));
}
#[test]
fn status_logout_and_doctor_are_prompt_free_and_safe() {
    for args in [
        ["auth", "status"].as_slice(),
        ["auth", "logout"].as_slice(),
        ["doctor"].as_slice(),
    ] {
        moff().args(args).assert().success().stdout(
            predicates::str::is_match("access_token|refresh_token|device_code")
                .unwrap()
                .not(),
        );
    }
}

#[test]
fn persisted_profile_preflights_typed_scope_before_touching_keychain_or_graph() {
    let root = std::env::temp_dir().join(format!("moff-cli-contract-{}", std::process::id()));
    std::fs::create_dir_all(root.join("profiles")).unwrap();
    std::fs::write(root.join("profiles/work.json"), r#"{"name":"work","client_id":"11111111-1111-1111-1111-111111111111","tenant":"common","scopes":["Mail.Read"],"account_display_name":null}"#).unwrap();
    moff()
        .env("MOFF_DATA_DIR", &root)
        .args([
            "--profile",
            "work",
            "outlook",
            "mail",
            "send",
            "--to",
            "a@example.test",
            "--body",
            "hello",
            "--confirm-send",
        ])
        .assert()
        .failure()
        .code(2)
        .stderr(contains("consent_required"))
        .stderr(contains("Mail.Send"));
    std::fs::remove_dir_all(root).unwrap();
}

#[test]
fn dry_run_dispatches_list_get_and_send_without_credentials_or_mutation() {
    for args in [
        ["--dry-run", "outlook", "mail", "list"].as_slice(),
        ["--dry-run", "outlook", "mail", "get", "--id", "id"].as_slice(),
        [
            "--dry-run",
            "outlook",
            "mail",
            "send",
            "--to",
            "a@example.test",
            "--body",
            "hello",
            "--confirm-send",
        ]
        .as_slice(),
    ] {
        moff()
            .args(args)
            .assert()
            .success()
            .stdout(contains("network_mutation").and(contains("false")));
    }
}
