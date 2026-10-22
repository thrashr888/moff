use assert_cmd::Command;
use predicates::prelude::PredicateBooleanExt;
use predicates::str::contains;
use serde_json::Value;

fn moff() -> Command {
    Command::cargo_bin("moff").unwrap()
}

#[test]
fn send_requires_explicit_content_opt_in_and_confirmation_without_browser_access() {
    let send = [
        "mail",
        "send",
        "--to",
        "recipient@example.com",
        "--subject",
        "Synthetic",
        "--body",
        "Text",
    ];
    moff()
        .arg("outlook")
        .args(send)
        .arg("--confirm")
        .assert()
        .failure()
        .stderr(contains("--allow-mutation"));
    moff()
        .args(["outlook", "--allow-mutation"])
        .args(send)
        .assert()
        .failure()
        .stderr(contains("--confirm"));
    moff()
        .args(["outlook", "mail", "send", "--help"])
        .assert()
        .success()
        .stdout(contains("moff outlook --allow-mutation mail send").and(contains("--attachment")));
    for missing in ["--to", "--subject", "--body"] {
        let mut args = vec!["outlook", "--allow-mutation", "mail", "send", "--confirm"];
        for (flag, value) in [
            ("--to", "recipient@example.com"),
            ("--subject", "Synthetic"),
            ("--body", "Text"),
        ] {
            if flag != missing {
                args.extend([flag, value]);
            }
        }
        moff()
            .args(args)
            .assert()
            .failure()
            .stderr(contains(missing));
    }
    for flag in ["--attachments", "--from", "--sender", "--html", "--id"] {
        moff()
            .args(["outlook", "--allow-mutation"])
            .args(send)
            .args(["--confirm", flag, "unsupported"])
            .assert()
            .failure()
            .stderr(contains("unexpected argument"));
    }
    for (to, subject, body) in [
        ("Sarah", "Synthetic", "Text"),
        ("a@example.com,b@example.com", "Synthetic", "Text"),
        ("recipient@example.com", "", "Text"),
        (
            "recipient@example.com",
            "Subject\nBcc: other@example.com",
            "Text",
        ),
        ("recipient@example.com", "Synthetic", ""),
    ] {
        moff()
            .args([
                "outlook",
                "--allow-mutation",
                "mail",
                "send",
                "--to",
                to,
                "--subject",
                subject,
                "--body",
                body,
                "--confirm",
            ])
            .assert()
            .failure()
            .stderr(contains("send requires"));
    }
    moff()
        .args(["outlook", "--allow-mutation"])
        .args(send)
        .args(["--cc", "RECIPIENT@example.com", "--confirm"])
        .assert()
        .failure()
        .stderr(contains("unique explicit SMTP"));
}

#[test]
fn attachment_send_validates_files_and_gates_before_browser_access() {
    let send = [
        "mail",
        "send",
        "--to",
        "recipient@example.com",
        "--subject",
        "Synthetic",
        "--body",
        "Text",
        "--attachment",
        "/not/read/avatar.png",
    ];
    moff()
        .arg("outlook")
        .args(send)
        .arg("--confirm")
        .assert()
        .failure()
        .stderr(contains("--allow-mutation"));
    moff()
        .args(["outlook", "--allow-mutation"])
        .args(send)
        .assert()
        .failure()
        .stderr(contains("--confirm"));
    moff()
        .args(["outlook", "--allow-mutation"])
        .args(send)
        .arg("--confirm")
        .assert()
        .failure()
        .stderr(contains("local file transfer failed"));
    let mut args = vec![
        "outlook",
        "--allow-mutation",
        "mail",
        "send",
        "--to",
        "recipient@example.com",
        "--subject",
        "Synthetic",
        "--body",
        "Text",
        "--confirm",
    ];
    for _ in 0..6 {
        args.extend(["--attachment", "/not/read/avatar.png"]);
    }
    moff()
        .args(args)
        .assert()
        .failure()
        .stderr(contains("at most 5 attachments"));
}

#[test]
fn existing_draft_resume_requires_attachments_and_both_mutation_gates() {
    let send = [
        "mail",
        "send",
        "--to",
        "recipient@example.com",
        "--subject",
        "Synthetic",
        "--body",
        "Text",
        "--resume-draft",
        "synthetic-draft",
    ];
    moff()
        .args(["outlook", "--allow-mutation"])
        .args(send)
        .arg("--confirm")
        .assert()
        .failure()
        .stderr(contains("--attachment"));
    moff()
        .arg("outlook")
        .args(send)
        .args(["--attachment", "/not/read/avatar.png", "--confirm"])
        .assert()
        .failure()
        .stderr(contains("--allow-mutation"));
    moff()
        .args(["outlook", "--allow-mutation"])
        .args(send)
        .args(["--attachment", "/not/read/avatar.png"])
        .assert()
        .failure()
        .stderr(contains("--confirm"));
}

#[test]
fn safari_is_the_only_transport_and_exports_are_not_accepted() {
    moff()
        .args(["outlook", "--session-curl", "/not/read", "mail", "list"])
        .assert()
        .failure()
        .stderr(contains("unexpected argument"));
    moff()
        .args(["outlook", "--safari-window", "1", "mail", "list"])
        .assert()
        .failure()
        .stderr(contains("--safari-tab"));
    moff()
        .args(["outlook", "--safari", "mail", "list", "--help"])
        .assert()
        .success();
    moff()
        .args(["outlook", "mail", "list", "--help"])
        .assert()
        .success();
}

#[test]
fn product_commands_remain_discoverable_without_browser_access() {
    for args in [
        vec!["people", "search"],
        vec!["people", "get", "--user", "fake"],
        vec!["org", "manager"],
        vec!["org", "reports"],
        vec!["onedrive", "list"],
        vec!["onedrive", "get", "--id", "fake"],
        vec!["onedrive", "get", "--id", "fake", "--drive-id", "drive"],
        vec!["onedrive", "permissions", "--id", "fake"],
        vec!["onedrive", "share"],
        vec![
            "onedrive",
            "search",
            "--library-url",
            "https://tenant.sharepoint.com/sites/Project/Documents",
            "--query",
            "pipeline",
        ],
        vec!["groups", "list"],
        vec!["loop", "list"],
        vec!["loop", "get", "--id", "fake"],
    ] {
        moff()
            .arg("outlook")
            .args(args)
            .arg("--help")
            .assert()
            .success();
    }
}

#[test]
fn file_sharing_requires_explicit_recipient_opt_in_and_confirmation() {
    let args = [
        "onedrive",
        "share",
        "--id",
        "file",
        "--email",
        "sarah@example.com",
    ];
    moff()
        .arg("outlook")
        .args(args)
        .arg("--confirm")
        .assert()
        .failure()
        .stderr(contains("--allow-mutation"));
    moff()
        .args(["outlook", "--allow-mutation"])
        .args(args)
        .assert()
        .failure()
        .stderr(contains("--confirm"));
    moff()
        .args([
            "outlook",
            "--allow-mutation",
            "onedrive",
            "share",
            "--id",
            "file",
            "--email",
            "not-an-email",
            "--confirm",
        ])
        .assert()
        .failure()
        .stderr(contains("explicit recipient email"));
    moff()
        .args([
            "outlook",
            "onedrive",
            "get",
            "--id",
            "file",
            "--drive-id",
            "drive",
            "--output",
            "/not/write",
        ])
        .assert()
        .failure()
        .stderr(contains("cannot be used with"));
    let output = moff().arg("schema").output().unwrap();
    let schema: Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(schema["commands"]["onedrive"]["share"], true);
    assert_eq!(schema["commands"]["onedrive"]["permissions"], true);
    assert_eq!(
        schema["product_writes"]["onedrive_sharing"]["roles"],
        serde_json::json!(["read"])
    );
}

#[test]
fn scoped_library_search_rejects_invalid_urls_before_browser_access() {
    moff()
        .args([
            "outlook",
            "onedrive",
            "search",
            "--query",
            "pipeline",
            "--library-url",
            "https://evil.example/sites/Project/Documents",
        ])
        .assert()
        .failure()
        .stderr(contains("HTTPS SharePoint library"));
    moff()
        .args([
            "outlook",
            "people",
            "contacts",
            "search",
            "--query",
            "Sarah",
            "--library-url",
            "https://tenant.sharepoint.com/sites/Project/Documents",
        ])
        .assert()
        .failure()
        .stderr(contains("only supported for OneDrive"));
}

#[test]
fn mutations_require_confirmation_and_opt_in_before_browser_access() {
    for args in [
        vec![
            "mail",
            "categories",
            "--id",
            "fake",
            "--add",
            "Needs Response",
            "--confirm",
        ],
        vec![
            "mail",
            "reply-draft",
            "--id",
            "fake",
            "--body-file",
            "/not/read",
            "--confirm",
        ],
        vec!["mail", "mark", "--id", "fake", "--unread", "--confirm"],
        vec![
            "calendar",
            "work-blocks",
            "apply",
            "--body-file",
            "/not/read",
            "--confirm",
        ],
    ] {
        moff()
            .arg("outlook")
            .args(args)
            .assert()
            .failure()
            .stderr(contains("--allow-mutation"));
    }
    for product in ["mail", "calendar"] {
        moff()
            .args(["outlook", product, "delete", "--id", "fake", "--confirm"])
            .assert()
            .failure()
            .stderr(contains("--allow-mutation"));
        moff()
            .args([
                "outlook",
                "--allow-mutation",
                product,
                "update",
                "--id",
                "fake",
                "--body-file",
                "/not/read",
            ])
            .assert()
            .failure()
            .stderr(contains("--confirm"));
    }
}

#[test]
fn meeting_mutations_require_opt_in_and_confirmation() {
    for args in [
        vec![
            "calendar",
            "respond",
            "--id",
            "fake",
            "--body-file",
            "/not/read",
            "--confirm",
        ],
        vec![
            "calendar",
            "move",
            "--id",
            "fake",
            "--body-file",
            "/not/read",
            "--confirm",
        ],
        vec!["calendar", "cancel", "--id", "fake", "--confirm"],
    ] {
        moff()
            .arg("outlook")
            .args(args)
            .assert()
            .failure()
            .stderr(contains("--allow-mutation"));
    }
    for args in [
        vec![
            "outlook",
            "--allow-mutation",
            "calendar",
            "respond",
            "--id",
            "fake",
            "--body-file",
            "/not/read",
        ],
        vec![
            "outlook",
            "--allow-mutation",
            "calendar",
            "move",
            "--id",
            "fake",
            "--body-file",
            "/not/read",
        ],
        vec![
            "outlook",
            "--allow-mutation",
            "calendar",
            "cancel",
            "--id",
            "fake",
        ],
    ] {
        moff()
            .args(args)
            .assert()
            .failure()
            .stderr(contains("--confirm"));
    }
}

#[test]
fn auth_has_root_and_outlook_aliases_and_an_explicit_refresh_gate() {
    for prefix in [vec!["auth"], vec!["outlook", "auth"]] {
        moff()
            .args(&prefix)
            .arg("refresh")
            .assert()
            .failure()
            .stderr(contains("--confirm"));
        moff()
            .args(&prefix)
            .args(["refresh", "--help"])
            .assert()
            .success()
            .stdout(contains("--wait-seconds"));
        moff()
            .args(&prefix)
            .args(["status", "--help"])
            .assert()
            .success();
        moff()
            .args(&prefix)
            .args(["refresh", "--wait-seconds", "121", "--confirm"])
            .assert()
            .failure()
            .stderr(contains("121"));
    }
}

#[test]
fn schema_is_honest_about_unavailable_products_and_safe_writes() {
    let output = moff()
        .arg("schema")
        .assert()
        .success()
        .get_output()
        .stdout
        .clone();
    let schema: Value = serde_json::from_slice(&output).unwrap();
    assert_eq!(schema["default_transport"], "safari");
    assert!(schema["transports"].get("curl-export").is_none());
    assert_eq!(schema["commands"]["auth"]["refresh"], true);
    assert_eq!(schema["commands"]["auth"]["status"], true);
    assert_eq!(schema["commands"]["mail"]["mark"], true);
    assert_eq!(schema["commands"]["mail"]["reply_all_draft"], true);
    assert_eq!(schema["commands"]["mail"]["send"], true);
    assert_eq!(schema["mail_send"]["attachments"], true);
    assert_eq!(
        schema["mail_send"]["max_attachment_bytes_combined"],
        500_000
    );
    assert_eq!(schema["mail_send"]["live_send_validated"], false);
    assert_eq!(schema["commands"]["calendar"]["respond"], true);
    assert_eq!(schema["commands"]["calendar"]["move"], true);
    assert_eq!(schema["commands"]["calendar"]["cancel"], true);
    for product in ["mail", "calendar"] {
        assert_eq!(schema["commands"][product]["delete"], true);
        assert_eq!(schema["commands"][product]["get-batch"], true);
    }
    assert_eq!(schema["unimplemented_products"], serde_json::json!([]));
    assert_eq!(schema["commands"]["people"]["search"], true);
    assert_eq!(schema["commands"]["people"]["contacts"]["create"], true);
    assert_eq!(schema["commands"]["groups"]["update"], true);
    assert_eq!(schema["commands"]["onedrive"]["delete"], true);
    assert_eq!(schema["commands"]["loop"]["create"], false);
    assert_eq!(schema["commands"]["loop"]["workspace"]["create"], true);
    assert_eq!(schema["commands"]["loop"]["workspace"]["update"], false);
    assert_eq!(schema["commands"]["loop"]["workspace"]["title"], false);
    assert_eq!(schema["product_writes"]["max_file_bytes"], 250_000_000);
    assert_eq!(schema["commands"]["org"]["delete"], false);
    assert_eq!(schema["full_product_management"], false);
}

#[test]
fn native_workspaces_have_aliases_and_permanent_delete_gates() {
    for prefix in [vec!["loop"], vec!["outlook", "loop"]] {
        for verb in ["list", "get", "create", "delete", "search"] {
            moff()
                .args(&prefix)
                .args(["workspace", verb, "--help"])
                .assert()
                .success();
        }
        moff()
            .args(&prefix)
            .args(["workspace", "create", "--confirm"])
            .assert()
            .failure()
            .stderr(contains("--allow-mutation"));
        moff()
            .args(&prefix)
            .args(["workspace", "delete", "--id", "fake", "--confirm"])
            .assert()
            .failure()
            .stderr(contains("--allow-mutation"));
        moff()
            .args(&prefix)
            .args(["workspace", "delete", "--id", "fake"])
            .assert()
            .failure()
            .stderr(contains("--confirm"));
        moff()
            .args(&prefix)
            .args(["workspace", "update"])
            .assert()
            .failure()
            .stderr(contains("unrecognized subcommand"));
        moff()
            .args(&prefix)
            .args(["workspace", "create", "--title", "Test", "--confirm"])
            .assert()
            .failure()
            .stderr(contains("unexpected argument"));
    }
    for verb in ["create", "update", "get"] {
        moff()
            .args(["outlook", "onedrive", verb, "--help"])
            .assert()
            .success()
            .stdout(contains("--max-bytes"));
    }
    moff()
        .args([
            "outlook",
            "onedrive",
            "get",
            "--id",
            "fake",
            "--max-bytes",
            "250000001",
        ])
        .assert()
        .failure()
        .stderr(contains("invalid value"));
}

#[test]
fn product_management_commands_have_gates_and_validate_input_before_browser_access() {
    for product in [vec!["people", "contacts"], vec!["onedrive"], vec!["groups"]] {
        for verb in ["create", "update", "delete", "search"] {
            moff()
                .arg("outlook")
                .args(&product)
                .args([verb, "--help"])
                .assert()
                .success();
        }
        for args in [
            vec!["create", "--body-file", "/not/read", "--confirm"],
            vec![
                "update",
                "--id",
                "fake",
                "--body-file",
                "/not/read",
                "--confirm",
            ],
            vec!["delete", "--id", "fake", "--confirm"],
        ] {
            moff()
                .arg("outlook")
                .args(&product)
                .args(args)
                .assert()
                .failure()
                .stderr(contains("--allow-mutation"));
        }
        moff()
            .args(["outlook", "--allow-mutation"])
            .args(&product)
            .args(["create", "--body-file", "/dev/stdin", "--confirm"])
            .write_stdin(r#"{"owners@odata.bind":["bad"]}"#)
            .assert()
            .failure()
            .stderr(contains("unsupported field").or(contains("accepts only")));
    }
    for verb in ["update", "delete"] {
        moff()
            .args(["outlook", "loop", verb, "--id", "fake", "--confirm"])
            .args(if verb == "update" {
                vec!["--body-file", "/not/read"]
            } else {
                vec![]
            })
            .assert()
            .failure()
            .stderr(contains("--allow-mutation"));
    }
    for verb in ["add-member", "remove-member"] {
        moff()
            .args([
                "outlook",
                "groups",
                verb,
                "--id",
                "fake",
                "--user",
                "fake",
                "--confirm",
            ])
            .assert()
            .failure()
            .stderr(contains("--allow-mutation"));
    }
    moff()
        .args([
            "outlook",
            "--allow-mutation",
            "onedrive",
            "create",
            "--file",
            "/dev/null",
            "--name",
            "bad/name",
            "--confirm",
        ])
        .assert()
        .failure()
        .stderr(contains("single path segments"));
    moff()
        .args([
            "outlook",
            "--allow-mutation",
            "groups",
            "create",
            "--file",
            "/not/read",
            "--confirm",
        ])
        .assert()
        .failure()
        .stderr(contains("--file is only supported"));
}

#[test]
fn read_escape_hatch_and_inputs_are_validated_before_browser_access() {
    moff()
        .args([
            "outlook",
            "service",
            "read",
            "--action",
            "SendItem",
            "--body-file",
            "/not/read",
        ])
        .assert()
        .failure()
        .stderr(contains("not a supported read operation"));
    for args in [
        vec!["mail", "create", "--confirm", "--body-file", "/dev/stdin"],
        vec![
            "calendar",
            "create",
            "--folder-id",
            "fake",
            "--confirm",
            "--body-file",
            "/dev/stdin",
        ],
    ] {
        moff()
            .args(["outlook", "--allow-mutation"])
            .args(args)
            .write_stdin(r#"{"attendees":[]}"#)
            .assert()
            .failure()
            .stderr(contains("--body-file must contain only"));
    }
    moff()
        .args([
            "outlook",
            "mail",
            "mark",
            "--id",
            "fake",
            "--read",
            "--unread",
            "--confirm",
        ])
        .assert()
        .failure()
        .stderr(contains("cannot be used"));
    moff()
        .args(["outlook", "calendar", "list"])
        .assert()
        .failure()
        .stderr(contains("--folder-id"))
        .stderr(contains("--start"))
        .stderr(contains("--end"));
    moff()
        .args([
            "outlook", "mail", "search", "--query", "fake", "--limit", "101",
        ])
        .assert()
        .failure()
        .stderr(contains("invalid value"));
    moff()
        .args([
            "outlook",
            "service",
            "write",
            "--action",
            "SendItem",
            "--confirm",
        ])
        .assert()
        .failure()
        .stderr(contains("unrecognized subcommand"));
}
