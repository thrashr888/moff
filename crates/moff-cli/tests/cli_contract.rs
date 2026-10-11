use assert_cmd::Command;
use predicates::str::contains;
use serde_json::Value;

fn moff() -> Command {
    Command::cargo_bin("moff").unwrap()
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
            .args([
                "outlook",
                "--allow-mutation",
                product,
                "delete",
                "--id",
                "fake",
                "--confirm",
            ])
            .assert()
            .failure()
            .stderr(contains("deletion is not supported"));
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
    assert_eq!(schema["commands"]["mail"]["send"], false);
    for product in ["mail", "calendar"] {
        assert_eq!(schema["commands"][product]["delete"], false);
        assert_eq!(schema["commands"][product]["get-batch"], true);
    }
    assert!(
        schema["unimplemented_products"]
            .as_array()
            .unwrap()
            .is_empty()
    );
    assert_eq!(schema["commands"]["people"]["search"], true);
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
    for args in [
        vec!["mail", "send", "--confirm"],
        vec!["service", "write", "--action", "SendItem", "--confirm"],
    ] {
        moff()
            .arg("outlook")
            .args(args)
            .assert()
            .failure()
            .stderr(contains("unrecognized subcommand"));
    }
}
