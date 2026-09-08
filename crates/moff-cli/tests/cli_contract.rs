use assert_cmd::Command;
use predicates::prelude::PredicateBooleanExt;
use predicates::str::contains;
use serde_json::Value;

fn moff() -> Command {
    Command::cargo_bin("moff").unwrap()
}

#[test]
fn help_exposes_all_browser_session_read_products() {
    moff()
        .arg("--help")
        .assert()
        .success()
        .stdout(contains("outlook"));
    moff()
        .args(["outlook", "--help"])
        .assert()
        .success()
        .stdout(contains("mail"))
        .stdout(contains("calendar"))
        .stdout(contains("people"))
        .stdout(contains("org"))
        .stdout(contains("onedrive"))
        .stdout(contains("groups"))
        .stdout(contains("loop"))
        .stdout(contains("service"))
        .stdout(contains("Graph"))
        .stdout(contains("doctor").not());
}

#[test]
fn mail_list_uses_the_default_session_export_path() {
    let home = std::env::temp_dir().join(format!("moff-home-{}", std::process::id()));
    std::fs::create_dir_all(&home).unwrap();
    moff()
        .env("HOME", &home)
        .args(["outlook", "mail", "list"])
        .assert()
        .failure()
        .code(2)
        .stderr(contains("session export could not be read"));
    std::fs::remove_dir_all(home).unwrap();
}

#[test]
fn schema_names_the_one_browser_session_provider() {
    let output = moff()
        .arg("schema")
        .assert()
        .success()
        .get_output()
        .stdout
        .clone();
    let value: Value = serde_json::from_slice(&output).unwrap();
    assert_eq!(value["schema"], "moff.schema/v1");
    assert_eq!(value["provider"], "browser-session");
    assert!(value["commands"]["mail"]["list"].as_bool().unwrap());
    assert!(value["commands"]["calendar"]["view"].as_bool().unwrap());
    assert!(value["commands"]["people"]["search"].as_bool().unwrap());
    assert!(value["commands"]["org"]["manager"].as_bool().unwrap());
    assert!(value["commands"]["onedrive"]["list"].as_bool().unwrap());
    assert!(value["commands"]["groups"]["list"].as_bool().unwrap());
    assert!(value["commands"]["loop"]["list"].as_bool().unwrap());
    assert!(value["commands"]["service"]["read"].as_bool().unwrap());
    assert!(
        value["unimplemented_products"]
            .as_array()
            .unwrap()
            .is_empty()
    );
}

#[test]
fn non_readonly_service_action_is_rejected_before_session_access() {
    moff()
        .args([
            "outlook",
            "service",
            "read",
            "--action",
            "SendItem",
            "--body-file",
            "/not/read.json",
        ])
        .assert()
        .failure()
        .code(2)
        .stderr(contains("not a supported read operation"));
}

#[test]
fn every_mutation_requires_confirmation_before_files_or_sessions() {
    for product in ["mail", "calendar"] {
        for operation in ["create", "update", "delete"] {
            let mut args = vec!["outlook", product, operation];
            if operation != "create" {
                args.extend(["--id", "synthetic-item"]);
            }
            if operation != "delete" {
                args.extend(["--body-file", "/not/read.json"]);
            }
            moff()
                .env("HOME", "/definitely/missing/moff-contract-home")
                .args(args)
                .assert()
                .failure()
                .code(2)
                .stderr(contains("--confirm"))
                .stderr(contains("session export").not());
        }
    }
}

#[test]
fn outlook_search_and_calendar_get_do_not_open_the_graph_export() {
    let home = std::env::temp_dir().join(format!("moff-routing-{}", std::process::id()));
    let config = home.join(".config/moff");
    std::fs::create_dir_all(&config).unwrap();
    std::fs::write(
        config.join("graph-session.curl"),
        "invalid synthetic export",
    )
    .unwrap();
    for args in [
        vec!["outlook", "mail", "search", "--query", "synthetic"],
        vec!["outlook", "calendar", "get", "--id", "synthetic-item"],
    ] {
        moff()
            .env("HOME", &home)
            .args(args)
            .assert()
            .failure()
            .code(2)
            .stderr(contains("session export could not be read"));
    }
    std::fs::remove_dir_all(home).unwrap();
}

#[test]
fn schema_advertises_confirmed_crud_without_send() {
    let output = moff()
        .arg("schema")
        .assert()
        .success()
        .get_output()
        .stdout
        .clone();
    let schema: Value = serde_json::from_slice(&output).unwrap();
    for product in ["mail", "calendar"] {
        for operation in ["list", "search", "get", "create", "update", "delete"] {
            assert_eq!(schema["commands"][product][operation], true);
        }
    }
    assert_eq!(schema["mutations"], true);
    assert_eq!(schema["mutation_confirmation"], "--confirm");
    assert_eq!(schema["commands"]["mail"]["send"], false);
}

#[test]
fn calendar_listing_requires_an_explicit_folder_and_range() {
    moff()
        .args(["outlook", "calendar", "list"])
        .assert()
        .failure()
        .code(2)
        .stderr(contains("--folder-id"))
        .stderr(contains("--start"))
        .stderr(contains("--end"));
}

#[test]
fn mail_search_rejects_invalid_limits_before_session_access() {
    for limit in ["0", "101"] {
        moff()
            .args([
                "outlook",
                "mail",
                "search",
                "--query",
                "synthetic",
                "--limit",
                limit,
            ])
            .assert()
            .failure()
            .code(2)
            .stderr(contains("invalid value"))
            .stderr(contains("session export").not());
    }
}

#[test]
fn arbitrary_graph_or_owa_payloads_are_rejected_before_session_access() {
    for (args, input) in [
        (
            vec![
                "outlook",
                "mail",
                "create",
                "--confirm",
                "--body-file",
                "/dev/stdin",
            ],
            r#"{"subject":"synthetic","toRecipients":[]}"#,
        ),
        (
            vec![
                "outlook",
                "calendar",
                "create",
                "--folder-id",
                "synthetic-folder",
                "--confirm",
                "--body-file",
                "/dev/stdin",
            ],
            r#"{"subject":"synthetic","attendees":[]}"#,
        ),
        (
            vec![
                "outlook",
                "mail",
                "update",
                "--id",
                "synthetic-item",
                "--confirm",
                "--body-file",
                "/dev/stdin",
            ],
            r#"{"Body":{"MessageDisposition":"SendAndSaveCopy"}}"#,
        ),
    ] {
        moff()
            .env("HOME", "/definitely/missing/moff-contract-home")
            .args(args)
            .write_stdin(input)
            .assert()
            .failure()
            .code(2)
            .stderr(contains("--body-file must contain only"))
            .stderr(contains("session export").not());
    }
}

#[test]
fn no_send_or_generic_write_command_is_exposed() {
    for args in [
        vec!["outlook", "mail", "send", "--confirm"],
        vec![
            "outlook",
            "service",
            "write",
            "--action",
            "SendItem",
            "--confirm",
        ],
    ] {
        moff()
            .args(args)
            .assert()
            .failure()
            .code(2)
            .stderr(contains("unrecognized subcommand"));
    }
}
