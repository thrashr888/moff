use assert_cmd::Command;
use predicates::prelude::PredicateBooleanExt;
use predicates::str::contains;
use serde_json::Value;

fn moff() -> Command {
    Command::cargo_bin("moff").unwrap()
}

#[test]
fn help_exposes_only_local_outlook_read_commands() {
    moff()
        .arg("--help")
        .assert()
        .success()
        .stdout(contains("outlook"))
        .stdout(contains("auth").not())
        .stdout(contains("--provider").not())
        .stdout(contains("--profile").not());
    moff()
        .args(["outlook", "--help"])
        .assert()
        .success()
        .stdout(contains("doctor"))
        .stdout(contains("mail"))
        .stdout(contains("Graph").not());
}

#[test]
fn graph_and_auth_commands_are_not_accepted() {
    for args in [
        ["auth", "login"].as_slice(),
        ["outlook", "--provider", "graph", "doctor"].as_slice(),
        ["outlook", "mail", "list"].as_slice(),
    ] {
        moff().args(args).assert().failure().code(2);
    }
}

#[test]
fn doctor_is_safe_and_contains_no_email_content() {
    let output = moff()
        .env(
            "MOFF_TEST_WINDOW_JSON",
            r#"[{"number":42,"title":"Inbox","x":0,"y":0,"width":800,"height":600}]"#,
        )
        .args(["outlook", "doctor"])
        .assert()
        .success()
        .get_output()
        .stdout
        .clone();
    let value: Value = serde_json::from_slice(&output).unwrap();
    assert_eq!(value["command"], "outlook.doctor");
    assert_eq!(value["source"], "outlook-screen-local");
    assert_eq!(value["email_content"], false);
    assert!(!String::from_utf8(output).unwrap().contains("Alice Example"));
}

#[test]
fn visible_mail_rejects_over_limit_before_capture() {
    moff()
        .args(["outlook", "mail", "visible", "--limit", "101"])
        .assert()
        .failure()
        .code(2)
        .stderr(contains("invalid_limit"));
}

#[test]
fn visible_mail_outputs_safe_provenance_and_cleans_injected_temp_dir() {
    let temp = std::env::temp_dir().join(format!("moff-cli-test-{}", std::process::id()));
    std::fs::create_dir_all(&temp).unwrap();
    let fixture = r#"[{"text":"Alice Example","x":0.10,"y":0.86,"w":0.2,"h":0.02},{"text":"Quarterly status","x":0.10,"y":0.83,"w":0.2,"h":0.02},{"text":"Agenda attached","x":0.10,"y":0.80,"w":0.2,"h":0.02},{"text":"10:42 AM","x":0.64,"y":0.86,"w":0.08,"h":0.02},{"text":"OPEN MAIL BODY MUST NOT LEAK","x":0.92,"y":0.50,"w":0.05,"h":0.02}]"#;
    let output = moff()
        .env(
            "MOFF_TEST_WINDOW_JSON",
            r#"[{"number":42,"title":"Inbox","x":0,"y":0,"width":800,"height":600}]"#,
        )
        .env("MOFF_TEST_OCR_JSON", fixture)
        .env("MOFF_TEST_TEMP_DIR", &temp)
        .args(["outlook", "mail", "visible"])
        .assert()
        .success()
        .get_output()
        .stdout
        .clone();
    let value: Value = serde_json::from_slice(&output).unwrap();
    assert_eq!(value["source"], "outlook-screen-local");
    assert_eq!(value["content_origin"], "external_email_untrusted");
    assert_eq!(value["value"][0]["locator"], "capture-row:1");
    assert!(
        !String::from_utf8(output)
            .unwrap()
            .contains("OPEN MAIL BODY")
    );
    assert!(std::fs::read_dir(&temp).unwrap().next().is_none());
}
