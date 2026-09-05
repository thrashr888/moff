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
fn unknown_outlook_source_fails_with_a_typed_code() {
    moff()
        .args(["outlook", "unknown"])
        .assert()
        .failure()
        .code(2)
        .stderr(contains("unknown_source"));
}
#[test]
fn schema_filters_to_the_requested_source() {
    let value = json_stdout(&["schema", "--source", "outlook"]);
    assert_eq!(value["schema"], "moff.schema/v1");
    assert_eq!(value["commands"].as_object().expect("command map").len(), 1);
    assert!(value["commands"]["outlook"].is_object());
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
