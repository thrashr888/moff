use assert_cmd::Command;
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
        .stdout(predicates::str::contains("auth"))
        .stdout(predicates::str::contains("doctor"))
        .stdout(predicates::str::contains("schema"))
        .stdout(predicates::str::contains("outlook"))
        .stdout(predicates::str::contains("--pretty"))
        .stdout(predicates::str::contains("--envelope"))
        .stdout(predicates::str::contains("--json"))
        .stdout(predicates::str::contains("--profile"))
        .stdout(predicates::str::contains("--dry-run"))
        .stdout(predicates::str::contains("--yes"))
        .stdout(predicates::str::contains("--no-color"));
}

#[test]
fn no_color_is_accepted_as_a_global_agent_flag() {
    moff().args(["--no-color", "doctor"]).assert().success();
}

#[test]
fn default_command_output_is_compact_json() {
    let value = json_stdout(&["doctor"]);
    assert_eq!(value["command"], "doctor");
    assert_eq!(value["status"], "scaffold");
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
        .stderr(predicates::str::contains("incompatible_output_modes"));
}

#[test]
fn unknown_outlook_source_fails_with_a_typed_code() {
    moff()
        .args(["outlook", "unknown"])
        .assert()
        .failure()
        .code(2)
        .stderr(predicates::str::contains("unknown_source"));
}

#[test]
fn schema_filters_to_the_requested_source() {
    let value = json_stdout(&["schema", "--source", "outlook"]);
    assert_eq!(value["schema"], "moff.schema/v1");
    assert_eq!(value["commands"].as_object().expect("command map").len(), 1);
    assert!(value["commands"]["outlook"].is_object());
}
