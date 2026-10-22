use assert_cmd::Command;
use serde_json::{Value, json};

#[test]
fn full_access_exposes_actions_while_retaining_read_only_discovery() {
    let requests = [
        json!({"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-03-26"}}),
        json!({"jsonrpc":"2.0","id":2,"method":"tools/list"}),
        json!({"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"moff","arguments":{"args":["schema"]}}}),
        json!({"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"moff_action","arguments":{"args":["outlook","calendar","respond","--help"]}}}),
        json!({"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"moff_action","arguments":{"args":["mcp-native"]}}}),
        json!({"jsonrpc":"2.0","id":6,"method":"tools/call","params":{"name":"moff","arguments":{"args":["auth","refresh","--confirm"]}}}),
        json!({"jsonrpc":"2.0","id":7,"method":"tools/call","params":{"name":"moff_action","arguments":{"args":["outlook","onedrive","share","--help"]}}}),
    ];
    let output = Command::cargo_bin("moff")
        .unwrap()
        .args(["mcp", "--full-access"])
        .write_stdin(
            requests
                .iter()
                .map(Value::to_string)
                .collect::<Vec<_>>()
                .join("\n")
                + "\n",
        )
        .output()
        .unwrap();
    assert!(output.status.success());
    let responses: Vec<Value> = String::from_utf8(output.stdout)
        .unwrap()
        .lines()
        .map(|line| serde_json::from_str(line).unwrap())
        .collect();
    assert_eq!(responses[1]["result"]["tools"].as_array().unwrap().len(), 2);
    assert_eq!(
        responses[1]["result"]["tools"][1]["annotations"]["readOnlyHint"],
        false
    );
    assert_eq!(
        responses[1]["result"]["tools"][1]["annotations"]["idempotentHint"],
        false
    );
    assert_eq!(responses[2]["result"]["isError"], false);
    let schema: Value = serde_json::from_str(
        responses[2]["result"]["content"][0]["text"]
            .as_str()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(schema["full_access_mcp"]["requires_host_approval"], true);
    assert_eq!(schema["commands"]["mail"]["send"], true);
    assert_eq!(schema["mail_send"]["attachments"], false);
    assert_eq!(schema["commands"]["onedrive"]["share"], true);
    assert_eq!(responses[3]["result"]["isError"], false);
    assert_eq!(responses[4]["result"]["isError"], true);
    assert_eq!(responses[5]["result"]["isError"], true);
    assert_eq!(responses[6]["result"]["isError"], false);
}

#[test]
fn send_is_discoverable_and_approval_gated_through_actions_without_browser_access() {
    for full_access in [false, true] {
        let send = json!([
            "outlook",
            "mail",
            "send",
            "--to",
            "recipient@example.com",
            "--subject",
            "Synthetic",
            "--body",
            "Text",
            "--confirm"
        ]);
        let requests = [
            json!({"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}),
            json!({"jsonrpc":"2.0","id":2,"method":"tools/list"}),
            json!({"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"moff","arguments":{"args":send}}}),
            json!({"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"moff_action","arguments":{"args":send}}}),
            json!({"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"moff_action","arguments":{"args":["outlook","mail","send","--help"]}}}),
            json!({"jsonrpc":"2.0","id":6,"method":"tools/call","params":{"name":"moff_action","arguments":{"args":["outlook","--allow-mutation","mail","send","--to","not-an-email","--subject","Synthetic","--body","Text","--confirm"]}}}),
        ];
        let mut command = Command::cargo_bin("moff").unwrap();
        command.arg("mcp");
        if full_access {
            command.arg("--full-access");
        }
        let output = command
            .write_stdin(
                requests
                    .iter()
                    .map(Value::to_string)
                    .collect::<Vec<_>>()
                    .join("\n")
                    + "\n",
            )
            .output()
            .unwrap();
        assert!(output.status.success());
        let responses: Vec<Value> = String::from_utf8(output.stdout)
            .unwrap()
            .lines()
            .map(|line| serde_json::from_str(line).unwrap())
            .collect();
        assert_eq!(responses[2]["result"]["isError"], true);
        assert_eq!(responses[3]["result"]["isError"], true);
        assert_eq!(responses[4]["result"]["isError"], !full_access);
        assert_eq!(responses[5]["result"]["isError"], true);
        if full_access {
            let tool = &responses[1]["result"]["tools"][1];
            assert!(
                tool["description"]
                    .as_str()
                    .unwrap()
                    .contains("mail sending")
            );
            assert_eq!(tool["annotations"]["readOnlyHint"], false);
            assert_eq!(tool["annotations"]["idempotentHint"], false);
            assert!(
                responses[3]["result"]["content"][0]["text"]
                    .as_str()
                    .unwrap()
                    .contains("--allow-mutation")
            );
            assert!(
                responses[5]["result"]["content"][0]["text"]
                    .as_str()
                    .unwrap()
                    .contains("unique explicit SMTP")
            );
        }
    }
}

#[test]
fn stdio_discovery_and_cli_defaults_without_browser_access() {
    let requests = [
        json!({"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-03-26"}}),
        json!({"jsonrpc":"2.0","method":"notifications/initialized"}),
        json!({"jsonrpc":"2.0","id":2,"method":"tools/list"}),
        json!({"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"moff","arguments":{"args":["schema"]}}}),
        json!({"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"moff","arguments":{"args":["outlook","mail","list","--help"]}}}),
        json!({"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"moff","arguments":{"args":["auth","refresh","--confirm"]}}}),
        json!({"jsonrpc":"2.0","id":6,"method":"tools/call","params":{"name":"moff","arguments":{"args":["mcp"]}}}),
    ];
    let input = requests
        .iter()
        .map(Value::to_string)
        .collect::<Vec<_>>()
        .join("\n")
        + "\n";
    let output = Command::cargo_bin("moff")
        .unwrap()
        .arg("mcp")
        .write_stdin(input)
        .output()
        .unwrap();
    assert!(output.status.success());
    let responses: Vec<Value> = String::from_utf8(output.stdout)
        .unwrap()
        .lines()
        .map(|line| serde_json::from_str(line).unwrap())
        .collect();
    assert_eq!(responses.len(), 6);
    assert_eq!(responses[0]["result"]["protocolVersion"], "2025-03-26");
    assert!(
        responses[0]["result"]["instructions"]
            .as_str()
            .unwrap()
            .contains("TotalItemsInView is the collection-wide pagination total")
    );
    assert_eq!(responses[1]["result"]["tools"][0]["name"], "moff");
    assert!(
        responses[1]["result"]["tools"][0]["description"]
            .as_str()
            .unwrap()
            .contains("NOT a bare array")
    );
    assert_eq!(responses[2]["result"]["isError"], false);
    let schema: Value = serde_json::from_str(
        responses[2]["result"]["content"][0]["text"]
            .as_str()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(schema["schema"], "moff.schema/v1");
    assert_eq!(schema["command_prefixes"]["default"], json!(["outlook"]));
    assert_eq!(schema["command_prefixes"]["auth"], json!(["auth"]));
    assert!(
        responses[3]["result"]["content"][0]["text"]
            .as_str()
            .unwrap()
            .contains("[default: 25]")
    );
    assert_eq!(responses[4]["result"]["isError"], true);
    assert_eq!(responses[5]["result"]["isError"], true);
}

#[test]
fn rejects_writes_invalid_calls_and_protocol_errors() {
    let mut lines = vec![
        "not json".to_owned(),
        json!({"jsonrpc":"2.0","id":1,"method":"tools/list"}).to_string(),
        json!({"jsonrpc":"2.0","id":2,"method":"initialize","params":{"protocolVersion":"2024-11-05"}}).to_string(),
        json!({"jsonrpc":"2.0","id":3,"method":"unknown"}).to_string(),
    ];
    for args in [
        vec!["outlook", "mail", "delete", "--id", "fake", "--confirm"],
        vec![
            "outlook",
            "onedrive",
            "share",
            "--id",
            "file",
            "--email",
            "sarah@example.com",
            "--confirm",
        ],
        vec!["outlook", "--allow-mutation=true", "mail", "list"],
        vec![
            "outlook",
            "onedrive",
            "download",
            "--id",
            "fake",
            "--out",
            "/tmp/not-written",
        ],
        vec!["outlook", "mail", "list", "--limit", "bad"],
    ] {
        lines.push(json!({"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"moff","arguments":{"args":args}}}).to_string());
    }

    let output = Command::cargo_bin("moff")
        .unwrap()
        .arg("mcp")
        .write_stdin(lines.join("\n") + "\n")
        .output()
        .unwrap();
    assert!(output.status.success());
    let responses: Vec<Value> = String::from_utf8(output.stdout)
        .unwrap()
        .lines()
        .map(|line| serde_json::from_str(line).unwrap())
        .collect();
    assert_eq!(responses[0]["error"]["code"], -32700);
    assert_eq!(responses[1]["error"]["code"], -32002);
    assert_eq!(responses[3]["error"]["code"], -32601);
    for response in &responses[4..] {
        assert_eq!(response["result"]["isError"], true);
    }
}

#[test]
fn archive_server_exposes_only_scoped_tools_and_rejects_invalid_plans_without_browser_access() {
    let mut requests = vec![
        json!({"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18"}}),
        json!({"jsonrpc":"2.0","id":2,"method":"tools/list"}),
    ];
    for (name, args) in [
        (
            "mail_archive",
            json!({"plan":{"archive_folder_id":"a","inbox_folder_id":"i","messages":[]}}),
        ),
        (
            "mail_archive_preview",
            json!({"ids":["same","same"],"archive_folder_id":"a"}),
        ),
        (
            "mail_archive_preview",
            json!({"ids":[],"archive_folder_id":"a"}),
        ),
        ("mail_archive_folders", json!({"unexpected":true})),
        (
            "moff",
            json!({"args":["outlook","mail","delete","--id","id","--confirm"]}),
        ),
    ] {
        requests.push(json!({"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":name,"arguments":args}}));
    }
    let output = Command::cargo_bin("moff")
        .unwrap()
        .arg("mcp-archive")
        .write_stdin(
            requests
                .iter()
                .map(Value::to_string)
                .collect::<Vec<_>>()
                .join("\n")
                + "\n",
        )
        .output()
        .unwrap();
    assert!(output.status.success());
    let responses: Vec<Value> = String::from_utf8(output.stdout)
        .unwrap()
        .lines()
        .map(|line| serde_json::from_str(line).unwrap())
        .collect();
    let tools = responses[1]["result"]["tools"].as_array().unwrap();
    assert_eq!(tools.len(), 3);
    assert_eq!(tools[0]["annotations"]["readOnlyHint"], true);
    assert_eq!(tools[1]["annotations"]["readOnlyHint"], true);
    assert_eq!(tools[2]["name"], "mail_archive");
    assert_eq!(tools[2]["annotations"]["readOnlyHint"], false);
    assert_eq!(tools[2]["annotations"]["destructiveHint"], true);
    for response in &responses[2..] {
        assert_eq!(response["result"]["isError"], true);
    }
}
