#![cfg(feature = "test-support")]

use std::io::{Read, Write};
use std::net::TcpListener;
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::{Duration, Instant};

use assert_cmd::Command;

fn moff_test() -> Command {
    Command::cargo_bin("moff-test").expect("test-only moff helper is built")
}

fn read_request(stream: &mut impl Read) -> String {
    let deadline = Instant::now() + Duration::from_secs(2);
    let mut request = Vec::new();
    let mut buffer = [0_u8; 1024];
    loop {
        let count = match stream.read(&mut buffer) {
            Ok(count) => count,
            Err(error) if error.kind() == std::io::ErrorKind::WouldBlock => {
                assert!(Instant::now() < deadline, "timed out reading request");
                thread::sleep(Duration::from_millis(10));
                continue;
            }
            Err(error) => panic!("read request: {error}"),
        };
        assert_ne!(count, 0, "request ended before headers");
        request.extend_from_slice(&buffer[..count]);
        let headers_end = request
            .windows(4)
            .position(|window| window == b"\r\n\r\n")
            .map(|index| index + 4);
        let Some(headers_end) = headers_end else {
            continue;
        };
        let headers = String::from_utf8_lossy(&request[..headers_end]);
        let content_length = headers
            .lines()
            .find_map(|line| line.strip_prefix("Content-Length: "))
            .and_then(|value| value.parse::<usize>().ok())
            .unwrap_or(0);
        if request.len() >= headers_end + content_length {
            return String::from_utf8(request).expect("request is utf-8");
        }
    }
}

#[test]
fn persisted_profile_refreshes_then_maps_token_to_graph_without_cli_token_leakage() {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind local test server");
    listener
        .set_nonblocking(true)
        .expect("make test server non-blocking");
    let address = listener.local_addr().expect("read test server address");
    let requests = Arc::new(Mutex::new(Vec::new()));
    let seen = Arc::clone(&requests);
    let server = thread::spawn(move || -> Result<(), String> {
        for _ in 0..2 {
            let deadline = Instant::now() + Duration::from_secs(2);
            let (mut stream, _) = loop {
                match listener.accept() {
                    Ok(connection) => break connection,
                    Err(error) if error.kind() == std::io::ErrorKind::WouldBlock => {
                        if Instant::now() >= deadline {
                            return Err("timed out waiting for CLI request".to_owned());
                        }
                        thread::sleep(Duration::from_millis(10));
                    }
                    Err(error) => return Err(error.to_string()),
                }
            };
            let request = read_request(&mut stream);
            let body = if request.starts_with("POST /common/oauth2/v2.0/token") {
                r#"{"access_token":"access-secret","expires_in":3600}"#
            } else {
                r#"{"value":[]}"#
            };
            seen.lock()
                .map_err(|_| "request lock poisoned".to_owned())?
                .push(request);
            write!(stream, "HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: {}\r\nConnection: close\r\n\r\n{}", body.len(), body).map_err(|error| error.to_string())?;
        }
        Ok(())
    });

    let root = std::env::temp_dir().join(format!("moff-cli-refresh-{}", std::process::id()));
    std::fs::create_dir_all(root.join("profiles")).expect("create profile directory");
    std::fs::write(root.join("profiles/default.json"), r#"{"name":"default","client_id":"11111111-1111-1111-1111-111111111111","tenant":"common","scopes":["Mail.Read"],"account_display_name":null}"#).expect("write persisted profile");
    let refresh_token = root.join("refresh-token");
    std::fs::write(&refresh_token, "refresh-secret").expect("write controlled refresh token");

    let output = moff_test()
        .env("MOFF_DATA_DIR", &root)
        .env("MOFF_ENTRA_BASE_URL", format!("http://{address}"))
        .env("MOFF_GRAPH_BASE_URL", format!("http://{address}/v1.0"))
        .env("MOFF_TEST_REFRESH_TOKEN_FILE", &refresh_token)
        .args(["outlook", "mail", "list"])
        .assert()
        .success()
        .get_output()
        .clone();

    std::fs::remove_dir_all(&root).expect("remove temporary profile");
    server
        .join()
        .expect("test server panicked")
        .expect("test server completed");
    let stdout = String::from_utf8(output.stdout).expect("stdout is utf-8");
    let stderr = String::from_utf8(output.stderr).expect("stderr is utf-8");
    assert!(!format!("{stdout}{stderr}").contains("refresh-secret"));
    assert!(!format!("{stdout}{stderr}").contains("access-secret"));

    let requests = requests.lock().expect("request lock");
    assert_eq!(requests.len(), 2);
    assert!(requests[0].starts_with("POST /common/oauth2/v2.0/token"));
    assert!(requests[0].contains("grant_type=refresh_token"));
    assert!(requests[0].contains("refresh_token=refresh-secret"));
    assert!(requests[1].starts_with("GET /v1.0/me/mailFolders/inbox/messages"));
    assert!(
        requests[1]
            .lines()
            .any(|line| line.eq_ignore_ascii_case("Authorization: Bearer access-secret"))
    );
}
