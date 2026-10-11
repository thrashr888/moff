use crate::{OutlookWebSession, http_client};
use serde_json::Value;
use std::{
    collections::BTreeMap,
    io::{Read, Write},
    net::TcpListener,
    sync::{Arc, Mutex},
    thread,
    time::{Duration, Instant},
};

#[allow(clippy::needless_pass_by_value)]
pub fn message(item: Value) -> Value {
    serde_json::json!({"Body":{"ResponseMessages":{"Items":[{"ResponseCode":"NoError","Items":[item]}]}}})
}

pub fn mock(
    replies: Vec<(&'static str, Value)>,
) -> (
    OutlookWebSession,
    Arc<Mutex<Vec<Value>>>,
    thread::JoinHandle<()>,
) {
    let listener = TcpListener::bind("127.0.0.1:0").unwrap();
    listener.set_nonblocking(true).unwrap();
    let origin = url::Url::parse(&format!("http://{}/", listener.local_addr().unwrap())).unwrap();
    let requests = Arc::new(Mutex::new(Vec::new()));
    let captured = requests.clone();
    let worker = thread::spawn(move || {
        for (action, response) in replies {
            let deadline = Instant::now() + Duration::from_secs(5);
            let mut socket = loop {
                if let Ok((socket, _)) = listener.accept() {
                    break socket;
                }
                assert!(Instant::now() < deadline, "expected request for {action}");
                thread::sleep(Duration::from_millis(5));
            };
            socket.set_nonblocking(false).unwrap();
            socket
                .set_read_timeout(Some(Duration::from_secs(5)))
                .unwrap();
            let mut bytes = Vec::new();
            let (offset, length) =
                loop {
                    let mut buffer = [0u8; 4096];
                    let read = socket.read(&mut buffer).unwrap();
                    assert!(read > 0);
                    bytes.extend_from_slice(&buffer[..read]);
                    if let Some(offset) = bytes.windows(4).position(|v| v == b"\r\n\r\n") {
                        let headers = std::str::from_utf8(&bytes[..offset]).unwrap();
                        let length = headers
                            .lines()
                            .find_map(|line| {
                                let (name, value) = line.split_once(':')?;
                                name.eq_ignore_ascii_case("content-length")
                                    .then(|| value.trim().parse::<usize>().unwrap())
                            })
                            .unwrap();
                        assert!(headers.starts_with(&format!(
                            "POST /owa/service.svc?action={action}&app=Mail "
                        )));
                        if bytes.len() >= offset + 4 + length {
                            break (offset + 4, length);
                        }
                    }
                };
            captured
                .lock()
                .unwrap()
                .push(serde_json::from_slice(&bytes[offset..offset + length]).unwrap());
            let text = response.to_string();
            write!(
                socket,
                "HTTP/1.1 200 OK\r\nContent-Length: {}\r\nConnection: close\r\n\r\n{}",
                text.len(),
                text
            )
            .unwrap();
        }
    });
    (
        OutlookWebSession {
            origin,
            headers: BTreeMap::new(),
            client: http_client().unwrap(),
            safari: None,
        },
        requests,
        worker,
    )
}
