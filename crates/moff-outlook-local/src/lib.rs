//! Read-only local macOS Outlook window capture and Vision OCR.
//! No network, credentials, Accessibility APIs, or persistent capture storage.
use std::fmt;
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::Command;

use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

const MAX_LIMIT: u16 = 100;
const OCR_HELPER: &str = include_str!("ocr.swift");

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LocalError {
    UnsupportedPlatform,
    OutlookNotRunning,
    WindowAmbiguous,
    ScreenRecordingDenied,
    UiChanged,
    InvalidLimit,
    CaptureFailed,
}
impl LocalError {
    #[must_use]
    pub const fn code(&self) -> &'static str {
        match self {
            Self::UnsupportedPlatform => "unsupported_platform",
            Self::OutlookNotRunning => "outlook_not_running",
            Self::WindowAmbiguous => "outlook_window_ambiguous",
            Self::ScreenRecordingDenied => "screen_recording_denied",
            Self::UiChanged => "ui_changed",
            Self::InvalidLimit => "invalid_limit",
            Self::CaptureFailed => "capture_failed",
        }
    }
}
impl fmt::Display for LocalError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::UnsupportedPlatform => "this provider is supported only on macOS",
            Self::OutlookNotRunning => "no visible Microsoft Outlook mail window was found",
            Self::WindowAmbiguous => {
                "exactly one visible Microsoft Outlook mail window is required"
            }
            Self::ScreenRecordingDenied => {
                "Screen Recording permission is required to capture the Outlook window"
            }
            Self::UiChanged => "Outlook's visible message-list layout was not recognized",
            Self::InvalidLimit => "limit must be between 1 and 100",
            Self::CaptureFailed => "unable to capture the selected Outlook window",
        })
    }
}
impl std::error::Error for LocalError {}

#[derive(Debug, Deserialize)]
struct Window {
    number: u64,
    title: String,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
}
#[derive(Debug, Deserialize)]
struct Observation {
    text: String,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
}
#[derive(Debug, Serialize)]
pub struct VisibleRow {
    pub unread: Option<bool>,
    pub sender: String,
    pub subject: String,
    pub preview: String,
    pub time: String,
    pub locator: String,
    pub locator_scope: &'static str,
}

fn checked_limit(limit: u16) -> Result<(), LocalError> {
    if (1..=MAX_LIMIT).contains(&limit) {
        Ok(())
    } else {
        Err(LocalError::InvalidLimit)
    }
}
fn windows() -> Result<Vec<Window>, LocalError> {
    if !cfg!(target_os = "macos") {
        return Err(LocalError::UnsupportedPlatform);
    }
    if let Some(value) = std::env::var_os("MOFF_TEST_WINDOW_JSON") {
        return serde_json::from_str(&value.to_string_lossy())
            .map_err(|_| LocalError::CaptureFailed);
    }
    let mut child = Command::new("/usr/bin/swift")
        .args(["-", "windows"])
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .spawn()
        .map_err(|_| LocalError::CaptureFailed)?;
    child
        .stdin
        .take()
        .ok_or(LocalError::CaptureFailed)?
        .write_all(OCR_HELPER.as_bytes())
        .map_err(|_| LocalError::CaptureFailed)?;
    let output = child
        .wait_with_output()
        .map_err(|_| LocalError::CaptureFailed)?;
    if !output.status.success() {
        return Err(LocalError::CaptureFailed);
    }
    serde_json::from_slice(&output.stdout).map_err(|_| LocalError::CaptureFailed)
}

fn selected_window() -> Result<Window, LocalError> {
    let windows = windows()?;
    let suitable: Vec<_> = windows
        .into_iter()
        .filter(|window| {
            window.width > 300.0
                && window.height > 200.0
                && (window.title.contains("Inbox")
                    || window.title.contains("Mail")
                    || window.title.contains("Outlook"))
        })
        .collect();
    match suitable.len() {
        0 => Err(LocalError::OutlookNotRunning),
        1 => Ok(suitable.into_iter().next().expect("one window")),
        _ => Err(LocalError::WindowAmbiguous),
    }
}

fn secure_capture_path() -> Result<PathBuf, LocalError> {
    let root =
        std::env::var_os("MOFF_TEST_TEMP_DIR").map_or_else(std::env::temp_dir, PathBuf::from);
    fs::create_dir_all(&root).map_err(|_| LocalError::CaptureFailed)?;
    let path = root.join(format!(
        "moff-outlook-{}-{}.png",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map_err(|_| LocalError::CaptureFailed)?
            .as_nanos()
    ));
    OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&path)
        .map_err(|_| LocalError::CaptureFailed)?;
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        fs::set_permissions(&path, fs::Permissions::from_mode(0o600))
            .map_err(|_| LocalError::CaptureFailed)?;
    }
    Ok(path)
}
struct Cleanup(PathBuf);
impl Drop for Cleanup {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.0);
    }
}

fn capture(window: &Window) -> Result<PathBuf, LocalError> {
    if std::env::var_os("MOFF_TEST_CAPTURE_DENIED").is_some() {
        return Err(LocalError::ScreenRecordingDenied);
    }
    let path = secure_capture_path()?;
    let cleanup = Cleanup(path.clone());
    if std::env::var_os("MOFF_TEST_OCR_JSON").is_some() {
        fs::write(&path, b"synthetic capture").map_err(|_| LocalError::CaptureFailed)?;
        std::mem::forget(cleanup);
        return Ok(path);
    }
    let result = Command::new("/usr/sbin/screencapture")
        .args([
            "-x",
            "-l",
            &window.number.to_string(),
            path.to_str().ok_or(LocalError::CaptureFailed)?,
        ])
        .status();
    match result {
        Ok(status) if status.success() && fs::metadata(&path).is_ok_and(|meta| meta.len() > 0) => {
            std::mem::forget(cleanup);
            Ok(path)
        }
        Ok(_) => Err(LocalError::ScreenRecordingDenied),
        Err(_) => Err(LocalError::CaptureFailed),
    }
}

/// # Errors
/// Returns `ui_changed` unless OCR proves a bounded, columnar message list.
///
/// Invariants: at least two exact date/time anchors share one right-hand column; each
/// emitted row has three ordered, non-overlapping text lines to that column's left; and
/// the list's right edge is derived from that anchor column. Text at or beyond that edge
/// is preview-pane content and is never considered. A mail header label anywhere in an
/// otherwise eligible row makes the whole capture ambiguous rather than risking a leak.
pub fn parse_visible_rows(ocr: &str, limit: u16) -> Result<Vec<VisibleRow>, LocalError> {
    checked_limit(limit)?;
    let items: Vec<Observation> = serde_json::from_str::<Vec<Observation>>(ocr)
        .map_err(|_| LocalError::UiChanged)?
        .into_iter()
        .filter(|item| {
            item.x >= 0.0
                && item.y >= 0.0
                && item.x + item.w <= 1.0
                && item.y + item.h <= 1.0
                && item.w > 0.0
                && item.h > 0.0
        })
        .collect();
    let mut anchors: Vec<_> = items
        .iter()
        .filter(|item| item.x >= 0.20 && item.x <= 0.80 && is_time(&item.text))
        .collect();
    if anchors.len() < 2 {
        return Err(LocalError::UiChanged);
    }
    anchors.sort_by(|a, b| b.y.total_cmp(&a.y));
    let min_anchor_x = anchors.iter().map(|item| item.x).fold(1.0, f64::min);
    let max_anchor_x = anchors.iter().map(|item| item.x).fold(0.0, f64::max);
    // A real list has one date column. A wider set can be document/body dates.
    if max_anchor_x - min_anchor_x > 0.08 {
        return Err(LocalError::UiChanged);
    }
    let list_right = max_anchor_x + 0.08;
    let mut rows = Vec::new();
    for (index, anchor) in anchors.iter().enumerate() {
        let next_anchor_y = anchors.get(index + 1).map_or(-1.0, |next| next.y);
        // Row bands must not overlap; this prevents an anchor from borrowing another row.
        if anchor.y - next_anchor_y < 0.08 && index + 1 < anchors.len() {
            return Err(LocalError::UiChanged);
        }
        let mut row_items: Vec<_> = items
            .iter()
            .filter(|item| {
                item.x >= 0.18
                    && item.x + item.w <= list_right
                    && item.x + item.w < anchor.x + 0.04
                    && item.y <= anchor.y + 0.05
                    && item.y >= anchor.y - 0.10
            })
            .collect();
        if row_items.iter().any(|item| forbidden_row_text(&item.text)) {
            return Err(LocalError::UiChanged);
        }
        row_items.sort_by(|a, b| b.y.total_cmp(&a.y));
        if row_items.len() != 3
            || row_items[0].y - row_items[1].y < 0.008
            || row_items[1].y - row_items[2].y < 0.008
            || row_items.iter().any(|item| is_time(&item.text))
        {
            return Err(LocalError::UiChanged);
        }
        rows.push(VisibleRow {
            unread: None,
            sender: row_items[0].text.clone(),
            subject: row_items[1].text.clone(),
            preview: row_items[2].text.clone(),
            time: anchor.text.clone(),
            locator: format!("capture-row:{}", rows.len() + 1),
            locator_scope: "ephemeral_local_screen_capture",
        });
    }
    if rows.is_empty() {
        return Err(LocalError::UiChanged);
    }
    rows.truncate(usize::from(limit));
    Ok(rows)
}

fn forbidden_row_text(value: &str) -> bool {
    let lowered = value.trim().to_ascii_lowercase();
    [
        "from:", "to:", "cc:", "date:", "subject:", "pinned", "today",
    ]
    .iter()
    .any(|label| lowered.starts_with(label))
}

fn is_time(value: &str) -> bool {
    let value = value.trim();
    if value.eq_ignore_ascii_case("yesterday") {
        return true;
    }
    let upper = value.to_ascii_uppercase();
    let clock = upper
        .strip_suffix(" AM")
        .or_else(|| upper.strip_suffix(" PM"))
        .unwrap_or(&upper);
    let clock_parts: Vec<_> = clock.split(':').collect();
    if clock_parts.len() == 2
        && clock_parts
            .iter()
            .all(|part| !part.is_empty() && part.chars().all(|c| c.is_ascii_digit()))
    {
        return true;
    }
    let date_parts: Vec<_> = value.split('/').collect();
    (date_parts.len() == 2 || date_parts.len() == 3)
        && date_parts
            .iter()
            .all(|part| !part.is_empty() && part.chars().all(|c| c.is_ascii_digit()))
}

fn ocr(path: &Path) -> Result<String, LocalError> {
    if let Some(value) = std::env::var_os("MOFF_TEST_OCR_JSON") {
        return Ok(value.to_string_lossy().into_owned());
    }
    let mut child = Command::new("/usr/bin/swift")
        .args(["-", "ocr", path.to_str().ok_or(LocalError::CaptureFailed)?])
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .spawn()
        .map_err(|_| LocalError::CaptureFailed)?;
    child
        .stdin
        .take()
        .ok_or(LocalError::CaptureFailed)?
        .write_all(OCR_HELPER.as_bytes())
        .map_err(|_| LocalError::CaptureFailed)?;
    let out = child
        .wait_with_output()
        .map_err(|_| LocalError::CaptureFailed)?;
    if out.status.success() {
        String::from_utf8(out.stdout).map_err(|_| LocalError::UiChanged)
    } else {
        Err(LocalError::UiChanged)
    }
}

/// # Errors
/// Returns a typed platform or visible-window selection error without reading email text.
pub fn doctor() -> Result<Value, LocalError> {
    let window = selected_window()?;
    // Permission is demonstrated with a window-only capture; its Cleanup guard
    // deletes it immediately and doctor never hands it to OCR.
    let path = capture(&window)?;
    drop(Cleanup(path));
    Ok(
        json!({"command":"outlook.doctor","provider":"local-macos-screen-ocr","source":"outlook-screen-local","platform":"macos","screen_recording":"ready","email_content":false,"network":false,"window":{"number":window.number,"frame":{"x":window.x,"y":window.y,"width":window.width,"height":window.height}}}),
    )
}
/// # Errors
/// Returns typed capture, permission, window, limit, or conservative OCR layout errors.
pub fn visible_mail(limit: u16) -> Result<Value, LocalError> {
    checked_limit(limit)?;
    let window = selected_window()?;
    let path = capture(&window)?;
    let cleanup = Cleanup(path.clone());
    let rows = parse_visible_rows(&ocr(&path)?, limit)?;
    drop(cleanup);
    Ok(
        json!({"command":"outlook.mail.visible","provider":"local-macos-screen-ocr","source":"outlook-screen-local","content_origin":"external_email_untrusted","value":rows,"capture_locator":{"window_number":window.number,"frame":{"x":window.x,"y":window.y,"width":window.width,"height":window.height},"scope":"ephemeral_current_capture"}}),
    )
}
