use std::io::{Cursor, Read, Write};

use serde_json::json;
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

use super::{
    DocumentEdit, DocumentLocation, WordPackage, upload_acknowledgement, verify_saved_with,
};
use crate::OutlookWebError;

#[test]
fn readback_rechecks_transient_or_stale_reads_without_ever_replaying_a_write() {
    let before = json!({"etag":"old"});
    let mut reads = 0;
    let after = verify_saved_with(
        &before,
        b"saved",
        || {
            reads += 1;
            match reads {
                1 => Err(OutlookWebError::SafariOperationFailed),
                2 => Ok((before.clone(), b"old".to_vec())),
                _ => Ok((json!({"etag":"new"}), b"saved".to_vec())),
            }
        },
        || {},
    )
    .unwrap();
    assert_eq!(after["etag"], "new");
    assert_eq!(reads, 3);
    reads = 0;
    assert!(
        verify_saved_with(
            &before,
            b"saved",
            || {
                reads += 1;
                Ok((json!({"etag":"new"}), b"different".to_vec()))
            },
            || {}
        )
        .is_err()
    );
    assert_eq!(reads, 3);
}

#[test]
fn rejected_writes_preserve_http_status_and_do_not_claim_success() {
    let failed = || {
        Err(OutlookWebError::OutcomeUnknown {
            action: "test update",
        })
    };
    assert!(matches!(
        upload_acknowledgement(failed(), &json!({"stage":"write","status":423})),
        Err(OutlookWebError::ProductRequestRejected(423))
    ));
    assert!(matches!(
        upload_acknowledgement(failed(), &json!({"stage":"write","status":412})),
        Err(OutlookWebError::PlanChanged)
    ));
    assert!(matches!(
        upload_acknowledgement(failed(), &json!({"stage":"write","status":503})),
        Err(OutlookWebError::OutcomeUnknown { .. })
    ));
    assert!(upload_acknowledgement(Ok(()), &json!(null)).is_err());
}

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn fixture(body: &str, settings: &str) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:custom="urn:retain" custom:flag="keep"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    );
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, bytes) in [
        ("word/document.xml", document.as_bytes()),
        ("word/settings.xml", settings.as_bytes()),
        ("word/media/retained.bin", &[0, 128, 255]),
        ("word/comments.xml", b"<comments>keep comments</comments>"),
        ("word/_rels/document.xml.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="old" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" TargetMode="External" Target="https://example.com/retained"/></Relationships>"#),
    ] {
        writer.start_file(name,SimpleFileOptions::default()).unwrap();
        writer.write_all(bytes).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

fn body() -> &'static str {
    r#"<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Next meeting</w:t></w:r></w:p>
       <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Existing note</w:t></w:r><w:hyperlink r:id="old"><w:r><w:t>Reference</w:t></w:r></w:hyperlink><w:del><w:r><w:t>Deleted text</w:t></w:r></w:del></w:p>
       <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Owner</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Milestone</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
       <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Older meeting</w:t></w:r></w:p>
       <w:p><w:r><w:t>Historic note</w:t></w:r></w:p>"#
}

fn part(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut value = Vec::new();
    archive
        .by_name(name)
        .unwrap()
        .read_to_end(&mut value)
        .unwrap();
    value
}

#[test]
fn reads_headings_tables_links_and_visible_text_without_deleted_content() {
    let package = WordPackage::open(fixture(body(), "<settings/>")).unwrap();
    let blocks = package.read().unwrap();
    assert_eq!(blocks[0].kind, "heading");
    assert_eq!(blocks[0].heading_level, Some(2));
    assert_eq!(blocks[1].text, "Existing noteReference");
    assert_eq!(blocks[1].links[0].url, "https://example.com/retained");
    assert_eq!(blocks[2].rows, vec![vec!["Owner", "Milestone"]]);
}

#[test]
fn appends_only_to_selected_section_and_preserves_original_parts_and_formatting() {
    let original = fixture(body(), "<settings/>");
    let package = WordPackage::open(original.clone()).unwrap();
    let edit: DocumentEdit = serde_json::from_value(json!({
        "paragraphs":[{"text":"A & B <review>","link":"https://example.com/new?a=1&b=2"}]
    }))
    .unwrap();
    let changed = package.append_section("Next meeting", &edit).unwrap();
    let blocks = WordPackage::open(changed.clone()).unwrap().read().unwrap();
    let inserted = blocks
        .iter()
        .position(|b| b.text == "A & B <review>")
        .unwrap();
    let historic = blocks
        .iter()
        .position(|b| b.text == "Older meeting")
        .unwrap();
    assert!(inserted < historic);
    assert_eq!(
        blocks[inserted].links[0].url,
        "https://example.com/new?a=1&b=2"
    );
    for name in [
        "word/settings.xml",
        "word/media/retained.bin",
        "word/comments.xml",
    ] {
        assert_eq!(part(&original, name), part(&changed, name));
    }
    let original_xml = String::from_utf8(part(&original, "word/document.xml")).unwrap();
    let changed_xml = String::from_utf8(part(&changed, "word/document.xml")).unwrap();
    let needle = "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Existing note</w:t>";
    assert!(original_xml.contains(needle) && changed_xml.contains(needle));
    assert!(changed_xml.contains("custom:flag=\"keep\""));
}

#[test]
fn new_section_is_inserted_before_exact_heading_and_ambiguous_edits_fail() {
    let package = WordPackage::open(fixture(body(), "<settings/>")).unwrap();
    let edit: DocumentEdit = serde_json::from_value(json!({
        "create_before":"Next meeting","paragraphs":[{"text":"Review proposed projects"}]
    }))
    .unwrap();
    let changed = package.append_section("Upcoming meeting", &edit).unwrap();
    let blocks = WordPackage::open(changed).unwrap().read().unwrap();
    assert_eq!(blocks[0].text, "Upcoming meeting");
    assert_eq!(blocks[0].heading_level, Some(2));
    assert_eq!(blocks[1].text, "Review proposed projects");
    assert_eq!(blocks[2].text, "Next meeting");
    assert!(package.append_section("Next meeting", &edit).is_err());
    let duplicate =
        WordPackage::open(fixture(&format!("{}{}", body(), body()), "<settings/>")).unwrap();
    assert!(duplicate.append_section("Upcoming meeting", &edit).is_err());
    let unsafe_link: DocumentEdit = serde_json::from_value(json!({
        "paragraphs":[{"text":"Unsafe","link":"javascript:alert(1)"}]
    }))
    .unwrap();
    assert!(
        package
            .append_section("Next meeting", &unsafe_link)
            .is_err()
    );
}

#[test]
fn protected_or_tracked_documents_are_readable_but_edits_are_refused() {
    let edit: DocumentEdit =
        serde_json::from_value(json!({"paragraphs":[{"text":"New note"}]})).unwrap();
    for setting in [
        r#"<w:documentProtection w:enforcement="1"/>"#,
        "<w:trackRevisions/>",
    ] {
        let package = WordPackage::open(fixture(
            body(),
            &format!(r#"<w:settings xmlns:w="{W}">{setting}</w:settings>"#),
        ))
        .unwrap();
        assert!(package.read().is_ok());
        assert!(package.append_section("Next meeting", &edit).is_err());
    }
}

#[test]
fn paragraph_tabs_and_line_breaks_use_word_elements_and_round_trip() {
    let package = WordPackage::open(fixture(body(), "<settings/>")).unwrap();
    let edit: DocumentEdit = serde_json::from_value(json!({
        "paragraphs":[{"text":"Before\tAfter\nNext & last"}]
    }))
    .unwrap();
    let changed = package.append_section("Next meeting", &edit).unwrap();
    let document = String::from_utf8(part(&changed, "word/document.xml")).unwrap();
    assert!(document.contains("<w:tab/>"));
    assert!(document.contains("<w:br/>"));
    let blocks = WordPackage::open(changed).unwrap().read().unwrap();
    assert!(
        blocks
            .iter()
            .any(|b| b.text == "Before\tAfter\nNext & last")
    );
}

#[test]
fn canonical_word_links_resolve_exact_identity_and_reject_foreign_or_ambiguous_inputs() {
    let id = "12345678-1234-1234-1234-123456789abc";
    for url in [
        format!(
            "https://tenant.sharepoint.com/sites/Project/_layouts/15/Doc.aspx?sourcedoc=%7B{id}%7D"
        ),
        format!(
            "https://tenant.sharepoint.com/:w:/r/sites/Project/Shared%20Documents/file.docx?d=w{}",
            id.replace('-', "")
        ),
    ] {
        let location = DocumentLocation::parse(&url).unwrap().unwrap();
        assert_eq!(location.id, id);
        assert_eq!(location.base, "https://tenant.sharepoint.com/sites/Project");
    }
    for url in [
        "https://tenant.sharepoint.com.evil.example/sites/Project/?sourcedoc=12345678-1234-1234-1234-123456789abc",
        "http://tenant.sharepoint.com/sites/Project/?sourcedoc=12345678-1234-1234-1234-123456789abc",
        "https://user@tenant.sharepoint.com/sites/Project/?sourcedoc=12345678-1234-1234-1234-123456789abc",
        "https://tenant.sharepoint.com/sites/Project/Doc.aspx?sourcedoc=bad",
    ] {
        assert!(DocumentLocation::parse(url).is_err());
    }
}
