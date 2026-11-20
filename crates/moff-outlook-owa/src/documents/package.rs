use std::{
    collections::{HashMap, HashSet},
    io::{Cursor, Read, Write},
};

use roxmltree::{Document, Node};
use serde::{Deserialize, Serialize};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

use crate::OutlookWebError;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOCUMENT: &str = "word/document.xml";
const RELATIONSHIPS: &str = "word/_rels/document.xml.rels";
pub const MAX_DOCUMENT_BYTES: u64 = 16_000_000;
const MAX_EXPANDED_BYTES: u64 = 64_000_000;

fn invalid(message: &'static str) -> OutlookWebError {
    OutlookWebError::InvalidInput(message)
}

fn xml(source: &str) -> Result<Document<'_>, OutlookWebError> {
    Document::parse(source).map_err(|_| invalid("invalid or unsupported document XML"))
}

pub struct WordPackage {
    bytes: Vec<u8>,
    parts: HashMap<String, String>,
    signed: bool,
}

#[derive(Debug, Serialize)]
pub struct DocumentBlock {
    pub kind: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub heading_level: Option<u8>,
    pub text: String,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub links: Vec<DocumentLink>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub rows: Vec<Vec<String>>,
}

#[derive(Debug, Serialize)]
pub struct DocumentLink {
    pub text: String,
    pub url: String,
}

#[derive(Debug, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct DocumentEdit {
    /// Create a new section immediately before this unique existing heading.
    pub create_before: Option<String>,
    pub paragraphs: Vec<DocumentParagraph>,
}

#[derive(Debug, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct DocumentParagraph {
    pub text: String,
    pub heading_level: Option<u8>,
    pub link: Option<String>,
    #[serde(default)]
    pub bold: bool,
}

impl WordPackage {
    pub fn open(bytes: Vec<u8>) -> Result<Self, OutlookWebError> {
        if bytes.len() as u64 > MAX_DOCUMENT_BYTES {
            return Err(invalid("Word documents must not exceed 16000000 bytes"));
        }
        let mut archive = ZipArchive::new(Cursor::new(&bytes))
            .map_err(|_| invalid("expected an unencrypted DOCX package"))?;
        if archive.len() > 2048 {
            return Err(invalid("DOCX package contains too many parts"));
        }
        let mut seen = HashSet::new();
        let mut total = 0_u64;
        let mut parts = HashMap::new();
        let mut signed = false;
        for index in 0..archive.len() {
            let mut file = archive
                .by_index(index)
                .map_err(|_| invalid("cannot read DOCX package part"))?;
            let name = file
                .name()
                .map_err(|_| invalid("invalid DOCX part name"))?
                .into_owned();
            if !seen.insert(name.clone()) {
                return Err(invalid("duplicate DOCX package parts are unsupported"));
            }
            total = total
                .checked_add(file.size())
                .filter(|total| *total <= MAX_EXPANDED_BYTES)
                .ok_or(invalid("expanded DOCX package exceeds 64000000 bytes"))?;
            signed |= name.starts_with("_xmlsignatures/")
                || name.to_ascii_lowercase().contains("vbaproject");
            if [
                DOCUMENT,
                RELATIONSHIPS,
                "word/styles.xml",
                "word/settings.xml",
            ]
            .contains(&name.as_str())
            {
                if file.size() > 8_000_000 {
                    return Err(invalid("document XML part exceeds 8000000 bytes"));
                }
                let mut source = String::new();
                file.by_ref()
                    .take(8_000_001)
                    .read_to_string(&mut source)
                    .map_err(|_| invalid("document XML must be valid UTF-8"))?;
                if source.len() > 8_000_000 {
                    return Err(invalid("document XML part exceeds 8000000 bytes"));
                }
                xml(&source)?;
                parts.insert(name, source);
            }
        }
        if !parts.contains_key(DOCUMENT) {
            return Err(invalid("DOCX package has no document body"));
        }
        drop(archive);
        Ok(Self {
            bytes,
            parts,
            signed,
        })
    }

    fn styles(&self) -> Result<HashMap<String, u8>, OutlookWebError> {
        let Some(source) = self.parts.get("word/styles.xml") else {
            return Ok(HashMap::new());
        };
        let doc = xml(source)?;
        let mut levels = HashMap::new();
        for style in doc.descendants().filter(|n| n.has_tag_name((W, "style"))) {
            let Some(id) = style.attribute((W, "styleId")) else {
                continue;
            };
            let level = style
                .descendants()
                .find(|n| n.has_tag_name((W, "outlineLvl")))
                .and_then(|n| n.attribute((W, "val")))
                .and_then(|v| v.parse::<u8>().ok())
                .filter(|level| *level < 9)
                .map(|level| level + 1)
                .or_else(|| heading_style(id));
            if let Some(level) = level {
                levels.insert(id.to_owned(), level);
            }
        }
        Ok(levels)
    }

    fn relationships(&self) -> Result<HashMap<String, String>, OutlookWebError> {
        let Some(source) = self.parts.get(RELATIONSHIPS) else {
            return Ok(HashMap::new());
        };
        let doc = xml(source)?;
        Ok(doc
            .descendants()
            .filter(|n| {
                n.has_tag_name((RELS, "Relationship"))
                    && n.attribute("Type") == Some(&format!("{R}/hyperlink"))
                    && n.attribute("TargetMode") == Some("External")
            })
            .filter_map(|n| {
                Some((
                    n.attribute("Id")?.to_owned(),
                    n.attribute("Target")?.to_owned(),
                ))
            })
            .collect())
    }

    pub fn read(&self) -> Result<Vec<DocumentBlock>, OutlookWebError> {
        let doc = xml(&self.parts[DOCUMENT])?;
        let body = body(&doc)?;
        let styles = self.styles()?;
        let relationships = self.relationships()?;
        let mut blocks = Vec::new();
        for node in body.children().filter(Node::is_element) {
            if node.has_tag_name((W, "p")) {
                blocks.push(paragraph(node, &styles, &relationships));
            } else if node.has_tag_name((W, "tbl")) {
                let rows = node
                    .children()
                    .filter(|n| n.has_tag_name((W, "tr")))
                    .map(|row| {
                        row.children()
                            .filter(|n| n.has_tag_name((W, "tc")))
                            .map(visible_text)
                            .collect::<Vec<_>>()
                    })
                    .collect::<Vec<_>>();
                blocks.push(DocumentBlock {
                    kind: "table",
                    heading_level: None,
                    text: rows
                        .iter()
                        .map(|r| r.join("\t"))
                        .collect::<Vec<_>>()
                        .join("\n"),
                    links: links(node, &relationships),
                    rows,
                });
            } else if node.has_tag_name((W, "sdt")) {
                for p in node.descendants().filter(|n| n.has_tag_name((W, "p"))) {
                    blocks.push(paragraph(p, &styles, &relationships));
                }
            }
        }
        Ok(blocks)
    }

    pub fn append_section(
        &self,
        section: &str,
        edit: &DocumentEdit,
    ) -> Result<Vec<u8>, OutlookWebError> {
        self.editable()?;
        validate_edit(section, edit)?;
        let source = &self.parts[DOCUMENT];
        let doc = xml(source)?;
        let body = body(&doc)?;
        let prefix = body
            .lookup_prefix(W)
            .filter(|prefix| !prefix.is_empty())
            .ok_or(invalid("editing requires a named Word XML namespace"))?;
        let styles = self.styles()?;
        let (offset, create_level) = section_offset(body, section, edit, &styles)?;
        let mut relationships = self
            .parts
            .get(RELATIONSHIPS)
            .cloned()
            .unwrap_or_else(|| format!("<Relationships xmlns=\"{RELS}\"></Relationships>"));
        let mut inserted = String::new();
        if let Some(level) = create_level {
            inserted.push_str(&render_paragraph(
                prefix,
                &DocumentParagraph {
                    text: section.to_owned(),
                    heading_level: Some(level),
                    link: None,
                    bold: false,
                },
                &styles,
                None,
            ));
        }
        for (index, p) in edit.paragraphs.iter().enumerate() {
            let id = if let Some(link) = &p.link {
                let mut id = format!("moffLink{index}");
                while relationships.contains(&format!("\"{id}\""))
                    || relationships.contains(&format!("'{id}'"))
                {
                    id.push('x');
                }
                let relationship = format!(
                    "<Relationship xmlns=\"{RELS}\" Id=\"{id}\" Type=\"{R}/hyperlink\" Target=\"{}\" TargetMode=\"External\"/>",
                    escape(link)
                );
                relationships = insert_before_close(&relationships, &relationship)?;
                Some(id)
            } else {
                None
            };
            inserted.push_str(&render_paragraph(prefix, p, &styles, id.as_deref()));
        }
        let mut changed = source.clone();
        changed.insert_str(offset, &inserted);
        xml(&changed)?;
        self.repack(&changed, &relationships)
    }

    fn editable(&self) -> Result<(), OutlookWebError> {
        if self.signed {
            return Err(invalid(
                "signed or macro-enabled documents cannot be edited",
            ));
        }
        if let Some(settings) = self.parts.get("word/settings.xml") {
            let settings = xml(settings)?;
            if settings.descendants().any(|n| {
                (n.has_tag_name((W, "documentProtection"))
                    && !matches!(n.attribute((W, "enforcement")), Some("0" | "false" | "off")))
                    || (n.has_tag_name((W, "trackRevisions"))
                        && !matches!(n.attribute((W, "val")), Some("0" | "false" | "off")))
            }) {
                return Err(invalid(
                    "protected documents or enabled track changes require editing in Word",
                ));
            }
        }
        Ok(())
    }

    fn repack(&self, document: &str, relationships: &str) -> Result<Vec<u8>, OutlookWebError> {
        let mut input = ZipArchive::new(Cursor::new(&self.bytes))
            .map_err(|_| OutlookWebError::InvalidResponse)?;
        let mut output = ZipWriter::new(Cursor::new(Vec::new()));
        for index in 0..input.len() {
            let file = input
                .by_index(index)
                .map_err(|_| OutlookWebError::InvalidResponse)?;
            let name = file
                .name()
                .map_err(|_| OutlookWebError::InvalidResponse)?
                .into_owned();
            let replacement = match name.as_str() {
                DOCUMENT => Some(document),
                RELATIONSHIPS => Some(relationships),
                _ => None,
            };
            if let Some(replacement) = replacement {
                output
                    .start_file(&name, SimpleFileOptions::default())
                    .map_err(|_| OutlookWebError::InvalidResponse)?;
                output
                    .write_all(replacement.as_bytes())
                    .map_err(OutlookWebError::FileTransferIo)?;
            } else {
                output
                    .raw_copy_file(file)
                    .map_err(|_| OutlookWebError::InvalidResponse)?;
            }
        }
        if !self.parts.contains_key(RELATIONSHIPS) {
            output
                .start_file(RELATIONSHIPS, SimpleFileOptions::default())
                .map_err(|_| OutlookWebError::InvalidResponse)?;
            output
                .write_all(relationships.as_bytes())
                .map_err(OutlookWebError::FileTransferIo)?;
        }
        let result = output
            .finish()
            .map_err(|_| OutlookWebError::InvalidResponse)?
            .into_inner();
        if result.len() as u64 > MAX_DOCUMENT_BYTES {
            return Err(invalid("edited document exceeds 16000000 bytes"));
        }
        Ok(result)
    }
}

fn body<'a>(doc: &'a Document<'a>) -> Result<Node<'a, 'a>, OutlookWebError> {
    doc.descendants()
        .find(|n| n.has_tag_name((W, "body")))
        .ok_or(invalid("Word document has no body"))
}

fn heading_style(style: &str) -> Option<u8> {
    style
        .to_ascii_lowercase()
        .strip_prefix("heading")?
        .parse::<u8>()
        .ok()
        .filter(|level| (1..=9).contains(level))
}

fn level(node: Node<'_, '_>, styles: &HashMap<String, u8>) -> Option<u8> {
    let properties = node.children().find(|n| n.has_tag_name((W, "pPr")))?;
    properties
        .children()
        .find(|n| n.has_tag_name((W, "outlineLvl")))
        .and_then(|n| n.attribute((W, "val")))
        .and_then(|value| value.parse::<u8>().ok())
        .filter(|level| *level < 9)
        .map(|level| level + 1)
        .or_else(|| {
            properties
                .children()
                .find(|n| n.has_tag_name((W, "pStyle")))
                .and_then(|n| n.attribute((W, "val")))
                .and_then(|id| styles.get(id).copied().or_else(|| heading_style(id)))
        })
}

fn visible_text(node: Node<'_, '_>) -> String {
    node.descendants()
        .filter(|n| {
            !n.ancestors()
                .any(|p| p.has_tag_name((W, "del")) || p.has_tag_name((W, "moveFrom")))
        })
        .filter_map(|n| {
            if n.has_tag_name((W, "t")) {
                n.text()
            } else if n.has_tag_name((W, "tab")) {
                Some("\t")
            } else if n.has_tag_name((W, "br")) {
                Some("\n")
            } else {
                None
            }
        })
        .collect()
}

fn links(node: Node<'_, '_>, relationships: &HashMap<String, String>) -> Vec<DocumentLink> {
    node.descendants()
        .filter(|n| n.has_tag_name((W, "hyperlink")))
        .filter_map(|n| {
            Some(DocumentLink {
                text: visible_text(n),
                url: relationships.get(n.attribute((R, "id"))?)?.clone(),
            })
        })
        .collect()
}

fn paragraph(
    node: Node<'_, '_>,
    styles: &HashMap<String, u8>,
    relationships: &HashMap<String, String>,
) -> DocumentBlock {
    let heading_level = level(node, styles);
    DocumentBlock {
        kind: if heading_level.is_some() {
            "heading"
        } else {
            "paragraph"
        },
        heading_level,
        text: visible_text(node),
        links: links(node, relationships),
        rows: Vec::new(),
    }
}

fn section_offset(
    body: Node<'_, '_>,
    section: &str,
    edit: &DocumentEdit,
    styles: &HashMap<String, u8>,
) -> Result<(usize, Option<u8>), OutlookWebError> {
    let headings = body
        .children()
        .filter_map(|node| Some((node, level(node, styles)?)))
        .collect::<Vec<_>>();
    let selected = headings
        .iter()
        .filter(|(node, _)| visible_text(*node).trim() == section)
        .collect::<Vec<_>>();
    if let Some(before) = &edit.create_before {
        if !selected.is_empty() {
            return Err(invalid(
                "new section already exists; omit create_before to append to it",
            ));
        }
        let anchor = headings
            .iter()
            .filter(|(node, _)| visible_text(*node).trim() == before)
            .collect::<Vec<_>>();
        if anchor.len() != 1 {
            return Err(invalid(
                "create_before must identify exactly one existing heading",
            ));
        }
        return Ok((anchor[0].0.range().start, Some(anchor[0].1)));
    }
    if selected.len() != 1 {
        return Err(invalid(
            "section must identify exactly one existing heading",
        ));
    }
    let (selected, selected_level) = *selected[0];
    let next = headings.iter().find(|(node, level)| {
        node.range().start > selected.range().start && *level <= selected_level
    });
    let offset = next.map_or_else(
        || {
            body.children()
                .find(|n| n.has_tag_name((W, "sectPr")))
                .map_or_else(
                    || {
                        body.range().end
                            - format!(
                                "</{}body>",
                                body.lookup_prefix(W)
                                    .map_or(String::new(), |p| format!("{p}:"))
                            )
                            .len()
                    },
                    |node| node.range().start,
                )
        },
        |(node, _)| node.range().start,
    );
    Ok((offset, None))
}

fn validate_edit(section: &str, edit: &DocumentEdit) -> Result<(), OutlookWebError> {
    if section.trim().is_empty()
        || section.trim() != section
        || section.len() > 500
        || edit.paragraphs.is_empty()
        || edit.paragraphs.len() > 100
        || edit.paragraphs.iter().map(|p| p.text.len()).sum::<usize>() > 50_000
    {
        return Err(invalid(
            "provide an exact section heading and 1-100 paragraphs, at most 50000 text bytes",
        ));
    }
    for p in &edit.paragraphs {
        if p.text.trim().is_empty()
            || p.text.chars().any(|c| c < ' ' && c != '\t' && c != '\n')
            || p.heading_level
                .is_some_and(|level| !(1..=9).contains(&level))
        {
            return Err(invalid("invalid document paragraph text or heading level"));
        }
        if let Some(link) = &p.link {
            let url =
                url::Url::parse(link).map_err(|_| invalid("document links must be HTTPS URLs"))?;
            if url.scheme() != "https"
                || url.host_str().is_none()
                || !url.username().is_empty()
                || url.password().is_some()
                || url.port().is_some()
            {
                return Err(invalid(
                    "document links must be HTTPS URLs without credentials",
                ));
            }
        }
    }
    Ok(())
}

fn escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn render_paragraph(
    prefix: &str,
    paragraph: &DocumentParagraph,
    styles: &HashMap<String, u8>,
    relationship: Option<&str>,
) -> String {
    let properties = if let Some(level) = paragraph.heading_level {
        let preferred = format!("Heading{level}");
        let style = if styles.get(&preferred) == Some(&level) {
            Some(preferred)
        } else {
            let mut matching = styles
                .iter()
                .filter(|(_, l)| **l == level)
                .map(|(id, _)| id.clone())
                .collect::<Vec<_>>();
            matching.sort();
            matching.into_iter().next()
        };
        style.map_or_else(
            || {
                format!(
                    "<{prefix}:pPr><{prefix}:outlineLvl {prefix}:val=\"{}\"/></{prefix}:pPr>",
                    level - 1
                )
            },
            |id| {
                format!(
                    "<{prefix}:pPr><{prefix}:pStyle {prefix}:val=\"{}\"/></{prefix}:pPr>",
                    escape(&id)
                )
            },
        )
    } else {
        String::new()
    };
    let bold = if paragraph.bold {
        format!("<{prefix}:b/>")
    } else {
        String::new()
    };
    let hyperlink = if relationship.is_some() {
        format!("<{prefix}:color {prefix}:val=\"0563C1\"/><{prefix}:u {prefix}:val=\"single\"/>")
    } else {
        String::new()
    };
    let run_properties = format!("{bold}{hyperlink}");
    let text = escape(&paragraph.text)
        .replace(
            '\t',
            &format!("</{prefix}:t><{prefix}:tab/><{prefix}:t xml:space=\"preserve\">"),
        )
        .replace(
            '\n',
            &format!("</{prefix}:t><{prefix}:br/><{prefix}:t xml:space=\"preserve\">"),
        );
    let run = format!(
        "<{prefix}:r><{prefix}:rPr>{run_properties}</{prefix}:rPr><{prefix}:t xml:space=\"preserve\">{text}</{prefix}:t></{prefix}:r>"
    );
    let content = relationship.map_or_else(|| run.clone(), |id| format!("<{prefix}:hyperlink xmlns:moffrel=\"{R}\" moffrel:id=\"{id}\">{run}</{prefix}:hyperlink>"));
    format!("<{prefix}:p>{properties}{content}</{prefix}:p>")
}

fn insert_before_close(source: &str, inserted: &str) -> Result<String, OutlookWebError> {
    let doc = xml(source)?;
    let root = doc.root_element();
    let range = root.range();
    let mut changed = source.to_owned();
    if source[range.clone()].trim_end().ends_with("/>") {
        let close = range.end - 2;
        let prefix = root.lookup_prefix(root.tag_name().namespace().unwrap_or(""));
        let name = prefix.map_or_else(
            || root.tag_name().name().to_owned(),
            |p| format!("{p}:{}", root.tag_name().name()),
        );
        changed.replace_range(close..range.end, &format!(">{inserted}</{name}>"));
    } else {
        let offset = source[..range.end]
            .rfind("</")
            .ok_or(OutlookWebError::InvalidResponse)?;
        changed.insert_str(offset, inserted);
    }
    xml(&changed)?;
    Ok(changed)
}
