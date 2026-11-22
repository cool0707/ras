//! A deliberately small XML tag scanner.
//!
//! The OPC parts this crate looks at (`[Content_Types].xml`, relationship
//! parts, `workbook.xml` and the first element of each sheet) are only
//! inspected for a few attributes and edited by splicing text at known
//! positions, so that every byte the crate does not have to change stays
//! exactly as it was. A full XML model is not needed for that.

use crate::error::OvbaError;

/// What a tag is.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum TagKind {
    /// `<name ...>`
    Open,
    /// `</name>`
    Close,
    /// `<name .../>`
    Empty,
}

/// A start, end or empty-element tag and its position in the document.
#[derive(Clone, Debug)]
pub(crate) struct Tag<'a> {
    pub kind: TagKind,
    /// Qualified name (with prefix, if any).
    pub name: &'a str,
    /// Byte offset of `<`.
    pub start: usize,
    /// Byte offset just after `>`.
    pub end: usize,
    /// Byte offset just after the tag name (where an attribute can be
    /// inserted).
    pub name_end: usize,
    doc: &'a str,
}

/// An attribute of a tag.
#[derive(Clone, Debug)]
pub(crate) struct Attribute<'a> {
    pub name: &'a str,
    /// Byte range of the value (between the quotes) in the document.
    pub value_start: usize,
    pub value_end: usize,
    doc: &'a str,
}

impl<'a> Attribute<'a> {
    /// The unescaped value.
    pub(crate) fn value(&self) -> String {
        unescape(&self.doc[self.value_start..self.value_end])
    }

    pub(crate) fn local_name(&self) -> &'a str {
        local_name(self.name)
    }
}

pub(crate) fn local_name(qualified: &str) -> &str {
    match qualified.rfind(':') {
        Some(i) => &qualified[i + 1..],
        None => qualified,
    }
}

impl<'a> Tag<'a> {
    pub(crate) fn local_name(&self) -> &'a str {
        local_name(self.name)
    }

    /// Namespace prefix including the colon (`"x:"`), or `""`.
    pub(crate) fn prefix(&self) -> &'a str {
        match self.name.rfind(':') {
            Some(i) => &self.name[..=i],
            None => "",
        }
    }

    /// The attributes of a start or empty-element tag.
    pub(crate) fn attributes(&self) -> Vec<Attribute<'a>> {
        let mut result = Vec::new();
        if self.kind == TagKind::Close {
            return result;
        }
        let bytes = self.doc.as_bytes();
        let limit = self.end.saturating_sub(1);
        let mut i = self.name_end;
        loop {
            while i < limit && (bytes[i].is_ascii_whitespace() || bytes[i] == b'/') {
                i += 1;
            }
            if i >= limit {
                break;
            }
            let name_start = i;
            while i < limit && bytes[i] != b'=' && !bytes[i].is_ascii_whitespace() {
                i += 1;
            }
            let name = &self.doc[name_start..i];
            while i < limit && bytes[i].is_ascii_whitespace() {
                i += 1;
            }
            if i >= limit || bytes[i] != b'=' {
                break;
            }
            i += 1;
            while i < limit && bytes[i].is_ascii_whitespace() {
                i += 1;
            }
            if i >= limit || (bytes[i] != b'"' && bytes[i] != b'\'') {
                break;
            }
            let quote = bytes[i];
            let value_start = i + 1;
            let mut j = value_start;
            while j < limit && bytes[j] != quote {
                j += 1;
            }
            if j >= limit {
                break;
            }
            result.push(Attribute {
                name,
                value_start,
                value_end: j,
                doc: self.doc,
            });
            i = j + 1;
        }
        result
    }

    /// Attribute with the given qualified name.
    pub(crate) fn attribute(&self, name: &str) -> Option<Attribute<'a>> {
        self.attributes().into_iter().find(|a| a.name == name)
    }

    /// Unescaped value of the attribute with the given qualified name.
    pub(crate) fn attr(&self, name: &str) -> Option<String> {
        self.attribute(name).map(|a| a.value())
    }
}

/// Iterates over the element tags of a document, skipping the XML
/// declaration, processing instructions, comments, CDATA sections and
/// DOCTYPE declarations.
pub(crate) struct Tags<'a> {
    doc: &'a str,
    pos: usize,
}

pub(crate) fn tags(doc: &str) -> Tags<'_> {
    Tags { doc, pos: 0 }
}

impl<'a> Iterator for Tags<'a> {
    type Item = Tag<'a>;

    fn next(&mut self) -> Option<Tag<'a>> {
        let bytes = self.doc.as_bytes();
        loop {
            let lt = self.pos + self.doc.get(self.pos..)?.find('<')?;
            let rest = &self.doc[lt..];
            if let Some(comment) = rest.strip_prefix("<!--") {
                self.pos = lt + 4 + comment.find("-->")? + 3;
                continue;
            }
            if let Some(cdata) = rest.strip_prefix("<![CDATA[") {
                self.pos = lt + 9 + cdata.find("]]>")? + 3;
                continue;
            }
            if let Some(instruction) = rest.strip_prefix("<?") {
                self.pos = lt + 2 + instruction.find("?>")? + 2;
                continue;
            }
            if rest.starts_with("<!") {
                self.pos = lt + rest.find('>')? + 1;
                continue;
            }
            // Find the closing `>` outside of quoted attribute values.
            let mut i = lt + 1;
            let mut quote: Option<u8> = None;
            let gt = loop {
                let b = *bytes.get(i)?;
                match quote {
                    Some(q) if b == q => quote = None,
                    Some(_) => {}
                    None if b == b'"' || b == b'\'' => quote = Some(b),
                    None if b == b'>' => break i,
                    None => {}
                }
                i += 1;
            };
            self.pos = gt + 1;
            let (kind, name_start) = if bytes.get(lt + 1) == Some(&b'/') {
                (TagKind::Close, lt + 2)
            } else if gt > lt + 1 && bytes[gt - 1] == b'/' {
                (TagKind::Empty, lt + 1)
            } else {
                (TagKind::Open, lt + 1)
            };
            let mut name_end = name_start;
            while name_end < gt
                && !bytes[name_end].is_ascii_whitespace()
                && bytes[name_end] != b'/'
                && bytes[name_end] != b'>'
            {
                name_end += 1;
            }
            if name_end == name_start {
                continue;
            }
            return Some(Tag {
                kind,
                name: &self.doc[name_start..name_end],
                start: lt,
                end: gt + 1,
                name_end,
                doc: self.doc,
            });
        }
    }
}

/// Resolves the five predefined entities and numeric character references.
pub(crate) fn unescape(text: &str) -> String {
    if !text.contains('&') {
        return text.to_string();
    }
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(amp) = rest.find('&') {
        out.push_str(&rest[..amp]);
        let tail = &rest[amp..];
        let Some(semi) = tail.find(';') else {
            out.push_str(tail);
            return out;
        };
        let entity = &tail[1..semi];
        let decoded = match entity {
            "amp" => Some('&'),
            "lt" => Some('<'),
            "gt" => Some('>'),
            "quot" => Some('"'),
            "apos" => Some('\''),
            _ => {
                let code =
                    if let Some(hex) = entity.strip_prefix("#x").or(entity.strip_prefix("#X")) {
                        u32::from_str_radix(hex, 16).ok()
                    } else if let Some(dec) = entity.strip_prefix('#') {
                        dec.parse::<u32>().ok()
                    } else {
                        None
                    };
                code.and_then(char::from_u32)
            }
        };
        match decoded {
            Some(c) => {
                out.push(c);
                rest = &tail[semi + 1..];
            }
            None => {
                out.push('&');
                rest = &tail[1..];
            }
        }
    }
    out.push_str(rest);
    out
}

/// Escapes text for use inside a double-quoted attribute value.
pub(crate) fn escape_attr(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for c in text.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            other => out.push(other),
        }
    }
    out
}

/// Decodes the bytes of an XML part: UTF-8 (with or without BOM) or
/// UTF-16 with BOM. Returns the text and whether it was UTF-8 (only UTF-8
/// parts are ever edited).
pub(crate) fn decode_part(bytes: &[u8], part: &str) -> Result<(String, bool), OvbaError> {
    if bytes.starts_with(&[0xFF, 0xFE]) || bytes.starts_with(&[0xFE, 0xFF]) {
        let big_endian = bytes[0] == 0xFE;
        let units: Vec<u16> = bytes[2..]
            .chunks_exact(2)
            .map(|c| {
                if big_endian {
                    u16::from_be_bytes([c[0], c[1]])
                } else {
                    u16::from_le_bytes([c[0], c[1]])
                }
            })
            .collect();
        return Ok((String::from_utf16_lossy(&units), false));
    }
    match std::str::from_utf8(bytes) {
        Ok(text) => Ok((text.to_string(), true)),
        Err(_) => Err(OvbaError::Format(format!(
            "package part `{part}` is not valid UTF-8 XML"
        ))),
    }
}

/// A set of text replacements applied to a document in one go.
#[derive(Default)]
pub(crate) struct Edits {
    edits: Vec<(usize, usize, String)>,
}

impl Edits {
    pub(crate) fn replace(&mut self, start: usize, end: usize, text: impl Into<String>) {
        self.edits.push((start, end, text.into()));
    }

    pub(crate) fn insert(&mut self, at: usize, text: impl Into<String>) {
        self.edits.push((at, at, text.into()));
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.edits.is_empty()
    }

    /// Applies the edits. Insertions at the same position keep the order in
    /// which they were added.
    pub(crate) fn apply(mut self, doc: &str) -> String {
        // Stable sort by position; overlapping replacements are not expected.
        self.edits.sort_by_key(|(start, _, _)| *start);
        let mut out = String::with_capacity(doc.len() + 256);
        let mut pos = 0usize;
        for (start, end, text) in self.edits {
            if start < pos
                || end > doc.len()
                || !doc.is_char_boundary(start)
                || !doc.is_char_boundary(end)
            {
                continue;
            }
            out.push_str(&doc[pos..start]);
            out.push_str(&text);
            pos = end;
        }
        out.push_str(&doc[pos..]);
        out
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scans_tags_and_attributes() {
        let doc = "<?xml version=\"1.0\"?>\n<!-- c <x> -->\n<a:root xmlns:a=\"u\" v='1 > 0'><b k=\"&lt;&amp;&#x41;&#66;\" /><![CDATA[<no>]]><c>t</c></a:root>";
        let all: Vec<Tag> = tags(doc).collect();
        let names: Vec<(&str, TagKind)> = all.iter().map(|t| (t.name, t.kind)).collect();
        assert_eq!(
            names,
            vec![
                ("a:root", TagKind::Open),
                ("b", TagKind::Empty),
                ("c", TagKind::Open),
                ("c", TagKind::Close),
                ("a:root", TagKind::Close),
            ]
        );
        assert_eq!(all[0].local_name(), "root");
        assert_eq!(all[0].prefix(), "a:");
        assert_eq!(all[0].attr("v").as_deref(), Some("1 > 0"));
        assert_eq!(all[0].attr("xmlns:a").as_deref(), Some("u"));
        assert_eq!(all[1].attr("k").as_deref(), Some("<&AB"));
        assert_eq!(all[1].attr("missing"), None);
        assert_eq!(
            &doc[all[1].start..all[1].end],
            "<b k=\"&lt;&amp;&#x41;&#66;\" />"
        );
    }

    #[test]
    fn broken_documents_do_not_panic() {
        for doc in [
            "<",
            "<a",
            "<a b=\"",
            "<!-- x",
            "<a b=c>",
            "</>",
            "<a 日本=\"x\"/>",
            "&#xZZ;",
        ] {
            let _: usize = tags(doc).map(|t| t.attributes().len()).sum();
            let _ = unescape(doc);
        }
    }

    #[test]
    fn edits_apply_in_order() {
        let doc = "0123456789";
        let mut e = Edits::default();
        e.insert(10, "!");
        e.replace(2, 4, "ab");
        e.insert(0, "<");
        e.insert(10, "?");
        assert_eq!(e.apply(doc), "<01ab456789!?");
        assert_eq!(escape_attr("a\"<&>"), "a&quot;&lt;&amp;&gt;");
    }
}
