//! XML read as text with positions: the parts of a workbook an edit changes are spliced at known
//! places, so that everything the edit does not concern stays byte for byte what it was.

use std::ops::Range;

/// A start, end or empty-element tag.
#[derive(Clone, Debug)]
pub(crate) struct Tag<'a> {
    /// `<name ...>`: 0, `</name>`: 1, `<name .../>`: 2.
    kind: u8,
    /// Qualified name.
    pub name: &'a str,
    pub start: usize,
    /// Just after `>`.
    pub end: usize,
}

impl Tag<'_> {
    pub fn is_open(&self) -> bool {
        self.kind == 0
    }
    pub fn is_close(&self) -> bool {
        self.kind == 1
    }
    pub fn is_empty(&self) -> bool {
        self.kind == 2
    }
    pub fn local(&self) -> &str {
        local(self.name)
    }
}

pub(crate) fn local(name: &str) -> &str {
    name.rsplit(':').next().unwrap_or(name)
}

/// The namespace prefix of a qualified name with its colon (`x:`), or nothing.
pub(crate) fn prefix(name: &str) -> &str {
    match name.rfind(':') {
        Some(i) => &name[..=i],
        None => "",
    }
}

/// The tags of `doc` from `from` on, in order (comments, processing instructions and CDATA
/// sections are passed over).
pub(crate) struct Tags<'a> {
    doc: &'a str,
    at: usize,
}

pub(crate) fn tags(doc: &str, from: usize) -> Tags<'_> {
    Tags { doc, at: from }
}

impl<'a> Iterator for Tags<'a> {
    type Item = Tag<'a>;

    fn next(&mut self) -> Option<Tag<'a>> {
        let bytes = self.doc.as_bytes();
        loop {
            let start = self.at + self.doc.get(self.at..)?.find('<')?;
            let rest = &self.doc[start..];
            let skip = |close: &str| rest.find(close).map(|i| start + i + close.len());
            if rest.starts_with("<!--") {
                self.at = skip("-->")?;
                continue;
            }
            if rest.starts_with("<![CDATA[") {
                self.at = skip("]]>")?;
                continue;
            }
            if rest.starts_with("<?") {
                self.at = skip("?>")?;
                continue;
            }
            if rest.starts_with("<!") {
                self.at = skip(">")?;
                continue;
            }
            let close = rest.starts_with("</");
            let name_start = start + if close { 2 } else { 1 };
            let mut i = name_start;
            while i < bytes.len() && !matches!(bytes[i], b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>')
            {
                i += 1;
            }
            let name = &self.doc[name_start..i];
            // To the `>` outside quotes.
            let mut quote = 0u8;
            while i < bytes.len() {
                match bytes[i] {
                    b'"' | b'\'' if quote == 0 => quote = bytes[i],
                    q if q == quote => quote = 0,
                    b'>' if quote == 0 => break,
                    _ => {}
                }
                i += 1;
            }
            if i >= bytes.len() {
                return None;
            }
            let end = i + 1;
            self.at = end;
            let kind = if close {
                1
            } else if bytes[i - 1] == b'/' {
                2
            } else {
                0
            };
            return Some(Tag {
                kind,
                name,
                start,
                end,
            });
        }
    }
}

/// The attributes of the start tag at `tag` (its span in `doc`): name and the span of the value.
pub(crate) fn attrs(doc: &str, tag: Range<usize>) -> Vec<(&str, Range<usize>)> {
    let bytes = doc.as_bytes();
    let limit = tag.end.saturating_sub(1);
    let mut out = Vec::new();
    let mut i = tag.start + 1;
    while i < limit && !bytes[i].is_ascii_whitespace() && bytes[i] != b'/' {
        i += 1;
    }
    loop {
        while i < limit && (bytes[i].is_ascii_whitespace() || bytes[i] == b'/') {
            i += 1;
        }
        let name_start = i;
        while i < limit && bytes[i] != b'=' && !bytes[i].is_ascii_whitespace() {
            i += 1;
        }
        let name = &doc[name_start..i];
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
        if j > limit || bytes.get(j) != Some(&quote) {
            break;
        }
        out.push((name, value_start..j));
        i = j + 1;
    }
    out
}

/// The unescaped value of the attribute `name` (a local name) of the start tag at `tag`.
pub(crate) fn attr(doc: &str, tag: Range<usize>, name: &str) -> Option<String> {
    attr_span(doc, tag, name).map(|r| unescape(&doc[r]))
}

pub(crate) fn attr_span(doc: &str, tag: Range<usize>, name: &str) -> Option<Range<usize>> {
    attrs(doc, tag)
        .into_iter()
        .find(|(n, _)| *n == name || (local(n) == name && !n.starts_with("xmlns")))
        .map(|(_, r)| r)
}

pub(crate) fn unescape(text: &str) -> String {
    if !text.contains('&') {
        return text.to_string();
    }
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(i) = rest.find('&') {
        out.push_str(&rest[..i]);
        rest = &rest[i..];
        let Some(j) = rest.find(';').filter(|j| *j <= 10) else {
            out.push('&');
            rest = &rest[1..];
            continue;
        };
        let entity = &rest[1..j];
        let ch = match entity {
            "amp" => Some('&'),
            "lt" => Some('<'),
            "gt" => Some('>'),
            "quot" => Some('"'),
            "apos" => Some('\''),
            _ => entity
                .strip_prefix("#x")
                .or_else(|| entity.strip_prefix("#X"))
                .map(|h| u32::from_str_radix(h, 16))
                .or_else(|| entity.strip_prefix('#').map(str::parse::<u32>))
                .and_then(|n| n.ok())
                .and_then(char::from_u32),
        };
        match ch {
            Some(c) => {
                out.push(c);
                rest = &rest[j + 1..];
            }
            None => {
                out.push('&');
                rest = &rest[1..];
            }
        }
    }
    out.push_str(rest);
    out
}

/// Text for element content and attribute values. Characters XML cannot hold are left out.
pub(crate) fn escape(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + 8);
    for c in text.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\t' | '\n' => out.push(c),
            // (A parser reads a carriage return as a line feed, unless it is written so.)
            '\r' => out.push_str("&#13;"),
            c if (c as u32) < 0x20 || c == '\u{FFFE}' || c == '\u{FFFF}' => {}
            c => out.push(c),
        }
    }
    out
}

/// An element and where it is.
#[derive(Debug)]
pub(crate) struct El<'a> {
    pub name: &'a str,
    /// Its start (or empty-element) tag.
    pub open: Range<usize>,
    /// Just after its end tag (after the start tag when it is empty).
    pub end: usize,
    /// What is between its tags.
    pub inner: Range<usize>,
    pub children: Vec<El<'a>>,
}

impl El<'_> {
    pub fn local(&self) -> &str {
        local(self.name)
    }
    pub fn span(&self) -> Range<usize> {
        self.open.start..self.end
    }
    pub fn attr(&self, doc: &str, name: &str) -> Option<String> {
        attr(doc, self.open.clone(), name)
    }
    pub fn child(&self, name: &str) -> Option<&Self> {
        self.children.iter().find(|c| c.local() == name)
    }
    /// The text between the tags, unescaped (of an element without child elements).
    pub fn text(&self, doc: &str) -> String {
        unescape(&doc[self.inner.clone()])
    }
}

/// The element tree of `doc`. The children of the elements named in `flat` (local names) are
/// not read (the cells of a sheet are read on their own). None when the tags do not nest.
pub(crate) fn tree<'a>(doc: &'a str, flat: &[&str]) -> Option<El<'a>> {
    let mut stack: Vec<El<'a>> = Vec::new();
    // Inside an element of `flat`: its name and how deep elements of that name nest.
    let mut skipping: Option<(&str, usize)> = None;
    for tag in tags(doc, 0) {
        if let Some((name, depth)) = skipping {
            if tag.name == name {
                if tag.is_open() {
                    skipping = Some((name, depth + 1));
                } else if tag.is_close() {
                    if depth > 0 {
                        skipping = Some((name, depth - 1));
                        continue;
                    }
                    skipping = None;
                } else {
                    continue;
                }
            } else {
                continue;
            }
            if skipping.is_some() {
                continue;
            }
        }
        if tag.is_close() {
            let mut el = stack.pop()?;
            if el.name != tag.name {
                return None;
            }
            el.inner = el.open.end..tag.start;
            el.end = tag.end;
            match stack.last_mut() {
                Some(parent) => parent.children.push(el),
                None => return Some(el),
            }
            continue;
        }
        let el = El {
            name: tag.name,
            open: tag.start..tag.end,
            end: tag.end,
            inner: tag.end..tag.end,
            children: Vec::new(),
        };
        if tag.is_empty() {
            match stack.last_mut() {
                Some(parent) => parent.children.push(el),
                None => return Some(el),
            }
        } else {
            if flat.contains(&tag.local()) {
                skipping = Some((tag.name, 0));
            }
            stack.push(el);
        }
    }
    None
}

/// Replacements of spans of a document. One inside another that is replaced is dropped.
#[derive(Default)]
pub(crate) struct Edits {
    list: Vec<(usize, usize, String)>,
}

impl Edits {
    pub fn replace(&mut self, span: Range<usize>, text: impl Into<String>) {
        self.list.push((span.start, span.end, text.into()));
    }
    pub fn insert(&mut self, at: usize, text: impl Into<String>) {
        self.list.push((at, at, text.into()));
    }
    pub fn is_empty(&self) -> bool {
        self.list.is_empty()
    }
    pub fn apply(mut self, doc: &str) -> String {
        // Wider spans first among those that start together.
        self.list.sort_by(|a, b| a.0.cmp(&b.0).then(b.1.cmp(&a.1)));
        let mut out = String::with_capacity(doc.len() + 256);
        let mut at = 0;
        for (start, end, text) in self.list {
            if start < at {
                continue;
            }
            out.push_str(&doc[at..start]);
            out.push_str(&text);
            at = end;
        }
        out.push_str(&doc[at..]);
        out
    }
}

/// Sets (or adds) an attribute of the start tag at `tag`.
pub(crate) fn set_attr(doc: &str, tag: Range<usize>, name: &str, value: &str, edits: &mut Edits) {
    match attr_span(doc, tag.clone(), name) {
        Some(span) => edits.replace(span, escape(value)),
        None => {
            let end = if doc[tag.clone()].ends_with("/>") {
                tag.end - 2
            } else {
                tag.end - 1
            };
            edits.insert(end, format!(" {name}=\"{}\"", escape(value)));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tags_attributes_and_trees() {
        let doc = "<?xml version=\"1.0\"?><a x=\"1 > 2\" y='&amp;b'><!-- <no> --><b:c r=\"A1\"/><d>t&lt;</d><e><f/><e>x</e></e></a>";
        let root = tree(doc, &["e"]).unwrap();
        assert_eq!(root.name, "a");
        assert_eq!(root.attr(doc, "x").as_deref(), Some("1 > 2"));
        assert_eq!(root.attr(doc, "y").as_deref(), Some("&b"));
        let names: Vec<&str> = root.children.iter().map(|c| c.local()).collect();
        assert_eq!(names, ["c", "d", "e"]);
        assert_eq!(root.children[1].text(doc), "t<");
        // The children of a flat element are not read; its span is whole.
        let e = &root.children[2];
        assert!(e.children.is_empty());
        assert_eq!(&doc[e.span()], "<e><f/><e>x</e></e>");
        assert!(tree("<a><b></a>", &[]).is_none());
    }

    #[test]
    fn edits_are_spliced() {
        let doc = "<a x=\"1\"><b/><c/></a>";
        let root = tree(doc, &[]).unwrap();
        let mut edits = Edits::default();
        set_attr(doc, root.open.clone(), "x", "2&", &mut edits);
        set_attr(doc, root.children[0].open.clone(), "y", "z", &mut edits);
        edits.replace(root.children[1].span(), "");
        assert_eq!(edits.apply(doc), "<a x=\"2&amp;\"><b y=\"z\"/></a>");
        // An edit inside a replaced span is dropped.
        let mut edits = Edits::default();
        edits.replace(root.children[0].span(), "<n/>");
        edits.insert(root.children[0].open.start + 2, " q=\"1\"");
        assert_eq!(edits.apply(doc), "<a x=\"1\"><n/><c/></a>");
    }
}
