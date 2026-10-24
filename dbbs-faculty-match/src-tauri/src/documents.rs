use crate::models::DocumentExtractionResult;
use docx_rs::{
    read_docx, DocumentChild, Insert, InsertChild, Paragraph, ParagraphChild, Run, RunChild,
    StructuredDataTag, StructuredDataTagChild, Table, TableCellContent, TableChild, TableRowChild,
};
use pdf_extract::extract_text_from_mem;
use rtf_parser::RtfDocument;
use std::fs;
use std::path::Path;

pub(crate) fn extract_document_prompt(path: &Path) -> Result<DocumentExtractionResult, String> {
    let data = fs::read(path)
        .map_err(|err| format!("Unable to read document '{}': {err}", path.display()))?;
    let extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase());

    let mut warnings = Vec::new();

    let raw_text = match extension.as_deref() {
        Some("txt") => decode_text_bytes(&data, &mut warnings),
        Some("pdf") => extract_pdf_text(&data)?,
        Some("docx") => extract_docx_text(&data)?,
        Some("rtf") => extract_rtf_text(&data)?,
        Some("doc") => extract_doc_text(&data, &mut warnings)?,
        _ => detect_and_extract_unknown_document(&data, &mut warnings)?,
    };

    let normalized = normalize_document_text(&raw_text);

    Ok(DocumentExtractionResult {
        text: normalized,
        warnings,
    })
}

fn detect_and_extract_unknown_document(
    data: &[u8],
    warnings: &mut Vec<String>,
) -> Result<String, String> {
    if looks_like_pdf(data) {
        return extract_pdf_text(data);
    }
    if looks_like_docx(data) {
        return extract_docx_text(data);
    }
    if looks_like_rtf(data) {
        return extract_rtf_text(data);
    }
    if std::str::from_utf8(data).is_ok() {
        return Ok(decode_text_bytes(data, warnings));
    }

    Err("The selected document format is not supported. Provide a PDF, Word document, or plain text file.".into())
}

fn extract_doc_text(data: &[u8], warnings: &mut Vec<String>) -> Result<String, String> {
    if looks_like_docx(data) {
        return extract_docx_text(data);
    }
    if looks_like_pdf(data) {
        return extract_pdf_text(data);
    }
    if looks_like_rtf(data) {
        return extract_rtf_text(data);
    }

    warnings.push(
        "The .doc file was treated as plain text. Save the document as .docx if formatting is important.".into(),
    );
    Ok(decode_text_bytes(data, warnings))
}

fn extract_pdf_text(data: &[u8]) -> Result<String, String> {
    extract_text_from_mem(data)
        .map(|text| text.trim().to_string())
        .map_err(|err| format!("Unable to extract text from the PDF document: {err}"))
}

fn extract_docx_text(data: &[u8]) -> Result<String, String> {
    let package =
        read_docx(data).map_err(|err| format!("Unable to read the DOCX document: {err}"))?;
    let mut segments = Vec::new();

    for child in &package.document.children {
        collect_docx_child_text(child, &mut segments);
    }

    Ok(segments.join("\n"))
}

fn extract_rtf_text(data: &[u8]) -> Result<String, String> {
    let content = String::from_utf8_lossy(data);
    let document = RtfDocument::try_from(content.as_ref())
        .map_err(|err| format!("Unable to parse the RTF document: {err}"))?;
    Ok(document.get_text())
}

fn decode_text_bytes(data: &[u8], warnings: &mut Vec<String>) -> String {
    match std::str::from_utf8(data) {
        Ok(text) => text.to_string(),
        Err(_) => {
            warnings.push(
                "The document contained invalid UTF-8 characters. Some characters were replaced during decoding.".into(),
            );
            String::from_utf8_lossy(data).into_owned()
        }
    }
}

fn normalize_document_text(text: &str) -> String {
    let mut normalized = text.replace('\u{0000}', "");
    normalized = normalized.trim_start_matches('\u{FEFF}').to_string();
    normalized = normalized.replace("\r\n", "\n");
    normalized = normalized.replace('\r', "\n");

    let lines: Vec<&str> = normalized.lines().map(|line| line.trim_end()).collect();
    lines.join("\n").trim().to_string()
}

fn looks_like_pdf(data: &[u8]) -> bool {
    data.starts_with(b"%PDF-")
}

fn looks_like_docx(data: &[u8]) -> bool {
    data.len() > 4 && data.starts_with(b"PK")
}

fn looks_like_rtf(data: &[u8]) -> bool {
    let sample = String::from_utf8_lossy(data);
    sample
        .trim_start_matches(|c: char| c.is_ascii_whitespace())
        .starts_with("{\\rtf")
}

fn collect_docx_child_text(child: &DocumentChild, segments: &mut Vec<String>) {
    match child {
        DocumentChild::Paragraph(paragraph) => {
            if let Some(text) = collect_docx_paragraph_text(paragraph.as_ref(), segments) {
                segments.push(text);
            }
        }
        DocumentChild::Table(table) => collect_docx_table_text(table.as_ref(), segments),
        DocumentChild::StructuredDataTag(tag) => {
            collect_docx_structured_data_tag(tag.as_ref(), segments)
        }
        _ => {}
    }
}

fn collect_docx_paragraph_text(
    paragraph: &Paragraph,
    segments: &mut Vec<String>,
) -> Option<String> {
    let mut buffer = String::new();
    for child in &paragraph.children {
        append_paragraph_child_text(child, &mut buffer, segments);
    }

    let trimmed = buffer.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

fn collect_docx_table_text(table: &Table, segments: &mut Vec<String>) {
    for row in &table.rows {
        let TableChild::TableRow(row) = row;

        for cell in &row.cells {
            let TableRowChild::TableCell(cell) = cell;

            for content in &cell.children {
                match content {
                    TableCellContent::Paragraph(paragraph) => {
                        if let Some(text) = collect_docx_paragraph_text(paragraph, segments) {
                            segments.push(text);
                        }
                    }
                    TableCellContent::Table(inner) => {
                        collect_docx_table_text(inner, segments);
                    }
                    TableCellContent::StructuredDataTag(tag) => {
                        collect_docx_structured_data_tag(tag.as_ref(), segments);
                    }
                    TableCellContent::TableOfContents(_) => {}
                }
            }
        }
    }
}

fn collect_docx_structured_data_tag(tag: &StructuredDataTag, segments: &mut Vec<String>) {
    let mut buffer = String::new();
    append_structured_data_tag_text(tag, &mut buffer, segments);
    let trimmed = buffer.trim();
    if !trimmed.is_empty() {
        segments.push(trimmed.to_string());
    }
}

fn append_paragraph_child_text(
    child: &ParagraphChild,
    buffer: &mut String,
    segments: &mut Vec<String>,
) {
    match child {
        ParagraphChild::Run(run) => append_run_text(run.as_ref(), buffer),
        ParagraphChild::Insert(insert) => append_insert_text(insert, buffer),
        ParagraphChild::Hyperlink(hyperlink) => {
            for inner in &hyperlink.children {
                append_paragraph_child_text(inner, buffer, segments);
            }
        }
        ParagraphChild::StructuredDataTag(tag) => {
            append_structured_data_tag_text(tag.as_ref(), buffer, segments);
        }
        ParagraphChild::BookmarkStart(_) | ParagraphChild::BookmarkEnd(_) => {}
        ParagraphChild::CommentStart(_) | ParagraphChild::CommentEnd(_) => {}
        ParagraphChild::Delete(_) => {}
        ParagraphChild::PageNum(_) | ParagraphChild::NumPages(_) => {}
    }
}

fn append_insert_text(insert: &Insert, buffer: &mut String) {
    for child in &insert.children {
        match child {
            InsertChild::Run(run) => append_run_text(run.as_ref(), buffer),
            InsertChild::Delete(_) => {}
            InsertChild::CommentStart(_) | InsertChild::CommentEnd(_) => {}
        }
    }
}

fn append_structured_data_tag_text(
    tag: &StructuredDataTag,
    buffer: &mut String,
    segments: &mut Vec<String>,
) {
    for child in &tag.children {
        match child {
            StructuredDataTagChild::Run(run) => append_run_text(run.as_ref(), buffer),
            StructuredDataTagChild::Paragraph(paragraph) => {
                if let Some(text) = collect_docx_paragraph_text(paragraph.as_ref(), segments) {
                    if !buffer.is_empty() && !buffer.ends_with('\n') && !buffer.ends_with(' ') {
                        buffer.push(' ');
                    }
                    buffer.push_str(&text);
                }
            }
            StructuredDataTagChild::Table(table) => {
                collect_docx_table_text(table.as_ref(), segments)
            }
            StructuredDataTagChild::StructuredDataTag(inner) => {
                append_structured_data_tag_text(inner.as_ref(), buffer, segments);
            }
            StructuredDataTagChild::BookmarkStart(_) | StructuredDataTagChild::BookmarkEnd(_) => {}
            StructuredDataTagChild::CommentStart(_) | StructuredDataTagChild::CommentEnd(_) => {}
        }
    }
}

fn append_run_text(run: &Run, buffer: &mut String) {
    for child in &run.children {
        match child {
            RunChild::Text(text) => buffer.push_str(&text.text),
            RunChild::Break(_) => buffer.push('\n'),
            RunChild::Tab(_) | RunChild::PTab(_) => buffer.push('\t'),
            RunChild::Sym(sym) => {
                if let Ok(value) = u32::from_str_radix(&sym.char, 16) {
                    if let Some(ch) = char::from_u32(value) {
                        buffer.push(ch);
                    }
                }
            }
            RunChild::InstrTextString(value) => buffer.push_str(value),
            RunChild::DeleteText(_) => {}
            RunChild::FieldChar(_) => {}
            RunChild::Drawing(_) => {}
            RunChild::Shape(_) => {}
            RunChild::CommentStart(_) | RunChild::CommentEnd(_) => {}
            RunChild::FootnoteReference(_) => {}
            RunChild::Shading(_) => {}
            RunChild::InstrText(_) => {}
            RunChild::DeleteInstrText(_) => {}
        }
    }
}

#[cfg(test)]
#[path = "tests/documents.rs"]
mod tests;
