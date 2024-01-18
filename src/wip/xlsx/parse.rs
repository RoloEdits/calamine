use crate::wip::{Cell, Font, Type, Worksheet, XlsxError};
use compact_str::CompactString;
use core::panic;
use quick_xml::{
    events::{attributes::Attribute, Event},
    name::QName,
    Reader,
};
use std::io::{BufReader, Read, Seek};
use zip::{result::ZipError, ZipArchive};

#[inline]
pub(super) fn shared_strings<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> Result<Vec<CompactString>, XlsxError> {
    let file = archive.by_name("xl/sharedStrings.xml").unwrap();
    let mut reader = Reader::from_reader(BufReader::new(file));

    reader
        .check_end_names(false)
        .trim_text(false)
        .check_comments(false)
        .expand_empty_elements(true);

    let mut buffer: Vec<u8> = Vec::with_capacity(1024);
    let mut inner_buffer: Vec<u8> = Vec::with_capacity(1024);
    let mut inner_inner_buffer: Vec<u8> = Vec::with_capacity(1024);

    let mut strings: Vec<CompactString> = Vec::new();
    // let mut rich_buffer: Option<CompactString> = None;
    let mut is_phonetic_text = false;

    loop {
        if let Ok(event) = reader.read_event_into(&mut buffer) {
            match event {
                Event::Start(ref start) if start.local_name().as_ref() == b"si" => {
                    if let Ok(inner_event) = reader.read_event_into(&mut inner_buffer) {
                        match inner_event {
                            // Event::Start(ref e) if e.local_name().as_ref() == b"r" => {
                            //     if rich_buffer.is_none() {
                            //         // use a buffer since richtext has multiples <r> and <t> for the same cell
                            //         rich_buffer = CompactString::default();
                            //     }
                            // }
                            Event::Start(ref e) if e.local_name().as_ref() == b"rPh" => {
                                is_phonetic_text = true;
                            }
                            // Event::End(ref e)
                            //     if e.local_name().as_ref() == start.local_name().as_ref() =>
                            // {
                            //     strings.push(rich_buffer);
                            // }
                            Event::End(ref e) if e.local_name().as_ref() == b"rPh" => {
                                is_phonetic_text = false;
                            }
                            Event::Start(ref e)
                                if e.local_name().as_ref() == b"t" && !is_phonetic_text =>
                            {
                                inner_inner_buffer.clear();

                                loop {
                                    match reader.read_event_into(&mut inner_inner_buffer).unwrap() {
                                        Event::Text(t) => {
                                            strings.push(CompactString::new(&t.unescape().unwrap()))
                                        }
                                        Event::End(end) if end.name() == e.name() => break,
                                        Event::Eof => return Err(XlsxError::XmlEof("t")),
                                        _ => (),
                                    }
                                }

                                // if let Some(ref mut s) = rich_buffer {
                                //     s.push_str(&value);
                                // } else {
                                //     inner_buffer.clear();
                                //     // consume any remaining events up to expected closing tag
                                //     reader
                                //         .read_to_end_into(start.name(), &mut inner_buffer)
                                //         .unwrap();
                                //     // return Ok(Some(value));
                                // }
                            }
                            _ => {}
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        buffer.clear();
    }

    Ok(strings)
}

#[inline]
pub(super) fn theme<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> Result<Vec<CompactString>, XlsxError> {
    let file = archive.by_name("xl/theme/theme1.xml").unwrap();
    let mut reader = Reader::from_reader(BufReader::new(file));

    reader
        .check_end_names(false)
        .trim_text(false)
        .check_comments(false)
        .expand_empty_elements(true);

    todo!()
}

#[inline]
pub(super) fn styles<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<Vec<Font>, XlsxError> {
    let file = archive.by_name("xl/styles.xml").unwrap();
    let mut reader = Reader::from_reader(BufReader::new(file));

    reader
        .check_end_names(false)
        .trim_text(false)
        .check_comments(false)
        .expand_empty_elements(true);

    let mut buffer = Vec::with_capacity(1024);
    let mut inner_buffer = Vec::with_capacity(1024);

    let mut fonts = Vec::new();
    loop {
        // TODO: validate document is utf8 from `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
        if let Ok(event) = reader.read_event_into(&mut buffer) {
            match event {
                // TODO: match to "fonts" and break when End is "fonts". Currently we only want the font information, nothing else.
                Event::Start(ref start) => {
                    if start.local_name().as_ref() == b"font" {
                        let mut font = Font::default();
                        // TODO: Need to go through the other elements. I was trying to get attributes for font but there are none.

                        // <font>
                        //     <b />
                        //     <sz val="17.000000" />
                        //     <color indexed="2" />
                        //     <name val="Arial" />
                        // </font>

                        loop {
                            inner_buffer.clear();

                            if let Ok(event) = reader.read_event_into(&mut inner_buffer) {
                                match event {
                                    Event::Start(ref start) => {
                                        if start.local_name().as_ref() == b"sz" {
                                            if let Some(Ok(Attribute {
                                                key: QName(b"val"),
                                                value: val,
                                            })) = start.attributes().next()
                                            {
                                                // SAFETY: document is known valid utf8
                                                font.size = unsafe {
                                                    std::str::from_utf8_unchecked(&val)
                                                        .parse()
                                                        .expect(
                                                            "font size should always be a float",
                                                        )
                                                };
                                            }
                                        }

                                        if start.local_name().as_ref() == b"color" {
                                            if let Some(Ok(Attribute {
                                                key: QName(b"rgb"),
                                                value: color,
                                            })) = start.attributes().next()
                                            {
                                                let rgb = if color.len() > 6 {
                                                    // removes `FF` from "FFD3D1A2"
                                                    &color.as_ref()[2..]
                                                } else {
                                                    // otherwise just return the a normal rgb hex.
                                                    color.as_ref()
                                                };

                                                // SAFETY: document is known valid utf8
                                                font.color = unsafe {
                                                    CompactString::from_utf8_unchecked(rgb)
                                                };
                                            }
                                            if let Some(Ok(Attribute {
                                                key: QName(b"theme"),
                                                value: _,
                                            })) = start.attributes().next()
                                            {
                                                font.color = CompactString::new("FFFFFF");
                                            }
                                            if let Some(Ok(Attribute {
                                                key: QName(b"indexed"),
                                                value: _,
                                            })) = start.attributes().next()
                                            {
                                                font.color = CompactString::new("FFFFFF");
                                            }
                                        }

                                        if start.local_name().as_ref() == b"name" {
                                            if let Some(Ok(Attribute {
                                                key: QName(b"val"),
                                                value: val,
                                            })) = start.attributes().next()
                                            {
                                                // SAFETY: document is known valid utf8
                                                font.font = unsafe {
                                                    CompactString::from_utf8_unchecked(val)
                                                };
                                            }
                                        }
                                    }
                                    Event::End(ref end) if end.local_name().as_ref() == b"b" => {
                                        continue;
                                    }
                                    Event::End(ref end) if end.local_name().as_ref() == b"font" => {
                                        break;
                                    }
                                    _ => {
                                        continue;
                                    }
                                }
                            }
                        }
                        fonts.push(font.clone());
                    }
                }
                Event::Eof => break,
                _ => continue,
            }
        }
    }

    Ok(fonts)
}

pub(super) fn relationships<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Vec<CompactString> {
    todo!()
}

pub(super) fn workbook<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Vec<CompactString> {
    todo!()
}

#[inline]
pub(super) fn worksheet<R: Read + Seek>(
    worksheet: &mut Worksheet,
    archive: &mut ZipArchive<R>,
    shared_strings: &[CompactString],
    styles: &[Font],
) -> Result<Option<()>, XlsxError> {
    let file = match archive.by_name(&format!("xl/worksheets/{}.xml", worksheet.name)) {
        Ok(ok) => ok,
        Err(ZipError::FileNotFound) => return Ok(None),
        Err(err) => return Err(err.into()),
    };

    let mut reader = Reader::from_reader(BufReader::new(file));

    reader
        .check_end_names(false)
        .trim_text(false)
        .check_comments(false)
        .expand_empty_elements(true);

    let mut buffer: Vec<u8> = Vec::with_capacity(1024);
    let mut cell_buffer: Vec<u8> = Vec::with_capacity(1024);
    let mut value_buffer: Vec<u8> = Vec::with_capacity(1024);

    loop {
        if let Ok(event) = reader.read_event_into(&mut buffer) {
            match event {
                Event::Start(ref start) if start.local_name().as_ref() == b"c" => {
                    let mut attributes = start.attributes();

                    let mut cell = Cell::default();

                    // Go through each attribute and take relevant data from them.
                    // Example: <c r="A1" s="5" t="s">
                    // Example:  <c r="E1" s="3">

                    if let Some(Ok(Attribute {
                        // Cell position.
                        // Example: "A1"
                        key: QName(b"r"),
                        value: ref pos,
                    })) = attributes.next()
                    {
                        // SAFETY: document is known valid utf8
                        let position = unsafe { std::str::from_utf8_unchecked(pos) };

                        let (column, row) = excel_column_row_to_tuple(position);

                        cell.column = column;
                        cell.row = row;
                    }

                    if let Some(Ok(Attribute {
                        // Style
                        key: QName(b"s"),
                        value: ref style_idx,
                    })) = attributes.next()
                    {
                        let idx: usize = unsafe {
                            // SAFETY: document is known valid utf8
                            std::str::from_utf8_unchecked(style_idx)
                                .parse()
                                .expect("should always be an integer")
                        };

                        // TODO: If one is not found, need to use the worksheet theme to default.
                        // This will be passed in in the future.
                        let font = match styles.get(idx) {
                            Some(some) => some.clone(),
                            None => {
                                // TODO: will get style from base theme
                                Font {
                                    font: CompactString::new("Arial"),
                                    size: 12.0,
                                    color: CompactString::new("000000"),
                                }
                            }
                        };

                        cell.font = font;
                    }

                    if let Some(Ok(Attribute {
                        // "t" = Datatype
                        key: QName(b"t"),
                        value: ref typ,
                    })) = attributes.next()
                    {
                        match typ.as_ref() {
                            // shared string
                            b"s" => {
                                cell.r#type = Some(Type::String);
                            }
                            // in-line string
                            b"is" => {}
                            // formula
                            b"f" => {}
                            // number
                            b"n" => {
                                cell.r#type = Some(Type::Number);
                            }
                            unknown => panic!(
                                "unknown attribute on cell: `{}`",
                                std::str::from_utf8(unknown).unwrap()
                            ),
                        }
                    }

                    if cell.r#type.is_none() {
                        cell.r#type = Some(Type::Number);
                    }

                    // Get relevant value in cell
                    cell_buffer.clear();
                    if let Ok(Event::Start(start)) = reader.read_event_into(&mut cell_buffer) {
                        // Get the value of `<v>VALUE</v>` if it exists. If it doesnt, then the cell will have a vlue of `None`
                        if start.local_name().as_ref() == b"v" {
                            if let Ok(Event::Text(text)) = reader.read_event_into(&mut value_buffer)
                            {
                                match cell
                                    .r#type
                                    .expect("must have a type if there is a value element found")
                                {
                                    Type::Number => {
                                        cell.value =
                                            Some(CompactString::new(&text.unescape().unwrap()))
                                    }
                                    Type::String => {
                                        let idx: usize = unsafe {
                                            // SAFETY: document is known valid utf8
                                            std::str::from_utf8_unchecked(&text).parse().expect(
                                                "should be integer to index into shared strings",
                                            )
                                        };

                                        cell.value = Some(shared_strings[idx].clone());
                                    }
                                    Type::Formula => todo!(),
                                };
                            }
                        }
                    }

                    worksheet.spreadsheet.insert(cell);
                }
                Event::End(ref end) if end.local_name().as_ref() == b"sheetData" => {
                    break;
                }
                Event::Eof => {
                    break;
                }
                _ => continue,
            }
        }
    }

    Ok(Some(()))
}

/// # Panics
///
/// Will panic if column format is not `a-z` or `A-Z`, or if there is no row number after the column letter/s.
#[inline]
fn excel_column_row_to_tuple(pos: &str) -> (u32, u32) {
    let mut idx = 0;

    for char in pos.chars() {
        if char.is_ascii_alphabetic() {
            idx += 1;
            continue;
        }
        break;
    }

    let column = excel_column_to_number(&pos[..idx]);
    let row: u32 = pos[idx..].parse().unwrap();

    // zero based position
    (column - 1, row - 1)
}

/// # Panics
///
/// Will panic if the provided string contains a letter other than `a-z` or `A-Z`.
#[inline]
fn excel_column_to_number(column: &str) -> u32 {
    let column = to_uppercase_compact_str(column);
    let mut result = 0;
    let mut multiplier = 1;

    for char in column.chars().rev() {
        if char.is_ascii_alphabetic() {
            let digit = char as u32 - 'A' as u32 + 1;
            result += digit * multiplier;
            multiplier *= 26;
        } else {
            // break;
            // If the string contains non-alphabetic characters panic
            panic!("`{char}` is not a valid column letter must be `A-Z`")
        }
    }

    result
}

#[inline]
fn to_uppercase_compact_str(string: &str) -> CompactString {
    // NOTE: maximum number of columns is 16,384 or `XFD`
    assert!(string.len() < 4);
    let len = string.len();

    let mut char_array: [u8; 3] = [0; 3];

    for (idx, char) in string.chars().enumerate() {
        char_array[idx] = char.to_ascii_uppercase() as u8;
    }

    // SAFETY: Input is a `&str` and `to_ascii_uppercase` returns a valid ascii.
    unsafe { CompactString::from_utf8_unchecked(&char_array[..len]) }
}

#[cfg(test)]
mod test {
    use super::*;

    #[test]
    fn should_convert_excel_letter_number_format_to_tuple() {
        let result = excel_column_row_to_tuple("A1");

        assert_eq!((0, 0), result);

        let result = excel_column_row_to_tuple("A2");

        assert_eq!((0, 1), result);

        let result = excel_column_row_to_tuple("B1");

        assert_eq!((1, 0), result);
    }
}
