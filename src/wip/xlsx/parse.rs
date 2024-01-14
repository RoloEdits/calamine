use crate::wip::{Cell, Font, Grid, Worksheet, XlsxError};
use compact_str::{CompactString, ToCompactString};
use core::panic;
use quick_xml::{
    events::{attributes::Attribute, Event},
    name::QName,
    Reader,
};
use std::io::{BufReader, Read, Seek};
use zip::{result::ZipError, ZipArchive};

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

pub(super) fn theme<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
) -> Result<Vec<CompactString>, XlsxError> {
    let file = zip.by_name("xl/theme/theme1.xml").unwrap();
    let mut reader = Reader::from_reader(BufReader::new(file));

    reader
        .check_end_names(false)
        .trim_text(false)
        .check_comments(false)
        .expand_empty_elements(true);

    todo!()
}

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
                                                // Split on `.` in "17.000000"
                                                let integer_part =
                                            // Take the first match of the split. "17"
                                            val.split(|val| *val == b'.').next().unwrap();

                                                font.size = CompactString::from_utf8(integer_part)
                                                    .expect("ascii is valid utf8")
                                                    .parse()?;
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

                                                font.color = CompactString::from_utf8(rgb)
                                                    .expect("ascii is valid utf8");
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
                                                font.font = CompactString::from_utf8(val)
                                                    .expect("ascii is valid utf8");
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

pub(super) fn relationships<R: Read + Seek>(zip: &mut ZipArchive<R>) -> Vec<CompactString> {
    todo!()
}

pub(super) fn workbook<R: Read + Seek>(zip: &mut ZipArchive<R>) -> Vec<CompactString> {
    todo!()
}

pub(super) fn worksheet<'a, R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    shared_strings: &[CompactString],
    styles: &'a [Font],
    worksheet: &str,
) -> Result<Option<Worksheet>, XlsxError> {
    let file = match archive.by_name(&format!("xl/worksheets/{worksheet}.xml")) {
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

    let mut cells = Grid::new();

    // FIX: looping over cells more than once
    // FIX: not always getting the attribute data
    // FIX: sometimes empty values when there shouldnt be

    let mut columns = 0;
    let mut rows = 0;
    loop {
        if let Ok(event) = reader.read_event_into(&mut buffer) {
            match event {
                Event::Start(ref start) if start.local_name().as_ref() == b"c" => {
                    let attributes = start.attributes();
                    let mut font: Option<Font> = None;
                    let mut column = 0;
                    let mut row = 0;
                    let mut value: Option<CompactString> = None;

                    for attribute in attributes {
                        if let Ok(Attribute {
                            key: QName(b"r"),
                            value: ref cell_pos,
                        }) = attribute
                        {
                            let cell_pos =
                                std::str::from_utf8(cell_pos).expect("ascii is valid utf8");

                            let (col, rw) = excel_column_row_to_tuple(cell_pos);

                            columns = columns.max(column);
                            rows = rows.max(row);

                            column = col;
                            row = rw;
                        }

                        // Style index
                        if let Ok(Attribute {
                            key: QName(b"s"),
                            value: ref style_idx,
                        }) = attribute
                        {
                            let idx: usize = std::str::from_utf8(style_idx)
                                .expect("ascii is valid utf8")
                                .parse()
                                .expect("should always be an integer");

                            // TODO: If one is not found, need to use the worksheet theme to default.
                            // This will be passed in in the future.
                            let style = match styles.get(idx) {
                                Some(some) => some,
                                None => {
                                    // dbg!(styles);
                                    // print!(
                                    //     "index out of bounds: len = {} idx = {idx}",
                                    //     styles.len()
                                    // );
                                    font = None;
                                    continue;
                                }
                            };

                            font = Some(style.clone());
                        }

                        // Datatype
                        if let Ok(Attribute {
                            key: QName(b"t"),
                            value: ref typ,
                        }) = attribute
                        {
                            // `s` = shared string
                            if typ.as_ref() == b"s" {
                                cell_buffer.clear();
                                if let Ok(Event::Start(start)) =
                                    reader.read_event_into(&mut cell_buffer)
                                {
                                    // Get the value of `<v>VALUE</v>`
                                    if start.local_name().as_ref() == b"v" {
                                        if let Ok(Event::Text(text)) =
                                            reader.read_event_into(&mut value_buffer)
                                        {
                                            // Shared string index
                                            let string_idx: usize = std::str::from_utf8(&text)
                                                    .expect("ascii is valid utf8")
                                                    .parse()
                                                    .expect("should be integer to index into shared strings");

                                            value = Some(shared_strings[string_idx].clone());
                                        }
                                    } else {
                                        // Defaulting to None as data working with in development time is only shared strings
                                        value = None;
                                    }
                                }
                            }
                        }
                    }

                    cells.insert(Cell {
                        value: value.clone(),
                        column,
                        row,
                        font,
                    });
                }
                Event::End(ref end) if end.local_name().as_ref() == b"sheetData" => {
                    break;
                }
                Event::Eof => {
                    break;
                }
                _ => {
                    continue;
                }
            }
        }
    }

    Ok(Some(Worksheet {
        name: worksheet.to_compact_string(),
        grid: cells,
        size: (columns, rows),
    }))
}

/// # Panics
///
/// Will panic if column format is not `a-z` or `A-Z`, or if there is no row number after the column letter/s.
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
    (column, row - 1)
}

/// # Panics
///
/// Will panic if the provided string contains a letter other than `a-z` or `A-Z`.
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
            // If the string contains non-alphabetic characters panic
            panic!("`{char}` is not a valid column letter must be `A-Z`")
        }
    }

    // zero based position
    result - 1
}

fn to_uppercase_compact_str(string: &str) -> CompactString {
    // NOTE: maximum number of columns is 16,384 or `XFD`
    let mut compact_str = CompactString::with_capacity(3);

    for char in string.chars() {
        compact_str.push(char.to_ascii_uppercase());
    }

    compact_str
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
