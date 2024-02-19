use crate::wip::{cell::Type, spreadsheet::Spreadsheet, Cell, Error, Worksheet};
use compact_str::CompactString;
use core::panic;
use quick_xml::{
    events::{attributes::Attribute, Event},
    name::QName,
    Reader,
};
use std::{borrow::Cow, fs::File, io::BufReader};
use zip::{result::ZipError, ZipArchive};

#[inline]
pub(super) fn workbook<'a>(
    archive: &mut ZipArchive<BufReader<File>>,
) -> Result<Vec<Worksheet<'a>>, Error> {
    let file = archive.by_name("xl/workbook.xml")?;
    let mut reader = Reader::from_reader(BufReader::new(file));

    reader.check_end_names(false);

    let mut buffer: Vec<u8> = Vec::with_capacity(1024);

    let mut worksheets = Vec::new();

    // <sheets>
    //     <sheet name="Sheet1" sheetId="1" state="visible" r:id="rId3" />
    // </sheets>
    loop {
        match &reader.read_event_into(&mut buffer)? {
            Event::Empty(sheet) if sheet.local_name().as_ref() == b"sheet" => {
                let mut attributes = sheet.attributes();

                let name = &attributes
                    .next()
                    .expect(r#"<sheet name="NAME"> should be first attribute"#)
                    .expect("attribute iter is infallible");

                // SAFETY: document should be valid UTF-8
                let name = unsafe { CompactString::from_utf8_unchecked(&name.value) };

                let id = &attributes
                    .next()
                    .expect(r#"<sheet name="NAME"> should be first attribute"#)
                    .expect("attribute iter is infallible");

                // SAFETY: document should be valid UTF-8
                let id = unsafe {
                    CompactString::from_utf8_unchecked(&id.value)
                        .parse::<u32>()
                        .expect("`sheetId` should always be parsable into a `u32`")
                };

                let spreadsheet = Spreadsheet::new();

                let worksheet = Worksheet {
                    id,
                    name,
                    spreadsheet,
                    reader: None,
                };
                worksheets.push(worksheet);
            }
            Event::End(end) if end.local_name().as_ref() == b"sheets" => break,
            _ => continue,
        }
    }

    Ok(worksheets)
}

#[inline]
pub(super) fn shared_strings(
    archive: &mut ZipArchive<BufReader<File>>,
) -> Result<Vec<CompactString>, Error> {
    let file = archive.by_name("xl/sharedStrings.xml")?;
    let mut reader = Reader::from_reader(BufReader::new(file));

    reader.check_end_names(false);

    let mut buffer: Vec<u8> = Vec::with_capacity(1024);
    let mut inner_buffer: Vec<u8> = Vec::with_capacity(1024);
    let mut inner_inner_buffer: Vec<u8> = Vec::with_capacity(1024);

    let mut strings: Vec<CompactString> = Vec::new();

    loop {
        buffer.clear();
        match &reader.read_event_into(&mut buffer)? {
            Event::Start(si) if si.local_name().as_ref() == b"si" => loop {
                inner_buffer.clear();
                if let Ok(si_inner) = reader.read_event_into(&mut inner_buffer) {
                    match &si_inner {
                        Event::Start(event) if event.local_name().as_ref() == b"t" => {
                            inner_inner_buffer.clear();
                            match reader.read_event_into(&mut inner_inner_buffer).unwrap() {
                                Event::Text(text) => {
                                    strings.push(CompactString::new(
                                        &text
                                            .unescape()
                                            .expect("text in `t` element should be valid"),
                                    ));
                                }
                                _ => continue,
                            }
                        }
                        // TODO: other tags in inner `si` will go here

                        // Reached end of inner `si` tag and can go on to the next one
                        Event::End(end) if end.local_name().as_ref() == b"si" => break,

                        _ => continue,
                    }
                }
            },
            Event::Eof => break,
            _ => continue,
        }
    }

    Ok(strings)
}

#[inline]
pub(super) fn _theme(
    archive: &mut ZipArchive<BufReader<File>>,
) -> Result<Vec<CompactString>, Error> {
    let file = archive.by_name("xl/theme/theme1.xml").unwrap();
    let mut reader = Reader::from_reader(BufReader::new(file));

    reader.check_end_names(false);

    todo!()
}

#[allow(clippy::too_many_lines, clippy::cognitive_complexity)]
#[inline]
pub(super) fn styles(archive: &mut ZipArchive<BufReader<File>>) -> Result<Styles, Error> {
    let file = archive.by_name("xl/styles.xml")?;
    let mut reader = Reader::from_reader(BufReader::new(file));

    reader.check_end_names(false);

    let mut buf_1 = Vec::with_capacity(1024);
    let mut buf_2 = Vec::with_capacity(1024);
    let mut buf_3 = Vec::with_capacity(1024);

    let mut styles = Styles::default();
    'outer: loop {
        // TODO: validate document is utf8 from `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
        if let Ok(event) = reader.read_event_into(&mut buf_1) {
            match event {
                // Event::Start(ref start) if start.local_name().as_ref() == b"numFmts" => {}
                Event::Start(ref start) if start.local_name().as_ref() == b"fonts" => {
                    // <fonts count="45">
                    if let Some(Ok(Attribute {
                        key: QName(b"count"),
                        value: fonts_count,
                    })) = start.attributes().next()
                    {
                        styles.fonts.count = unsafe {
                            std::str::from_utf8_unchecked(&fonts_count)
                                .parse::<u32>()
                                .expect("font count should always be a valid ascii number")
                        };
                    }

                    let mut f_count = 0;
                    loop {
                        buf_2.clear();
                        if let Ok(event) = reader.read_event_into(&mut buf_2) {
                            match event {
                                Event::Start(start) if start.local_name().as_ref() == b"font" => {
                                    //     <font>
                                    //         <b />
                                    //         <sz val="17.000000" />
                                    //         <color indexed="2" />
                                    //         <name val="Arial" />
                                    //     </font>

                                    let mut font = Font::default();

                                    loop {
                                        buf_3.clear();
                                        if let Ok(event) = reader.read_event_into(&mut buf_3) {
                                            match event {
                                                Event::Empty(ref start)
                                                    if start.local_name().as_ref() == b"sz" =>
                                                {
                                                    if let Some(Ok(Attribute {
                                                        key: QName(b"val"),
                                                        value: val,
                                                    })) = start.attributes().next()
                                                    {
                                                        font.sz = unsafe {
                                                            // SAFETY: document is known valid utf8
                                                            std::str::from_utf8_unchecked(&val)
                                                                    .parse()
                                                                    .expect("font size should always be a float")
                                                        };
                                                    }
                                                }
                                                Event::Empty(ref start)
                                                    if start.local_name().as_ref() == b"color" =>
                                                {
                                                    if let Some(Ok(Attribute {
                                                        key: QName(b"rgb"),
                                                        value: color,
                                                    })) = start.attributes().next()
                                                    {
                                                        font.color = unsafe {
                                                            Some(Color::Argb(
                                                                // SAFETY: document is known valid utf8
                                                                CompactString::from_utf8_unchecked(
                                                                    &color,
                                                                ),
                                                            ))
                                                        };
                                                    }

                                                    if let Some(Ok(Attribute {
                                                        key: QName(b"theme"),
                                                        value: theme,
                                                    })) = start.attributes().next()
                                                    {
                                                        let theme: u32 = unsafe {
                                                            std::str::from_utf8_unchecked(&theme)
                                                                .parse()?
                                                        };

                                                        let mut tint: Option<f32> = None;

                                                        if let Some(Ok(Attribute {
                                                            key: QName(b"tint"),
                                                            value: val,
                                                        })) = start.attributes().next()
                                                        {
                                                            tint = Some(unsafe {
                                                                std::str::from_utf8_unchecked(&val)
                                                                    .parse()
                                                                    .expect("tint must be float")
                                                            });
                                                        }

                                                        font.color =
                                                            Some(Color::Theme { theme, tint });
                                                    }

                                                    if let Some(Ok(Attribute {
                                                        key: QName(b"indexed"),
                                                        value: idx,
                                                    })) = start.attributes().next()
                                                    {
                                                        let idx = unsafe {
                                                            std::str::from_utf8_unchecked(&idx)
                                                                .parse()?
                                                        };
                                                        font.color = Some(Color::Indexed(idx));
                                                    }
                                                }
                                                Event::Empty(ref start)
                                                    if start.local_name().as_ref() == b"name" =>
                                                {
                                                    if let Some(Ok(Attribute {
                                                        key: QName(b"val"),
                                                        value: val,
                                                    })) = start.attributes().next()
                                                    {
                                                        font.name = unsafe {
                                                            // SAFETY: document is known valid utf8
                                                            CompactString::from_utf8_unchecked(val)
                                                        };
                                                    }
                                                }
                                                Event::End(end)
                                                    if end.local_name().as_ref() == b"font" =>
                                                {
                                                    break
                                                }
                                                _ => continue,
                                            }
                                        }
                                    }
                                    styles.fonts.add(font);
                                    f_count += 1;
                                }
                                Event::End(end) if end.local_name().as_ref() == b"fonts" => {
                                    // Gone through all `font` elements, start to go through other collections.
                                    // Check that the expected font count is achieved.
                                    debug_assert_eq!(styles.fonts.count, f_count);
                                    continue 'outer;
                                }
                                _ => continue,
                            }
                        }
                    }
                }
                // TODO: Go through other collections
                // Event::Start(ref start) if start.local_name().as_ref() == b"fills" => {}
                // Event::Start(ref start) if start.local_name().as_ref() == b"borders" => {}
                // Event::Start(ref start) if start.local_name().as_ref() == b"cellStyleXfs" => {}
                Event::Start(ref start) if start.local_name().as_ref() == b"cellXfs" => {
                    if let Some(Ok(Attribute {
                        key: QName(b"count"),
                        value: val,
                    })) = start.attributes().next()
                    {
                        styles.cells_xfs.count = unsafe {
                            // SAFETY: document is known utf-8
                            std::str::from_utf8_unchecked(&val).parse::<u32>()?
                        };
                    }

                    loop {
                        buf_2.clear();
                        if let Ok(event) = reader.read_event_into(&mut buf_2) {
                            match event {
                                Event::Start(start) if start.local_name().as_ref() == b"xf" => {
                                    // <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="false"
                                    //     applyBorder="false" applyAlignment="false" applyProtection="false">
                                    //     <alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false"
                                    //         indent="0" shrinkToFit="false" />
                                    //     <protection locked="true" hidden="false" />
                                    // </xf>

                                    let mut xf = Xf::default();

                                    for attribute in start.attributes() {
                                        let attribute =
                                            attribute.expect("attribute iter should be infallible");

                                        match attribute {
                                            Attribute {
                                                key: QName(b"numFmtId"),
                                                value: val,
                                            } => {
                                                xf.num_fmt_id = unsafe {
                                                    // SAFETY: document is known utf-8
                                                    std::str::from_utf8_unchecked(&val)
                                                        .parse::<u32>()?
                                                };
                                            }

                                            Attribute {
                                                key: QName(b"fontId"),
                                                value: val,
                                            } => {
                                                xf.font_id = unsafe {
                                                    // SAFETY: document is known utf-8
                                                    std::str::from_utf8_unchecked(&val)
                                                        .parse::<usize>()?
                                                };
                                            }

                                            Attribute {
                                                key: QName(b"fillId"),
                                                value: val,
                                            } => {
                                                xf.fill_id = unsafe {
                                                    // SAFETY: document is known utf-8
                                                    std::str::from_utf8_unchecked(&val)
                                                        .parse::<u32>()?
                                                };
                                            }

                                            Attribute {
                                                key: QName(b"borderId"),
                                                value: val,
                                            } => {
                                                xf.border_id = unsafe {
                                                    // SAFETY: document is known utf-8
                                                    std::str::from_utf8_unchecked(&val)
                                                        .parse::<u32>()?
                                                };
                                            }

                                            Attribute {
                                                key: QName(b"xfId"),
                                                value: val,
                                            } => {
                                                xf.xf_id = unsafe {
                                                    // SAFETY: document is known utf-8
                                                    std::str::from_utf8_unchecked(&val)
                                                        .parse::<u32>()?
                                                };
                                            }

                                            Attribute {
                                                key: QName(b"applyFont"),
                                                value: val,
                                            } => {
                                                let str = unsafe {
                                                    // SAFETY: document is known utf-8
                                                    std::str::from_utf8_unchecked(&val)
                                                };

                                                // NOTE: can be "1" or "0"
                                                xf.apply_font = if str.len() == 1 {
                                                    let int = str.parse::<u8>().unwrap();

                                                    matches!(int, 1)
                                                } else {
                                                    str.parse()
                                                        .unwrap_or_else(|err| {
                                                            panic!("should be `true` or `false`, got `{str}`: {err}")
                                                        })
                                                };
                                            }

                                            Attribute {
                                                key: QName(b"applyBorder"),
                                                value: val,
                                            } => {
                                                let str = unsafe {
                                                    // SAFETY: document is known utf-8
                                                    std::str::from_utf8_unchecked(&val)
                                                };

                                                // NOTE: can be "1" or "0"
                                                xf.apply_border = if str.len() == 1 {
                                                    let int = str.parse::<u8>().unwrap();

                                                    matches!(int, 1)
                                                } else {
                                                    str.parse()
                                                        .unwrap_or_else(|err| {
                                                            panic!("should be `true` or `false`, got `{str}`: {err}")
                                                        })
                                                };
                                            }

                                            Attribute {
                                                key: QName(b"applyAlignment"),
                                                value: val,
                                            } => {
                                                let str = unsafe {
                                                    // SAFETY: document is known utf-8
                                                    std::str::from_utf8_unchecked(&val)
                                                };

                                                // NOTE: can be "1" or "0"
                                                xf.apply_alignment = if str.len() == 1 {
                                                    let int = str.parse::<u8>().unwrap();

                                                    matches!(int, 1)
                                                } else {
                                                    str.parse()
                                                        .unwrap_or_else(|err| {
                                                            panic!("should be `true` or `false`, got `{str}`: {err}")
                                                        })
                                                };
                                            }

                                            Attribute {
                                                key: QName(b"applyProtection"),
                                                value: val,
                                            } => {
                                                let str = unsafe {
                                                    // SAFETY: document is known utf-8
                                                    std::str::from_utf8_unchecked(&val)
                                                };

                                                // NOTE: can be "1" or "0"
                                                xf.apply_protection = if str.len() == 1 {
                                                    let int = str.parse::<u8>().unwrap();

                                                    matches!(int, 1)
                                                } else {
                                                    str.parse()
                                                        .unwrap_or_else(|err| {
                                                            panic!("should be `true` or `false`, got `{str}`: {err}")
                                                        })
                                                };
                                            }

                                            _ => unreachable!(
                                                "unknown attribute found on `xf` element"
                                            ),
                                        }
                                    }

                                    // TODO: alignment and protection elements

                                    styles.cells_xfs.add(xf);
                                }
                                Event::End(end) if end.local_name().as_ref() == b"cellXfs" => break,
                                _ => {
                                    continue;
                                }
                            }
                        }
                    }
                }
                Event::Start(ref start) if start.local_name().as_ref() == b"indexedColors" => {
                    // <indexedColors>
                    loop {
                        buf_2.clear();
                        if let Ok(event) = reader.read_event_into(&mut buf_2) {
                            match event {
                                Event::Empty(start)
                                    if start.local_name().as_ref() == b"rgbColor" =>
                                {
                                    //   <rgbColor rgb="FF000000" />
                                    //   <rgbColor rgb="FFEDEDED" />

                                    if let Some(Ok(Attribute {
                                        key: QName(b"rgb"),
                                        value: rgb,
                                    })) = start.attributes().next()
                                    {
                                        styles.indexed_colors.add(RgbColor {
                                            rgb: unsafe {
                                                // SAFETY: document is known utf-8
                                                CompactString::from_utf8_unchecked(&rgb)
                                            },
                                        });
                                    }
                                }
                                Event::End(end)
                                    if end.local_name().as_ref() == b"indexedColors" =>
                                {
                                    // Parsed all `rgbColor` elements
                                    break;
                                }
                                _ => continue,
                            }
                        }
                    }
                }
                Event::Eof => break,
                _ => continue,
            }
        }
    }

    debug_assert_eq!(styles.fonts.fonts.len(), styles.fonts.count as usize);
    debug_assert_eq!(styles.cells_xfs.xfs.len(), styles.cells_xfs.count as usize);

    Ok(styles)
}

#[derive(Debug, Default)]
pub(super) struct Styles {
    fonts: Fonts,
    cells_xfs: CellXfs,
    indexed_colors: IndexedColors,
}

impl Styles {
    // PERF: Lots of cloning. Even if they are only copy, thats a lot of operations.
    // The challenge is that cannot move value out as other cells might need the same data.
    fn font(&self, index: usize) -> crate::wip::style::Font {
        // dbg!(index);
        let cell_xfs = &self.cells_xfs.xfs[index];

        // dbg!(&cell_xfs);

        let font = &self.fonts.fonts[cell_xfs.font_id];

        // dbg!(&font);

        let color = font.color.as_ref().map_or_else(
            || CompactString::new("FF000000"),
            |some| match some {
                Color::Argb(color) => color.clone(),
                Color::Indexed(idx) => self
                    .indexed_colors
                    .rgb_colors
                    .get(*idx)
                    .map_or_else(|| CompactString::new("FF000000"), |some| some.rgb.clone()),
                Color::Theme { .. } => CompactString::new("FF000000"),
            },
        );

        crate::wip::style::Font {
            name: font.name.clone(),
            size: font.sz,
            color,
        }
    }
}

#[derive(Debug, Default)]
struct Fonts {
    fonts: Vec<Font>,
    count: u32,
}

impl Fonts {
    fn add(&mut self, font: Font) {
        self.fonts.push(font);
    }
}

// <font>
//     <b val="true" />
//     <sz val="17" />
//     <color rgb="FF11806A" />
//     <name val="Arial" />
//     <family val="0" />
//     <charset val="1" />
// </font>
#[derive(Debug, Default)]
struct Font {
    sz: f64,
    color: Option<Color>,
    name: CompactString,
    family: u8,
    charset: Option<u8>,
    b: Option<bool>,
}

#[derive(Debug)]
enum Color {
    Argb(CompactString),
    Indexed(usize),
    Theme { theme: u32, tint: Option<f32> },
}

impl Default for Color {
    fn default() -> Self {
        Self::Argb(CompactString::new("FFFFFFF"))
    }
}

#[derive(Debug, Default)]
struct Fills {
    fills: Vec<Fill>,
    count: u32,
}

#[derive(Debug, Default)]
struct Fill {
    pattern_fill: CompactString,
}

// struct Borders {
//     borders: Vec<Border>,
//     count: u32,
// }

// struct Border {

// }

#[derive(Debug, Default)]
struct CellStyleXfs {
    xfs: Vec<Xf>,
    count: u32,
}

#[derive(Debug, Default)]
struct CellXfs {
    xfs: Vec<Xf>,
    count: u32,
}

impl CellXfs {
    fn add(&mut self, xf: Xf) {
        self.xfs.push(xf);
    }
}

#[derive(Debug, Default)]
struct Xf {
    num_fmt_id: u32,
    font_id: usize,
    fill_id: u32,
    border_id: u32,
    xf_id: u32,
    apply_font: bool,
    apply_border: bool,
    apply_alignment: bool,
    apply_protection: bool,
    alignment: Option<Alignment>,
    protection: Option<Protection>,
}

#[derive(Debug, Default)]
struct Alignment {
    horizontal: CompactString,
    vertical: CompactString,
    text_rotation: u32,
    wrap_text: bool,
    indent: u32,
    shrink_ro_fit: bool,
}

#[derive(Debug, Default)]
struct Protection {
    locked: bool,
    hidden: bool,
}

#[derive(Debug, Default)]
struct IndexedColors {
    rgb_colors: Vec<RgbColor>,
}

impl IndexedColors {
    fn add(&mut self, rgb_color: RgbColor) {
        self.rgb_colors.push(rgb_color);
    }
}

#[derive(Debug, Default)]
struct RgbColor {
    rgb: CompactString,
}

pub(super) fn _relationships(_archive: &mut ZipArchive<BufReader<File>>) -> Vec<CompactString> {
    todo!()
}

#[inline]
#[allow(clippy::too_many_lines)]
pub(super) fn worksheet<'a>(
    worksheet: &mut Worksheet<'a>,
    archive: &mut ZipArchive<BufReader<File>>,
    shared_strings: &'a [CompactString],
    styles: &'a Styles,
) -> Result<Option<()>, Error> {
    let file = match archive.by_name(&format!("xl/worksheets/sheet{}.xml", worksheet.id)) {
        Ok(ok) => ok,
        Err(ZipError::FileNotFound) => return Ok(None),
        Err(err) => return Err(err.into()),
    };

    let mut reader = Reader::from_reader(BufReader::new(file));

    reader.check_end_names(false);

    let mut buf_1: Vec<u8> = Vec::with_capacity(1024);
    let mut buf_2: Vec<u8> = Vec::with_capacity(1024);
    let mut buf_3: Vec<u8> = Vec::with_capacity(1024);

    let mut dimensions_are_known = false;

    loop {
        if let Ok(event) = &reader.read_event_into(&mut buf_1) {
            match event {
                Event::Empty(start) if start.local_name().as_ref() == b"dimension" => {
                    if let Some(Ok(Attribute {
                        key: QName(b"ref"),
                        value: dimensions,
                    })) = start.attributes().next()
                    {
                        let mut parts = dimensions.split(|char| *char == b':');

                        let top_left =
                            unsafe { excel_column_row_to_tuple_unchecked(parts.next().unwrap()) };
                        let bottom_right =
                            unsafe { excel_column_row_to_tuple_unchecked(parts.next().unwrap()) };

                        worksheet.spreadsheet.resize(top_left, bottom_right);

                        dimensions_are_known = true;
                    }
                }
                Event::Start(ref start) if start.local_name().as_ref() == b"c" => {
                    let mut attributes = start.attributes();

                    let mut cell = Cell::default();

                    // Go through each attribute and take relevant data from them.
                    // Example: <c r="A1" s="5" t="s">
                    // If no `t` then default to number
                    // Example:  <c r="E1" s="3">

                    if let Some(Ok(Attribute {
                        // Cell position.
                        // Example: "A1"
                        key: QName(b"r"),
                        value: ref pos,
                    })) = attributes.next()
                    {
                        // SAFETY: document is known valid utf8
                        // `r="A1"` getting the `A1` part
                        let (column, row) = unsafe { excel_column_row_to_tuple_unchecked(pos) };

                        cell.column = column;
                        cell.row = row;
                    }

                    if let Some(Ok(Attribute {
                        // Style
                        key: QName(b"s"),
                        value: ref style_idx,
                    })) = attributes.next()
                    {
                        let idx = unsafe {
                            std::str::from_utf8_unchecked(style_idx)
                                .parse::<usize>()
                                .unwrap()
                        };

                        cell.font = Cow::Owned(styles.font(idx));
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
                            b"is" => cell.r#type = Some(Type::String),
                            // formula
                            b"f" => cell.r#type = Some(Type::Formula),
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

                    // If there is no `t` attribute, then default to `Type::Number`
                    if cell.r#type.is_none() {
                        cell.r#type = Some(Type::Number);
                    }

                    // Get relevant value in cell
                    buf_2.clear();
                    if let Ok(Event::Start(start)) = reader.read_event_into(&mut buf_2) {
                        // Get the value of `<v>VALUE</v>` if it exists. If it doesnt, then the cell will have a vlue of `None`
                        if start.local_name().as_ref() == b"v" {
                            if let Ok(Event::Text(text)) = reader.read_event_into(&mut buf_3) {
                                match cell
                                    .r#type
                                    .expect("must have a type if there is a value element found")
                                {
                                    Type::Number => {
                                        cell.value = Some(Cow::Owned(CompactString::new(
                                            &text.unescape().unwrap(),
                                        )));
                                    }
                                    Type::String => {
                                        let idx = unsafe {
                                            std::str::from_utf8_unchecked(&text)
                                                .parse::<usize>()
                                                .unwrap()
                                        };

                                        cell.value = Some(Cow::Borrowed(&shared_strings[idx]));
                                    }
                                    Type::Formula => todo!(),
                                };
                            }
                        }
                    }

                    if dimensions_are_known {
                        worksheet.spreadsheet.insert_exact(cell);
                    } else {
                        worksheet.spreadsheet.insert(cell);
                    }
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

/// # Safety
///
/// Must only pass in valid ascii `A-Z` columns and `0-9` rows.
#[inline(always)]
unsafe fn excel_column_row_to_tuple_unchecked(pos: &[u8]) -> (u32, u32) {
    let mut idx = 0;

    for char in pos {
        if *char > 64 {
            idx += 1;
            continue;
        }
        break;
    }

    let column = excel_column_to_number_unchecked(&pos[..idx]);
    let row = std::str::from_utf8_unchecked(&pos[idx..])
        .parse::<u32>()
        .unwrap();

    // zero based position
    (column - 1, row - 1)
}

/// # Panics
///
/// Will panic if the provided string contains a letter other than `A-Z`.
#[inline(always)]
unsafe fn excel_column_to_number_unchecked(column: &[u8]) -> u32 {
    let mut result = 0;
    let mut multiplier = 1;

    for char in column.iter().rev() {
        let digit = u32::from(*char) - 'A' as u32 + 1;
        result += digit * multiplier;
        multiplier *= 26;
    }

    result
}

#[cfg(test)]
mod test {

    use super::*;

    // extern crate test;
    // use test::{black_box, Bencher};

    #[test]
    fn should_convert_excel_letter_number_format_to_tuple() {
        let result = unsafe { excel_column_row_to_tuple_unchecked(b"A1") };

        assert_eq!((0, 0), result);

        let result = unsafe { excel_column_row_to_tuple_unchecked(b"A2") };

        assert_eq!((0, 1), result);

        let result = unsafe { excel_column_row_to_tuple_unchecked(b"B1") };

        assert_eq!((1, 0), result);
    }
}
