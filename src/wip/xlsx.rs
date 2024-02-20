mod error;
mod parse;

use compact_str::CompactString;
use quick_xml::Reader;
use std::{fs::File, io::BufReader};
use zip::{result::ZipError, ZipArchive};

pub use self::error::Error;
use self::parse::Styles;

use super::{spreadsheet::Rows, Workbook, Worksheet};

const MAX_ROWS: u32 = 1_048_576;
const MAX_COLUMNS: u16 = 16_384;

/// 255 characters
const MAX_COLUMN_WIDTH: u8 = 255;

// 409 points
const MAX_ROW_HEIGHT: u16 = 409;

/// 1,026 horizontal and vertical
const MAX_PAGE_BREAKS: u16 = 1_026;

/// Total number of characters that a cell can contain
const MAX_CELL_CHARACTERS: u16 = 32_767;

/// Characters in a header or footer
const MAX_HEADER_FOOTER_CHARACTERS: u8 = 255;

/// Maximum number of line feeds per cell
const MAX_CELL_LINE_FEEDS: u8 = 253;

/// Unique cell formats/cell styles
const MAX_CELL_STYLES: u16 = 65_490;

const MAX_FILL_STYLES: u16 = 256;

const MAX_LINE_WEIGHT_STYLES: u16 = 256;

const MAX_UNIQUE_GLOBAL_FONTS: u16 = 1_024;
const MAX_UNIQUE_WORKBOOK_FONTS: u16 = 512;

/// Between 200 and 250, depending on the language version of Excel that you have installed
const MAX_NUMBER_FORMATS: u8 = 250;

/// Hyperlinks in a worksheet
const MAX_WORKSHEET_HYPERLINKS: u16 = 65_530;

const MAX_WINDOW_PANES: u8 = 4;

/// Changing cells in a scenario
const MAX_CHANGE_CELLS: u8 = 32;

/// Adjustable cells in Solver
const MAX_ADJUSTABLE_CELLS: u8 = 200;

const MAX_ZOOM: f64 = 400.0;
const MIN_ZOOM: f64 = 10.0;

const MAX_DATA_FORM_FIELDS: u8 = 32;

const MIN_DATE: i64 = -2_208_960_000;
const MIN_1904_DATE: i64 = -2_082_816_000;
const MAX_DATE: i64 = 253_402_329_599;

mod format {
    pub struct Xlsx;
}

pub struct Xlsx<'a> {
    archive: ZipArchive<BufReader<File>>,
    worksheets: Vec<Worksheet<'a, Self>>,
    shared_strings: Vec<CompactString>,
    theme: Vec<CompactString>,
    styles: Styles,
    relationships: Vec<CompactString>,
}

impl<'a> Workbook<'a> for Xlsx<'a> {
    type Workbook = Xlsx<'a>;
    type Error = Error;

    #[inline]
    fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Self::Workbook, Self::Error> {
        let file = BufReader::new(File::open(path)?);
        let mut archive = ZipArchive::new(file)?;

        let worksheets = parse::workbook(&mut archive)?;
        let styles = parse::styles(&mut archive)?;
        let shared_strings = parse::shared_strings(&mut archive)?;

        Ok(Xlsx {
            archive,
            worksheets,
            shared_strings,
            theme: Vec::new(),
            styles,
            relationships: Vec::new(),
        })
    }

    #[inline]
    fn worksheet<'b: 'a>(&'a mut self, name: &str) -> Option<&'b mut Worksheet<Self::Workbook>> {
        for worksheet in &mut self.worksheets {
            if worksheet.name == name {
                let file = match self
                    .archive
                    // INFO: All sheet file names are in the form of `sheet` and then the `sheetId`.
                    // Before this function is called, the worksheet is filtered to the matching `name`
                    // and then the id is used to get the sheet xml file handle.
                    .by_name(&format!("xl/worksheets/sheet{}.xml", worksheet.id))
                {
                    Ok(file) => file,
                    Err(ZipError::FileNotFound) => return None,
                    _ => panic!("error trying to get file handle for `{name}`"),
                };

                let mut reader = Reader::from_reader(BufReader::new(file));
                reader.check_end_names(false);

                worksheet.reader = Some(reader);

                return Some(worksheet);
            }
        }

        None
    }
}

impl<'a> Worksheet<'a, Xlsx<'a>> {
    pub fn rows(&'a mut self) -> Result<Rows<'a>, Box<dyn std::error::Error>> {
        parse::worksheet(
            self.reader
                .as_mut()
                .expect("reader is set when worksheet handle is gotten"),
            &mut self.spreadsheet,
            self.workbook,
        )?;

        Ok(Rows {
            spreadsheet: &self.spreadsheet,
            row: 0,
        })
    }
}
