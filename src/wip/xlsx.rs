mod error;
mod parse;

use compact_str::CompactString;
use std::{fs::File, io::BufReader};
use zip::ZipArchive;

pub use self::error::Error;
use self::parse::Styles;

use super::{Spreadsheet, WorkbookImpl, Worksheet};

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

pub struct Xlsx<'a> {
    archive: ZipArchive<BufReader<File>>,
    theme: Vec<CompactString>,
    relationships: Vec<CompactString>,
    styles: Styles,
    shared_strings: Vec<CompactString>,
    worksheets: Vec<Worksheet<'a>>,
}

impl<'a> WorkbookImpl<'a> for Xlsx<'a> {
    type Workbook = Xlsx<'a>;
    type Error = Error;

    fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Self::Workbook, Self::Error> {
        let file = BufReader::new(File::open(path)?);
        let mut archive = ZipArchive::new(file)?;

        let shared_strings = parse::shared_strings(&mut archive)?;

        let styles = parse::styles(&mut archive)?;

        // println!("styles = {styles:#?}");
        // for (idx, style) in styles.fonts.iter().enumerate() {
        //     println!("styles[{}] = `{}`", idx + 1, style.color());
        // }

        // let relationships = parse::relationships(&mut archive);

        let mut worksheets = Vec::new();

        // Gets names of worksheets
        // for worksheet in parse::relationships(&mut archive) {
        //     worksheets.push(Worksheet {
        //         name: worksheet,
        //         spreadsheet: Spreadsheet::new(),
        //     })
        // }

        // TEMP:
        worksheets.push(Worksheet {
            name: CompactString::new("sheet1"),
            spreadsheet: Spreadsheet::new(),
        });

        Ok(Xlsx {
            archive,
            theme: Vec::new(),
            relationships: Vec::new(),
            styles,
            shared_strings,
            worksheets,
        })
    }

    fn worksheet(
        &'a mut self,
        worksheet: impl AsRef<str>,
    ) -> Result<Option<&Worksheet>, Self::Error> {
        for __worksheet in &mut self.worksheets {
            if __worksheet.name == worksheet.as_ref().to_lowercase() {
                if let Ok(Some(())) = parse::worksheet(
                    __worksheet,
                    &mut self.archive,
                    &self.shared_strings,
                    &self.styles,
                ) {
                    return Ok(Some(__worksheet));
                };
            }
        }

        Ok(None)
    }

    fn worksheets(&self) -> &[Worksheet] {
        &self.worksheets
    }

    fn add_worksheet<'w: 'a>(
        &mut self,
        worksheet: Worksheet<'w>,
    ) -> Result<(), Box<dyn std::error::Error>> {
        if self
            .worksheets
            .iter()
            .any(|__worksheet| __worksheet.name == worksheet.name)
        {
            return Err("Sheet with name already exists".into());
        }

        self.worksheets.push(worksheet);
        Ok(())
    }
}
