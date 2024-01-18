mod error;
mod parse;

use compact_str::CompactString;
use std::{fs::File, io::BufReader};
use zip::ZipArchive;

pub use self::error::XlsxError;

use super::{Font, Spreadsheet, WorkbookTrait, Worksheet};

pub struct Xlsx {
    archive: ZipArchive<BufReader<File>>,
    theme: Vec<CompactString>,
    relationships: Vec<CompactString>,
    styles: Vec<Font>,
    shared_strings: Vec<CompactString>,
    worksheets: Vec<Worksheet>,
}

impl WorkbookTrait for Xlsx {
    type Workbook = Xlsx;
    type Error = XlsxError;

    fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Self::Workbook, Self::Error> {
        let file = BufReader::new(File::open(path)?);
        let mut archive = ZipArchive::new(file)?;

        println!("acquired handle");

        let shared_strings = parse::shared_strings(&mut archive)?;

        println!("parsed shared strings");

        let styles = parse::styles(&mut archive)?;

        println!("parsed styles");

        // let relationships = parse::relationships(&mut zip);

        let mut worksheets = Vec::new();

        // Gets names of worksheets
        // for worksheet in parse::relationships(&mut archive) {
        //     worksheets.push(Worksheet {
        //         name: worksheet,
        //         spreadsheet: Spreadsheet::new(),
        //     })
        // }

        // TEMP
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

    fn worksheet(&mut self, worksheet: impl AsRef<str>) -> Result<Option<&Worksheet>, Self::Error> {
        for _worksheet in self.worksheets.iter_mut() {
            if _worksheet.name == worksheet.as_ref().to_lowercase() {
                if let Ok(Some(_)) = parse::worksheet(
                    _worksheet,
                    &mut self.archive,
                    &self.shared_strings,
                    &self.styles,
                ) {
                    return Ok(Some(_worksheet));
                };
            }
        }

        println!("found none");

        Ok(None)
    }

    fn worksheets(&self) -> &[Worksheet] {
        &self.worksheets
    }
}
