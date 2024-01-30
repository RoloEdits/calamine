mod error;
mod parse;

use compact_str::CompactString;
use std::{fs::File, io::BufReader};
use zip::ZipArchive;

pub use self::error::Error;
use self::parse::Styles;

use super::{Spreadsheet, WorkbookImpl, Worksheet};

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
