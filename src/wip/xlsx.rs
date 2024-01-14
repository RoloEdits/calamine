mod error;
mod parse;

use compact_str::CompactString;
use std::{fs::File, io::BufReader};
use zip::ZipArchive;

pub use self::error::XlsxError;

use super::{Font, Workbook, Worksheet};

pub struct Xlsx {
    archive: ZipArchive<BufReader<File>>,
    theme: Vec<CompactString>,
    relationships: Vec<CompactString>,
    styles: Vec<Font>,
    shared_strings: Vec<CompactString>,
    // sheets: Vec<Worksheet>,
}

impl Workbook for Xlsx {
    type Workbook = Xlsx;
    type Error = XlsxError;

    // NOTE: Would need a sheets name list that has all the sheets names in it. This is used bu Nushell to list them.
    fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Self::Workbook, Self::Error> {
        let file = BufReader::new(File::open(path)?);
        let mut archive = ZipArchive::new(file)?;

        let shared_strings = parse::shared_strings(&mut archive)?;
        let styles = parse::styles(&mut archive)?;
        // let relationships = parse::relationships(&mut zip);

        Ok(Xlsx {
            archive,
            theme: Vec::new(),
            relationships: Vec::new(),
            styles,
            shared_strings,
            // sheets: Vec::new(),
        })
    }

    fn worksheet(&mut self, worksheet: &str) -> Result<Option<Worksheet>, Self::Error> {
        parse::worksheet(
            &mut self.archive,
            &self.shared_strings,
            &self.styles,
            worksheet,
        )
    }
}
