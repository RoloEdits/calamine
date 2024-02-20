pub mod xlsx;
use quick_xml::Reader;
pub use xlsx::{Error, Xlsx};
mod cell;
mod spreadsheet;
mod style;

use cell::Cell;
use compact_str::CompactString;
use spreadsheet::Spreadsheet;
use std::{io::BufReader, path::Path};
use zip::read::ZipFile;

// Here so that there is a cleaner API.
// Rather than having the logic fall to the user of library, matching on an exstention
// we can just do it here and allow a simple `Workbook::open()` that will take any of the supported file exstentions.
// pub enum Workbook<'a> {
//     Xlsx(Xlsx<'a>),
// }

// impl<'a> Workbook<'a> {
//     #[inline]
//     pub fn open<P: AsRef<Path>>(path: P) -> Result<Self, Box<dyn std::error::Error>> {
//         let workbook = match path
//             .as_ref()
//             .extension()
//             .expect("file must have extension")
//             .to_str()
//             .unwrap()
//         {
//             "xlsx" => {
//                 let xslx = Xlsx::open(path)?;
//                 Workbook::Xlsx(xslx)
//             }
//             unsupported => panic!("unsupported file type `{unsupported}`"),
//         };

//         Ok(workbook)
//     }

//     pub fn worksheets(&'a mut self) -> &mut [Worksheet] {
//         match self {
//             Workbook::Xlsx(xlsx) => xlsx.worksheets(),
//         }
//     }

//     /// Returns `Some` if worksheet exists in the workbook, otherwise returns `None`.
//     #[inline]
//     pub fn worksheet(&'a mut self, worksheet: impl AsRef<str>) -> Option<&mut Worksheet> {
//         let worksheet = match self {
//             Workbook::Xlsx(xlsx) => xlsx.worksheet(worksheet),
//         };

//         worksheet
//     }

// }

pub trait Workbook<'a> {
    type Workbook;
    type Error: std::error::Error;

    fn open<P: AsRef<Path>>(path: P) -> Result<Self::Workbook, Self::Error>;

    fn worksheet<'b: 'a>(
        &'a mut self,
        worksheet: &str,
    ) -> Option<&'b mut Worksheet<Self::Workbook>>
    where
        <Self as Workbook<'a>>::Workbook: Workbook<'a>;
}

// IDEA: When trying to implement a `rows_lazy()`, being able to reuse a buffer would be more performant.
//     - Can look to `File::lines()` and `File::read(&mut Vec<u8>)` as options to add for different use cases depending on the situation.
//         - Something like `Worksheet::rows()` would be the full read all at once.
//         - `Worksheet::rows_into(&mut Vec<Cell>)` would be a more memory effiecient way to read very very large files.
//            The worksheet would get parsed till the end of each row and then return an iterator over that row and then `.clear().`
//            the buffer to reuse the allocated space. Spreadsheet would need a `Spreadsheet::new_with_buffer(&mut Vec<Cell>)`.
//            Would also need to implement the optimization for min and max cell positions and abstracting the cells position over the buffer size.
//            If the buffer can fit the inserted cell in the current min max cell positions, then the buffer can be used to trim excess empty cells.
//         - This change would mean that the parsing would need to happen outside of the current `.worksheet()` function, and in the various accesseors.
//           Potentially could even have different Workbook methods to create lazy versions of opening used throughout the rest of its usage.
//           A read only variant. `WorkbookLazyImpl` for a `XlsxLazy` struct. Its `.rows()` would then need a buffer, `.rows(&mut Vec<Cell>)`.
//           `Worksheet::rows(&mut Vec<Cell>)`
pub struct Worksheet<'a, W: Workbook<'a>> {
    id: u32,
    name: CompactString,
    reader: Option<Reader<BufReader<ZipFile<'a>>>>,
    spreadsheet: Spreadsheet<'a>,
    workbook: &'a W,
}

impl<'a, W: Workbook<'a>> Worksheet<'a, W> {
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[inline]
    pub fn cells(&self) -> impl Iterator<Item = &Cell> {
        self.spreadsheet.cells.iter()
    }

    #[inline]
    pub fn size(&self) -> (u32, u32) {
        self.spreadsheet.size()
    }

    #[inline]
    pub fn insert_cell(&mut self, cell: Cell<'a>) {
        self.spreadsheet.insert(cell);
    }

    #[inline]
    pub fn insert_cell_exact(&mut self, cell: Cell<'a>) {
        self.spreadsheet.insert_exact(cell);
    }

    #[inline]
    pub fn insert_cells(&mut self, cells: Vec<Cell<'a>>) {
        for cell in cells {
            self.spreadsheet.insert(cell);
        }
    }

    #[inline]
    pub fn insert_cells_exact(&mut self, cells: Vec<Cell<'a>>) {
        for cell in cells {
            self.spreadsheet.insert_exact(cell);
        }
    }
}
