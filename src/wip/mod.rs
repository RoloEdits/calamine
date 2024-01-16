pub mod xlsx;
pub use xlsx::{Xlsx, XlsxError};

use compact_str::{CompactString, ToCompactString};
use std::{fmt::Display, path::Path};

pub trait Workbook {
    type Workbook;
    type Error: std::error::Error;

    fn open<P: AsRef<Path>>(path: P) -> Result<Self::Workbook, Self::Error>;

    fn worksheet(&mut self, sheet: &str) -> Result<Option<&Worksheet>, Self::Error>;

    fn worksheets(&self) -> &[Worksheet];
}

#[derive(Debug)]
pub struct Worksheet {
    name: CompactString,
    spreadsheet: Spreadsheet,
}

impl Worksheet {
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    // pub fn window<'a>(&self, column: u32, row: u32) -> &Window<'a> {
    //     todo!()
    // }

    #[inline]
    pub fn rows(&self) -> Rows<'_> {
        Rows {
            spreadsheet: &self.spreadsheet,
            row: 0,
        }
    }

    #[inline]
    pub fn column(&self, column: u32) -> Column<'_> {
        Column {
            spreadsheet: &self.spreadsheet,
            column,
            row: 0,
        }
    }

    // pub fn row(&self, row: u32) -> Row<'_> {
    //     Row {
    //         cells: self.spreadsheet,
    //         columns: 0,
    //         row,
    //     }
    // }

    #[inline]
    pub fn cells(&self) -> impl Iterator<Item = &Cell> {
        self.spreadsheet.cells.iter()
    }

    #[inline]
    pub fn size(&self) -> (u32, u32) {
        self.spreadsheet.size()
    }
}

#[derive(Debug, Clone)]
pub struct Cell {
    value: Option<CompactString>,
    r#type: Type,
    column: u32,
    row: u32,
    font: Font,
}

#[derive(Debug)]
struct CellBuilder {
    value: Option<CompactString>,
    r#type: Option<Type>,
    column: Option<u32>,
    row: Option<u32>,
    font: Option<Font>,
}

// TODO: Remove and just have the setters on `Cell`
impl CellBuilder {
    fn new() -> Self {
        Self {
            value: None,
            r#type: None,
            column: None,
            row: None,
            font: None,
        }
    }

    fn value(&mut self, value: impl Display) -> &mut Self {
        self.value = Some(value.to_compact_string());
        self
    }

    fn data_type(&mut self, data_type: Type) -> &mut Self {
        self.r#type = Some(data_type);
        self
    }

    fn position(&mut self, column: u32, row: u32) -> &mut Self {
        self.column = Some(column);
        self.row = Some(row);
        self
    }

    fn font(&mut self, font: Font) -> &mut Self {
        self.font = Some(font);
        self
    }

    fn build(&self) -> Cell {
        Cell {
            value: self.value.clone(),
            r#type: self.r#type.unwrap_or(Type::String),
            column: self.column.unwrap(),
            row: self.row.unwrap(),
            font: self
                .font
                .as_ref()
                .unwrap_or(&Font {
                    font: "Arial".to_compact_string(),
                    size: 12.0,
                    color: "000000".to_compact_string(),
                })
                .clone(),
        }
    }
}

impl Cell {
    // pub fn default_with_position(column: u32, row: u32) -> Self {
    //     Self {
    //         value: None,
    //         column,
    //         row,
    //         font: None,
    //         r#type: ,
    //     }
    // }

    #[inline]
    pub fn value(&self) -> Option<&str> {
        match self.value.as_ref() {
            Some(value) => Some(value.as_str()),
            None => None,
        }
    }

    #[inline]
    pub fn font(&self) -> &Font {
        self.font.as_ref()
    }
}

#[derive(Debug, Clone, Copy)]
pub enum Type {
    Number,
    String,
    Formula,
}

pub struct Rows<'a> {
    spreadsheet: &'a Spreadsheet,
    row: u32,
}

impl<'a> Iterator for Rows<'a> {
    type Item = Row<'a>;

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        let item = Some(Row {
            cells: self.spreadsheet.row(self.row)?,
            columns: self.spreadsheet.buffer_columns,
            row: self.row,
        });

        self.row += 1;

        item
    }
}

#[derive(Debug)]
pub struct Row<'a> {
    cells: &'a [Cell],
    columns: u32,
    row: u32,
}

impl<'a> Row<'a> {
    #[inline]
    pub fn column(&self, column: usize) -> Option<&Cell> {
        self.cells.get(column)
    }
}

// FIX: never reaches None and always returns the same cell
impl<'a> Iterator for Row<'a> {
    type Item = &'a Cell;

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        self.cells.iter().next()
    }
}

pub struct Column<'a> {
    spreadsheet: &'a Spreadsheet,
    column: u32,
    row: u32,
}

impl<'a> Iterator for Column<'a> {
    type Item = &'a Cell;

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        let item = self.spreadsheet.column(self.column, self.row);
        self.row += 1;
        item
    }
}

pub struct Window<'a> {
    spreadsheet: &'a Spreadsheet,
}

// PERF: can add variables that hold the lowest cell row and column position and use that as an offset for the buffer size.
// This would be an optimization for spreadsheets who's first actual cell is somewhere far off the origin (0, 0)
//
// This would bring into question how we determine the `size` of a spreadsheet:
//     - Do we only report the top left most cell to the bottom right most?
//     - Do we report the origin to the bottom right most cell?
#[derive(Debug)]
struct Spreadsheet {
    cells: Vec<Cell>,
    // `rows` and `columns` represents the cells max positions
    rows: u32,
    columns: u32,
    // buffer_* represent the underlying Vec size.
    // may be larger than what cell positions would indicate to help with frequent balancing/allocations.
    buffer_columns: u32,
    buffer_rows: u32,
}

impl Spreadsheet {
    pub fn new() -> Self {
        Self {
            cells: Vec::new(),
            buffer_columns: 0,
            buffer_rows: 0,
            rows: 0,
            columns: 0,
        }
    }

    pub fn with_dimensions(columns: u32, rows: u32) -> Self {
        let mut cells = Vec::with_capacity((columns * rows) as usize);

        for row in 0..rows {
            for column in 0..columns {
                let cell = CellBuilder::new().position(column, row).build();
                cells.push(cell);
            }
        }

        Self {
            cells,
            rows,
            columns,
            buffer_columns: columns,
            buffer_rows: rows,
        }
    }

    // WARN: Off-by-one hell.
    // Cells use a zero based position, but the columns and rows count is 1 based.
    // 0 columns and 0 rows indicate no cells in the spreadsheet and is therefore used in the `new` implementation.
    #[inline]
    pub fn insert(&mut self, cell: Cell) {
        self.rows = self.rows.max(cell.row + 1);
        self.columns = self.columns.max(cell.column + 1);

        if self.buffer_columns < cell.column + 1 {
            // cell.row can be larger than current known row count, so take the larger of the two.
            self.buffer_rows = self.buffer_rows.max(cell.row + 1);
            // cell.column is now the largest known cell, having been greater than the previous largest, thus serving as the total number of columns
            self.buffer_columns = self.buffer_columns.max(cell.column + 1);

            let mut cells: Vec<Cell> =
                Vec::with_capacity((self.buffer_columns * self.buffer_rows) as usize);

            // Fill with empty cells with the correct position set
            for row in 0..self.buffer_rows {
                for column in 0..self.buffer_columns {
                    cells.push(CellBuilder::new().position(column, row).build())
                }
            }

            debug_assert!(self.buffer_rows * self.buffer_columns == cells.len() as u32);
            debug_assert!(cells.capacity() == cells.len());

            // Put previous cells in their position
            for cell in self.cells.drain(..) {
                // Take old cell position and offset the index by the new max columns amount
                let idx = (cell.row * self.buffer_columns + cell.column) as usize;

                cells[idx] = cell;
            }

            let idx = (cell.row * self.buffer_columns + cell.column) as usize;
            cells[idx] = cell;

            self.cells = cells;

            // Having checked that the columns are the within bounds, if the incoming cells row + column is greater than the current capacity
            // then that can only mean that there is more rows needed.
        } else if self.buffer_rows < cell.row + 1 {
            self.cells.reserve(self.cells.len());

            debug_assert!(self.cells.capacity() == self.cells.len() * 2);

            // Speculatively grow to avoid frequent reallocations
            let rows = self.cells.capacity() as u32 / self.buffer_columns;

            for row in self.buffer_rows..rows {
                for column in 0..self.buffer_columns {
                    self.cells
                        .push(CellBuilder::new().position(column, row).build());
                }
            }

            self.buffer_rows = rows;

            // NOTE: asserts must come after rows reassign
            debug_assert!(self.buffer_rows * self.buffer_columns == self.cells.len() as u32);
            debug_assert!(self.cells.capacity() == self.cells.len());

            let idx = (cell.row * self.buffer_columns + cell.column) as usize;

            self.cells[idx] = cell;

            // Cell fits within bounds and can be added directly
        } else {
            let idx = (cell.row * self.buffer_columns + cell.column) as usize;

            self.cells[idx] = cell;
        }
    }

    #[inline]
    pub fn row(&self, row: u32) -> Option<&[Cell]> {
        // If self.columns is 0, then index becomes `0..0` which is in range for a `Spreadsheet::new()` spreadsheet.
        // Meaning that it would always return `Some([])` causing an infinite loop if used as a iterator.
        if self.cells.is_empty() {
            return None;
        }

        let idx = (row * self.columns) as usize..((row + 1) * self.columns) as usize;
        self.cells.get(idx)
    }

    #[inline]
    pub fn column(&self, column: u32, row: u32) -> Option<&Cell> {
        self.cells.get((row * self.columns + column) as usize)
    }

    pub fn size(&self) -> (u32, u32) {
        (self.columns, self.rows)
    }
}

#[derive(Debug, Default, Clone)]
pub struct Font {
    font: CompactString,
    size: f64,
    color: CompactString,
}

impl Font {
    #[inline]
    pub fn color(&self) -> &str {
        &self.color
    }
}

impl AsRef<Font> for Font {
    fn as_ref(&self) -> &Font {
        self
    }
}
