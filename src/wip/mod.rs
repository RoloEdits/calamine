pub mod xlsx;
pub use xlsx::{Xlsx, XlsxError};

use compact_str::CompactString;
use std::path::Path;

pub trait Workbook {
    type Workbook;
    type Error: std::error::Error;

    fn open<P: AsRef<Path>>(path: P) -> Result<Self::Workbook, Self::Error>;

    fn worksheet(&mut self, sheet: &str) -> Result<Option<Worksheet>, Self::Error>;
}

#[derive(Debug)]
pub struct Worksheet {
    name: CompactString,
    grid: Grid,
    size: (u32, u32),
}

impl Worksheet {
    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn cells(&self) -> impl Iterator<Item = &Cell> {
        self.grid.cells.iter()
    }

    pub fn rows(&self) -> Rows<'_> {
        Rows {
            grid: &self.grid,
            row: 0,
        }
    }

    // pub fn window<'a>(&self, column: u32, row: u32) -> &Window<'a> {
    //     todo!()
    // }

    pub fn column(&self, column: u32) -> Column<'_> {
        Column {
            grid: &self.grid,
            column,
            row: 0,
        }
    }

    // pub fn row(&self, row: u32) -> Row<'_> {
    //     Row {
    //         cells: self.grid,
    //         columns: 0,
    //         row,
    //     }
    // }

    pub fn dimensions(&self) -> (u32, u32) {
        self.size
    }
}

pub struct Rows<'a> {
    grid: &'a Grid,
    row: u32,
}

impl<'a> Iterator for Rows<'a> {
    type Item = Row<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        let item = Some(Row {
            cells: self.grid.row(self.row)?,
            columns: self.grid.buffer_columns,
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
    pub fn column(&self, column: usize) -> Option<&Cell> {
        self.cells.get(column)
    }
}

impl<'a> Iterator for Row<'a> {
    type Item = &'a Cell;

    fn next(&mut self) -> Option<Self::Item> {
        self.cells.iter().next()
    }
}

pub struct Column<'a> {
    grid: &'a Grid,
    column: u32,
    row: u32,
}

impl<'a> Iterator for Column<'a> {
    type Item = &'a Cell;

    fn next(&mut self) -> Option<Self::Item> {
        let item = self.grid.column(self.column, self.row);
        self.row += 1;
        item
    }
}

pub struct Window<'a> {
    grid: &'a Grid,
}

#[derive(Debug, Clone)]
pub struct Cell {
    value: Option<CompactString>,
    column: u32,
    row: u32,
    font: Option<Font>,
}

impl Cell {
    pub fn default_with_position(column: u32, row: u32) -> Self {
        Self {
            value: None,
            column,
            row,
            font: None,
        }
    }

    pub fn value(&self) -> Option<&str> {
        let Some(value) = self.value.as_ref() else {
            return None;
        };

        Some(value.as_str())
    }

    pub fn font(&self) -> Option<&Font> {
        self.font.as_ref()
    }
}

#[derive(Debug)]
struct Grid {
    cells: Vec<Cell>,
    // `rows` and `columns` represents the cells max positions
    rows: u32,
    columns: u32,
    // real_* represent the underlying Vec size.
    // may be larger than what cell positions would indicate to help with frequent balancing/allocations.
    buffer_columns: u32,
    buffer_rows: u32,
    // populated: u32,
    // empty: u32,
}

impl Grid {
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
                cells.push(Cell::default_with_position(column, row))
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
                    cells.push(Cell::default_with_position(column, row))
                }
            }

            assert!(self.buffer_rows * self.buffer_columns == cells.len() as u32);
            assert!(cells.capacity() == cells.len());

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

            assert!(self.cells.capacity() == self.cells.len() * 2);

            // Speculatively grow to avoid frequent reallocations
            let rows = self.cells.capacity() as u32 / self.buffer_columns;

            for row in self.buffer_rows..rows {
                for column in 0..self.buffer_columns {
                    self.cells.push(Cell::default_with_position(column, row));
                }
            }

            self.buffer_rows = rows;

            // NOTE: asserts must come after rows reassign
            assert!(self.buffer_rows * self.buffer_columns == self.cells.len() as u32);
            assert!(self.cells.capacity() == self.cells.len());

            let idx = (cell.row * self.buffer_columns + cell.column) as usize;

            self.cells[idx] = cell;

            // Cell fits within bounds and can be added directly
        } else {
            let idx = (cell.row * self.buffer_columns + cell.column) as usize;

            self.cells[idx] = cell;
        }
    }

    pub fn row(&self, row: u32) -> Option<&[Cell]> {
        let idx = (row * self.columns) as usize..((row + 1) * self.columns) as usize;
        self.cells.get(idx)
    }

    pub fn column(&self, column: u32, row: u32) -> Option<&Cell> {
        self.cells.get((row * self.columns + column) as usize)
    }
}

#[derive(Debug, Default, Clone)]
pub struct Font {
    font: CompactString,
    size: u16,
    color: CompactString,
}

impl Font {
    pub fn color(&self) -> &str {
        &self.color
    }
}

impl AsRef<Font> for Font {
    fn as_ref(&self) -> &Font {
        self
    }
}
