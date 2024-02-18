use super::cell::Cell;

// IDEA: For merged cells, might be able to use the first cell instance as the holder of data, and an Option<Cell> with any `None` indicating
// that its part of a merged cell and backtrack until the first Some(Cell).
//     - Won't work well with rows that are merged.
// PERF: can add variables that hold the lowest cell row and column position and use that as an offset for the buffer size.
// This would be an optimization for spreadsheets who's first actual cell is somewhere far off the origin (0, 0)
//
// This would bring into question how we determine the `size` of a spreadsheet:
//     - Do we only report the top left most cell to the bottom right most?
//     - Do we report the origin to the bottom right most cell?
#[derive(Debug)]
pub(crate) struct Spreadsheet<'a> {
    pub(crate) cells: Vec<Cell<'a>>,
    // `rows` and `columns` represents the cells max positions
    rows: u32,
    columns: u32,
    // buffer_* represent the underlying Vec size.
    // may be larger than what cell positions would indicate to help with frequent balancing/allocations.
    buffer_columns: u32,
    buffer_rows: u32,
}

impl<'a> Spreadsheet<'a> {
    #[inline]
    pub fn new() -> Self {
        Self {
            cells: Vec::new(),
            buffer_columns: 0,
            buffer_rows: 0,
            rows: 0,
            columns: 0,
        }
    }

    #[inline]
    pub fn with_size(columns: u32, rows: u32) -> Self {
        let mut cells = Vec::with_capacity((columns * rows) as usize);

        for row in 0..rows {
            for column in 0..columns {
                let cell = Cell::new(column, row);
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

    // TODO: As this might be used after its fill, should truncate off old cells but fill in ones that fit in the new dimensions.
    //     Currently deletes all cells and starts with a clean slate.
    // NOTE: Currently the min cell position, 0, 0 (A1), is always the considered the min.
    #[inline]
    pub fn resize(&mut self, _: (u32, u32), (max_col, max_row): (u32, u32)) {
        // incoming cell position is zero based, but the rows and columns count is 1 based.
        self.columns = max_col + 1;
        self.rows = max_row + 1;

        self.buffer_columns = max_col + 1;
        self.buffer_rows = max_row + 1;

        let mut cells: Vec<Cell> =
            Vec::with_capacity((self.buffer_columns * self.buffer_rows) as usize);

        // Fill with empty cells with the correct position set
        for row in 0..self.buffer_rows {
            for column in 0..self.buffer_columns {
                let cell = Cell::new(column, row);
                cells.push(cell);
            }
        }

        debug_assert!((self.buffer_rows * self.buffer_columns) as usize == cells.len());
        debug_assert!(cells.capacity() == cells.len());

        self.cells = cells;
    }

    // PERF: If the new dimensions are the same as the old, just different columns and rows, can reuse the existing buffer.
    // WARN: Off-by-one hell.
    // Cells use a zero based position, but the columns and rows count is 1 based.
    // 0 columns and 0 rows indicate no cells in the spreadsheet and is therefore used in the `new` implementation.
    #[inline(always)]
    pub fn insert(&mut self, cell: Cell<'a>) {
        self.columns = self.columns.max(cell.column + 1);
        self.rows = self.rows.max(cell.row + 1);

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
                    let cell = Cell::new(column, row);
                    cells.push(cell);
                }
            }

            debug_assert!((self.buffer_rows * self.buffer_columns) as usize == cells.len());
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
            // Speculatively grow to avoid frequent reallocations
            self.cells.reserve(self.cells.len());

            debug_assert!(self.cells.capacity() == self.cells.len() * 2);

            let rows = u32::try_from(self.cells.capacity() / self.buffer_columns as usize)
                .expect("rows can never overgrow a u32 ");

            for row in self.buffer_rows..rows {
                for column in 0..self.buffer_columns {
                    let cell = Cell::new(column, row);
                    self.cells.push(cell);
                }
            }

            self.buffer_rows = rows;

            // NOTE: asserts must come after rows reassign
            debug_assert!((self.buffer_rows * self.buffer_columns) as usize == self.cells.len());
            debug_assert!(self.cells.capacity() == self.cells.len());

            let idx = (cell.row * self.buffer_columns + cell.column) as usize;

            self.cells[idx] = cell;

            // Cell fits within bounds and can be added directly
        } else {
            let idx = (cell.row * self.buffer_columns + cell.column) as usize;

            self.cells[idx] = cell;
        }
    }

    /// # Panics
    ///
    /// When inserted cells' position is out of bounds of the current spreadsheet dimensions
    #[inline(always)]
    pub fn insert_exact(&mut self, cell: Cell<'a>) {
        let idx = (cell.row * self.buffer_columns + cell.column) as usize;
        self.cells[idx] = cell;
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

    // #[inline]
    // pub fn row_mut(&mut self, row: u32) -> Option<&mut [Cell]> {
    //     // If self.columns is 0, then index becomes `0..0` which is in range for a `Spreadsheet::new()` spreadsheet.
    //     // Meaning that it would always return `Some([])` causing an infinite loop if used as a iterator.
    //     if self.cells.is_empty() {
    //         return None;
    //     }

    //     let idx = (row * self.columns) as usize..((row + 1) * self.columns) as usize;
    //     self.cells.get_mut(idx)
    // }

    #[inline]
    pub fn column(&self, column: u32, row: u32) -> Option<&Cell> {
        self.cells.get((row * self.columns + column) as usize)
    }

    #[inline]
    pub fn size(&self) -> (u32, u32) {
        (self.columns, self.rows)
    }
}
// pub struct Window<'a> {
//     spreadsheet: &'a Spreadsheet,
// }

pub struct Rows<'a> {
    pub(crate) spreadsheet: &'a Spreadsheet<'a>,
    pub(crate) row: u32,
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

// pub struct RowsMut<'a> {
//     spreadsheet: &'a mut Spreadsheet,
//     row: u32,
// }

// impl<'a> Iterator for RowsMut<'a> {
//     type Item = RowMut<'a>;

//     #[inline]
//     fn next(&mut self) -> Option<Self::Item> {
//         let item = Some(RowMut {
//             cells: &mut self.spreadsheet.row_mut(self.row)?,
//             columns: self.spreadsheet.buffer_columns,
//             row: self.row,
//         });

//         self.row += 1;

//         item
//     }
// }

#[derive(Debug)]
pub struct Row<'a> {
    pub(crate) cells: &'a [Cell<'a>],
    pub(crate) columns: u32,
    pub(crate) row: u32,
}

impl<'a> Row<'a> {
    #[inline]
    #[must_use]
    pub fn column(&self, column: usize) -> Option<&Cell> {
        self.cells.get(column)
    }
}

// FIX: never reaches None and always returns the same cell
impl<'a> Iterator for Row<'a> {
    type Item = &'a Cell<'a>;

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        self.cells.iter().next()
    }
}

// #[derive(Debug)]
// pub struct RowMut<'a> {
//     cells: &'a mut [Cell],
//     columns: u32,
//     row: u32,
// }

pub struct Column<'a> {
    pub(crate) spreadsheet: &'a Spreadsheet<'a>,
    pub(crate) column: u32,
    pub(crate) row: u32,
}

impl<'a> Iterator for Column<'a> {
    type Item = &'a Cell<'a>;

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        let item = self.spreadsheet.column(self.column, self.row);
        self.row += 1;
        item
    }
}
