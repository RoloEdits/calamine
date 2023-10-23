
# Changes for v1.0.0

## Summary

## Things of Note

- [ ] Get rid of `Seek` requirement https://github.com/tafia/calamine/issues/164
- [ ] maintain/improve support for wasm
- [ ] async api https://github.com/tafia/calamine/issues/286
- [ ] lazy iterator over rows/columns https://github.com/tafia/calamine/issues/27
- [ ] Associating data and formula https://github.com/tafia/calamine/issues/247
- [ ] Ignore hidden rows https://github.com/tafia/calamine/issues/237
- [ ] Support for Google Sheets https://github.com/tafia/calamine/issues/162
- [ ] Getting the name of the author https://github.com/tafia/calamine/issues/149
- [ ] Support SPSS files https://github.com/tafia/calamine/issues/112
- [ ] XLSX with passwords https://github.com/tafia/calamine/issues/102 https://github.com/nushell/nushell/issues/10546
- [ ] Get sub-view of sheet https://github.com/tafia/calamine/issues/147
- [ ] Multi-threaded reading https://github.com/tafia/calamine/issues/346
- [ ] Get cell style/color info https://github.com/tafia/calamine/issues/191
- [ ] VBA from open docs
- [ ] Writing files out
- [ ] Experiment in the overhead of dynamic dispatch for cells types
- [ ] File based backend with either a big data format or a file database

## Goals

The main focus of the proposed changes is to:
- Be more hierarchical in data access. Having a clear order of Workbook -> Sheet -> Cell access pattern; The data should belong to the logical scope.  

  - Example: A Workbook has a theme, and so does a sheet, but the colors of a cell can be changed in scope of the cell. Therefore, the color should then be accessed through the Cell, as thats the final link in the chain that can alter the thing we want.

- Ergonomics, with lower exposure of internals. And on top of conforming to data scope access, there should be a focus on providing straight forward ways to handle various data access patterns.
  
  - Example: When looking into implementing [#362](https://github.com/tafia/calamine/issues/362), it would require a pretty direct breaking change as the `String` type is directly exposed when getting the value out. Instead, for example, exposing a `&str` would allow the backing memory to be an implementation detail. Or having getters on the cell to get the value out without ever have to destruct an enum and go through all possible values.  
 
  - Example: The current `fn worksheets(&mut self) -> Vec<(String, Range<DataType>)>;` returns path information and the cells. These should just be in a `Worksheet` type with fields for `name`, `path`, `cells`, etc., accessed through the thing that should logically own it: `sheet.name()`, `sheet.path()` and `sheet.cells()`. Even further, a `Cell` should be its own type, further encapsulation the internals.
 
  - Example: Exposing different ways to open workbooks, based on the conditions of the data one has.
      -  `open()` which will maintain the structure of the workbook and sheets and load it all up front. A method that is very aggressive in both parsing and iterating over the data as fast as possible to matter the memory usage. 
      -  `open_sparse()` that will remove empty values of cells and make the structure as compact as possible. This will break row-column access but it would be lighter on memory, but might be best for the type of access a user might want. 
      -  `open_lazy()` that reads and build only the 
      
- Performance. This is in terms of wall times, but also memory "performance", i.e. its usage. The goal should be to provide the best performance in our solutions. 
  
  - Example: In `xlsx` files, there is a shared string component that is linked by an index in the corresponding sheet/s. In a worse case scenario the sheet could have nothing but strings, with each string being unique. This means that the shared string file would be at least as long as any given sheet. The parsing of these can files can be done in parallel. The backing string `Vec` can be of an SSO string type, like [`compact_str`](https://github.com/ParkMyCar/compact_str), and only needs to share a reference of `&str`, to the `DataType` enum. The string `Vec` acting as an arena. The `Cell` then only needs to store a `&str` with the index into the string arena, mapping very well to the indexing of the `xlsx` spec. Having a `Vec` where possible should also improve general performance as look-ups are `O(1)` and they are more cache friendly.

- Continuing and expanding outside integrations. Along with `serde`, include support for `rayon`, `polars` and anything else where it makes sense in the scope of various domains where excel might be used in and along side.   

## Proposals Through Examples
### Read
#### Imports

The first part of the proposed changes is to simplify the direct imports. A `Workbook` trait and the relevant implementing struct and its errors:

```rust
use calamine::{Workbook, Xlsx, XlsxError};
```

Using the traits fully qualified syntax to open the workbook:

```rust
let workbook: Result<Xlsx<'_, '_>, XlsxError> = Workbook::open("file.xlsx");
```

This would leave the reading implementation scoped to the workbook type, solving this issue of the present API:
```rust  
// FIXME `Reader` must only be seek `Seek` for `Xls::xls`. Because of the present API, this limits the kinds of readers (other) data in formats can be read from.
```

The current `open_workbook_auto()` would go away completely in the new API, opting for user explicitness to help enforce correctness.

#### Workbook

```rust
// Gets all the sheets in the workbook. If there are no sheets then return `None`.
let sheets: Option<Worksheets<'_>> = workbook.worksheets();
let sheets: Worksheets<'_> = workbook.worksheets_by_name(&["Employees", "Supplies"]);

let author: &str = workbook.author();

// Using traits we can move to support both strings and numbers to get the
// desired sheet.
//
// If the sheet does not exist in the workbook then return `None`
let sheet: Option<&Workheet<'_>> = workbook.worksheet("Sheet1");
let sheet: Option<&Workheet<'_>> = workbook.worksheet(0);
```
#### Sheet

```rust
// It contains methods to get info, like the amount of sheets.
let count: size = sheets.count();

// `Sheets<'_>` is also an iterator over all the sheets, yielding a `Sheet<'_>`.
for sheet in workbook.worksheets() {
	// Gets the name of the sheet
	let name: &str = sheet.name();
	
	// Can also look into returning a `&Path` instead
	let path: &str = sheet.path();
	
	// Returns the columns and rows count to make up the size of the sheet
	// This only counting the used cells, not from A1.
	let dimensions: (u32, u32) = sheet.dimensions();
	
	// Returns the top left most column-row pair
	let start: (u32, u32) = sheet.start();
	// Returns the bottom right most column-row pair
	let end: (u32, u32) = sheet.end();
	
	// Gets the specified cell from the sheet
	//
	// Using traits we can support both number syntax and excel syntax
	let cell: Option<Cell<'_>> = sheet.cell((0,0));
	let cell: Option<Cell<'_>> = sheet.cell([0, 0]);
	let cell: Option<Cell<'_>> = sheet.cell("A1");
	
	// Returns an 'iterator' over all the cells in the sheet
	while let Some(cell) in sheet.cells() {}
	
	// but also has metadata like the count
	let count: usize = cells.count();
	// how many cells are used
	let used: usize = cells.used();
	// and empty
	let empty: usize = cells.empty();
	
	// `.rows()` returns an iterator, `Rows<'_>`, which yields a `Row<'_>`. 
	for row in sheet.rows() {
		// effectively a `.len()`
		let columns: usize = row.columns();
		
		// Gets the cell from the specified column of the row
		let cell: Option<Cell<'_>> = row.column("A");
		let cell: Option<Cell<'_>> = row.column(0);
	}
	
	// Get a subview into the sheet
	// Returns `None` if the range is out of bounds from the parent sheet
	let window: Option<Window<'_>> = sheet.window([9, 9], [14, 19]);
	let window: Option<Window<'_>> = sheet.window((9, 9),(14, 19));
	let window: Option<Window<'_>> = sheet.window("J10", "O20");
	
	// Would shares much the same with `Sheet<'_>`
	// with the main difference being that the indexing is relative to the new window.
	//
	// This gets the top left most cell of the window.
	let cell: Option<Cell<'_>> = window.cell(0, 0);
	
	let count: usize = window.count();
	let used: usize = window.used();
	let empty: usize = window.empty();
	
	// Creating another window from a window
	let another_window: Option<Window<'_>> = window.window("A1", "C4");
	
	// Get the rows over the window
	for another_row in another_window.rows() {}
}

// A `.columns()` function which returns an iterator `Columns<'_>` that returns a 
// `Column<'_>`
for column in sheet.columns() {
	// Get the some given rows data
	let cell: Option<Cell<'_>> = column.row("A");
	let cell: Option<Cell<'_>> = column.row(5);
}
```

#### Cell

```rust

let cell: Cell<'_> = sheet.cell("A1");

// Get the column and row of the cell
let column: u32 = cell.column();
let row: u32 = cell.column();
// Or maybe just a position
let (column, row) = cell.position();

// Returns the formula if there was one in the cell 
let formula: Option<&str> = cell.formula();

// Gets the solid fill color, defaulting to the theme color if no manual color was set
let color: &Color = cell.fill();
// Gets the hex representation
// 'FFFFFF'
let rgb: &str = color.rgb();
let rgbs: &str = color.rgba();
let argb: &str = color.argb();

let red: u8 = color.red();
let green: u8 = color.green();
let blue: u8 = color.blue();
let alpha: u8 = color.alpha();

let (red, green, blue, alpha) = color.raw();

// Gets info about the font styling
let font: &Font = cell.font();

let color: &Color = font.color();
// example: 'Arial'
let name: &str = font.name();
// example: 11
let size: usize = font.size();

// bool checks
font.is_bold();
font.is_italic();
font.is_underscore();
font.is_strikethrough();

let value: Option<Value<'_>> = cell.value();


// Given the expected usage of spreadsheets, the columns are most likley all going to have the same type bar a header.
// With this knowledge, the value syntax could be cleaned up so you then put the expected type you want, rather than having to match against all the types.
match value {
    None => println!("no value in cell"),
    Some(value) => match value {
    Value::Int(int) => println!("Int: {}", int),
    Value::Float(float) => println!("Float: {}", float),
    Value::String(string) => println!("String: {}", string),
    //...
    }
}
```

Here is what a slightly modified `Reader: Simple` example would look like with the new API:
```rust
use calamine::{Workbook, Xslx, XslxError};

fn main() -> Result<(), XlsxError> {
	// +-----------+--------------+---------------+---------+
	// |     A     |      B       |       C       |    D    |
	// +-----------+--------------+---------------+---------+
	// | franchise |   creator    |     value     | created |
	// +-----------+--------------+---------------+---------+
	// | Star Wars | George Lucas |   5749978736  |   1977  |
	// +-----------+--------------+---------------+---------+
	let workbook: Xslx<'_, '_>> = Workbook::open("file.xlsx")?;
	
	if let Some(worksheet): Option<&Worksheet<'_>> = workbook.worksheet("Sheet1") {
	    let rows = worksheet.rows();
		
		let header: Option<Row<'_>> = rows.next();
		
	    let Some(Value::String(franchise)) = header.column(0) else {
		    panic!("not a header row");
	    };
		
	    let Some(Value::String(value)) = header.column("C") else {
		    panic!("not a header row");
	    };
		
		println!("{franchise}, {value}");
		
	    for row in rows {
		    let franchise: Cell<'_> = .unwrap();
		    let valuation: Cell<'_> = .unwrap();
		    
		let Some(Value::String(franchise)) = row.column("A") else {
		    panic!("not a string");
	    };
		
	    let Some(Value::Int(value)) = row.column(2) else {
		    panic!("not an int");
	    };
		
		println!("column(\"A\")={:?}, column(2)={:?}", franchise.value(), valuation.value());
		// Output:
		//      franchise, value
		//      column("A")="Star Wars", column(2)=5749987736 
		}
	}
	
	Ok(())
}
```

## Write
#### Imports

```rust
use clamimne::{Workbook, Xlsx, XlsxError};
```

```rust
// Open existing workbook
let mut workbook: Result<Workbook<'_>> = Workbook::open("file.xslx");
// From scratch
let mut workbook_builder: Result<WorkbookBuilder> = Workbook::new();

let mut workbook: Workbook<'_, '_> = workbook_builder
									.author("I Wrote This")
									.theme()
									.build();

// Worksheets inheret the workbook theme if one is not set, so we can only
// have them be made through the workbook.
let new_sheet: Result<&mut WorksheetBuilder> = workbook.new_worksheet("Sheet1");

new_sheet.size("A1", "E100").;

// Get mutable worksheet
let mut worksheet: Option<&mut Worksheet> = workbook.worksheet_mut("Sheet1");

// Make a new worksheet using excel-range syntax
//
// Can fail if dimensions entered are invalid.
let mut worksheet: Result<Worksheet<'_>> = Worksheet::new("A1", "E100");
// Using number syntax
let mut worksheet: Result<Worksheet<'_>> = Worksheet::new((0, 0), (5, 99));
// or
let mut worksheet: Result<Worksheet<'_>> = Worksheet::new([0, 0], [5, 99]);
```
## Internals

### Workbook

```rust
  struct Xlsx<'path,'cell> {
	path: &'path Path,
	author: CompactString,
	worksheets: Vec<Worksheet<'cell>>,
  }
```

```rust
trait Workbook {
	type Workbook;
	type Error;

    /// Will read all sheets into memory. The empty value cells are kept in position.
    /// For opening with empy cells removed and the other cells compacted look to
    /// `open_sparse`.
	fn open<P: AsRef<Path>>(path: P) -> Result<Self::Workbook, Self::Error>;

    /// Will read all sheets into memory but the resulting sheets will be compacted
    /// with any empty value cells being removed from the structure. 
    fn open_sparse<P: AsRef<Path>>(path: P) -> Result<Self::Workbook, Self::Error>;

    /// Opens the workbook while only getting the metadata needed to form a skeleton
    /// for the rest. Only when used will sheets be read and formed. This will allow
    /// targeted sheet   
    fn open_lazy<P: AsRef<Path>>(path: P) -> Result<Self::Workbook, Self::Error>;

    /// Will read all worksheets into a file on disk.
    ///
    /// This is slower than in-memory options, but if the data cannot fit in memory
    ///  then this will still allow you to read and write values. 
    ///
    /// The file type used is not considered part of the API and changes to it can
    // happen at any time. 
    fn open_to_disk<P: AsRef<Path>>(path: P) -> Result<Self::Workbook, Self::Error>;

	fn new() -> Workbook;

    /// Returns all worksheets in the workbook.
	fn worksheets(&self) -> Option<Worksheets<'_>>;

    /// Returns all asked for sheets. If a sheet asked for does not exist, it will return `None`
    /// in the same position as it was passed in.  
    fn worksheets_by_name<S: AsRef<str>>(&self, sheets: &[S]) -> Option<Worksheets<'_>>;

    /// Gets a single worksheet. Will return `None` if the worksheet does not exist.
	fn worksheet<I: SheetIndex>(&self, sheet: I) -> Option<Worksheet<'_>>  {
        self.worksheets.get(sheet.into()?)
    }
	
	fn add_worksheet(worksheet: Worksheet);

    fn save(self);

    fn save_with_writer<W: Writer>(self, writer: W);
} 
```

```rust
  impl<'path, 'cell> Workbook for Xlsx<'path, 'cell> {

    type Workbook = Xlsx<'path, 'cell>;
    type Error = XlsxError;

    fn open<P: AsRef<Path>>(path: P) -> Result<Self::Workbook, Self::Error> {
	    todo!()
    }
	
	fn new() -> Self::Workbook {
		todo!()
	}
	  
    fn worksheets(&self) -> Option<Worksheets<'cell>> {
	    if self.sheets.is_empty() {
	        None
		} else {
	        &self.sheets
	    }
    }
  }
```

```rust
  trait ToSheetIndex {
    fn into(&self) -> Option<usize>;
  }
```

```rust
  impl ToSheetIndex for &str {

    /// # Performance
    ///
    /// Using a `&str` an `O(n)` operation. using a `usize` or `u32` is `O(1)`.
    fn into(self) -> Option<usize> {
     let mut idx = 0;
       for sheet in self.sheets {
        if sheet.name() == self {
          return Some(idx);
        }
	
        idx += 1; 
      }
	
      None
    }
  }

  impl ToSheetIndex for usize {
    fn into(self) -> Option<usize> {
      Some(self)
    }
  }
```

### Sheet

 Max sheet cells 17,179,869,184. 
  
```rust
pub struct Sheet<'value> {
  name: CompactString,
  path: CompactString,
  // (column, row)
  start: (u32, u32),
  end: (u32, u32),
  delta: u32, // end.0 - start.0
  cells: Vec<Cell<'value>>,
  default_cell: Cell<'value>,
  #[cfg(feature = "pictures")]
  pictures: Vec<Vec<u8>>,
}
```

```rust
impl Worksheet {
	fn resize<R: ToRangeIndex>(&mut self, R) {
		
	}

    fn columns_by_name<C: IntoColumnIndex>(&self, &[C]) -> Columns<'_>;
}
```

```rust
  pub struct Cells<'cell> {
    cells: &'cell [Cell<'cell'>],
  }

  impl<'cell> Cells<'cell> {
    pub fn count(&self) -> usize {
      self.cells.len()
    }

    pub fn used(&self) -> usize {
      self.cells.iter().reduce(|cell, used| u32::from(!cell.is_empty()) + used)
    }
    pub fn empty(&self) -> usize {
      self.cells.iter().reduce(|cell, empty| u32::from(cell.is_empty()) + empty)
    }
  }


  impl<'cell> Iterator for Cells<'cell> {
    type Item = Cell<'cell>;
    
    fn next(&mut self) -> Option<Self::Item> {
      self.cells.next()
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
      (0, Some(self.cells.len()))
    }
  }
```

```rust
impl<'cell> Sheet<'cell> {
  pub fn name(&self) -> &str {
    self.name.as_str()
  }

  pub fn path(&self) -> &str {
    self.name.as_str()
  }

  pub fn cell<P: ToCellPosition>(&self, pos: P) -> Option<&Cell> {
    self.cells.get(pos.into())
  }

  pub fn cells(&self) -> Cells<'cell> {
    Cells {
      
    }
  }

  /// # Returns
  ///
  /// Returns and iterator over the sheets rows.
  pub fn rows(&self) -> Rows<'cell> {
    Rows {
      cells: &self.cells,
      delta: self.delta,
      row: u32,
    }
  }

  
  /// # Returns
  ///
  /// Returns `Some(&Sheet<'_>)` if the given range is within the bounds of the current sheets dimensions, or `None` of it goes outside.
  pub fn window<R: ToCellRangeIndex>(&self, range: R) -> Option<&Window<'cell>> {

    // TODO: Implement ToCellRangeIndex for ((u32, u32), (u32, u32)), [[u32, u32],[u32, u32]], [u32, u32, u32, u32], (u32, u32, u32, u32) and &str

    let ((start_col, start_row), (end_col, end_row)) = range.into();

    // TODO: Calc a new delta.

    self.cells.get(self.delta )
  }
}
  
```

```rust
  trait CellIndex {
    fn into(self) -> (u32, u32);
  }
```

```rust
  impl CellIndex for (u32, u32) {
    #[inline]
    fn into(self) -> (u32, u32) {
      self
    }
  }
  
  impl CellIndex for &[u32, u32] {
    #[inline]
    fn into(self) -> (u32, u32) {
      [self.0, self.1]
    }
  }
  
impl CellIndex for (usize, usize) {
    #[inline]
    fn into(self) -> (u32, u32) {
      (self.0 as u32, self.1 as u32)
    }
  }
  
impl CellIndex for &str {
    /// # Panics
    ///
    /// Will panic if column format is not `a-z` or `A-Z`, or if there is no row number after the column letter/s.
    #[inline]
    fn into(self) -> (u32, u32) {
      excel_column_row_to_tuple(self)
    }
  }
```

```rust
  /// # Panics
  ///
  /// Will panic if column format is not `a-z` or `A-Z`, or if there is no row number after the column letter/s.
  #[inline]
  fn excel_column_row_to_tuple(pos: &str) -> (u32, u32) {
      let mut idx = 0;
      
      for c in pos.chars() {
        if c.is_ascii_alphabetic() {
          idx += 1;
        }
        break;
      }

      let column = excel_column_to_number(&pos[..=idx]);
      let row: u32 = &pos[idx + 1..].parse().unwrap();

      (column, row)
  }
```

```rust
  /// # Panics
  ///
  /// Will panic if the providec string contains a letter other than `a-z` or `A-Z`.
  #[inline]
  fn excel_column_to_number(column: &str) -> u32 {
    // PERF: Can hand roll a `to_upper(&mut self)` to prevent an allocation
    let column = column.to_uppercase();
    let mut result = 0;
    let mut multiplier = 1;

    for c in column.chars().rev() {
        if c.is_ascii_alphabetic() {
            let digit = c as u32 - 'A' as u32 + 1;
            result += digit * multiplier;
            multiplier *= 26;
        } else {
            // If the string contains non-alphabetic characters panic
            panic!("`{c}` is not a valid column letter must be `A-Z`")
        }
    }
    result
}
```

```rust
#[inline]
fn number_to_excel_column(n: u32) -> CompactString {
    let mut result = CompactString::default();
    let mut n = n;

    while n > 0 {
        let remainder = (n - 1) % 26;
        result.push((b'A' + remainder as u8) as char);
        n = (n - 1) / 26;
    }

    result.chars().rev().collect()
}```

```rust
  pub struct Window<'cell> {
    
    cells: &'cell [Cell<'cell>],
    start: (u32, u32),
    end: (u32, u32),
    delta: u32,
  }
```

```rust
  impl<'cell> Window<'cell> {
    // TODO: add Sheet functions to Window
  }
```

It could be possible to support both types of indexes with some kind of `IntoColumnIndex` trait.

```rust
  trait IntoColumnIndex {
     fn into(&self) -> u32 {
        // u32:
        *self
        // &str: Some kind match over a static map, or dynamically calc it
        ALPHABET_INDEX[self]
    }
  }
```

Storing the cells as a vec would permit great cache locality on both reads and writes.

Storing 2-dimensional data in a 1-dimensional vec requires a delta to then offset the index. In this case we will delta the length of each row.
Put another way we take the column amount, end.0, and index into a multiple of it to access the row we want.
To get the column of the row we want, we just use a standard additive offset.

Example:
```rust
pub struct Rows<'cell> {
  cells: &'row [Cell<'cell>],
  columns: u32,
  rows: u32,
  row: u32,
}
```

```rust
  impl<'cell> Iterator for Rows<'cell> {
    type Item = Row<'cell>
    
    fn next(&mut self) -> Option<Self::Item> {

      let start = self.row * self.columns;
      let end = start + self.columns;
          
      let row = Row { 
        row: self.cells.get(start..end)?,
        len: *self.columns,
        number: *self.row,
        };

        self.row += 1;

        Some(row)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
      (0, Some(self.rows)
    }
  }

  impl<'cell> ExactSizeIterator for Rows<'cell> {
    fn len(&self) -> usize {
      (self.rows - self.row) as usize
    }
  }
```

```rust
pub struct Row<'cell> {
  row: &'cell [Cell<'cell>],
  len: u32,
  number: u32,
}
```

```rust
impl<'cell> Row<'cell> {
  pub fn column<C: ToColumnIndex>(&self, column: C) -> &Cell {
    
  }

  pub fn len(&self) -> usize {
    *self.len
  }
}  
```

```rust
pub fn next(&mut self) -> Option<&[Cell<'_>]> {
    let row_start = self.curr * self.delta;
    let row_end = row_start + self.delta;
  
    self.curr += 1;
    
    self.get(row_start..row_end)
}
```

### Cell

```rust
  pub struct Cell<'value> {
    pos: (u32, u32),
    val: Option<Value<'value>>,
    formula: Option<&'value str>,
    font: Font,
    fill: Color,
    format: CompactString, 
  }

  impl<'string> Cell<'value> {
    pub fn value(&self) -> Option<Value<'value>>
  }
```

### Value

```rust
// Leaves open the ability to specialize with future updates
// like a Python() for example, extracting the python code from the formula.
#[non_exhaustive]
pub enum Value<'value> {
  /// Signed integer
  Int(i64),
  /// Float
  Float(f64),
  /// String
  Text(&'value str),
  /// Boolean
  Bool(bool),
  /// Date or Time
  DateTime(f64),
  /// Duration
  Duration(f64),
  /// Date, Time or DateTime in ISO 8601
  DateTimeIso(&'value str),
  /// Duration in ISO 8601
  DurationIso(&'value str),
  /// Error
  Error(CellErrorType),
  // Hyperlink
  Hyperlink(&'value str),
}
```

Passing around the lifetime for strings means that we can choose what ever backing memory we want, for example `compact_str::CompactString` or `smartstring::String`.

We get the benefit of SSO while also keeping the exposed API to a standard `&str`. 

If we directly stored those types in the Value enum then we would be at the mercy of relying on that types changes as well as compatibility issues when users use the value.
  
```rust
pub enum CellErrorType {
  /// Division by 0 error
  Div0,
  /// Unavailable value error
  NA,
  /// Invalid name error
  Name,
  /// Null value error
  Null,
  /// Number error
  Num,
  /// Invalid cell reference error
  Ref,
  /// Value error
  Value,
  /// Getting data
  GettingData,
}
```

Another option is to pass around a `Box<dyn Value>` and have types that represent the cells. This way the types can impliment their own `.value()` method.
```rust
trait CellValue {
    type Value;
    
    fn value(&self) -> Option<Self::Value>;
}

Value(Box<dyn CellValue>);

Int(i64);
Text(CompactStr);
Float(f64);

impl CellValue for Int {
    type Value = i64;
    fn value(&self) -> Option<Self::Value> {
        self.0
    }
}

// Cells would then look like
struct Cell {
    value: Value,
}

// and usage
let value: Option<i64> = cell.value();

```

### Font

```rust
 pub struct Font {
  name: String,
  size: u32,
  color: Color,
  is_bold: bool,
  is_italic: bool,
  is_underscore: bool,
  is_strikethrough: bool,
}
```

```rust
  let name: &str = font.name();
  font.size();
  let color: Color = font.color();
  font.is_bold();
```

### Color

```rust
  // ARGB representation
  pub struct Color {
  red: u8,
  green: u8,
  blue: u8,
  alpha: u8,
  }
```

```rust
  // "FF00FF00"
  let rgb: &str = color.rgb();
  let rgba: &str = color.rgba();
  let argb: &str = color.argb();
```
