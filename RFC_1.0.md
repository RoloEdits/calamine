# Changes for v1.0.0

## Summary

## TODO

- [ ] Decide if library will stay read-only

## API

The main focus of the proposed changes is to:
- Be more hierarchical in data access. Having a clear ownership of Workbook -> Sheet -> Cell access pattern. The data should belong to the logical scope. Even if potentially out of spec.
  - Example: A Workbook has a theme, but the colors of a cell can be changed outside of that. The color value should then come through the Cell.
  - Example: Pictures are related to a sheet. Even if the spec has a `media` path directly under the Workbook, the pictures should be accessed through the Sheet.
- Focus on an ergonomic API, with lower exposure of internals. Contrasting the current API to others from other languages it can be seen that there is some rough edges. There are some areas that can be polished up during the refactor. Just doing the above point would help here.
  - Example: When looking into implimenting https://github.com/tafia/calamine/issues/362, it would require a pretty direct breaking change as the `String` type is directly exposed when getting the value out. Instead, exposing a `&str` would allow the backing memory to be an implimentation detail.
  - Example: The current `fn worksheets(&mut self) -> Vec<(String, Range<DataType>)>;` returns paht information and the size of the sheet. These should just be in a Sheet type with fields for `name` and `dimensions`, accessed through the thing that should logicaly own it; `sheet.name()` and `sheet.dimensions()`.
- And using the opportunity in the resructure to squeeze more performance.
  - Example: In `xlsx` files, as an example, there is a shared string component that is linked to by an index in the corrosponding sheet/s. In a worse case scenario the sheet have nothing but strings and each string is unique, with no sharing occuring. This means that teh shared string file would be at least as long as any given sheet. The parsing of these can files can be done in parallel. The backing string `Vec<T>` can be of an SSO string type, like [`compact_str`](https://github.com/ParkMyCar/compact_str), and only need to share a refernce of `&str`, to the `DataType` enum. The string Vec acting as an arena. The `Cell<'value>` then only need to store a `&str` with the index into the string arena maping very well to the indexing of the `xlsx` spec. Changing to Vec where possible should also improve general performance as look-ups are `O(1)` and they are more cache friendly.
  - Example: Parallel iteration using [`rayon`](https://github.com/rayon-rs/rayon). This could be an easy performance boost as the backing `Vec`s can be very friendly to this kind of work. The most obvious use case is when iterating over rows. If the hierarchal change above is stuck to, then all context of where a cell is, is stored within the `Cell` itself, and thus even with out of order iteration, you still have the context of the row the cell is in.
  
### Imports

The first part of the proposed changes is to simplify the direct imports.

```rust
use calamine::{Workbook, Xlsx};
```


```rust
let workbook: Result<Xlsx<'_>> = Workbook::open("tests/simple.xlsx");
```

This would leave the reading implimentation scoped to the workbook type, solving this issue of the present API:
```rust  
// FIXME `Reader` must only be seek `Seek` for `Xls::xls`. Because of the present API this limits
// the kinds of readers (other) data in formats can be read from.
```

The `open_workbook_auto()` would go away completely in the new API. Matching a file to opening it should be left to the library user prioritizing explictness to help enforce correctness.


Here is what the `Reader: Simple` example would look like with the new API:

```rust
  use calamine::{Workbook, Xslx, XslxError};

  let workbook: Result<Xslx<'_>, XlsxError> = Workbook::open("file.xlsx");

  let worksheet: Option<Worksheet> = workbook?.worksheet("Sheet1");

  if let Some(worksheet): Worksheet = worksheet {
    // Row<'_>           Rows<'_> 
    for row in worksheet.rows() {
      //                                        Cell<'_>
      println!("row={:?}, column(0)={:?}", row, row.column("A"));
    }
  }
  
  
```



```rust
// Gets a specified sheet from the workbook
let sheet: Option<Sheet<'_>> = workbook.worksheet("Sheet1");
let sheet: Option<Sheet<'_>> = workbook.worksheet(0);   

// Gets all sheets in the workbook
let sheets: &[Sheet<'_>] = workbook.worksheets();
```
### Sheet

```rust
let name: &str = sheet.name();
let path: &str = sheet.path();
let (column, row): (u32, u32) = sheet.start();
let (column, row): (u32, u32) = sheet.end();
let (columns, rows): (u32, 32) = sheet.dimensions();  
```

```rust
// returns an iterator, `Rows<'_>`, over the rows, `Row<'_>`, which wraps a `&[Cell<'_>]`.
for row in sheet.rows() {}
```

```rust
// returns an iterator `Column<'_>` that yeilds a `&Cell`
  for cell in sheet.column("A") {}
```
or
```rust
  for cell in sheet.column(0) {}
```

It could be possible to support both types of indexes with some kind of `IntoColumnIndex` trait.

```rust
  static ALPHABET_INDEX: HashMap<&str, u32> = hashmap!["A": 0, "B": 2];

  trait IntoColumnIndex {
     fn into(&self) -> u32 {
        // u32:
        *self
        // &str: Some kind match over a static map, or dynamicly calc it
        ALPHABET_INDEX[self]
    }
  }
```

Getting a cell from a co-ordinate:
```rust
let cell: Cell = sheet.cell(0, 0);
```

```rust
  // &[Cell]
  for cell in sheet.cells() {}
```

A subsection of a square range
```rust
// could also call `window`
// Sheet<'_> or CellWindow<'_>
  for cell in sheet.cell_range("A1:B100") {}
  // &[Cell]
  for cell in sheet.window((0, 0), (2, 100)) {}
```

```rust
  pub struct Sheet<'sheet> {
    cells: &'sheet [Cell<'sheet>],
    start: (u32, u32),
    end: (u32, u32),
  }
```

### Cell 

```rust
match cell.value() {
    None => println!("empty cell"),
    Some(value) => match value {
    // i64
    Value::Int(int) => println!("Int: {}", int),
    // f64
    Value::Float(float) => println!("Float: {}", float),
    // &str
    Value::String(str) => println!("String: {}", str),
    //...
    }
}
```

```rust
let fill: Color = cell.fill();
let (column, row): (u32, u32) = cell.position(); 
let row: u32 = cell.row();
let column: u32 = cell.column();

let formula: Option<&str> = cell.formula();

let font: &Font = cell.font();
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
  pub struct Color(u8, u8, u8, u8)
```

```rust
  // "FF00FF00"
  let rgb: &str = color.rgb();
  let argb: &str = color.argb();
```

## Internals

### Workbook

```rust
  struct Xlsx<'cell> {
    sheets: Cell<'cell>,
  }
```
//  TODO: GATs

```rust
trait Workbook {
  type Workbook<'cell> where Self: 'cell;
  type Error;

  fn open() -> Result<Self::Workbook<'cell>, Self::Error>;

  fn worksheets(&self) -> Option<&[Sheet]>;

  fn worksheet<I: SheetIndex>(&self, sheet: T) -> Option<&Sheet>;
} 
```

```rust
  impl<'cell> Workbook for Xlsx<'cell> {

    type Workbook<'cell> = Xlsx<'cell>;
    type Error = XlsxError;
  
    fn open() -> Result<Self::Workbook<'cell>, Self::Error> {
      todo!()
    }
  
    fn worksheets(&self) -> Option<&[Sheet]> {
      if self.sheets.is_empty() {
        None
      } else {
        &self.sheets
      }
    }

    fn worksheet<I: SheetIndex>(&self, sheet: I) -> Option<&Sheet> {
        self.sheets.get(sheet.into()?)
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
  delta: u32,
  cells: Vec<Cell<'value>>,
  default_cell: Cell<'value>,
  #[cfg(feature = "pictures")]
  pictures: Vec<Vec<u8>>,
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

    // TODO: Impliment ToCellRangeIndex for ((u32, u32), (u32, u32)), [[u32, u32],[u32, u32]], [u32, u32, u32, u32], (u32, u32, u32, u32) and &str

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
fn number_to_excel_column(n: u32) -> String {
    let mut result = String::new();
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


Storing the cells as a vec would permit great cache locality on both reads and writes.

Storing 2-dimensional data in a 1-dimensional vec requires a delta to then offset the index. In this case we will delta the lengh of each row.
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
  }

  impl<'string> Cell<'value> {
    pub fn value(&self) -> Option<Value<'value>>
  }
```

### Value

```rust
pub enum Value<'value> {
  /// Signed integer
  Int(i64),
  /// Float
  Float(f64),
  /// String
  String(&'value str),
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

We get the benifit of SSO while also keeping the exposed API to a standard `&str`. 

If we directly stored those types in the Value enum then we would be at the mercy of relying on that types changes as well as compatablity issues when users use the value.
  
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
