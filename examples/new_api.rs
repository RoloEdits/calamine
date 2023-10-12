/*

Api design proposal



*/

fn main() {
    let workbook: Xlsx<'_> = Workbook::open("tests/simple.xlsx");

    // returns a slice of `Sheet`
    let sheets: &[Sheet] = workbook.sheets();

    // returns an Option, Some if the sheet exists and None if not.
    let sheet: Sheet = workbook.sheet("Sheet1").unwrap();

    println!("Sheet Name: {}", sheet.name());

    let cell: Cell = sheet.cell("A1");
    let cell: Cell = sheet.cell(0, 0);

    // Range is non-inclusive
    // Might also try something like `sheet.cell_range("A1:B100")`
    let cells: &[Cell] = sheet.cell_range("A1:B100");
    let cells: &[Cell] = sheet.cell_range((0, 0), (1, 100));

    // Will return the underlying slice of all the cells from left to right.
    let cells: &[Cell] = sheet.cells();

    // `.rows()` returns an iterator, `Rows<'_>`, over the rows of the current sheet.
    // row is a `Row` thats wraps a `&[Cell]`
    for row in sheet.rows() {
        // `.column()` returns a `Option<Cell>` and will default to returning a `None`, even if picking outside the actual sheets range.
        // `is_empty()` returns a true if the cells `DataType` is `Empty`.
        // TODO: Need to figure out how a color change, but no value works is represented in the xml.
        assert!(!row.column("A").is_empty());
        // `.column()` would support both numbers and letters
        let cell: Option<Cell> = row.column(0);

        // `.value` returns an Option<T> where T is the data type of the cell.
        println!("Value: {}", cell.value());

        // `.formula()` returns an Option<&str>
        let formula: Option<&str> = row.column("B").formula();

        // `.text()` returns a `&Text` that has metadata about the font, font color, styles, etc.
        // this defaults to the theme colors if there is none set
        let font: &Font = row.column("A").font();

        // returns a `Color` with methods like `.argb()`, `.rgb()`, '.hsl()' that return a `&str` of the values in hex form.
        let color: &str = font.color().rgb();
        // returns `&str` of the name of the font.
        let font: &str = font.name();

        // returns a Color
        let fill: &str = row.column("A").fill().rgb();

        // row: Row also implments `Iterator` that yields the cells in the columns of the row
        for column in row {}
    }

    // returns a `Column` that is an iterator over the `Cell`s of that column.
    //for row in sheet.column(0) {}
    for cell in sheet.column("A") {
        let value: Option<Value<'_>> = cell.value();
    }
}

pub struct Cell<'value> {
    pos: (u32, u32),
    // There could be some discussion on whether to have `empty` cells return the default values of a cell in the workbook
    // and provide an `Empty` variant in the `Data` enum or if it should return `None`.
    // Part of this debate will also be decided if the font and the color values, for example, matter outside of the value itself.
    val: Option<Value<'value>>,
    formula: Option<String>,
    font: Font,
    fill: Color,
}

// Currently getting a value out of an enum needs to put all the burden on the caller, but this may be desired behavior
// as types must be known to the user.
// Another option is maybe use dynamic dispatch and type erasure.
impl<'value> Cell<'value> {
    // (column, row)
    pub fn position(&self) -> (u32, u32) {
        todo!()
    }

    pub fn row(&self) -> u32 {
        todo!()
    }

    pub fn column(&self) -> u32 {
        todo!()
    }

    pub fn font(&self) -> &Font {
        todo!()
    }

    pub fn fill(&self) -> &Color {
        todo!()
    }

    pub fn value(&self) -> Option<Value<'value>> {
        self.val
    }
}

struct Font {
    name: String,
    size: u8,
    color: Color,
    is_italic: bool,
    is_bold: bool,
    is_underscore: bool,
    is_strikethrough: bool,
}

/// Represents an ARGB layout.
pub struct Color(u8, u8, u8, u8);

pub struct Sheet<'cell> {
    name: String,
    // Can make a &Path instead
    path: String,
    // (column, row)
    start: (u32, u32),
    end: (u32, u32),
    // The size of the max row length, the end column value end.0, will act as an offset inside the Vec when indexing; The stride.
    // Every new row will start at use the (stride * row) to offset to the correct row, returning the slice `&[Cell]`.
    // This same slice is used for the `.rows()` `Rows<'_>` iterator
    cells: Vec<Cell<'cell>>,
}

struct Row<'row> {
    row: &'row [Cell<'row>],
    stride: u32,
    number: u32,
}

struct Column<'row> {
    cell: &'row Cell<'row>,
    stride: u32,
    row: u32,
}

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
