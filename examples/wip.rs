use calamine::wip::{Workbook, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = r#"NYC_311_SR_2010-2020-sample-1M.xlsx"#;

    let mut workbook = Xlsx::open(path)?;

    let worksheet = workbook.worksheet("sheet1")?.unwrap();

    for row in worksheet.rows() {}

    Ok(())
}
