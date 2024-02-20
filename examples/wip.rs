use calamine::wip::{Workbook, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = "NYC_311_SR_2010-2020-sample-1M.xlsx";

    let mut workbook = Xlsx::open(path)?;

    let worksheet = workbook
        .worksheet("NYC_311_SR_2010-2020-sample-1M")
        .unwrap();

    for _row in worksheet.rows() {}

    Ok(())
}
