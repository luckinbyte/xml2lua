extern crate calamine;

use calamine::{open_workbook, Reader, Xlsx};

fn main() {
    let path = format!("./ee.xlsx");
    let mut excel:Xlsx<_> = open_workbook(&path).unwrap();
    let sheets:Vec<String> = excel.sheet_names().to_owned();
    for s in sheets{
        let mut cells_reader = excel.worksheet_cells_reader(&s).unwrap();
    }
}
