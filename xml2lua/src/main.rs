extern crate calamine;

use calamine::{open_workbook, Reader, Xlsx, Xls};

fn main() {
    let path = format!("./ee.xlsx");
    let mut excel:Xlsx<_> = open_workbook(&path).unwrap();
    // let sheets:Vec<String> = excel.sheet_names().to_owned();
    // for s in sheets{
    //     let mut cells_reader = excel.worksheet_cells_reader(&s).unwrap();
    // }
    let sheets = excel.worksheets();
    for sheet in sheets{
        let sheet_data = sheet.1;
        println!("table: {}", sheet.0);
        let mut row = 0;
        for temp_row in sheet_data.rows(){
            let mut col = 0;
            for cell in temp_row{
                println!("{}", cell.to_string());
                match excel.getcolor(row,col) {
                    Some(number) => println!("Calling Daniel: {}", number),
                    _ => println!("Don't have Daniel's number."),
                }
                col = col+1;
            }
            row = row +1;
        }
    }
    // let (_sheetname, firstrange) = sheets.first().expect("bad format");
    // let mut init_flag = false;
    // for temp_row in firstrange.rows(){
    //     if !init_flag {
    //         init_flag = true;
    //         for cell in temp_row{
    //             println!("first line {}", cell.to_string());
    //         }
    //     }else{
    //         for cell in temp_row{
    //             println!("{}", cell.to_string());
    //         }
    //     }
    // }
}
