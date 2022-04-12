use calamine::{open_workbook, Reader, Xlsx};
extern crate simple_excel_writer;
use excel::*;
use simple_excel_writer as excel;

fn main() {
    let mut wb = Workbook::create("newfile.xlsx");
    let mut sheet = wb.create_sheet("filtered");

    let mut excel: Xlsx<_> = open_workbook("file.xlsx").unwrap();
    if let Some(Ok(r)) = excel.worksheet_range("Sheet1") {
        wb.write_sheet(&mut sheet, |sheet_writer| -> Result<(), std::io::Error> {
            let sw = sheet_writer;
            Ok(for row in r.rows() {
                if &row[1] == "Changes done successfully".trim() {
                    let _append_row = sw.append_row(row![row[0].to_string()]);
                }
            })
        })
        .expect("write excel error!");
    }
}
