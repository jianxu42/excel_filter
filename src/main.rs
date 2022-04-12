use calamine::{open_workbook, Reader, Xlsx};
extern crate simple_excel_writer;
use clap::{Arg, Command};
use excel::*;
use simple_excel_writer as excel;

fn main() {
    // println!("{:?}", std::env::args());
    let matches = Command::new("excel_filter")
        .version("0.1.0")
        .author("None")
        .about("Excel File Filter")
        .arg(
            Arg::new("source_path")
                .short('s')
                .value_name("Source Path")
                .help("Input source path.")
                .required(true)
                .min_values(1),
        )
        .arg(
            Arg::new("source_sheet_name")
                .short('k')
                .value_name("Source Sheet Name")
                .help("Input source sheet name.")
                .required(true)
                .min_values(1),
        )
        .arg(
            Arg::new("target_path")
                .short('t')
                .value_name("Target Path")
                .help("Input target path.")
                .required(true)
                .min_values(1),
        )
        .arg(
            Arg::new("target_sheet_name")
                .short('l')
                .value_name("Target Sheet Name")
                .help("Input target sheet name.")
                .required(true)
                .min_values(1),
        )
        .arg(
            Arg::new("matching_column")
                .short('c')
                .value_name("Matching Column")
                .help("Input matching column, starting from 0.")
                .required(true)
                .min_values(1),
        )
        .arg(
            Arg::new("matching_string")
                .short('m')
                .value_name("Matching String")
                .help("Input matching string.")
                .required(true)
                .min_values(1),
        )
        .get_matches();
    let source_path = matches
        .values_of("source_path")
        .unwrap()
        .collect::<Vec<&str>>();
    let source_sheet_name = matches
        .values_of("source_sheet_name")
        .unwrap()
        .collect::<Vec<&str>>();
    let target_path = matches
        .values_of("target_path")
        .unwrap()
        .collect::<Vec<&str>>();
    let target_sheet_name = matches
        .values_of("target_sheet_name")
        .unwrap()
        .collect::<Vec<&str>>();
    let matching_column = matches
        .values_of("matching_column")
        .unwrap()
        .collect::<Vec<&str>>();
    let matching_string = matches
        .values_of("matching_string")
        .unwrap()
        .collect::<Vec<&str>>();

    let mut wb = Workbook::create(target_path.get(0).unwrap());
    let mut sheet = wb.create_sheet(target_sheet_name.get(0).unwrap());

    let mut excel: Xlsx<_> = open_workbook(source_path.get(0).unwrap()).unwrap();
    if let Some(Ok(r)) = excel.worksheet_range(source_sheet_name.get(0).unwrap()) {
        wb.write_sheet(&mut sheet, |sheet_writer| -> Result<(), std::io::Error> {
            let sw = sheet_writer;
            for row in r.rows() {
                if &row[matching_column.get(0).unwrap().parse::<usize>().unwrap()]
                    == //"Changes done successfully"
                matching_string.get(0).unwrap().trim()
                {
                    let _append_row = sw.append_row(row![row[0].to_string()]);
                }
            }
            Ok(())
        })
        .expect("write excel error!");
    }
}
