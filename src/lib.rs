use std::error::Error;

use calamine::{open_workbook, Reader, Xlsx};
extern crate simple_excel_writer;
use clap::{Arg, Command};
use excel::*;
use simple_excel_writer as excel;

type MyResult<T> = Result<T, Box<dyn Error>>;

pub fn run(config: Config) -> MyResult<()> {
    let mut wb = Workbook::create(config.target_path.get(0).unwrap());
    let mut sheet = wb.create_sheet(config.target_sheet_name.get(0).unwrap());
    let mut excel: Xlsx<_> = open_workbook(config.source_path.get(0).unwrap()).unwrap();
    if let Some(Ok(r)) = excel.worksheet_range(config.source_sheet_name.get(0).unwrap()) {
        wb.write_sheet(&mut sheet, |sheet_writer| -> Result<(), std::io::Error> {
            let sw = sheet_writer;
            for row in r.rows() {
                if &row[config
                    .matching_column
                    .get(0)
                    .unwrap()
                    .parse::<usize>()
                    .unwrap()]
                    == config.matching_string.get(0).unwrap().trim()
                {
                    let append_row = sw.append_row(row![row
                        [config.copy_column.get(0).unwrap().parse::<usize>().unwrap()]
                    .to_string()
                    .trim_start_matches('0')]);
                    match append_row {
                        Ok(_) => (),
                        Err(err) => eprintln!("{:?}", err),
                    };
                }
            }
            Ok(())
        })
        .expect("write excel error!");
        wb.close().expect("close excel error!");
        println!("write excel success!");
    }
    Ok(())
}

#[derive(Debug, Clone)]
pub struct Config {
    pub source_path: Vec<String>,
    pub source_sheet_name: Vec<String>,
    pub target_path: Vec<String>,
    pub target_sheet_name: Vec<String>,
    pub matching_column: Vec<String>,
    pub copy_column: Vec<String>,
    pub matching_string: Vec<String>,
}

pub fn get_args() -> MyResult<Config> {
    let matches = Command::new("rust excel filter")
        .version("0.1.0")
        .author("None")
        .about("Excel File Filter")
        .arg(
            Arg::new("source_path")
                .allow_invalid_utf8(true)
                .short('s')
                .value_name("Source Path")
                .help("Input source path.")
                .required(true)
                .min_values(1),
        )
        .arg(
            Arg::new("source_sheet_name")
                .allow_invalid_utf8(true)
                .short('k')
                .value_name("Source Sheet Name")
                .help("Input source sheet name.")
                .required(true)
                .min_values(1),
        )
        .arg(
            Arg::new("target_path")
                .allow_invalid_utf8(true)
                .short('t')
                .value_name("Target Path")
                .help("Input target path.")
                .required(true)
                .min_values(1),
        )
        .arg(
            Arg::new("target_sheet_name")
                .allow_invalid_utf8(true)
                .short('l')
                .value_name("Target Sheet Name")
                .help("Input target sheet name.")
                .required(true)
                .min_values(1),
        )
        .arg(
            Arg::new("matching_column")
                .allow_invalid_utf8(true)
                .short('c')
                .value_name("Matching Column")
                .help("Input matching column, starting from 0.")
                .required(true)
                .min_values(1),
        )
        .arg(
            Arg::new("copy_column")
                .allow_invalid_utf8(true)
                .short('p')
                .value_name("Copy Column")
                .help("Input copy column, starting from 0.")
                .required(true)
                .min_values(1),
        )
        .arg(
            Arg::new("matching_string")
                .allow_invalid_utf8(true)
                .short('m')
                .value_name("Matching String")
                .help("Input matching string.")
                .required(true)
                .min_values(1),
        )
        .get_matches();
    let source_path = matches.values_of_lossy("source_path").unwrap();
    let source_sheet_name = matches.values_of_lossy("source_sheet_name").unwrap();
    let target_path = matches.values_of_lossy("target_path").unwrap();
    let target_sheet_name = matches.values_of_lossy("target_sheet_name").unwrap();
    let matching_column = matches.values_of_lossy("matching_column").unwrap();
    let copy_column = matches.values_of_lossy("copy_column").unwrap();
    let matching_string = matches.values_of_lossy("matching_string").unwrap();
    Ok(Config {
        source_path,
        source_sheet_name,
        target_path,
        target_sheet_name,
        matching_column,
        copy_column,
        matching_string,
    })
}
