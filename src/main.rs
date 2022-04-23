fn main() {
    if let Err(e) = excel_filter::get_args().and_then(excel_filter::run) {
        eprintln!("{}", e);
        std::process::exit(1);
    }
}
