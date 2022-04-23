use assert_cmd::Command;
type TestResult = Result<(), Box<dyn std::error::Error>>;

#[test]
fn runs() -> TestResult {
    let mut cmd = Command::cargo_bin("excel_filter")?;
    cmd.arg("-s ..\\..\\file.xlsx")
        .arg("-k Sheet1")
        .arg("-t ..\\..\\newfile.xlsx")
        .arg("-c 1")
        .arg("-p 0")
        .arg("-m \"Changes done successfully\"")
        .arg("-l filtered");
    cmd.assert().success().stdout("");
    Ok(())
}
