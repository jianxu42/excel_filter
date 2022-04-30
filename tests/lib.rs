use assert_cmd::Command;
type TestResult = Result<(), Box<dyn std::error::Error>>;

#[test]
fn runs() -> TestResult {
    let mut cmd = Command::cargo_bin("excel_filter")?;
    cmd.arg("-s ..\\..\\tests\\inputs\\file.xlsx")
        .arg("-k Sheet1")
        .arg("-t ..\\..\\tests\\newfile.xlsx")
        .arg("-c 1")
        .arg("-p 0")
        .arg("-m Done")
        .arg("-l Sheet1");
    cmd.assert().success();
    Ok(())
}