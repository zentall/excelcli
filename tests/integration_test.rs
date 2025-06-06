use assert_cmd::Command;
use std::fs;
use std::path::Path;

fn setup_test_xlsx() {
    // 既にファイルがあれば何もしない
    if Path::new("tests/data/test.xlsx").exists() {
        return;
    }
    // サンプルのExcelファイルを生成
    fs::create_dir_all("tests/data").unwrap();
    fs::write("tests/data/test.xlsx", b"").unwrap();
}

#[test]
fn test_extract_col_with_filename() {
    setup_test_xlsx();

    let output_csv = "tests/data/outputs/test_extract_col_with_filename.csv";
    let _ = fs::remove_file(output_csv);

    let mut cmd = Command::cargo_bin("excelcli").unwrap();
    cmd.arg("extract-col")
        .arg("tests/data/test.xlsx")
        .arg("--col-range").arg("A:C")
        .arg("--rows").arg("1")
        .arg("--headers").arg("A列,B列,C列")
        .arg("--output").arg(output_csv)
        .arg("--with-filename");

    let assert = cmd.assert();
    assert.success();

    // 出力ファイルができているか
    assert!(Path::new(output_csv).exists());
    // 内容の検証（ここではファイル名ヘッダーが含まれているかのみ確認）
    let content = fs::read_to_string(output_csv).unwrap();
    assert!(content.contains("ファイル名,A列,B列,C列"));
}

#[test]
fn test_extract_row_with_filename() {
    setup_test_xlsx();

    let output_csv = "tests/data/outputs/test_extract_row_with_filename.csv";
    let _ = fs::remove_file(output_csv);

    let mut cmd = Command::cargo_bin("excelcli").unwrap();
    cmd.arg("extract-row")
        .arg("tests/data/test.xlsx")
        .arg("--range").arg("1:1")
        .arg("--columns").arg("A,B,C")
        .arg("--headers").arg("A列,B列,C列")
        .arg("--output").arg(output_csv)
        .arg("--with-filename");

    let assert = cmd.assert();
    assert.success();

    // 出力ファイルができているか
    assert!(Path::new(output_csv).exists());
    let content = fs::read_to_string(output_csv).unwrap();
    assert!(content.contains("ファイル名,A列,B列,C列"));
}