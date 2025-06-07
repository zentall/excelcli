use clap::{Parser, Subcommand};
use calamine::{Reader, open_workbook_auto, DataType};
use glob::glob;
use csv::Writer;
// use std::path::PathBuf;
use std::error::Error;

#[derive(Parser)]
#[command(name = "excelcli")]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    ExtractRow {
        #[arg(help = "ファイルパターン（例: ./reports/*.xlsx）")]
        file_pattern: String,

        #[arg(long, default_value = "Sheet1")]
        sheet: String,

        #[arg(long, help = "行範囲（例: 5:20）")]
        range: String,

        #[arg(long, help = "列リスト（例: B,D,E）")]
        columns: String,

        #[arg(long, help = "CSVヘッダー（例: 名前,売上,部署）")]
        headers: String,

        #[arg(long, help = "フィルター対象列（nullじゃない行のみ対象）", default_value = "")]
        filter_col: String,

        #[arg(long, help = "出力CSVファイル")]
        output: String,

        #[arg(long, help = "CSVにファイル名を含める")]
        with_filename: bool,
    },

    ExtractCol {
        #[arg(help = "ファイルパターン（例: ./reports/*.xlsx）")]
        file_pattern: String,

        #[arg(long, default_value = "Sheet1")]
        sheet: String,

        #[arg(long, help = "列範囲（例: B:E）")]
        col_range: String,

        #[arg(long, help = "行番号リスト（例: 3,5,7）")]
        rows: String,

        #[arg(long, help = "CSVヘッダー（例: 東京,大阪,名古屋）")]
        headers: String,

        #[arg(long, help = "フィルター対象行番号（nullじゃない列のみ対象）", default_value = "")]
        filter_row: String,

        #[arg(long, help = "出力CSVファイル")]
        output: String,

        #[arg(long, help = "CSVにファイル名を含める")]
        with_filename: bool,
    },
}

// A1形式の列名を0-based列番号に変換（例: A→0, B→1）
fn col_str_to_index(col: &str) -> usize {
    let mut index = 0;
    for (i, c) in col.chars().rev().enumerate() {
        let val = (c.to_ascii_uppercase() as u8 - b'A' + 1) as usize;
        index += val * 26_usize.pow(i as u32);
    }
    index - 1
}


/// ファイルパターンからファイルリストを取得し、件数を表示
fn collect_files(pattern: &str) -> Result<Vec<std::path::PathBuf>, Box<dyn Error>> {
    let entries: Vec<_> = glob(pattern)?.collect::<Result<_, _>>()?;
    Ok(entries)
}

/// ワークブックとシートデータを取得
fn open_sheet<'a>(
    path: &std::path::Path,
    sheet_name: &str,
) -> Result<calamine::Range<DataType>, Box<dyn Error>> {
    let mut workbook = open_workbook_auto(path)?;
    let range = workbook
        .worksheet_range(sheet_name)
        .ok_or("指定したシートが存在しません")??;
    Ok(range)
}

/// CSVヘッダーを書き込む
fn write_csv_header(wtr: &mut Writer<std::fs::File>, headers: &str, with_filename: bool) -> Result<(), Box<dyn Error>> {
    let mut headers_vec: Vec<&str> = headers.split(',').collect();
    if with_filename {
        headers_vec.insert(0, "source_file");
    }
    wtr.write_record(headers_vec)?;
    Ok(())
}

// 列範囲・行リストで抽出
fn extract_col(
    file_pattern: &str,
    sheet_name: &str,
    col_range: &str,
    rows: &str,
    headers: &str,
    filter_row: &str,
    output: &str,
    with_filename: bool,
) -> Result<(), Box<dyn Error>> {
    let mut wtr = Writer::from_path(output)?;
    write_csv_header(&mut wtr, headers, with_filename)?;

    let parts: Vec<&str> = col_range.split(':').collect();
    let col_start = col_str_to_index(parts[0]);
    let col_end = col_str_to_index(parts[1]);
    let row_nums: Vec<usize> = rows.split(',').map(|r| r.parse::<usize>().unwrap()).collect();

    let entries = collect_files(file_pattern)?;
    println!("対象ファイル数: {}", entries.len());

    for path in entries {
        let range_data = open_sheet(&path, sheet_name)?;

        for &row_idx in &row_nums {
            let mut row_vals = Vec::new();
            for col_idx in col_start..=col_end {
                let val = range_data.get((row_idx - 1, col_idx)).unwrap_or(&DataType::Empty);
                row_vals.push(match val {
                    DataType::String(s) => s.to_string(),
                    DataType::Float(f) => f.to_string(),
                    DataType::Int(i) => i.to_string(),
                    DataType::Bool(b) => b.to_string(),
                    _ => "".to_string(),
                });
            }

            // フィルター
            let should_include = if filter_row.is_empty() {
                true
            } else if let Ok(filter_row_num) = filter_row.parse::<usize>() {
                row_idx == filter_row_num && row_vals.iter().any(|v| !v.is_empty())
            } else {
                true
            };
            if !should_include {
                continue;
            }

            if with_filename {
                let fname = path.file_name().and_then(|s| s.to_str()).unwrap_or("");
                let mut record = vec![fname.to_string()];
                record.extend(row_vals);
                wtr.write_record(record.iter().map(|s| s.as_str()))?;
            } else {
                wtr.write_record(row_vals.iter().map(|s| s.as_str()))?;
            }
        }
    }
    wtr.flush()?;
    println!("列ベース抽出結果を {} に出力しました。", output);
    Ok(())
}


// 行範囲・列リストで抽出
fn extract_row(
    file_pattern: &str,
    sheet_name: &str,
    range: &str,
    columns: &str,
    headers: &str,
    filter_col: &str,
    output: &str,
    with_filename: bool,
) -> Result<(), Box<dyn Error>> {
    let mut wtr = Writer::from_path(output)?;
    write_csv_header(&mut wtr, headers, with_filename)?;

    let (row_start, row_end) = {
        let parts: Vec<&str> = range.split(':').collect();
        (parts[0].parse::<usize>()?, parts[1].parse::<usize>()?)
    };
    let cols: Vec<&str> = columns.split(',').collect();

    let entries = collect_files(file_pattern)?;
    println!("対象ファイル数: {}", entries.len());

    for path in entries {
        let range_data = open_sheet(&path, sheet_name)?;

        for row_idx in row_start..=row_end {
            let mut row_vals = Vec::new();
            for &col in &cols {
                let col_idx = col_str_to_index(col);
                let val = range_data.get((row_idx - 1, col_idx)).unwrap_or(&DataType::Empty);
                row_vals.push(match val {
                    DataType::String(s) => s.to_string(),
                    DataType::Float(f) => f.to_string(),
                    DataType::Int(i) => i.to_string(),
                    DataType::Bool(b) => b.to_string(),
                    _ => "".to_string(),
                });
            }

            // フィルター
            let should_include = if filter_col.is_empty() {
                true
            } else if let Some(pos) = cols.iter().position(|&c| c == filter_col) {
                !row_vals[pos].is_empty()
            } else {
                true
            };
            if !should_include {
                continue;
            }

            if with_filename {
                let fname = path.file_name().and_then(|s| s.to_str()).unwrap_or("");
                let mut record = vec![fname.to_string()];
                record.extend(row_vals);
                wtr.write_record(record.iter().map(|s| s.as_str()))?;
            } else {
                wtr.write_record(row_vals.iter().map(|s| s.as_str()))?;
            }
        }
    }
    wtr.flush()?;
    println!("行ベース抽出結果を {} に出力しました。", output);
    Ok(())
}


fn main() -> Result<(), Box<dyn Error>> {
    let cli = Cli::parse();

    match &cli.command {
        Commands::ExtractRow {
            file_pattern,
            sheet,
            range,
            columns,
            headers,
            filter_col,
            output,
            with_filename,
        } => {
            extract_row(
                file_pattern,
                sheet,
                range,
                columns,
                headers,
                filter_col,
                output,
                *with_filename,
            )?;
        }
        Commands::ExtractCol {
            file_pattern,
            sheet,
            col_range,
            rows,
            headers,
            filter_row,
            output,
            with_filename,
        } => {
            extract_col(
                file_pattern,
                sheet,
                col_range,
                rows,
                headers,
                filter_row,
                output,
                *with_filename,
            )?;
        }
    }

    Ok(())
}


