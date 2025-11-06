use std::error::Error;
use std::io::Write;
use std::path::Path;

use calamine::{Reader, Xlsx, open_workbook};
use serde_json::json;

pub fn excel_to_jsonl<P: AsRef<Path>, W: Write>(
    path: P,
    sheet: &str,
    mut writer: W,
) -> Result<(), Box<dyn Error>> {
    let mut excel: Xlsx<_> =
        open_workbook(path).map_err(|e| format!("Failed to open excel file: {e}"))?;

    let range = excel
        .worksheet_range(sheet)
        .map_err(|e| format!("Failed to get sheet or range: {e}"))?;

    let mut rows = range.rows();
    let header = rows
        .next()
        .ok_or_else(|| "Failed to get header row from the sheet".to_string())?;

    for row in rows {
        let mut json_row = json!({});
        for (i, cell) in row.iter().enumerate() {
            let key = header.get(i).map(|c| c.to_string()).unwrap_or_default();
            let value = match cell {
                calamine::Data::Int(v) => json!(v),
                calamine::Data::Float(v) => json!(v),
                calamine::Data::String(v) => json!(v),
                calamine::Data::Bool(v) => json!(v),
                calamine::Data::DateTime(v) => json!(v.to_string()),
                calamine::Data::DateTimeIso(v) => json!(v),
                calamine::Data::DurationIso(v) => json!(v),
                calamine::Data::Error(v) => json!(v.to_string()),
                calamine::Data::Empty => json!(null),
            };
            json_row[key] = value;
        }
        serde_json::to_writer(&mut writer, &json_row)?;
        writeln!(writer)?;
    }

    Ok(())
}
