
use pyo3::prelude::*;
use uuid::Uuid;
use xlsxwriter::Workbook;



fn write_to_xlsx(vals: Vec<Vec<String>>) -> String {
    let uuid_f = Uuid::new_v4().to_string();
    let workbook = Workbook::new_with_options(&((&uuid_f).to_string() + ".xlsx"), true, None, false);

    let mut sheet1 = workbook.add_worksheet(None).expect("error");

    sheet1.set_header("&LCiao").expect("error");
    for (idx, _row) in vals.iter().enumerate(){
        for (idx_r, _rr) in _row.iter().enumerate(){
            sheet1.write_string(idx as u32, idx_r as u16, &*_rr, None).expect("error");

        }

    }
    workbook.close().expect("error");
    uuid_f + ".xlsx"
}

/// Formats the sum of two numbers as string.
#[pyfunction]
fn write_xlsx(vals: Vec<Vec<String>>) -> PyResult<String> {
    Ok(write_to_xlsx(vals))
}

/// A Python module implemented in Rust.
#[pymodule]
fn xlsxtopy(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(write_xlsx, m)?)?;
    Ok(())
}