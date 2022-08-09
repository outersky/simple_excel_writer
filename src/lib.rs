//! # simple_excel_writer
//! Simple Excel Writer
//!
//! #### Install
//! Just include another `[dependencies.*]` section into your Cargo.toml:
//!
//! ```toml
//! [dependencies]
//! simple_excel_writer="0.1.4"
//! ```
//! #### Sample
//!
//! ```
//! #[macro_use]
//! extern crate simple_excel_writer;
//! use simple_excel_writer as excel;
//!
//! use excel::*;
//!
//! fn main() {
//!     let mut wb = Workbook::create("/tmp/b.xlsx");
//!     let mut sheet = wb.create_sheet("SheetName");
//!
//!     // set column width
//!     sheet.add_column(Column { width: 30.0 });
//!     sheet.add_column(Column { width: 30.0 });
//!     sheet.add_column(Column { width: 80.0 });
//!     sheet.add_column(Column { width: 60.0 });
//!
//!     wb.write_sheet(&mut sheet, |sheet_writer| {
//!         let sw = sheet_writer;
//!         sw.append_row(row!["Name", "Title","Success","XML Remark"])?;
//!         sw.append_row(row!["Amy", (), true,"<xml><tag>\"Hello\" & 'World'</tag></xml>"])?;
//!         sw.append_blank_rows(2);
//!         sw.append_row(row!["Tony", blank!(30), "retired"])
//!     }).expect("write excel error!");
//!
//!     let mut sheet = wb.create_sheet("Sheet2");
//!     wb.write_sheet(&mut sheet, |sheet_writer| {
//!         let sw = sheet_writer;
//!         sw.append_row(row!["Name", "Title","Success","Remark"])?;
//!         sw.append_row(row!["Amy", "Manager", true])
//!     }).expect("write excel error!");
//!
//!     let euro_fmt_idx = wb.add_number_format("\"€\"#,##0.00".to_string());
//!     let weight_fmt_idx = wb.add_number_format("#,##0.0\" KG\"".to_string());
//!     let euro_fmt = wb.add_cell_xf(CellXf { num_fmt: Some(euro_fmt_idx), ..Default::default() });
//!     let weight_fmt = wb.add_cell_xf(CellXf { num_fmt: Some(weight_fmt_idx), ..Default::default() });
//!     let mut sheet_num_fmt = wb.create_sheet("SheetNumFormatted");
//!     sheet_num_fmt.add_column(Column { width: 30.0 });
//!     sheet_num_fmt.add_column(Column { width: 30.0 });
//!     wb.write_sheet(&mut sheet_num_fmt, |sheet_writer| {
//!         let sw = sheet_writer;
//!         sw.append_row(row!["Weight", "Price"])?;
//!         sw.append_row(row![(700.5, weight_fmt), (12045.99, euro_fmt)])?;
//!         sw.append_row(row![(1525.0, weight_fmt), (25999.00, euro_fmt)])
//!     }).expect("write excel error!");
//!
//!     wb.close().expect("close excel error!");
//! }
//! ```
//!

#![crate_name = "simple_excel_writer"]

extern crate zip;

#[cfg(feature = "chrono")]
extern crate chrono;

pub use sheet::*;
pub use workbook::*;

pub mod sheet;
pub mod workbook;

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {}
}
