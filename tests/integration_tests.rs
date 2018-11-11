extern crate simple_excel_writer;
use simple_excel_writer as excel;
use excel::*;
#[test]
fn creates_and_saves_an_excel_sheet() {
    let mut wb = excel::Workbook::create("test.xlsx");
    let mut ws = wb.create_sheet("test_sheet");

    wb.write_sheet(&mut ws, |sw| {
        sw.append_row(row!["Name", "Title", "Success"])

    }).expect("Write excel error!");

    wb.close().expect("Close excel error!");

}
