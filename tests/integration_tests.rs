extern crate simple_excel_writer;

use excel::*;
use simple_excel_writer as excel;

#[test]
fn creates_and_saves_an_excel_sheet() {
    let mut wb = excel::Workbook::create("test.xlsx");
    let mut ws = wb.create_sheet("test_sheet");
    ws.add_column(Column { width: 60.0 });
    ws.add_column(Column { width: 30.0 });
    ws.add_column(Column { width: 10.0 });
    ws.add_column(Column { width: 60.0 });
    wb.write_sheet(&mut ws, |sw| {
        sw.append_row(row!["Name", "Title", "Success", "Remark"]).unwrap();
        sw.append_row(row!["Mary", "Acountant", false, r#"<xml><tag>"" & 'World'</tag></xml>"#]).unwrap();
        sw.append_row(row!["Mary", "Programmer", true, "<xml><tag>\"Hello\" & 'World'</tag></xml>"]).unwrap();
        sw.append_row(row!["Marly", "Mary",  "Success", "Success", true, 500.])
    })
    .expect("Write excel error!");

    let mut ws = wb.create_sheet("test_sheet<2");

    wb.write_sheet(&mut ws, |sw| {
        sw.append_row(row!["Name", "Title", "Success"]).unwrap();
        sw.append_row(row!["Mary", "This", true])
    })
    .expect("Write excel error!");

    let mut ws = wb.create_sheet("test_sheet3 is very long and breaks the limit of 31 charcters");

    wb.write_sheet(&mut ws, |sw| {
        sw.append_row(row!["Name", "Title", "Success"]).unwrap();
        sw.append_row(row!["Mary", "Sgt Monkey", true])
    })
    .expect("Write excel error!");

    wb.close().expect("Close excel error!");
}
