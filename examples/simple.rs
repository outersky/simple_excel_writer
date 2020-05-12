extern crate simple_excel_writer as excel;

use excel::*;
fn main() {
    let mut wb = Workbook::create("./tmp/b.xlsx");
    let mut sheet = wb.create_sheet("SheetName");

    // set column width
    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 80.0 });
    sheet.add_column(Column { width: 60.0 });

    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;
        sw.append_row(row!["Name", "Title", "Success", "XML Remark"])?;
        sw.append_row(row![
            "Amy",
            (),
            true,
            "<xml><tag>\"Hello\" & 'World'</tag></xml>"
        ])?;
        sw.append_blank_rows(2);
        sw.append_row(row!["Tony", blank!(720), "retired"]) // A5: Tony , AAT5 : retired
    })
    .expect("write excel error!");

    let mut sheet = wb.create_sheet("Sheet2");
    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;
        sw.append_row(row!["Name", "Title", "Success", "Remark"])?;
        sw.append_row(row!["Amy", "Manager", true])?;
        sw.append_row(row![1.0, 2.0, 3.0, 4.1, "=sum(a3:d3)"])
    })
    .expect("write excel error!");

    wb.close().expect("close excel error!");
}
