extern crate simple_excel_writer;

use excel::*;
use simple_excel_writer as excel;

fn creates_and_saves_an_excel_sheet_driver(filename: Option<&str>) -> Option<Vec<u8>> {
    let mut wb = if let Some(name) = filename {
        excel::Workbook::create(name)
    } else {
        excel::Workbook::create_in_memory()
    };

    let mut ws = wb.create_sheet("test_sheet");
    ws.add_column(Column { width: 60.0 });
    ws.add_column(Column { width: 30.0 });
    ws.add_column(Column { width: 10.0 });
    ws.add_column(Column { width: 60.0 });
    wb.write_sheet(&mut ws, |sw| {
        sw.append_row(row!["Name", "Title", "Success", "Remark"])
            .unwrap();
        sw.append_row(row![
            "Mary",
            "Acountant",
            false,
            r#"<xml><tag>"" & 'World'</tag></xml>"#
        ])
        .unwrap();
        sw.append_row(row![
            "Mary",
            "Programmer",
            true,
            "<xml><tag>\"Hello\" & 'World'</tag></xml>"
        ])
        .unwrap();
        sw.append_row(row!["Marly", "Mary", "Success", "Success", true, 500.])
    })
    .expect("Write excel error!");

    let mut ws = wb.create_sheet("test_sheet<2");

    wb.write_sheet(&mut ws, |sw| {
        sw.append_row(row!["Name", "Title", "Success"]).unwrap();
        sw.append_row(row!["Mary", "This", true]).unwrap();

        #[cfg(feature = "chrono")]
        sw.append_row(row![
            chrono::NaiveDate::from_ymd(2020, 10, 15).and_hms(18, 27, 11),
            chrono::NaiveDate::from_ymd(2020, 10, 16)
        ])
        .unwrap();
        Ok(())
    })
    .expect("Write excel error!");

    let mut ws = wb.create_sheet("test_sheet3 is very long and breaks the limit of 31 charcters");

    wb.write_sheet(&mut ws, |sw| {
        sw.append_row(row!["Name", "Title", "Success"]).unwrap();
        sw.append_row(row!["Mary", "Sgt Monkey", true])
    })
    .expect("Write excel error!");

    wb.close().expect("Close excel error!")
}

#[test]
fn creates_and_saves_an_excel_sheet() {
    let file_test = creates_and_saves_an_excel_sheet_driver(Some("test.xlsx"));
    assert!(file_test.is_none());

    let in_memory_test = creates_and_saves_an_excel_sheet_driver(None);
    assert!(in_memory_test.is_some());
}
