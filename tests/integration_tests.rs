extern crate simple_excel_writer;
use excel::*;
use simple_excel_writer as excel;
use std::io::{Cursor, Read};

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

fn get_file_as_str_from_zip(mem_file: &Vec<u8>, file_name: &str) -> String {
    let mut archive = zip::read::ZipArchive::new(Cursor::new(mem_file)).unwrap();
    let mut zip_file = archive.by_name(file_name).unwrap();
    let mut temp_buf = vec![];
    let _ = zip_file.read_to_end(&mut temp_buf).unwrap();
    std::str::from_utf8(&temp_buf[..]).unwrap().to_string()
}

#[test]
fn creates_and_saves_an_excel_sheet() {
    let file_test = creates_and_saves_an_excel_sheet_driver(Some("test.xlsx"));
    assert!(file_test.is_none());

    let in_memory_test = creates_and_saves_an_excel_sheet_driver(None);
    assert!(in_memory_test.is_some());
}

const DEFAULT_STYLE_XML: &str = include_str!("default_style.xml");

#[test]
fn creates_file_and_checks_default_style() {
    let mem_file = creates_and_saves_an_excel_sheet_driver(None).unwrap();
    let result = get_file_as_str_from_zip(&mem_file, "xl/styles.xml");
    assert_eq!(
        DEFAULT_STYLE_XML.to_string(),
        result,
        "The style sheet should match!"
    );
}

#[test]
fn creates_file_with_custom_number_format_and_checks_style() {
    let expected_xml = include_str!("creates_file_with_custom_number_format_and_checks_style.xml");
    let mut wb = excel::Workbook::create_in_memory();
    let dollar_idx = wb.add_number_format("\"$\"#,##0.00".to_string());
    let dollar_fmt = wb.add_cell_xf(CellXf {
        num_fmt: Some(dollar_idx),
        ..Default::default()
    });
    let weight_idx = wb.add_number_format("#,##0.00\" KG\"".to_string());
    let weight_fmt = wb.add_cell_xf(CellXf {
        num_fmt: Some(weight_idx),
        ..Default::default()
    });
    let diamond_idx = wb.add_number_format("#,##0.0\"<>\"".to_string());
    let diamond_fmt = wb.add_cell_xf(CellXf {
        num_fmt: Some(diamond_idx),
        ..Default::default()
    });

    assert_eq!(dollar_fmt.value(), 3);
    assert_eq!(weight_fmt.value(), 4);
    assert_eq!(diamond_fmt.value(), 5);

    let mut ws = wb.create_sheet("test_sheet");
    ws.add_column(Column { width: 20.0 });
    ws.add_column(Column { width: 20.0 });
    ws.add_column(Column { width: 20.0 });
    wb.write_sheet(&mut ws, |sw| {
        sw.append_row(row!["Cost", "Weight", "Symbol"])
            .expect("Should append header!");
        sw.append_row(row![
            (20.1, dollar_fmt),
            (50.12, weight_fmt),
            (700.0, diamond_fmt)
        ])
    })
    .expect("Write excel error!");

    let mem_file = wb
        .close()
        .expect("No error on workbook close!")
        .expect("Should have file in memory!");
    let result = get_file_as_str_from_zip(&mem_file, "xl/styles.xml");
    assert_eq!(expected_xml, result, "The style sheet should match!");

    let sheet1 = get_file_as_str_from_zip(&mem_file, "xl/worksheets/sheet1.xml");
    assert!(
        dbg!(&sheet1).contains(format!("<c r=\"A2\" s=\"{}\"><v>20.1</v></c>", 3).as_str()),
        "First cell should reference the 3rd index of the cellXfs list"
    );
    assert!(
        sheet1.contains(format!("<c r=\"B2\" s=\"{}\"><v>50.12</v></c>", 4).as_str()),
        "First cell should reference the 4th index of the cellXfs list"
    );
    assert!(
        sheet1.contains(format!("<c r=\"C2\" s=\"{}\"><v>700</v></c>", 5).as_str()),
        "First cell should reference the 5th (last) index of the cellXfs list"
    );
}

#[cfg(feature = "chrono")]
#[test]
fn chrono_check_default_style() {
    let mut wb = excel::Workbook::create_in_memory();
    let mut ws = wb.create_sheet("test_sheet");
    ws.add_column(Column { width: 20.0 });
    ws.add_column(Column { width: 20.0 });
    ws.add_column(Column { width: 20.0 });
    wb.write_sheet(&mut ws, |sw| {
        sw.append_row(row!["Date", "Datetime"])
            .expect("Should append header!");
        sw.append_row(row![
            chrono::NaiveDate::from_ymd(2012, 11, 10),
            chrono::NaiveDate::from_ymd(2014, 9, 8).and_hms(21, 12, 44)
        ])
    })
    .expect("Write excel error!");

    let mem_file = wb
        .close()
        .expect("No error on workbook close!")
        .expect("Should have file in memory!");
    let result = get_file_as_str_from_zip(&mem_file, "xl/styles.xml");
    assert_eq!(
        DEFAULT_STYLE_XML.to_string(),
        result,
        "The style sheet should match!"
    );

    let sheet1 = get_file_as_str_from_zip(&mem_file, "xl/worksheets/sheet1.xml");
    let expected_date_format_idx = 1;
    let expected_datetime_format_idx = 2;
    assert!(
        sheet1.contains(
            format!(
                "<c r=\"A2\" s=\"{}\"><v>41223</v></c>",
                expected_date_format_idx
            )
            .as_str()
        ),
        "Date contains correct reference to date format"
    );
    assert!(
        sheet1.contains(
            format!(
                "<c r=\"B2\" s=\"{}\"><v>41890.88384259259</v></c>",
                expected_datetime_format_idx
            )
            .as_str()
        ),
        "Date contains correct reference to date format"
    );
}
