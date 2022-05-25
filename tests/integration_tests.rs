extern crate simple_excel_writer;
use std::io::{Cursor, Read};
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

#[test]
fn creates_file_and_checks_default_style() {
    let expected_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <fonts count="1">
        <font>
            <sz val="12"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>
    <fills count="2">
        <fill>
            <patternFill patternType="none"/>
        </fill>
        <fill>
            <patternFill patternType="gray125"/>
        </fill>
    </fills>
    <borders count="1">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="3">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>
</styleSheet>"#.to_string();
    let mem_file = creates_and_saves_an_excel_sheet_driver(None).unwrap();
    let result = get_file_as_str_from_zip(&mem_file, "xl/styles.xml");
    assert_eq!(expected_xml, result, "The style sheet should match!");
}

#[test]
fn creates_file_with_custom_number_format_and_checks_style() {
    let expected_xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <numFmts count="3">
        <numFmt numFmtId="165" formatCode="&quot;$&quot;#,##0.00"/>
        <numFmt numFmtId="166" formatCode="#,##0.00&quot; KG&quot;"/>
        <numFmt numFmtId="167" formatCode="#,##0.0&quot;&lt;&gt;&quot;"/>
    </numFmts>
    <fonts count="1">
        <font>
            <sz val="12"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>
    <fills count="2">
        <fill>
            <patternFill patternType="none"/>
        </fill>
        <fill>
            <patternFill patternType="gray125"/>
        </fill>
    </fills>
    <borders count="1">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="6">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>
</styleSheet>"##.to_string();
    let mut wb = excel::Workbook::create_in_memory();
    let dollar_id = wb.add_cust_number_format("\"$\"#,##0.00".to_string());
    let weight_id = wb.add_cust_number_format("#,##0.00\" KG\"".to_string());
    let diamond_id = wb.add_cust_number_format("#,##0.0\"<>\"".to_string());

    assert_eq!(dollar_id, 3);
    assert_eq!(weight_id, 4);
    assert_eq!(diamond_id, 5);

    let mut ws = wb.create_sheet("test_sheet");
    ws.add_column(Column { width: 20.0 });
    ws.add_column(Column { width: 20.0 });
    ws.add_column(Column { width: 20.0 });
    wb.write_sheet(&mut ws, |sw| {
        sw.append_row(row!["Cost", "Weight", "Symbol"]).unwrap();
        sw.append_row(row![(20.1, Some(dollar_id)), (50.12,  Some(weight_id)), (700.0,  Some(diamond_id))])
    })
    .expect("Write excel error!");

    let mem_file = wb.close().unwrap().unwrap();
    let result = get_file_as_str_from_zip(&mem_file, "xl/styles.xml");
    assert_eq!(expected_xml, result, "The style sheet should match!");

    let sheet = get_file_as_str_from_zip(&mem_file, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains(format!("<c r=\"A2\" s=\"{}\"><v>20.1</v></c>", dollar_id).as_str()), "First cell should reference the 3rd index of the cellXfs list");
    assert!(sheet.contains(format!("<c r=\"B2\" s=\"{}\"><v>50.12</v></c>", weight_id).as_str()), "First cell should reference the 4th index of the cellXfs list");
    assert!(sheet.contains(format!("<c r=\"C2\" s=\"{}\"><v>700</v></c>", diamond_id).as_str()), "First cell should reference the 5th (last) index of the cellXfs list");
}