extern crate simple_excel_writer;

use simple_excel_writer as excel;

#[test]
fn adds_autofilter_to_worksheet() {
    let mut wb = excel::Workbook::create_in_memory();
    let mut ws = wb.create_sheet("test_sheet");

    ws.add_auto_filter(0, 0, 0, 0);
    assert!(
        ws.auto_filter.is_none(),
        "An invalid range should not create an autofilter"
    );

    ws.add_auto_filter(1, 3, 3, 2);
    assert!(
        ws.auto_filter.is_none(),
        "An invalid range should not create an autofilter"
    );

    ws.add_auto_filter(24, 3, 1, 2);
    assert!(
        ws.auto_filter.is_none(),
        "An invalid range should not create an autofilter"
    );

    ws.add_auto_filter(1, 1, 1, 1);
    assert!(ws.auto_filter.is_some(), "No autofilter was created!");
    assert_eq!("A1:A1", ws.auto_filter.as_ref().unwrap().to_string());

    ws.add_auto_filter(1, 4, 1, 21);
    assert_eq!("A1:D21", ws.auto_filter.as_ref().unwrap().to_string());

    ws.add_auto_filter(4, 20, 1, 455);
    assert_eq!("D1:T455", ws.auto_filter.as_ref().unwrap().to_string());

    _ = wb.close();
}
