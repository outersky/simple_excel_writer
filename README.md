# simple_excel_writer
simple excel writer in Rust

[![Build Status](https://travis-ci.org/outersky/simple_excel_writer.png?branch=master)](https://travis-ci.org/outersky/simple_excel_writer) 
[Documentation](https://docs.rs/simple_excel_writer/)

## Example

```rust,no_run
#[macro_use]
extern crate simple_excel_writer as excel;

use excel::*;

fn main() {
    let mut wb = Workbook::create("/tmp/b.xlsx");
    let mut sheet = wb.create_sheet("SheetName");

    // set column width
    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 80.0 });
    sheet.add_column(Column { width: 60.0 });

    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;
        sw.append_row(row!["Name", "Title","Success","XML Remark"])?;
        sw.append_row(row!["Amy", (), true,"<xml><tag>\"Hello\" & 'World'</tag></xml>"])?;
        sw.append_blank_rows(2);
        sw.append_row(row!["Tony", blank!(2), "retired"])
    }).expect("write excel error!");

    let mut sheet = wb.create_sheet("Sheet2");
    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;
        sw.append_row(row!["Name", "Title","Success","Remark"])?;
        sw.append_row(row!["Amy", "Manager", true])
    }).expect("write excel error!");

    wb.close().expect("close excel error!");
}
```

## Todo

- support style

## Change Log

### 0.1.9 (2021-10-28)
- support formula 
- support NaiveDate & NaiveDateTime
- format dates and date times
- Sheet name validation
- remove unndecessary bzip2 dependency

many thanks to all contributors !

#### 0.1.7 (2020-04-29)
- support create-in-memory mode, thanks to Maxburke.

```
This change creates all worksheet files in-memory and only writes them
to disk once the XLSX file is closed.

A new option for creating a version that is in-memory only is available
with `Worksheet::create_in_memory()` which returns the buffer holding
the completed XLSX file contents when closed.
```

#### 0.1.6 (2020-04-06)
- support shared strings between worksheets, thanks to Mikael Edlund.

#### 0.1.5 (2019-03-21)
- support Windows platform, thanks to Carl Fredrik Samson.

#### 0.1.4 (2017-03-24)
- escape xml characters.

#### 0.1.3 (2017-01-03)
- support 26+ columns .
- fix column width bug.

#### 0.1.2 (2017-01-02)
- support multiple sheets

#### 0.1 (2017-01-01)
- generate the basic xlsx file

## License
Apache-2.0
