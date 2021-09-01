use std::io::{Error, ErrorKind, Result, Write};

#[macro_export]
macro_rules! row {
    ($( $x:expr ),*) => {
        {
            let mut row = Row::new();
            $(row.add_cell($x);)*
            row
        }
    };
}

#[macro_export]
macro_rules! blank {
    ($x:expr) => {{
        CellValue::Blank($x)
    }};
    () => {{
        CellValue::Blank(1)
    }};
}

#[derive(Default)]
pub struct Sheet {
    pub id: usize,
    pub name: String,
    pub columns: Vec<Column>,
    max_row_index: usize,
    pub calc_chain: Vec<String>,
    pub merged_cells: Vec<MergedCell>,
}

#[derive(Default)]
pub struct Row {
    pub cells: Vec<Cell>,
    row_index: usize,
    max_col_index: usize,
    calc_chain: Vec<String>,
}

pub struct Cell {
    pub column_index: usize,
    pub value: CellValue,
}

pub struct MergedCell {
    pub start_ref: String,
    pub end_ref: String,
}

pub struct Column {
    pub width: f32,
}

#[derive(Clone)]
pub enum CellValue {
    Bool(bool),
    Number(f64),
    #[cfg(feature = "chrono")]
    Date(f64),
    #[cfg(feature = "chrono")]
    Datetime(f64),
    String(String),
    Formula(String),
    Blank(usize),
    SharedString(String),
}

pub struct SheetWriter<'a, 'b>
where
    'b: 'a,
{
    sheet: &'a mut Sheet,
    writer: &'b mut Vec<u8>,
    shared_strings: &'b mut crate::SharedStrings,
}

pub trait ToCellValue {
    fn to_cell_value(&self) -> CellValue;
}

impl ToCellValue for bool {
    fn to_cell_value(&self) -> CellValue {
        CellValue::Bool(self.to_owned())
    }
}

impl ToCellValue for f64 {
    fn to_cell_value(&self) -> CellValue {
        CellValue::Number(self.to_owned())
    }
}

impl ToCellValue for String {
    fn to_cell_value(&self) -> CellValue {
        if self.starts_with('=') {
            return CellValue::Formula(self.to_owned());
        }
        CellValue::String(self.to_owned())
    }
}

impl<'a> ToCellValue for &'a str {
    fn to_cell_value(&self) -> CellValue {
        if self.starts_with('=') {
            return CellValue::Formula(self.to_string());
        }
        CellValue::String(self.to_string())
    }
}

impl ToCellValue for () {
    fn to_cell_value(&self) -> CellValue {
        CellValue::Blank(1)
    }
}

#[cfg(feature = "chrono")]
impl ToCellValue for chrono::NaiveDateTime {
    fn to_cell_value(&self) -> CellValue {
        let seconds = self.timestamp();
        let nanos = f64::from(self.timestamp_subsec_nanos()) * 1e-9;
        let unix_seconds = seconds as f64 + nanos;
        let unix_days = unix_seconds / 86400.;
        CellValue::Datetime(unix_days + 25569.)
    }
}

#[cfg(feature = "chrono")]
impl ToCellValue for chrono::NaiveDate {
    fn to_cell_value(&self) -> CellValue {
        use chrono::Datelike;
        const UNIX_EPOCH_DAY: i32 = 719_163;

        let unix_days: f64 = (self.num_days_from_ce() - UNIX_EPOCH_DAY).into();
        CellValue::Date(unix_days + 25569.)
    }
}

impl Row {
    pub fn new() -> Row {
        Row {
            ..Default::default()
        }
    }

    pub fn from_iter<T>(iter: impl Iterator<Item = T>) -> Row
    where
        T: ToCellValue + Sized,
    {
        let mut row = Row::new();

        for val in iter {
            row.add_cell(val)
        }

        row
    }

    pub fn add_cell<T>(&mut self, value: T)
    where
        T: ToCellValue + Sized,
    {
        let value = value.to_cell_value();
        match &value {
            CellValue::Formula(f) => {
                self.calc_chain.push(f.to_owned());
                self.max_col_index += 1;
                self.cells.push(Cell {
                    column_index: self.max_col_index,
                    value,
                })
            }
            CellValue::Blank(cols) => self.max_col_index += cols,
            _ => {
                self.max_col_index += 1;
                self.cells.push(Cell {
                    column_index: self.max_col_index,
                    value,
                })
            }
        }
    }

    pub fn add_empty_cells(&mut self, cols: usize) {
        self.max_col_index += cols
    }

    pub fn join(&mut self, row: Row) {
        for cell in row.cells.into_iter() {
            self.inner_add_cell(cell)
        }
    }

    fn inner_add_cell(&mut self, cell: Cell) {
        self.max_col_index += 1;
        self.cells.push(Cell {
            column_index: self.max_col_index,
            value: cell.value,
        })
    }

    pub fn write(&mut self, writer: &mut dyn Write) -> Result<()> {
        let head = format!("<row r=\"{}\">\n", self.row_index);
        writer.write_all(head.as_bytes())?;
        for c in self.cells.iter() {
            c.write(self.row_index, writer)?;
        }
        writer.write_all(b"\n</row>\n")
    }

    pub fn replace_strings(mut self, shared: &mut crate::SharedStrings) -> Self {
        if !shared.used() {
            return self;
        }
        for cell in self.cells.iter_mut() {
            cell.value = match &cell.value {
                CellValue::String(val) => shared.register(&escape_xml(val)),
                x => x.to_owned(),
            };
        }
        self
    }
}

impl ToCellValue for CellValue {
    fn to_cell_value(&self) -> CellValue {
        self.clone()
    }
}

fn write_value(cv: &CellValue, ref_id: String, writer: &mut dyn Write) -> Result<()> {
    match cv {
        CellValue::Bool(b) => {
            let v = if *b { 1 } else { 0 };
            let s = format!("<c r=\"{}\" t=\"b\"><v>{}</v></c>", ref_id, v);
            writer.write_all(s.as_bytes())?;
        }
        &CellValue::Number(num) => write_number(&ref_id, num, None, writer)?,
        #[cfg(feature = "chrono")]
        &CellValue::Date(num) => write_number(&ref_id, num, Some(1), writer)?,
        #[cfg(feature = "chrono")]
        &CellValue::Datetime(num) => write_number(&ref_id, num, Some(2), writer)?,
        CellValue::String(ref s) => {
            let s = format!(
                "<c r=\"{}\" t=\"str\"><v>{}</v></c>",
                ref_id,
                escape_xml(&s)
            );
            writer.write_all(s.as_bytes())?;
        }
        CellValue::Formula(ref s) => {
            let s = format!(
                "<c r=\"{}\" t=\"str\"><f>{}</f></c>",
                ref_id,
                escape_xml(&s)
            );
            writer.write_all(s.as_bytes())?;
        }
        CellValue::SharedString(ref s) => {
            let s = format!("<c r=\"{}\" t=\"s\"><v>{}</v></c>", ref_id, s);
            writer.write_all(s.as_bytes())?;
        }
        CellValue::Blank(_) => {}
    }
    Ok(())
}

fn write_number(
    ref_id: &str,
    value: f64,
    style: Option<u16>,
    writer: &mut dyn Write,
) -> Result<()> {
    match style {
        Some(style) => write!(
            writer,
            r#"<c r="{}" s="{}"><v>{}</v></c>"#,
            ref_id, style, value
        ),
        None => write!(writer, r#"<c r="{}"><v>{}</v></c>"#, ref_id, value),
    }
}

fn escape_xml(str: &str) -> String {
    let str = str.replace("&", "&amp;");
    let str = str.replace("<", "&lt;");
    let str = str.replace(">", "&gt;");
    let str = str.replace("'", "&apos;");
    str.replace("\"", "&quot;")
}

impl Cell {
    fn write(&self, row_index: usize, writer: &mut dyn Write) -> Result<()> {
        write_value(&self.value, ref_id(self.column_index, row_index), writer)
    }
}

impl MergedCell {
    fn write(&self, writer: &mut dyn Write) -> Result<()> {
        write!(
            writer,
            "<mergeCell ref=\"{}:{}\" />",
            self.start_ref, self.end_ref
        )?;

        Ok(())
    }
}

pub fn ref_id(column_index: usize, row_index: usize) -> String {
    format!("{}{}", column_letter(column_index), row_index)
}

/**
 * column_index : 1-based
 */
pub fn column_letter(column_index: usize) -> String {
    let mut column_index = (column_index - 1) as isize; // turn to 0-based;
    let single = |n: u8| {
        // n : 0-based
        (b'A' + n) as char
    };
    let mut result = vec![];
    while column_index >= 0 {
        result.push(single((column_index % 26) as u8));
        column_index = column_index / 26 - 1;
    }

    let result = result.into_iter().rev();

    use std::iter::FromIterator;
    String::from_iter(result)
}

pub fn validate_name(name: &str) -> String {
    escape_xml(name).replace("/", "-")
}

impl Sheet {
    pub fn new(id: usize, sheet_name: &str) -> Sheet {
        Sheet {
            id,
            name: validate_name(sheet_name), //sheet_name.to_owned(),//escape_xml(sheet_name),
            ..Default::default()
        }
    }

    pub fn add_column(&mut self, column: Column) {
        self.columns.push(column)
    }

    fn write_row<W>(&mut self, writer: &mut W, mut row: Row) -> Result<()>
    where
        W: Write + Sized,
    {
        self.max_row_index += 1;
        row.row_index = self.max_row_index;
        self.calc_chain.append(&mut row.calc_chain);
        row.write(writer)
    }

    fn write_blank_rows(&mut self, rows: usize) {
        self.max_row_index += rows;
    }

    fn write_head(&self, writer: &mut dyn Write) -> Result<()> {
        let header = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        "#;
        writer.write_all(header.as_bytes())?;
        /*
                let dimension = format!("<dimension ref=\"A1:{}{}\"/>", column_letter(self.dimension.columns), self.dimension.rows);
                writer.write_all(dimension.as_bytes())?;
        */

        if self.columns.is_empty() {
            return Ok(());
        }

        writer.write_all(b"\n<cols>\n")?;
        let mut i = 1;
        for col in self.columns.iter() {
            writer.write_all(
                format!(
                    "<col min=\"{}\" max=\"{}\" width=\"{}\" customWidth=\"1\"/>\n",
                    &i, &i, col.width
                )
                .as_bytes(),
            )?;
            i += 1;
        }
        writer.write_all(b"</cols>\n")
    }

    fn write_merged_cells(&self, writer: &mut dyn Write) -> Result<()> {
        if !self.merged_cells.is_empty() {
            write!(writer, "<mergeCells count=\"{}\">", self.merged_cells.len())?;
            for merged_cell in self.merged_cells.iter() {
                merged_cell.write(writer)?;
            }
            write!(writer, "</mergeCells>")?;
        }
        writer.flush()?;

        Ok(())
    }

    fn write_data_begin(&self, writer: &mut dyn Write) -> Result<()> {
        writer.write_all(b"\n<sheetData>\n")
    }

    fn write_data_end(&self, writer: &mut dyn Write) -> Result<()> {
        writer.write_all(b"\n</sheetData>\n")
    }

    fn close(&self, writer: &mut dyn Write) -> Result<()> {
        writer.write_all(b"</worksheet>\n")
    }
}

impl<'a, 'b> SheetWriter<'a, 'b> {
    pub fn new(
        sheet: &'a mut Sheet,
        writer: &'b mut Vec<u8>,
        shared_strings: &'b mut crate::SharedStrings,
    ) -> SheetWriter<'a, 'b> {
        SheetWriter {
            sheet,
            writer,
            shared_strings,
        }
    }

    pub fn append_row(&mut self, row: Row) -> Result<()> {
        self.sheet
            .write_row(self.writer, row.replace_strings(&mut self.shared_strings))
    }

    pub fn append_blank_rows(&mut self, rows: usize) {
        self.sheet.write_blank_rows(rows)
    }

    /// Merges the range between `start` and `end` cells, specified as 1-based `(column, row)` pairs.
    /// For example, `(1, 2)` is equivalent to cell `A2`.
    pub fn merge_cells(&mut self, start: (usize, usize), end: (usize, usize)) -> Result<()> {
        if end.0 >= start.0 && end.1 >= start.1 {
            self.sheet.merged_cells.push(MergedCell {
                start_ref: ref_id(start.0, start.1),
                end_ref: ref_id(end.0, end.1),
            });

            Ok(())
        } else {
            Err(Error::new(ErrorKind::Other, "invalid range"))
        }
    }

    /// Merges the range between `start_ref` and `end_ref` cells, specified as cell ref IDs (e.g.
    /// `B3`).
    pub fn merge_range(&mut self, start_ref: String, end_ref: String) -> Result<()> {
        self.sheet
            .merged_cells
            .push(MergedCell { start_ref, end_ref });

        Ok(())
    }

    /// Merges cells in a `width` by `height` range beginning at `start`.
    /// Arguments `width` and `height` specify the final size of the merged range, so specifying
    /// `1` for each would result in a single cell with no change, and specifying `0` for either is
    /// invalid.
    pub fn merge_area(&mut self, start: (usize, usize), width: usize, height: usize) -> Result<()> {
        self.merge_cells(start, (start.0 + width - 1, start.1 + height - 1))
    }

    pub fn write<F>(&mut self, write_data: F) -> Result<()>
    where
        F: FnOnce(&mut SheetWriter) -> Result<()> + Sized,
    {
        self.sheet.write_head(self.writer)?;

        self.sheet.write_data_begin(self.writer)?;

        write_data(self)?;

        self.sheet.write_data_end(self.writer)?;
        self.sheet.write_merged_cells(self.writer)?;
        self.sheet.close(self.writer)
    }
}

#[cfg(test)]
#[cfg(feature = "chrono")]
mod chrono_tests {
    use chrono::NaiveDate;

    use super::*;

    #[test]
    fn chrono_datetime() {
        const EXPECTED: f64 = 41223.63725694444;
        let cell = NaiveDate::from_ymd(2012, 11, 10)
            .and_hms(15, 17, 39)
            .to_cell_value();

        match cell {
            CellValue::Datetime(n) if n == EXPECTED => {}
            CellValue::Datetime(n) => panic!(
                "invalid chrono::NaiveDateTime conversion to CellValue. {} is expected, found {}",
                EXPECTED, n
            ),
            _ => panic!("invalid chrono::NaiveDateTime conversion to CellValue"),
        }
    }

    #[test]
    fn chrono_date() {
        const EXPECTED: f64 = 41223.;
        let cell = NaiveDate::from_ymd(2012, 11, 10).to_cell_value();

        match cell {
            CellValue::Date(n) if n == EXPECTED => {}
            CellValue::Date(n) => panic!(
                "invalid chrono::NaiveDate conversion to CellValue. {} is expected, found {}",
                EXPECTED, n
            ),
            _ => panic!("invalid chrono::NaiveDate conversion to CellValue"),
        }
    }
}
