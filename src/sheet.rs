use std::{io::{Error, ErrorKind, Result, Write}, iter::FromIterator};

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

pub struct AutoFilter {
    pub start_col: String,
    pub end_col: String,
    pub start_row: usize,
    pub end_row: usize,
}

impl ToString for AutoFilter {
    fn to_string(&self) -> String {
        format!(
            "{}{}:{}{}",
            self.start_col, self.start_row, self.end_col, self.end_row
        )
    }
}

#[derive(Default)]
pub struct Sheet {
    pub id: usize,
    pub name: String,
    pub columns: Vec<Column>,
    max_row_index: usize,
    pub calc_chain: Vec<String>,
    pub merged_cells: Vec<MergedCell>,
    pub auto_filter: Option<AutoFilter>,
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
    pub value: CellData,
}

pub struct MergedCell {
    pub start_ref: String,
    pub end_ref: String,
}

pub struct Column {
    pub width: f32,
}

#[derive(Clone)]
pub struct CellData {
    pub value: CellValue,
    pub xf_format: Option<XfFormat>,
}
impl From<CellValue> for CellData {
    fn from(value: CellValue) -> Self {
        CellData {
            value,
            xf_format: None,
        }
    }
}
impl CellData {
    pub fn with_xf_format(mut self, xf_format: XfFormat) -> Self {
        self.xf_format = Some(xf_format);
        self
    }
}

#[derive(Copy, Clone, Debug)]
pub struct XfFormat(pub(crate) u16);
impl XfFormat {
    pub fn value(&self) -> u16 {
        self.0
    }
}

#[derive(Copy, Clone, Debug)]
pub struct NumberFormat(pub(crate) u16);
impl NumberFormat {
    pub fn value(&self) -> u16 {
        self.0
    }
}

#[derive(Copy, Clone, Debug)]
pub struct Border(pub(crate) u16);
impl Border {
    pub fn value(&self) -> u16 {
        self.0
    }
}

#[derive(Copy, Clone, Debug)]
pub struct Font(pub(crate) u16);
impl Font {
    pub fn value(&self) -> u16 {
        self.0
    }
}

#[derive(Copy, Clone, Debug)]
pub struct Fill(pub(crate) u16);
impl Fill {
    pub fn value(&self) -> u16 {
        self.0
    }
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

pub trait ToCellData {
    fn to_cell_data(&self) -> CellData;
}

impl ToCellData for CellData {
    fn to_cell_data(&self) -> CellData {
        self.clone()
    }
}

impl ToCellData for CellValue {
    fn to_cell_data(&self) -> CellData {
        self.clone().into()
    }
}

impl ToCellData for bool {
    fn to_cell_data(&self) -> CellData {
        CellValue::Bool(self.to_owned()).into()
    }
}

impl ToCellData for (f64, XfFormat) {
    fn to_cell_data(&self) -> CellData {
        let (number, format) = self;
        CellData::from(CellValue::Number(*number)).with_xf_format(*format)
    }
}

impl ToCellData for f64 {
    fn to_cell_data(&self) -> CellData {
        CellValue::Number(self.to_owned()).into()
    }
}

impl ToCellData for String {
    fn to_cell_data(&self) -> CellData {
        if self.starts_with('=') {
            return CellValue::Formula((&self[1..]).to_string()).into();
        }
        CellValue::String(self.to_owned()).into()
    }
}

impl<'a> ToCellData for &'a str {
    fn to_cell_data(&self) -> CellData {
        if self.starts_with('=') {
            return CellValue::Formula((&self[1..]).to_string()).into();
        }
        CellValue::String(self.to_string()).into()
    }
}

impl ToCellData for () {
    fn to_cell_data(&self) -> CellData {
        CellValue::Blank(1).into()
    }
}

#[cfg(feature = "chrono")]
impl ToCellData for chrono::NaiveDateTime {
    fn to_cell_data(&self) -> CellData {
        let seconds = self.timestamp();
        let nanos = f64::from(self.timestamp_subsec_nanos()) * 1e-9;
        let unix_seconds = seconds as f64 + nanos;
        let unix_days = unix_seconds / 86400.;
        CellValue::Datetime(unix_days + 25569.).into()
    }
}

#[cfg(feature = "chrono")]
impl ToCellData for chrono::NaiveDate {
    fn to_cell_data(&self) -> CellData {
        use chrono::Datelike;
        const UNIX_EPOCH_DAY: i32 = 719_163;

        let unix_days: f64 = (self.num_days_from_ce() - UNIX_EPOCH_DAY).into();
        CellValue::Date(unix_days + 25569.).into()
    }
}

impl<A: ToCellData> FromIterator<A> for Row {
    fn from_iter<T: IntoIterator<Item = A>>(iter: T) -> Self {
        let mut row = Row::new();
        for val in iter {
            row.add_cell(val)
        }
        row
    }
}

impl Row {
    pub fn new() -> Row {
        Row {
            ..Default::default()
        }
    }

    pub fn add_cell<T>(&mut self, value: T)
    where
        T: ToCellData + Sized,
    {
        let value = &value.to_cell_data();
        match &value.value {
            CellValue::Formula(f) => {
                self.calc_chain.push(f.to_owned());
                self.max_col_index += 1;
                self.cells.push(Cell {
                    column_index: self.max_col_index,
                    value: value.clone(),
                })
            }
            CellValue::Blank(cols) => self.max_col_index += cols,
            _ => {
                self.max_col_index += 1;
                self.cells.push(Cell {
                    column_index: self.max_col_index,
                    value: value.clone(),
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
            cell.value.value = match &cell.value.value {
                CellValue::String(val) => shared.register(&escape_xml(val)),
                x => x.to_owned(),
            };
        }
        self
    }
}

fn write_value(cv: &CellData, ref_id: String, writer: &mut dyn Write) -> Result<()> {
    let cell_style = cv
        .xf_format
        .map(|XfFormat(xf_format)| format!(" s=\"{xf_format}\""))
        .unwrap_or_default();
    match &cv.value {
        CellValue::Bool(b) => {
            let v = if *b { 1 } else { 0 };
            let s = format!("<c r=\"{}\" t=\"b\"{cell_style}><v>{}</v></c>", ref_id, v);
            writer.write_all(s.as_bytes())?;
        }
        &CellValue::Number(num) => write_number(&ref_id, num, cv.xf_format, writer)?,
        #[cfg(feature = "chrono")]
        &CellValue::Date(num) => write_number(&ref_id, num, Some(XfFormat(1)), writer)?,
        #[cfg(feature = "chrono")]
        &CellValue::Datetime(num) => write_number(&ref_id, num, Some(XfFormat(2)), writer)?,
        CellValue::String(ref s) => {
            let s = format!(
                "<c r=\"{}\" t=\"str\"{cell_style}><v>{}</v></c>",
                ref_id,
                escape_xml(s)
            );
            writer.write_all(s.as_bytes())?;
        }
        CellValue::Formula(ref s) => {
            let s = format!(
                "<c r=\"{}\" t=\"str\"{cell_style}><f>{}</f></c>",
                ref_id,
                escape_xml(s)
            );
            writer.write_all(s.as_bytes())?;
        }
        CellValue::SharedString(ref s) => {
            let s = format!("<c r=\"{}\" t=\"s\"{cell_style}><v>{}</v></c>", ref_id, s);
            writer.write_all(s.as_bytes())?;
        }
        CellValue::Blank(_) => {}
    }
    Ok(())
}

fn write_number(
    ref_id: &str,
    value: f64,
    style: Option<XfFormat>,
    writer: &mut dyn Write,
) -> Result<()> {
    write!(writer, r#"<c r="{ref_id}""#)?;
    if let Some(XfFormat(format)) = style {
        write!(writer, r#" s="{format}""#)?;
    }
    write!(writer, r#"><v>{value}</v></c>"#)
}

pub fn escape_xml(str: &str) -> String {
    let str = str.replace('&', "&amp;");
    let str = str.replace('<', "&lt;");
    let str = str.replace('>', "&gt;");
    let str = str.replace('\'', "&apos;");
    str.replace('\"', "&quot;")
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

    String::from_iter(result)
}

pub fn validate_name(name: &str) -> String {
    escape_xml(name).replace('/', "-")
}

impl Sheet {
    pub fn new(id: usize, sheet_name: &str) -> Sheet {
        Sheet {
            id,
            name: validate_name(sheet_name), //sheet_name.to_owned(),//escape_xml(sheet_name),
            ..Default::default()
        }
    }

    /// Adds the "AutoFilter" feature to the specified range of columns and rows (1-indexed).
    /// The arguments are used to construct the range of columns and rows used by the "AutoFilter"
    /// feature. For example: Column 1, Row 1 to Column 2, Row 2 will create the range "A1:B2".
    /// If invalid parameters are provided, the "AutoFilter" is not created.
    pub fn add_auto_filter(
        &mut self,
        start_col: usize,
        end_col: usize,
        start_row: usize,
        end_row: usize,
    ) {
        if start_col > 0 && start_row > 0 && start_col <= end_col && start_row <= end_row {
            self.auto_filter = Some(AutoFilter {
                start_col: column_letter(start_col),
                end_col: column_letter(end_col),
                start_row,
                end_row,
            });
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
        if let Some(auto_filter) = &self.auto_filter {
            writer.write_all(
                format!(
                    "\n</sheetData>\n<autoFilter ref=\"{}\"/>\n",
                    auto_filter.to_string()
                )
                .as_bytes(),
            )
        } else {
            writer.write_all(b"\n</sheetData>\n")
        }
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
            .write_row(self.writer, row.replace_strings(self.shared_strings))
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
            .to_cell_data();

        match cell.value {
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
        let cell = NaiveDate::from_ymd(2012, 11, 10).to_cell_data();

        match cell.value {
            CellValue::Date(n) if n == EXPECTED => {}
            CellValue::Date(n) => panic!(
                "invalid chrono::NaiveDate conversion to CellValue. {} is expected, found {}",
                EXPECTED, n
            ),
            _ => panic!("invalid chrono::NaiveDate conversion to CellValue"),
        }
    }
}
