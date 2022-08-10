use std::collections::BTreeMap;
use std::io::*;
use std::path::*;
use std::{borrow::Borrow, fs::File};

use crate::Border;
use crate::Fill;
use crate::Font;
use crate::{NumberFormat, XfFormat};

use super::{escape_xml, Sheet, SheetWriter};

struct ArchiveFile {
    name: PathBuf,
    data: Vec<u8>,
}

fn path_format(path: &std::path::Path) -> String {
    let buf = String::with_capacity(path.as_os_str().len());

    path.components().fold(buf, |mut buf, comp| {
        if let std::path::Component::Normal(s) = comp {
            if !buf.is_empty() {
                buf.push('/');
            }

            buf.push_str(s.to_string_lossy().borrow());
        }

        buf
    })
}

#[derive(Debug, Copy, Clone, Eq, PartialEq, Hash)]
pub enum BorderStyle {
    None,
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}
impl BorderStyle {
    pub fn xml(&self) -> String {
        format!(
            "style=\"{}\"",
            match self {
                BorderStyle::None => "none",
                BorderStyle::Thin => "thin",
                BorderStyle::Medium => "medium",
                BorderStyle::Dashed => "dashed",
                BorderStyle::Dotted => "dotted",
                BorderStyle::Thick => "thick",
                BorderStyle::Double => "double",
                BorderStyle::Hair => "hair",
                BorderStyle::MediumDashed => "mediumDashed",
                BorderStyle::DashDot => "dashDot",
                BorderStyle::MediumDashDot => "mediumDashDot",
                BorderStyle::DashDotDot => "dashDotDot",
                BorderStyle::MediumDashDotDot => "mediumDashDotDot",
                BorderStyle::SlantDashDot => "slantDashDot",
            }
        )
    }
}

#[derive(Debug, Copy, Clone, Eq, PartialEq, Hash, Default)]
pub struct BorderFormat {
    pub top: Option<(BorderStyle, Color)>,
    pub right: Option<(BorderStyle, Color)>,
    pub bottom: Option<(BorderStyle, Color)>,
    pub left: Option<(BorderStyle, Color)>,
}

#[derive(Debug, Copy, Clone, Eq, PartialEq, Hash)]
pub enum Color {
    Theme(u16),
    Argb(u32),
}
impl Color {
    fn xml(&self) -> String {
        match self {
            Color::Theme(theme) => format!("theme=\"{theme}\""),
            Color::Argb(rgb) => format!("rgb=\"{rgb:0>8X}\""),
        }
    }
}

#[derive(Default)]
pub struct Workbook {
    xlsx_file: Option<String>,
    archive_files: Vec<ArchiveFile>,
    max_sheet_index: usize,
    shared_strings: SharedStrings,
    sheets: Vec<SheetRef>,
    calc_chain: Vec<(String, usize)>,
    saved: bool,
    cell_formats: CellFormats,
}

pub const BASE_NUM_FMT_CUSTOM: u16 = 165;
pub const BASE_FILL: u16 = 2;
pub const BASE_BORDER: u16 = 1;

#[derive(Default)]
struct CellFormats {
    pub num_fmts: BTreeMap<u16, String>,
    pub fills: Vec<(Color, Color)>,
    pub borders: Vec<BorderFormat>,
    pub cell_xfs: Vec<CellXf>,
}

impl CellFormats {
    pub fn new() -> CellFormats {
        let cell_xfs = vec![
            CellXf {
                num_fmt: Some(NumberFormat(0)),
                ..Default::default()
            },
            CellXf {
                num_fmt: Some(NumberFormat(14)),
                ..Default::default()
            },
            CellXf {
                num_fmt: Some(NumberFormat(22)),
                ..Default::default()
            },

        ];
        CellFormats {
            num_fmts: BTreeMap::new(),
            fills: Vec::new(),
            borders: Vec::new(),
            cell_xfs,
        }
    }

    pub fn add_fill(&mut self, fg_color: Color, bg_color: Color) -> Fill {
        let result = BASE_FILL + self.fills.len() as u16;
        self.fills.push((fg_color, bg_color));
        Fill(result)
    }

    pub fn add_border(&mut self, border_format: BorderFormat) -> Border {
        let result = BASE_BORDER + self.borders.len() as u16;
        self.borders.push(border_format);
        Border(result)
    }

    pub fn add_number_format(&mut self, pattern: String) -> NumberFormat {
        let new_id = BASE_NUM_FMT_CUSTOM + self.num_fmts.len() as u16;
        self.num_fmts.insert(new_id, pattern);
        NumberFormat(new_id)
    }

    pub fn add_cell_xf(&mut self, cell_xf: CellXf) -> XfFormat {
        let result = self.cell_xfs.len() as u16;
        self.cell_xfs.push(cell_xf);
        XfFormat(result)
    }
}

#[derive(Default, Clone)]
pub struct CellXf {
    pub num_fmt: Option<NumberFormat>,
    pub font: Option<Font>,
    pub fill: Option<Fill>,
    pub border: Option<Border>,
}

#[derive(Default, Clone)]
pub struct SharedStrings {
    count: usize,
    used: bool,
    strings: Vec<String>,
}

struct SheetRef {
    id: usize,
    name: String,
}

impl SharedStrings {
    pub fn new() -> Self {
        SharedStrings {
            used: true,
            ..Default::default()
        }
    }
    pub fn new_unused() -> Self {
        SharedStrings {
            used: false,
            ..Default::default()
        }
    }
    pub fn used(&self) -> bool {
        self.used
    }
    pub fn set_used(&mut self, using: bool) {
        self.used = using;
    }
    pub fn add_count(&mut self) {
        self.count += 1;
    }
    /// Takes a string value checks if it's present in shared strings and returns a CellValue with the index
    pub fn register(&mut self, val: &str) -> crate::CellValue {
        self.add_count();

        match self.strings.iter().position(|v| v == val) {
            Some(idx) => crate::sheet::CellValue::SharedString(format!("{}", idx)),
            None => {
                self.strings.push(val.to_owned());
                crate::sheet::CellValue::SharedString(format!("{}", (self.strings.len() - 1)))
            }
        }
    }
}

impl Workbook {
    /// Creates a workbook using shared strings
    pub fn create(xlsx_file: &str) -> Self {
        Self {
            xlsx_file: Some(xlsx_file.to_owned()),
            archive_files: Vec::new(),
            max_sheet_index: 0,
            shared_strings: SharedStrings::new(),
            sheets: Vec::new(),
            calc_chain: Vec::new(),
            saved: false,
            cell_formats: CellFormats::new(),
        }
    }
    /// Creates a workbook not using shared strings
    pub fn create_simple(xlsx_file: &str) -> Self {
        Self {
            xlsx_file: Some(xlsx_file.to_owned()),
            archive_files: Vec::new(),
            max_sheet_index: 0,
            shared_strings: SharedStrings::new_unused(),
            sheets: Vec::new(),
            calc_chain: Vec::new(),
            saved: false,
            cell_formats: CellFormats::new(),
        }
    }

    pub fn create_in_memory() -> Self {
        Self {
            xlsx_file: None,
            archive_files: Vec::new(),
            max_sheet_index: 0,
            shared_strings: SharedStrings::new_unused(),
            sheets: Vec::new(),
            calc_chain: Vec::new(),
            saved: false,
            cell_formats: CellFormats::new(),
        }
    }

    pub fn create_sheet(&mut self, sheet_name: &str) -> Sheet {
        self.max_sheet_index += 1;

        let validated_name = crate::validate_name(sheet_name);

        self.sheets.push(SheetRef {
            id: self.max_sheet_index,
            name: validated_name.clone(),
        });

        // `Sheet` has a private field, so we can't just construct it here with needed values.
        // So we must create default sheet first and then mutate it.
        let mut sheet = Sheet::default();

        sheet.id = self.max_sheet_index;
        sheet.name = validated_name;

        sheet
    }

    pub fn close(&mut self) -> Result<Option<Vec<u8>>> {
        self.create_files()?;

        let mut buf = Vec::new();
        {
            let mut cursor = Cursor::new(&mut buf);
            let mut writer = zip::ZipWriter::new(&mut cursor);
            for archive_file in self.archive_files.iter() {
                let options = zip::write::FileOptions::default();
                writer.start_file(path_format(&archive_file.name), options)?;
                writer.write_all(&archive_file.data)?;
            }

            writer.finish()?;
        }

        if let Some(xlsx_file) = &self.xlsx_file {
            let mut file = File::create(xlsx_file)?;
            file.write_all(&buf)?;
            self.saved = true;

            Ok(None)
        } else {
            Ok(Some(buf))
        }
    }

    pub fn add_fill(&mut self, fg_color: Color, bg_color: Color) -> Fill {
        self.cell_formats.add_fill(fg_color, bg_color)
    }

    pub fn add_border(&mut self, border_format: BorderFormat) -> Border {
        self.cell_formats.add_border(border_format)
    }

    pub fn add_number_format(&mut self, format_str: String) -> NumberFormat {
        self.cell_formats.add_number_format(format_str)
    }

    pub fn add_cell_xf(&mut self, cell_xf: CellXf) -> XfFormat {
        self.cell_formats.add_cell_xf(cell_xf)
    }

    fn create_files(&mut self) -> Result<()> {
        let mut root = PathBuf::new();

        // [Content_Types].xml
        root.push("[Content_Types].xml");
        let mut writer = Vec::new();
        self.create_content_types(&mut writer)?;
        self.archive_files.push(ArchiveFile {
            name: root.clone(),
            data: writer,
        });
        root.pop();

        // _rels/.rels
        root.push("_rels");
        root.push(".rels");
        let mut writer = Vec::new();
        Self::create_rels(&mut writer)?;
        self.archive_files.push(ArchiveFile {
            name: root.clone(),
            data: writer,
        });
        root.pop();
        root.pop();

        // docProps
        root.push("docProps");
        root.push("app.xml");
        let mut writer = Vec::new();
        self.create_app(&mut writer)?;
        self.archive_files.push(ArchiveFile {
            name: root.clone(),
            data: writer,
        });
        root.pop();
        root.push("core.xml");
        let mut writer = Vec::new();
        Self::create_core(&mut writer)?;
        self.archive_files.push(ArchiveFile {
            name: root.clone(),
            data: writer,
        });
        root.pop();
        root.pop();

        // xl
        root.push("xl");
        root.push("styles.xml");
        let mut writer = Vec::new();
        self.create_styles(&mut writer)?;
        self.archive_files.push(ArchiveFile {
            name: root.clone(),
            data: writer,
        });
        root.pop();
        //if self.shared_strings.used(){
        root.push("sharedStrings.xml");
        let mut writer = Vec::new();
        self.create_shared_strings(&mut writer)?;
        self.archive_files.push(ArchiveFile {
            name: root.clone(),
            data: writer,
        });
        root.pop();
        //}
        root.push("workbook.xml");
        let mut writer = Vec::new();
        self.create_workbook(&mut writer)?;
        self.archive_files.push(ArchiveFile {
            name: root.clone(),
            data: writer,
        });
        root.pop();

        root.push("calcChain.xml");
        let mut writer = Vec::new();
        self.create_calc_chain(&mut writer)?;
        self.archive_files.push(ArchiveFile {
            name: root.clone(),
            data: writer,
        });
        root.pop();

        // xl/_rels
        root.push("_rels");
        root.push("workbook.xml.rels");
        let mut writer = Vec::new();
        self.create_xl_rels(&mut writer)?;
        self.archive_files.push(ArchiveFile {
            name: root.clone(),
            data: writer,
        });
        root.pop();
        root.pop();

        // xl/theme
        root.push("theme");
        root.push("theme1.xml");
        let mut writer = Vec::new();
        Self::create_xl_theme(&mut writer)?;
        self.archive_files.push(ArchiveFile {
            name: root.clone(),
            data: writer,
        });
        root.pop();
        root.pop();

        /*
                // xl/worksheets
                root.push("worksheets");
                root.push("sheet1.xml");
                let mut writer = Vec::new();
                Self::create_sample_sheet(&mut writer)?;
                self.archive_files.push(ArchiveFile {name: root.clone(), data: writer});
                root.pop();
                root.pop();
        */
        Ok(())
    }

    pub fn write_sheet<F>(&mut self, sheet: &mut Sheet, write_data: F) -> Result<()>
    where
        F: FnOnce(&mut SheetWriter) -> Result<()> + Sized,
    {
        let mut root = PathBuf::new();
        root.push("xl");
        root.push("worksheets");
        root.push(format!("sheet{}.xml", sheet.id));
        for cc in &sheet.calc_chain {
            self.calc_chain.push((cc.to_owned(), sheet.id));
        }

        let mut writer = Vec::new();
        let sw = &mut SheetWriter::new(sheet, &mut writer, &mut self.shared_strings);
        sw.write(write_data)?;
        self.archive_files.push(ArchiveFile {
            name: root,
            data: writer,
        });
        Ok(())
    }

    fn create_calc_chain(&mut self, writer: &mut dyn Write) -> Result<()> {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#;
        writer.write_all(xml)?;
        for x in &self.calc_chain {
            let wb = format!("<c r=\"{}\" i=\"{}\"/>", x.0, x.1);
            writer.write_all(wb.as_bytes())?;
        }
        let tail = br#"</calcChain>"#;
        writer.write_all(tail)
    }

    fn create_content_types(&mut self, writer: &mut dyn Write) -> Result<()> {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" xmlns:xsd="http://www.w3.org/2001/XMLSchema"
       xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/xl/workbook.xml"
              ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
              "#;
        writer.write_all(xml)?;
        for sf in self.sheets.iter() {
            let wb = format!(
                "<Override PartName=\"/xl/worksheets/sheet{}.xml\"
                    ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
                sf.id
            );
            writer.write_all(wb.as_bytes())?;
        }

        let tail = br#"
    <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    <Override PartName="/xl/styles.xml"
              ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml"
              ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
        "#;

        writer.write_all(tail)
    }

    fn create_rels(writer: &mut dyn Write) -> Result<()> {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
        "#;
        writer.write_all(xml)
    }

    fn create_app(&mut self, writer: &mut dyn Write) -> Result<()> {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>SheetJS</Application>
    <HeadingPairs>
        <vt:vector size="2" baseType="variant">
            <vt:variant>
                <vt:lpstr>Worksheets</vt:lpstr>
            </vt:variant>
            <vt:variant>
                <vt:i4>1</vt:i4>
            </vt:variant>
        </vt:vector>
    </HeadingPairs>
    <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
    <vt:lpstr>SheetJS</vt:lpstr>
    "#;
        writer.write_all(xml)?;

        /* let vector = format!(
            "<vt:vector size=\"{}\" baseType=\"lpstr\">",
            self.sheets.len()
        );
        writer.write_all(vector.as_bytes())?;
        for sf in self.sheets.iter() {
            let str = format!(
                "<vt:lpstr>{}</vt:lpstr>",
                sf.name
            );
            writer.write_all(str.as_bytes())?;
        } */
        let tail = r#"</vt:vector>
    </TitlesOfParts>
</Properties>
        "#;

        writer.write_all(tail.as_bytes())
    }

    fn create_core(writer: &mut dyn Write) -> Result<()> {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>
        "#;
        writer.write_all(xml)
    }

    fn create_styles(&self, writer: &mut dyn Write) -> Result<()> {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
"#;
        writer.write_all(xml)?;
        if !self.cell_formats.num_fmts.is_empty() {
            let num_fmts = format!(
                "    <numFmts count=\"{}\">",
                self.cell_formats.num_fmts.len()
            );
            writer.write_all(num_fmts.as_bytes())?;
            // Sort the map for consistent XML format
            let mut fmts_sorted: Vec<(&u16, &String)> = self.cell_formats.num_fmts.iter().collect();
            fmts_sorted.sort_by_key(|item| item.0);
            for (fmt_id, value) in fmts_sorted {
                let fmt = format!(
                    "\n        <numFmt numFmtId=\"{}\" formatCode=\"{}\"/>",
                    fmt_id,
                    escape_xml(value)
                );
                writer.write_all(fmt.as_bytes())?;
            }
            let fmt_tail = "\n    </numFmts>\n".as_bytes();
            writer.write_all(fmt_tail)?;
        }
        let mid = format!(
            r#"    <fonts count="1">
        <font>
            <sz val="12"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>
    <fills count="{}">
        <fill>
            <patternFill patternType="none"/>
        </fill>
        <fill>
            <patternFill patternType="gray125"/>
        </fill>
"#,
            self.cell_formats.fills.len() as u16 + BASE_FILL
        );
        writer.write_all(mid.as_bytes())?;
        for (fg, bg) in self.cell_formats.fills.iter() {
            let fill = format!(
                r#"        <fill>
            <patternFill patternType="solid">
                <fgColor {}/>
                <bgColor {}/>
            </patternFill>
        </fill>
"#,
                fg.xml(),
                bg.xml()
            );
            writer.write_all(fill.as_bytes())?;
        }
        let mid2 = format!(
            r#"    </fills>
    <borders count="{}">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
"#,
            self.cell_formats.borders.len() as u16 + BASE_BORDER
        );
        writer.write_all(mid2.as_bytes())?;
        for border in self.cell_formats.borders.iter() {
            writer.write_all(b"        <border>\n")?;
            let mappings = [
                ("left", border.left.as_ref()),
                ("right", border.right.as_ref()),
                ("top", border.top.as_ref()),
                ("bottom", border.bottom.as_ref()),
            ];
            for (side, border) in mappings {
                if let Some((style, color)) = border {
                    writeln!(writer, "            <{side} {}>", style.xml())?;
                    writeln!(writer, "                <color {}/>", color.xml())?;
                    writeln!(writer, "            </{side}>")?;
                } else {
                    writeln!(writer, "            <{side}/>")?;
                }
            }
            writer.write_all(b"            <diagonal/>\n")?;
            writer.write_all(b"        </border>\n")?;
        }
        let mid3 = br#"    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>"#;
        writer.write_all(mid3)?;
        if !self.cell_formats.cell_xfs.is_empty() {
            let cell_xfs_start = format!(
                "\n    <cellXfs count=\"{}\">",
                self.cell_formats.cell_xfs.len()
            );
            writer.write_all(cell_xfs_start.as_bytes())?;
            for xf in &self.cell_formats.cell_xfs {
                let num_fmt_id = xf.num_fmt.unwrap_or(NumberFormat(0)).0;
                let font_id = xf.font.unwrap_or(Font(0)).0;
                let fill_id = xf.fill.unwrap_or(Fill(0)).0;
                let border_id = xf.border.unwrap_or(Border(0)).0;
                write!(writer, "\n        <xf")?;
                write!(writer, " numFmtId=\"{num_fmt_id}\"")?;
                write!(writer, " fontId=\"{font_id}\"")?;
                write!(writer, " fillId=\"{fill_id}\"")?;
                write!(writer, " borderId=\"{border_id}\"")?;
                write!(writer, " xfId=\"0\"")?;
                if num_fmt_id > 0 {
                    write!(writer, " applyNumberFormat=\"1\"")?;
                }
                if font_id > 0 {
                    write!(writer, " applyFont=\"1\"")?;
                }
                if fill_id > 0 {
                    write!(writer, " applyFill=\"1\"")?;
                }
                if border_id > 0 {
                    write!(writer, " applyBorder=\"1\"")?;
                }
                write!(writer, "/>")?;
            }
            let cell_xfs_end = "\n    </cellXfs>\n".as_bytes();
            writer.write_all(cell_xfs_end)?;
        }
        let tail = br#"    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>
</styleSheet>"#;
        writer.write_all(tail)
    }

    fn create_workbook(&mut self, writer: &mut dyn Write) -> Result<()> {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <workbookPr date1904="false"/>
    <sheets>"#;
        let tail = r#"
    </sheets>
</workbook>
        "#;
        writer.write_all(xml.as_bytes())?;
        for sf in self.sheets.iter() {
            let str = format!(
                "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"rId{}\"/>",
                sf.name,
                sf.id,
                sf.id + 2
            );
            writer.write_all(str.as_bytes())?;
        }
        writer.write_all(tail.as_bytes())
    }
    fn create_shared_strings(&mut self, writer: &mut dyn Write) -> Result<()> {
        let shared_strings = &self.shared_strings; //.as_ref().unwrap();
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        "#;
        let sst = format!("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{}\" uniqueCount=\"{}\">", shared_strings.count, shared_strings.strings.len());
        let tail = r#"</sst>"#;
        writer.write_all(xml.as_bytes())?;
        writer.write_all(sst.as_bytes())?;
        // might have to use a vector instead to ensure index
        for sf in shared_strings.strings.iter() {
            let space = if sf.trim().len() < sf.len() {
                "t xml:space=\"preserve\""
            } else {
                "t"
            };
            let xmls = format!("<si><{}>{}</{}></si>", space, sf, "t");
            writer.write_all(xmls.as_bytes())?;
        }

        writer.write_all(tail.as_bytes())
    }

    fn create_xl_rels(&mut self, writer: &mut dyn Write) -> Result<()> {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        "#;
        let tail = br#"
</Relationships>
        "#;
        writer.write_all(xml)?;
        let mut rid = 0;
        for sf in self.sheets.iter() {
            let str = format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.xml\"/>", sf.id + 2, sf.id);
            writer.write_all(str.as_bytes())?;
            rid = sf.id + 2;
        }
        let ss = format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>", rid + 1);
        writer.write_all(ss.as_bytes())?;
        writer.write_all(tail)
    }

    fn create_xl_theme(writer: &mut dyn Write) -> Result<()> {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
    <a:themeElements>
        <a:clrScheme name="Office">
            <a:dk1>
                <a:sysClr val="windowText" lastClr="000000"/>
            </a:dk1>
            <a:lt1>
                <a:sysClr val="window" lastClr="FFFFFF"/>
            </a:lt1>
            <a:dk2>
                <a:srgbClr val="1F497D"/>
            </a:dk2>
            <a:lt2>
                <a:srgbClr val="EEECE1"/>
            </a:lt2>
            <a:accent1>
                <a:srgbClr val="4F81BD"/>
            </a:accent1>
            <a:accent2>
                <a:srgbClr val="C0504D"/>
            </a:accent2>
            <a:accent3>
                <a:srgbClr val="9BBB59"/>
            </a:accent3>
            <a:accent4>
                <a:srgbClr val="8064A2"/>
            </a:accent4>
            <a:accent5>
                <a:srgbClr val="4BACC6"/>
            </a:accent5>
            <a:accent6>
                <a:srgbClr val="F79646"/>
            </a:accent6>
            <a:hlink>
                <a:srgbClr val="0000FF"/>
            </a:hlink>
            <a:folHlink>
                <a:srgbClr val="800080"/>
            </a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Office">
            <a:majorFont>
                <a:latin typeface="Cambria"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
                <a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
                <a:font script="Hang" typeface="맑은 고딕"/>
                <a:font script="Hans" typeface="宋体"/>
                <a:font script="Hant" typeface="新細明體"/>
                <a:font script="Arab" typeface="Times New Roman"/>
                <a:font script="Hebr" typeface="Times New Roman"/>
                <a:font script="Thai" typeface="Tahoma"/>
                <a:font script="Ethi" typeface="Nyala"/>
                <a:font script="Beng" typeface="Vrinda"/>
                <a:font script="Gujr" typeface="Shruti"/>
                <a:font script="Khmr" typeface="MoolBoran"/>
                <a:font script="Knda" typeface="Tunga"/>
                <a:font script="Guru" typeface="Raavi"/>
                <a:font script="Cans" typeface="Euphemia"/>
                <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                <a:font script="Thaa" typeface="MV Boli"/>
                <a:font script="Deva" typeface="Mangal"/>
                <a:font script="Telu" typeface="Gautami"/>
                <a:font script="Taml" typeface="Latha"/>
                <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                <a:font script="Orya" typeface="Kalinga"/>
                <a:font script="Mlym" typeface="Kartika"/>
                <a:font script="Laoo" typeface="DokChampa"/>
                <a:font script="Sinh" typeface="Iskoola Pota"/>
                <a:font script="Mong" typeface="Mongolian Baiti"/>
                <a:font script="Viet" typeface="Times New Roman"/>
                <a:font script="Uigh" typeface="Microsoft Uighur"/>
                <a:font script="Geor" typeface="Sylfaen"/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="Calibri"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
                <a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
                <a:font script="Hang" typeface="맑은 고딕"/>
                <a:font script="Hans" typeface="宋体"/>
                <a:font script="Hant" typeface="新細明體"/>
                <a:font script="Arab" typeface="Arial"/>
                <a:font script="Hebr" typeface="Arial"/>
                <a:font script="Thai" typeface="Tahoma"/>
                <a:font script="Ethi" typeface="Nyala"/>
                <a:font script="Beng" typeface="Vrinda"/>
                <a:font script="Gujr" typeface="Shruti"/>
                <a:font script="Khmr" typeface="DaunPenh"/>
                <a:font script="Knda" typeface="Tunga"/>
                <a:font script="Guru" typeface="Raavi"/>
                <a:font script="Cans" typeface="Euphemia"/>
                <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                <a:font script="Thaa" typeface="MV Boli"/>
                <a:font script="Deva" typeface="Mangal"/>
                <a:font script="Telu" typeface="Gautami"/>
                <a:font script="Taml" typeface="Latha"/>
                <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                <a:font script="Orya" typeface="Kalinga"/>
                <a:font script="Mlym" typeface="Kartika"/>
                <a:font script="Laoo" typeface="DokChampa"/>
                <a:font script="Sinh" typeface="Iskoola Pota"/>
                <a:font script="Mong" typeface="Mongolian Baiti"/>
                <a:font script="Viet" typeface="Arial"/>
                <a:font script="Uigh" typeface="Microsoft Uighur"/>
                <a:font script="Geor" typeface="Sylfaen"/>
            </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
            <a:fillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"/>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="50000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="35000">
                            <a:schemeClr val="phClr">
                                <a:tint val="37000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:tint val="15000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="16200000" scaled="1"/>
                </a:gradFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="100000"/>
                                <a:shade val="100000"/>
                                <a:satMod val="130000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:tint val="50000"/>
                                <a:shade val="100000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="16200000" scaled="0"/>
                </a:gradFill>
            </a:fillStyleLst>
            <a:lnStyleLst>
                <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr">
                            <a:shade val="95000"/>
                            <a:satMod val="105000"/>
                        </a:schemeClr>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                </a:ln>
                <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                </a:ln>
                <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                </a:ln>
            </a:lnStyleLst>
            <a:effectStyleLst>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="38000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="35000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="35000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                    <a:scene3d>
                        <a:camera prst="orthographicFront">
                            <a:rot lat="0" lon="0" rev="0"/>
                        </a:camera>
                        <a:lightRig rig="threePt" dir="t">
                            <a:rot lat="0" lon="0" rev="1200000"/>
                        </a:lightRig>
                    </a:scene3d>
                    <a:sp3d>
                        <a:bevelT w="63500" h="25400"/>
                    </a:sp3d>
                </a:effectStyle>
            </a:effectStyleLst>
            <a:bgFillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"/>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="40000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="40000">
                            <a:schemeClr val="phClr">
                                <a:tint val="45000"/>
                                <a:shade val="99000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:shade val="20000"/>
                                <a:satMod val="255000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:path path="circle">
                        <a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
                    </a:path>
                </a:gradFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="80000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:shade val="30000"/>
                                <a:satMod val="200000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:path path="circle">
                        <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
                    </a:path>
                </a:gradFill>
            </a:bgFillStyleLst>
        </a:fmtScheme>
    </a:themeElements>
    <a:objectDefaults>
        <a:spDef>
            <a:spPr/>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:style>
                <a:lnRef idx="1">
                    <a:schemeClr val="accent1"/>
                </a:lnRef>
                <a:fillRef idx="3">
                    <a:schemeClr val="accent1"/>
                </a:fillRef>
                <a:effectRef idx="2">
                    <a:schemeClr val="accent1"/>
                </a:effectRef>
                <a:fontRef idx="minor">
                    <a:schemeClr val="lt1"/>
                </a:fontRef>
            </a:style>
        </a:spDef>
        <a:lnDef>
            <a:spPr/>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:style>
                <a:lnRef idx="2">
                    <a:schemeClr val="accent1"/>
                </a:lnRef>
                <a:fillRef idx="0">
                    <a:schemeClr val="accent1"/>
                </a:fillRef>
                <a:effectRef idx="1">
                    <a:schemeClr val="accent1"/>
                </a:effectRef>
                <a:fontRef idx="minor">
                    <a:schemeClr val="tx1"/>
                </a:fontRef>
            </a:style>
        </a:lnDef>
    </a:objectDefaults>
    <a:extraClrSchemeLst/>
</a:theme>
        "#;
        writer.write_all(xml.as_bytes())
    }

    /*
        fn create_sample_sheet(writer: &mut Write) -> Result<()> {
            let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <dimension ref="A1:E4"/>
        <cols>
            <col min="1" max="1" width="6.7109375" customWidth="1"/>
            <col min="2" max="2" width="7.7109375" customWidth="1"/>
            <col min="3" max="3" width="10.7109375" customWidth="1"/>
            <col min="4" max="4" width="20.7109375" customWidth="1"/>
            <col min="5" max="5" width="20.7109375" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1">
                <c r="A1">
                    <v>1</v>
                </c>
                <c r="B1">
                    <v>2</v>
                </c>
                <c r="C1">
                    <v>3</v>
                </c>
            </row>
            <row r="2">
                <c r="A2" t="b">
                    <v>1</v>
                </c>
                <c r="B2" t="b">
                    <v>0</v>
                </c>
                <c r="D2" t="str">
                    <v>sheetjs</v>
                </c>
            </row>
            <row r="3">
                <c r="A3" t="str">
                    <v>foo</v>
                </c>
                <c r="B3" t="str">
                    <v>bar</v>
                </c>
                <c r="C3" s="1">
                    <v>41689.604166666664</v>
                </c>
                <c r="D3" t="str">
                    <v>0.3</v>
                </c>
            </row>
            <row r="4">
                <c r="A4" t="str">
                    <v>baz2222</v>
                </c>
                <c r="C4" t="str">
                    <v>qux姓名</v>
                </c>
            </row>
            <row r="5">
                <c r="A5" t="str">
                    <v>baz2222</v>
                </c>
                <c r="E5" t="str">
                    <v>qux姓名</v>
                </c>
            </row>
        </sheetData>
    </worksheet>
            "#;
            writer.write_all(xml.as_bytes())
        }
    */
}

impl Drop for Workbook {
    fn drop(&mut self) {
        if !self.saved && self.xlsx_file.is_some() {
            self.close().expect("Workbook saving error");
        }
    }
}
