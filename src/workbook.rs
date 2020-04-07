use std::fs::{self, File};
use std::io::*;
use std::path::*;

use utilities::zip_files;

use super::{Sheet, SheetWriter};

#[derive(Default)]
pub struct Workbook {
    pub file: String,
    tmp_dir: String,
    max_sheet_index: usize,
    shared_strings: SharedStrings,
    sheets: Vec<SheetRef>,
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
    pub fn register(&mut self, val: &String) -> crate::CellValue {
        self.add_count();
        
        match self.strings.iter().position(|v|v==val) {
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
    pub fn create(xlsx_file: &str) -> Workbook {
        let target_dir = format!("{}_tmp", &xlsx_file);

        let workbook = Workbook {
            file: xlsx_file.to_owned(),
            tmp_dir: target_dir,
            max_sheet_index: 0,
            shared_strings: SharedStrings::new(),
            ..Default::default()
        };

        if Path::new(&workbook.tmp_dir).is_dir() {
            fs::remove_dir_all(&workbook.tmp_dir).unwrap();
        }
        fs::create_dir(&workbook.tmp_dir).unwrap();

        workbook
    }
    /// Creates a workbook not using shared strings
    pub fn create_simple(xlsx_file: &str) -> Workbook {
        let target_dir = format!("{}_tmp", &xlsx_file);

        let workbook = Workbook {
            file: xlsx_file.to_owned(),
            tmp_dir: target_dir,
            max_sheet_index: 0,
            shared_strings: SharedStrings::new_unused(),
            ..Default::default()
        };

        if Path::new(&workbook.tmp_dir).is_dir() {
            fs::remove_dir_all(&workbook.tmp_dir).unwrap();
        }
        fs::create_dir(&workbook.tmp_dir).unwrap();

        workbook
    }

    pub fn create_sheet(&mut self, sheet_name: &str) -> Sheet {
        self.max_sheet_index += 1;

        self.sheets.push(SheetRef {
            id: self.max_sheet_index,
            name: crate::validate_name(sheet_name), //sheet_name.to_owned(),
        });
        Sheet::new(self.max_sheet_index, sheet_name)
    }

    pub fn close(&mut self) -> Result<()> {
        self.create_files().expect("Create files error!");

        // zip the files in the working directory
        zip_files(&self.tmp_dir, &self.file)?;

        fs::remove_dir_all(&self.tmp_dir)
    }

    fn create_files(&mut self) -> Result<()> {
        let mut root = PathBuf::from(&self.tmp_dir);

        // [Content_Types].xml
        root.push("[Content_Types].xml");
        let writer = &mut File::create(root.as_path()).unwrap();
        self.create_content_types(writer)?;
        root.pop();

        // _rels/.rels
        root.push("_rels");
        fs::create_dir(root.as_path())?;
        root.push(".rels");
        let writer = &mut File::create(root.as_path())?;
        Self::create_rels(writer)?;
        root.pop();
        root.pop();

        // docProps
        root.push("docProps");
        fs::create_dir(root.as_path())?;
        root.push("app.xml");
        let writer = &mut File::create(root.as_path())?;
        self.create_app(writer)?;
        root.pop();
        root.push("core.xml");
        let writer = &mut File::create(root.as_path())?;
        Self::create_core(writer)?;
        root.pop();
        root.pop();

        // xl
        root.push("xl");
        if !root.as_path().exists() {
            fs::create_dir(root.as_path()).unwrap();
        }
        root.push("styles.xml");
        let writer = &mut File::create(root.as_path())?;
        Self::create_styles(writer)?;
        root.pop();
        //if self.shared_strings.used(){
        root.push("sharedStrings.xml");
        let writer = &mut File::create(root.as_path())?;
        self.create_shared_strings(writer)?;
        root.pop();
        //}
        root.push("workbook.xml");
        let writer = &mut File::create(root.as_path())?;
        self.create_workbook(writer)?;
        root.pop();

        // xl/_rels
        root.push("_rels");
        fs::create_dir(root.as_path())?;
        root.push("workbook.xml.rels");
        let writer = &mut File::create(root.as_path())?;
        self.create_xl_rels(writer)?;
        root.pop();
        root.pop();

        // xl/theme
        root.push("theme");
        fs::create_dir(root.as_path())?;
        root.push("theme1.xml");
        let writer = &mut File::create(root.as_path())?;
        Self::create_xl_theme(writer)?;
        root.pop();
        root.pop();

        /*
                // xl/worksheets
                root.push("worksheets");
                fs::create_dir(root.as_path())?;
                root.push("sheet1.xml");
                let writer = &mut File::create(root.as_path())?;
                Self::create_sample_sheet(writer)?;
                root.pop();
                root.pop();
        */
        Ok(())
    }

    pub fn write_sheet<F>(&mut self, sheet: &mut Sheet, write_data: F) -> Result<()>
    where
        F: FnOnce(&mut SheetWriter<File>) -> Result<()> + Sized,
    {
        let mut root = PathBuf::from(&self.tmp_dir);
        root.push("xl");
        root.push("worksheets");
        fs::create_dir_all(root.as_path())?;
        root.push(format!("sheet{}.xml", sheet.id));
        let writer = &mut File::create(root.as_path()).unwrap();

        let sw = &mut SheetWriter::new(sheet, writer, &mut self.shared_strings);
        sw.write(write_data)
    }

    fn create_content_types(&mut self, writer: &mut dyn Write) -> Result<()> {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" xmlns:xsd="http://www.w3.org/2001/XMLSchema"
       xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/xl/workbook.xml"
              ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#;
        writer.write_all(xml)?;
        for x in 1..self.sheets.len() {
            let wb = format!(
                "<Override PartName=\"/xl/worksheets/sheet{}.xml\"
                    ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
                x
            );
            writer.write_all(wb.as_bytes())?;
        }

        let tail = br#"<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
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

    fn create_styles(writer: &mut dyn Write) -> Result<()> {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>
</styleSheet>
        "#;
        writer.write_all(xml)
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
            let space = match sf.trim().len() < sf.len() {
                true => format!("t xml:space=\"preserve\""),
                false => format!("t"),
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
