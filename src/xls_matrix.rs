use chrono::NaiveDate;
use rust_decimal::Decimal;
use rust_decimal::prelude::ToPrimitive;
use rust_xlsxwriter::{ColNum, Color, Format, FormatBorder, IntoExcelData, RowNum, Worksheet};


pub enum XlsCellFormat {
    FontColor(Color),
    Background(Color),
    NumFormat(&'static str),
    Bordered
}

pub enum XlsCellValue {
    None,
    Decimal(Decimal),
    Date(NaiveDate),
    String(String)
}
pub struct XlsCell {
    pub cell_value: XlsCellValue,
    pub formats: Vec<XlsCellFormat>
}

pub struct XlsMatrix {
    pub rows: Vec<Vec<XlsCell>>
}

pub(crate) fn write_to_cell<T: IntoExcelData>(sheet: &mut Worksheet, row_num: usize, col_num: usize, value: T, formats: &Vec<XlsCellFormat>) {
    let mut format = Format::new();
    for f in formats {
        match f {
            XlsCellFormat::FontColor(color) => {format = format.set_font_color(*color)}
            XlsCellFormat::Background(color) => {format = format.set_background_color(*color)}
            XlsCellFormat::NumFormat(nf) => {format = format.set_num_format(*nf)}
            XlsCellFormat::Bordered => {format = format.set_border(FormatBorder::Thin)}
        }
    }
    sheet.write_with_format(row_num as RowNum, col_num as ColNum, value, &format).expect("Ошибка при записи ячейки");
}

impl XlsMatrix {

    pub fn new() -> Self {
        XlsMatrix{
            rows: vec![]
        }
    }

    pub fn write_to_worksheet(&self, sheet: &mut Worksheet) {
        for row_num in 0..self.rows.len() {
            for col_num in 0..self.rows[row_num].len() {
                let cell = &self.rows[row_num][col_num];

                match &cell.cell_value {
                    XlsCellValue::None => { write_to_cell(sheet, row_num+1, col_num+1, None::<String>, &cell.formats) }
                    XlsCellValue::Decimal(d) => { write_to_cell(sheet, row_num+1, col_num+1, d.to_f64(), &cell.formats) }
                    XlsCellValue::Date(d) => { write_to_cell(sheet, row_num+1, col_num+1, d, &cell.formats) }
                    XlsCellValue::String(s) => { write_to_cell(sheet, row_num+1, col_num+1, s, &cell.formats) }
                };
            }
        }
    }
}