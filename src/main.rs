
mod init_data;
mod orders_plan;
mod xls_matrix;
mod xlsx;

use std::error::Error;
use std::{env};
use std::collections::HashMap;
use std::slice::SliceIndex;
use calamine::{DataType, Reader};
use chrono::{Datelike, Duration, NaiveDate, Utc, Weekday};
use rust_decimal::Decimal;
use rust_decimal::prelude::{FromPrimitive, ToPrimitive, Zero};
use rust_xlsxwriter::{ColNum, Color, RowNum, Workbook};

use crate::init_data::{InitialData, DeliveryTime, PurchaseOrder, PurchasePlanItem, Specification, SpecificationItem};
use crate::orders_plan::{calculate_stocks_plan, calculate_need_for_materials, MaterialInfo};
use crate::xls_matrix::{XlsCell, XlsCellFormat, XlsCellValue, XlsMatrix};
use crate::xlsx::*;


#[inline]
fn check_date(date: &NaiveDate, file_name: &String) {
    if date.weekday() != Weekday::Mon {
        panic!("{} не понедельник, в файле: {}",date, file_name)
    }
}

#[inline]
fn collect_dates_and_materials(stocks_plan: &HashMap<(NaiveDate, &String), Decimal>) -> (Vec<NaiveDate>, Vec<String>) {
    let mut dates = vec![];
    let mut materials = vec![];

    for item in stocks_plan.iter() {
        if !dates.contains(&item.0.0) {
            dates.push(item.0.0);
        }
        if !materials.contains(item.0.1) {
            materials.push(item.0.1.clone());
        }
    }
    dates.sort();
    materials.sort();
    (dates, materials)
}

fn read_sp_material_names() -> Vec<String> {
    let mut items = vec![];
    for sp in read_specifications() {
        for spi in sp.items {
            if !items.contains(&spi.material_name) {
                items.push(spi.material_name.clone())
            }
        }
    }
    items
}

fn correct_stocks_file()-> Result<(), Box<dyn Error>> {
    let stocks = read_stocks();

    let mut result: HashMap<String, Decimal> = HashMap::new();

    let names = read_sp_material_names();
    for mn in names {
        for smi in stocks.iter() {
            if mn.trim() == smi.material.trim() {
                let mut value = Decimal::zero();
                if result.contains_key(&mn) {
                    value = result.get(&mn).unwrap().clone();// value + smi.qty;
                }
                result.insert(smi.material.clone(), value+smi.qty);
            }
        }
    }


    let mut workbook = Workbook::new();
    let mut _worksheet = workbook.add_worksheet().set_name("стр")?;

    for (i,(key, val)) in result.iter().enumerate() {
        _worksheet.write((i + 1) as RowNum, 1 as ColNum, key).expect("TODO: panic message");
        _worksheet.write((i + 1) as RowNum, 2 as ColNum, val.to_f64()).expect("TODO: panic message");
    }



    _worksheet.autofit();

    workbook.save(format!("{}/{}",env::current_exe()?.parent().unwrap().to_str().unwrap(), "Остатки (кор.).xlsx"))?;
    println!("Создание файла завершено, смотрите файл \"{}\"","Остатки (кор.).xlsx");



    Ok(())
}

fn create_empty_stocks() -> Result<(), Box<dyn Error>>{
    println!("Создание файла остатков");

    let mut items = read_sp_material_names();

    let mut workbook = Workbook::new();
    let mut _worksheet = workbook.add_worksheet().set_name("стр")?;

    for i in 0..items.len(){
        _worksheet.write((i + 1) as RowNum, 1 as ColNum, &items[i]).expect("TODO: panic message");

    }


    _worksheet.autofit();

    workbook.save(format!("{}/{}",env::current_exe()?.parent().unwrap().to_str().unwrap(), "Остатки (авто).xlsx"))?;
    println!("Создание файла завершено, смотрите файл \"{}\"","Остатки (авто).xlsx");
    Ok(())
}

fn  main() -> Result<(), Box<dyn Error>> {

    for argument in env::args() {
        if argument.eq("-correct_stocks") {
            correct_stocks_file()?;
            return Ok(());
        }
        if argument.eq("-empty_stocks") {
            create_empty_stocks()?;
            return Ok(())
        }
    }

    println!("Расчет состояния заказов...");

    let init_data = InitialData {
        purchase_orders: read_purchase_orders(),
        delivery_times: read_delivery_time_items(),
        specifications: read_specifications(),
        stocks: read_stocks(),
        purchase_plan_items: read_purchase_plan_items()
    };

    let need_for_materials = calculate_need_for_materials(&init_data)?;
    let stocks_plan = calculate_stocks_plan(&need_for_materials, &init_data);
    let (dates, materials) = collect_dates_and_materials(&stocks_plan);


    let mut matrix = XlsMatrix::new();
    let mut dates_row: Vec<XlsCell> = Vec::new();

    dates_row.push(XlsCell{ cell_value: XlsCellValue::None, formats: vec![] });
    for date in dates.iter() {
        dates_row.push(XlsCell{ cell_value: XlsCellValue::Date(date.clone()), formats: vec![] })
    }
    matrix.rows.push(dates_row);
    for m in materials.iter() {
        matrix.rows.push(vec![XlsCell{ cell_value: XlsCellValue::String(m.clone()), formats: vec![] }])
    }

    for row_num in 0..materials.len() {
        let mut cur_value = Decimal::zero();

        for col_num in 0..dates.len() {
            let acc_val = stocks_plan.get(&(dates[col_num], &materials[row_num]));

            if let Some(cell_val) = acc_val {
                cur_value += cell_val;
                matrix.rows[row_num+1].push(XlsCell{ cell_value: XlsCellValue::Decimal(cur_value), formats: vec![] });
            } else {
                matrix.rows[row_num+1].push(XlsCell{ cell_value: XlsCellValue::Decimal(cur_value), formats: vec![] });
            }
        }
    }

    for cell in &mut matrix.rows[0] {
        cell.formats.push(XlsCellFormat::Bordered);
        cell.formats.push(XlsCellFormat::NumFormat("dd.mm.yyyy"));
    }

    let now = Utc::now().naive_utc().date();
    let mut now_index = 0;
    for d in 0..dates.len() {
        if dates[d]<=now && dates[d]+Duration::days(7) > now {
            now_index = d;
            break;
        }
    };

    for row in &mut matrix.rows {
        row[now_index].formats.push(XlsCellFormat::Background(Color::RGB(0xEEEEEE)))
    }
    for row_num in 0..matrix.rows.len() {
        let row = &mut matrix.rows[row_num];
        for col_num in 0..row.len() {
            let mut cell = &mut row[col_num];
            cell.formats.push(XlsCellFormat::Bordered);
            match cell.cell_value {
                XlsCellValue::Decimal(d) => {
                    if(col_num<=now_index+init_data.get_delivery_weeks(&materials[row_num-1])?){
                        if d<Decimal::zero() {
                            cell.formats.push(XlsCellFormat::FontColor(Color::Red));
                        } else {
                            cell.formats.push(XlsCellFormat::FontColor(Color::Green));
                        }
                    } else {
                        cell.formats.push(XlsCellFormat::FontColor(Color::RGB(0xDDDDDD)));
                    }
                }
                _ => ()
            }
        }
    }

    let mut workbook = Workbook::new();
    let mut _worksheet = workbook.add_worksheet().set_name("Состояние заказов")?;
    matrix.write_to_worksheet(&mut _worksheet);
    _worksheet.autofit();

    workbook.save(format!("{}/{}",env::current_exe()?.parent().unwrap().to_str().unwrap(), "Состояние заказов.xlsx"))?;
    println!("Расчет завершен, смотрите файл \"{}\"","Состояние заказов.xlsx");
    Ok(())
}
