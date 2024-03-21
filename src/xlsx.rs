use std::{env, fs};
use calamine::{Data, DataType, open_workbook, Range, Reader, Xlsx};
use chrono::NaiveDate;
use rust_decimal::Decimal;
use rust_decimal::prelude::FromPrimitive;
use crate::check_date;
use crate::init_data::{DeliveryTime, PurchaseOrder, PurchasePlanItem, Specification, SpecificationItem};
use crate::orders_plan::MaterialInfo;

fn get_template_path(file_name: &str) -> String {
    format!("{}/{}",env::current_exe().unwrap().parent().unwrap().join("Исходные данные").to_str().unwrap(), file_name)
}

/// F - function for read data. args: row, col, &Range
///
/// R - function for determinate range, returns tuple (row, height, col, width) of range.
///
/// ### Example
///
/// ```
///
/// fn read_data()
/// {
///     let mut data = Data::new();
///     read_xlsx("file.xlsx".to_string(),
///     |row, col, range| {
///         data.push(range.get(row, col));
///     },
///     |range| {
///         (0, range.height(), 0, range.width())
///     }
/// }
///
/// ```
pub fn read_xlsx<R, B>(path: String, mut read_func: R, bounds_func: B)
    where
        R: FnMut(usize, usize,  &Range<Data>),
        B: Fn(&Range<Data>) -> (usize, usize, usize, usize)
{
    let mut workbook: Xlsx<_> = open_workbook(path).unwrap();
    let binding = workbook.worksheets();
    let worksheet = binding.get(0).unwrap();
    let range = &(*worksheet).1;
    let b = bounds_func(range);
    for row in b.0..b.1 {
        for col in b.2..b.3 {
            read_func(row, col, range);
        }
    }
}

pub fn read_purchase_plan_items() -> Vec<PurchasePlanItem> {
    let mut purchase_plan = vec![];
    let p = get_template_path("План обеспечения.xlsx");
    read_xlsx(p, |row, col, range| {
        let plan_item = PurchasePlanItem {
            product_name: range.get((row, 0usize)).unwrap().as_string().unwrap(),
            date: range.get((0usize, col)).unwrap().as_date().unwrap(),
            qty: Decimal::from_f64(range.get((row, col)).unwrap().as_f64().expect("Не получилось прочитать значение ячейки как число в файле План обеспечения")).unwrap(),
        };
        purchase_plan.push(plan_item);
    }, |range| {
        (1, range.height(), 1, range.width())
    });
    return purchase_plan
}

pub fn read_delivery_time_items() -> Vec<DeliveryTime> {
    let mut delivery_times = vec![];
    read_xlsx(get_template_path("Сроки доставки.xlsx"), |row, _, range| {
        let delivery_time_item = DeliveryTime {
            material_name:  range.get((row, 0usize)).unwrap().as_string().unwrap(),
            weeks:          range.get((row, 1usize)).unwrap().as_i64().unwrap() as u32,
        };
        delivery_times.push(delivery_time_item);
    }, |range|{
        (0,range.height(),0,1)
    });
    return delivery_times
}

pub fn read_decimal(data: &Data) -> Option<Decimal> {
    match data {
        Data::Int(i) => { Some(Decimal::from_i64(*i).unwrap()) }
        Data::Float(f) => { Some(Decimal::from_f64(*f).unwrap()) }
        _ => None
    }
}

pub fn read_stocks() -> Vec<MaterialInfo> {
    let mut stocks = vec![];
    read_xlsx(get_template_path("Остатки.xlsx"), |row, col, range|{
        let qty_opt = range.get((row, col));
        if qty_opt.is_some() {
            let dec = read_decimal(qty_opt.unwrap());

            if dec.is_some() {
                let mi = MaterialInfo{
                    date: range.get((0usize, col)).unwrap().as_date().unwrap(),
                    material: range.get((row, 0usize)).unwrap().as_string().unwrap(),
                    qty: dec.unwrap()
                };
                check_date(&mi.date,&"Остатки.xlsx".to_string());
                stocks.push(mi);

            }
        }
    }, |range|{
        (1, range.height(), 1, range.width())
    });
    stocks
}

pub fn read_specifications() -> Vec<Specification> {
    let mut result = vec![];

    for entry in fs::read_dir(get_template_path("Спецификации")).unwrap() {
        let dir_e = entry.unwrap();
        let path = dir_e.path().into_os_string().to_str().unwrap().to_string();
        let file_name = dir_e.file_name().to_str().unwrap().to_string();
        let name_parts = file_name.split(".").collect::<Vec<&str>>()[0].split("_").collect::<Vec<&str>>();
        if name_parts.len() != 2 {
            panic!("{} invalid file name", path)
        }
        let name = name_parts[0].to_string();
        let date = NaiveDate::parse_from_str(name_parts[1],"%Y%m%d").expect(&format!("Не удалось извлечь дату из файла спецификации ({})", file_name));
        let mut sp = Specification{
            product_name: name,
            date_from: date,
            items: vec![],
        };

        read_xlsx(path.to_string(), |row, _, range|{
            let sp_item = SpecificationItem{
                material_name: range.get((row,0usize)).unwrap().as_string().unwrap(),
                qty: read_decimal(range.get((row, 1usize)).unwrap()).unwrap(),
            };
            sp.items.push(sp_item);
        }, |range|{
            (0,range.height(),0,1)
        });
        result.push(sp);
    }
    result
}

pub fn read_purchase_orders() -> Vec<PurchaseOrder> {
    let mut result = vec![];
    for dir_e in fs::read_dir(get_template_path("Заказы поставщикам")).unwrap() {
        let dir_e = dir_e.unwrap();
        let path = dir_e.path().into_os_string().to_str().unwrap().to_string();
        let file_name = dir_e.file_name().to_str().unwrap().to_string();
        let name_parts = file_name.split(".").collect::<Vec<&str>>()[0].split("_").collect::<Vec<&str>>();
        if name_parts.len() != 2 {
            panic!("{} invalid file name", path)
        }
        let name = name_parts[1].to_string();
        let mut po = PurchaseOrder {
            name,
            items: vec![],
        };
        read_xlsx(path.to_string(), |row, col, range| {
            let qty_o = read_decimal(range.get((row, col)).unwrap());
            if qty_o.is_some() {
                let po_item = MaterialInfo {
                    material: range.get((row, 0usize)).unwrap().as_string().unwrap(),
                    date: range.get((0usize, col)).unwrap().as_date().unwrap(),
                    qty: qty_o.unwrap(),
                };
                check_date(&po_item.date,&file_name);
                po.items.push(po_item);
            }
        }, |range| {
            (3, range.height(), 1, range.width())
        });
        result.push(po);
    }
    result
}