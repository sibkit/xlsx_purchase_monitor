use std::collections::HashMap;
use chrono::NaiveDate;
use rust_decimal::Decimal;
use crate::init_data::{InitialData, Specification};

fn find_specification(product_name: String, specifications: &Vec<Specification>) -> Result<&Specification,String> {
    for sp in specifications {
        if sp.product_name == product_name { return Ok(&sp) }
    }
    Err(format!("Спецификация для товара \"{}\" не найдена.", product_name))
}

pub fn calculate_need_for_materials(data_set: &InitialData) -> Result<Vec<MaterialInfo>, String> {
    let mut map: HashMap<(NaiveDate, String),Decimal> = HashMap::new();

    for ppi in &data_set.purchase_plan_items {
        let sp = find_specification(ppi.product_name.clone(),&data_set.specifications)?;
        for spi in &sp.items {
            let tuple = (ppi.date.clone(),spi.material_name.clone());
            if !map.contains_key(&tuple){
                map.insert(tuple, Decimal::from(-1)*ppi.qty*spi.qty);
            }
            else {
                let exist_val = map.get(&tuple).unwrap();
                map.insert(tuple,Decimal::from(-1)*ppi.qty*spi.qty + exist_val);
            }
        }
    }
    let mut result = vec![];
    for hmv in map.iter() {
        result.push(MaterialInfo {
            date: hmv.0.0.clone(),
            material: hmv.0.1.clone(),
            qty: hmv.1.clone(),
        })
    }
    Ok(result)
}

pub fn append_qty<'a, 'b, 'c>(map: &'a mut HashMap<(NaiveDate, &'c String),Decimal>, mis: &'b Vec<MaterialInfo>)
where 'b: 'a, 'b: 'c
{
    for mi in mis {
        let tuple = (mi.date.clone(),&mi.material);
        if !map.contains_key(&tuple){
            let ir = map.insert(tuple, mi.qty);
            if ir.is_some() {panic!("ir is some")}
        }
        else {
            let exist_val = map.get(&tuple).unwrap();
            let ir = map.insert(tuple, mi.qty + exist_val);
            if ir.is_none() {panic!("ir is none")}
        }
    }
}

pub fn calculate_stocks_plan<'a>(requirements: &'a Vec<MaterialInfo>, data_set: &'a InitialData) -> HashMap<(NaiveDate, &'a String), Decimal> {
    let mut map: HashMap<(NaiveDate, &String), Decimal> = HashMap::new();
    append_qty(&mut map, &data_set.stocks);
    append_qty(&mut map, &requirements);
    for po in &data_set.purchase_orders {
        append_qty(&mut map, &po.items)
    }
    map
}

#[derive(Debug)]
pub struct MaterialInfo {
    pub date: NaiveDate,
    pub material: String,
    pub qty: Decimal
}