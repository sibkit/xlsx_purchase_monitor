use chrono::NaiveDate;
use rust_decimal::Decimal;
use crate::orders_plan::MaterialInfo;

#[derive(Debug)]
pub struct PurchasePlanItem {
    pub(crate) product_name: String,
    pub(crate) date: NaiveDate,
    pub(crate) qty: Decimal
}
#[derive(Debug)]
pub struct Specification {
    pub(crate) product_name: String,
    pub(crate) date_from: NaiveDate,
    pub(crate) items: Vec<SpecificationItem>
}

#[derive(Debug)]
pub struct SpecificationItem{
    pub(crate) material_name: String,
    pub(crate) qty: Decimal
}

#[derive(Debug)]
pub struct PurchaseOrder{
    pub(crate) name: String,
    pub(crate) items: Vec<MaterialInfo>
}

#[derive(Debug)]
pub struct DeliveryTime {
    pub(crate) material_name: String,
    pub(crate) weeks: u32
}

pub struct InitialData {
    pub(crate) purchase_orders: Vec<PurchaseOrder>,
    pub(crate) delivery_times: Vec<DeliveryTime>,
    pub(crate) specifications: Vec<Specification>,
    pub(crate) stocks: Vec<MaterialInfo>,
    pub(crate) purchase_plan_items: Vec<PurchasePlanItem>
}

impl InitialData {
    pub fn get_delivery_weeks(&self, material_name: &str) -> Result<usize,String> {
        for dt in &self.delivery_times {
            if dt.material_name == material_name {
                return Ok(dt.weeks as usize);
            }
        }
        Err(format!(r#"Не найден срок доставки для материала '{}'"#,material_name))
    }
}
