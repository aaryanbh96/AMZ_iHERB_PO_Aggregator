#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use calamine::{open_workbook_auto, Reader, Data};
use csv::ReaderBuilder;
use indexmap::IndexMap;
use rust_xlsxwriter::*;
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};

const PREMIER_SKUS: &[&str] = &[
    "TOT91451", "TOT91471", "TOT91461",
    "TOT91411", "TOT91421", "TOT91491", "TOT91441",
];
const BIG_TICKET_THRESHOLD: f64 = 2000.0;

fn units_per_case(sku: &str) -> f64 {
    if PREMIER_SKUS.contains(&sku) { 6.0 } else { 12.0 }
}

#[derive(Serialize, Deserialize, Clone, Debug)]
pub struct AggRow {
    pub id: String,
    pub title: String,
    pub qty_ordered: f64,
    pub case_qty: Option<f64>,
    pub total_cost: f64,
}

#[derive(Serialize, Deserialize, Clone, Debug)]
pub struct CombinedRow {
    pub id: String,
    pub title: String,
    pub amz_qty: Option<f64>,
    pub amz_cost: Option<f64>,
    pub iherb_cases: Option<f64>,
    pub iherb_cost: Option<f64>,
    pub total_cases: Option<f64>,
    pub total_cost: f64,
    pub is_pack: bool,
}

#[derive(Serialize, Deserialize, Clone, Debug)]
pub struct PoDetailRow {
    pub po: String,
    pub date: String,
    pub ship_to: String,
    pub qty: f64,
    pub total: f64,
}

#[derive(Serialize, Deserialize, Clone, Debug)]
pub struct BigTicketSection {
    pub id: String,
    pub name: String,
    pub combined_total: f64,
    pub amz_rows: Vec<PoDetailRow>,
    pub iherb_rows: Vec<PoDetailRow>,
}

#[derive(Serialize, Deserialize, Debug)]
pub struct ProcessResult {
    pub rows: Vec<AggRow>,
    pub grand_total: f64,
    pub total_cases: f64,
    pub sku_count: usize,
}

#[derive(Serialize, Deserialize, Debug)]
pub struct CombinedResult {
    pub rows: Vec<CombinedRow>,
    pub grand_total: f64,
}

fn cell_str(cell: &Data) -> String {
    match cell {
        Data::String(s) => s.trim().to_string(),
        Data::Float(f)  => if f.fract() == 0.0 { format!("{:.0}", f) } else { f.to_string() },
        Data::Int(i)    => i.to_string(),
        Data::Bool(b)   => b.to_string(),
        _               => String::new(),
    }
}

fn cell_f64(cell: &Data) -> f64 {
    match cell {
        Data::Float(f)  => *f,
        Data::Int(i)    => *i as f64,
        Data::String(s) => s.trim().parse().unwrap_or(0.0),
        _               => 0.0,
    }
}

fn load_excel(path: &str) -> Result<(Vec<String>, Vec<Vec<Data>>), String> {
    let mut wb = open_workbook_auto(path).map_err(|e| format!("Cannot open: {e}"))?;
    let names = wb.sheet_names().to_vec();
    let sheet = wb.worksheet_range(&names[0]).map_err(|e| format!("Sheet error: {e}"))?;
    let mut iter = sheet.rows();
    let headers = match iter.next() {
        Some(r) => r.iter().map(cell_str).collect(),
        None    => return Err("File is empty".to_string()),
    };
    let rows: Vec<Vec<Data>> = iter.map(|r| r.to_vec()).collect();
    Ok((headers, rows))
}

fn load_csv(path: &str) -> Result<(Vec<String>, Vec<Vec<String>>), String> {
    let mut rdr = ReaderBuilder::new().has_headers(true).from_path(path)
        .map_err(|e| format!("Cannot open CSV: {e}"))?;
    let headers: Vec<String> = rdr.headers().map_err(|e| e.to_string())?
        .iter().map(|s| s.trim().to_string()).collect();
    let mut rows = Vec::new();
    for result in rdr.records() {
        let record = result.map_err(|e| e.to_string())?;
        rows.push(record.iter().map(|s| s.trim().to_string()).collect());
    }
    Ok((headers, rows))
}

fn col(headers: &[String], name: &str) -> Option<usize> {
    headers.iter().position(|h| h.trim().eq_ignore_ascii_case(name.trim()))
}

#[tauri::command]
fn process_amazon(path: String) -> Result<ProcessResult, String> {
    let (headers, rows) = load_excel(&path)?;
    let ia = col(&headers, "ASIN").ok_or("ASIN not found")?;
    let it = col(&headers, "Title").ok_or("Title not found")?;
    let iq = col(&headers, "Quantity Requested").ok_or("Quantity Requested not found")?;
    let ic = col(&headers, "Case Cost").ok_or("Case Cost not found")?;

    let mut map: IndexMap<String, (String, f64, f64)> = IndexMap::new();
    for row in &rows {
        let asin  = cell_str(&row[ia]);
        let title = cell_str(&row[it]);
        let qty   = cell_f64(&row[iq]);
        let cost  = cell_f64(&row[ic]);
        let entry = map.entry(asin).or_insert((title, 0.0, 0.0));
        entry.1 += qty;
        entry.2 += qty * cost;
    }

    let mut result: Vec<AggRow> = map.into_iter().map(|(id, (title, qty, total))| {
        AggRow { id, title, qty_ordered: qty, case_qty: None, total_cost: total }
    }).collect();
    result.sort_by(|a, b| b.total_cost.partial_cmp(&a.total_cost).unwrap());
    let grand_total = result.iter().map(|r| r.total_cost).sum();
    let total_cases = result.iter().map(|r| r.qty_ordered).sum();
    let sku_count   = result.len();
    Ok(ProcessResult { rows: result, grand_total, total_cases, sku_count })
}

#[tauri::command]
fn process_iherb(path: String) -> Result<ProcessResult, String> {
    let (headers, rows) = load_csv(&path)?;
    let is = col(&headers, "Buyers Catalog or Stock Keeping #").ok_or("SKU not found")?;
    let it = col(&headers, "Product/Item Description").ok_or("Title not found")?;
    let iq = col(&headers, "Qty Ordered").ok_or("Qty Ordered not found")?;
    let iu = col(&headers, "Unit Price").ok_or("Unit Price not found")?;

    let mut map: IndexMap<String, (String, f64, f64, f64)> = IndexMap::new();
    for row in &rows {
        let sku = row.get(is).cloned().unwrap_or_default();
        if sku.is_empty() || sku == "nan" { continue; }
        let title = row.get(it).cloned().unwrap_or_default();
        let qty: f64  = row.get(iq).and_then(|v| v.parse().ok()).unwrap_or(0.0);
        let unit: f64 = row.get(iu).and_then(|v| v.parse().ok()).unwrap_or(0.0);
        let upc   = units_per_case(&sku);
        let entry = map.entry(sku).or_insert((title, 0.0, 0.0, upc));
        entry.1 += qty;
        entry.2 += qty / upc;
        entry.3 += qty * unit;
    }

    let mut result: Vec<AggRow> = map.into_iter().map(|(id, (title, qty, cases, total))| {
        AggRow { id, title, qty_ordered: qty, case_qty: Some(cases), total_cost: total }
    }).collect();
    result.sort_by(|a, b| b.total_cost.partial_cmp(&a.total_cost).unwrap());
    let grand_total = result.iter().map(|r| r.total_cost).sum();
    let total_cases = result.iter().filter_map(|r| r.case_qty).sum();
    let sku_count   = result.len();
    Ok(ProcessResult { rows: result, grand_total, total_cases, sku_count })
}

struct Mapping { sku: String, asin: String, name: String, is_pack: bool }

fn load_mapping(path: &str) -> Result<Vec<Mapping>, String> {
    let (headers, rows) = load_excel(path)?;
    let is  = col(&headers, "Buyers Catalog or Stock Keeping #").ok_or("SKU not found")?;
    let ia  = col(&headers, "ASIN").ok_or("ASIN not found")?;
    let iname = col(&headers, "Product Description").ok_or("Product Description not found")?;
    let ip  = col(&headers, "Pack?");
    let mut result = Vec::new();
    for row in &rows {
        let sku  = cell_str(&row[is]);
        let asin = cell_str(&row[ia]);
        let name = cell_str(&row[iname]);
        if sku.is_empty() && asin.is_empty() { continue; }
        let pack = ip.map(|i| cell_str(&row[i])).unwrap_or_default();
        result.push(Mapping { sku, asin, name, is_pack: pack.to_lowercase().contains("pack") });
    }
    Ok(result)
}

fn agg_amazon(paths: &[String]) -> Result<HashMap<String, (String, f64, f64)>, String> {
    let mut map: HashMap<String, (String, f64, f64)> = HashMap::new();
    for path in paths {
        let (headers, rows) = load_excel(path)?;
        let ia = col(&headers, "ASIN").ok_or("ASIN not found")?;
        let it = col(&headers, "Title").ok_or("Title not found")?;
        let iq = col(&headers, "Quantity Requested").ok_or("Qty not found")?;
        let ic = col(&headers, "Case Cost").ok_or("Case Cost not found")?;
        for row in &rows {
            let asin = cell_str(&row[ia]);
            if asin.is_empty() { continue; }
            let qty  = cell_f64(&row[iq]);
            let cost = cell_f64(&row[ic]);
            let entry = map.entry(asin).or_insert((cell_str(&row[it]), 0.0, 0.0));
            entry.1 += qty; entry.2 += qty * cost;
        }
    }
    Ok(map)
}

fn agg_iherb(paths: &[String]) -> Result<HashMap<String, (String, f64, f64)>, String> {
    let mut map: HashMap<String, (String, f64, f64)> = HashMap::new();
    for path in paths {
        let (headers, rows) = load_csv(path)?;
        let is = col(&headers, "Buyers Catalog or Stock Keeping #").ok_or("SKU not found")?;
        let it = col(&headers, "Product/Item Description").ok_or("Title not found")?;
        let iq = col(&headers, "Qty Ordered").ok_or("Qty not found")?;
        let iu = col(&headers, "Unit Price").ok_or("Unit Price not found")?;
        for row in &rows {
            let sku = row.get(is).cloned().unwrap_or_default();
            if sku.is_empty() || sku == "nan" { continue; }
            let qty: f64  = row.get(iq).and_then(|v| v.parse().ok()).unwrap_or(0.0);
            let unit: f64 = row.get(iu).and_then(|v| v.parse().ok()).unwrap_or(0.0);
            let upc = units_per_case(&sku);
            let entry = map.entry(sku).or_insert((row.get(it).cloned().unwrap_or_default(), 0.0, 0.0));
            entry.1 += qty / upc; entry.2 += qty * unit;
        }
    }
    Ok(map)
}

#[tauri::command]
fn build_combined(amz_paths: Vec<String>, iherb_paths: Vec<String>, mapping_path: String) -> Result<CombinedResult, String> {
    let mapping = load_mapping(&mapping_path)?;
    let amz_agg = agg_amazon(&amz_paths)?;
    let ih_agg  = agg_iherb(&iherb_paths)?;
    let mut rows: Vec<CombinedRow> = Vec::new();
    let mut used_asins: HashSet<String> = HashSet::new();
    let mut used_skus:  HashSet<String> = HashSet::new();

    for m in &mapping {
        let ad = if !m.asin.is_empty() { amz_agg.get(&m.asin) } else { None };
        let id = ih_agg.get(&m.sku);
        if ad.is_none() && id.is_none() { continue; }
        let amz_qty     = ad.map(|d| d.1);
        let amz_cost    = ad.map(|d| d.2);
        let iherb_cases = id.map(|d| d.1);
        let iherb_cost  = id.map(|d| d.2);
        let tc = amz_qty.unwrap_or(0.0) + iherb_cases.unwrap_or(0.0);
        let tt = amz_cost.unwrap_or(0.0) + iherb_cost.unwrap_or(0.0);
        rows.push(CombinedRow {
            id: if !m.asin.is_empty() { m.asin.clone() } else { m.sku.clone() },
            title: m.name.clone(), amz_qty, amz_cost, iherb_cases, iherb_cost,
            total_cases: if m.is_pack { None } else { Some(tc) },
            total_cost: tt, is_pack: m.is_pack,
        });
        if !m.asin.is_empty() { used_asins.insert(m.asin.clone()); }
        used_skus.insert(m.sku.clone());
    }
    for (asin, (title, qty, cost)) in &amz_agg {
        if used_asins.contains(asin) { continue; }
        rows.push(CombinedRow { id: asin.clone(), title: title.clone(), amz_qty: Some(*qty), amz_cost: Some(*cost), iherb_cases: None, iherb_cost: None, total_cases: Some(*qty), total_cost: *cost, is_pack: false });
        used_asins.insert(asin.clone());
    }
    for (sku, (title, cases, cost)) in &ih_agg {
        if used_skus.contains(sku) || sku == "nan" { continue; }
        rows.push(CombinedRow { id: sku.clone(), title: title.clone(), amz_qty: None, amz_cost: None, iherb_cases: Some(*cases), iherb_cost: Some(*cost), total_cases: Some(*cases), total_cost: *cost, is_pack: false });
        used_skus.insert(sku.clone());
    }
    rows.sort_by(|a, b| b.total_cost.partial_cmp(&a.total_cost).unwrap());
    let grand_total = rows.iter().map(|r| r.total_cost).sum();
    Ok(CombinedResult { rows, grand_total })
}

#[tauri::command]
fn build_bigticket(amz_paths: Vec<String>, iherb_paths: Vec<String>, mapping_path: String) -> Result<Vec<BigTicketSection>, String> {
    let mapping = load_mapping(&mapping_path)?;
    let mut amz_detail: HashMap<String, Vec<PoDetailRow>> = HashMap::new();
    for path in &amz_paths {
        let (headers, rows) = load_excel(path)?;
        let ia = col(&headers, "ASIN").ok_or("ASIN not found")?;
        let iq = col(&headers, "Quantity Requested").ok_or("Qty not found")?;
        let ic = col(&headers, "Case Cost").ok_or("Case Cost not found")?;
        let ipo = col(&headers, "PO");
        let iex = col(&headers, "Expected date");
        let ish = col(&headers, "Ship to location");
        for row in &rows {
            let asin = cell_str(&row[ia]);
            if asin.is_empty() { continue; }
            let qty = cell_f64(&row[iq]); let cost = cell_f64(&row[ic]);
            amz_detail.entry(asin).or_default().push(PoDetailRow {
                po:      ipo.map(|i| cell_str(&row[i])).unwrap_or_default(),
                date:    iex.map(|i| cell_str(&row[i])).unwrap_or_default(),
                ship_to: ish.map(|i| cell_str(&row[i])).unwrap_or_default(),
                qty, total: qty * cost,
            });
        }
    }
    let mut ih_detail: HashMap<String, Vec<PoDetailRow>> = HashMap::new();
    for path in &iherb_paths {
        let (headers, rows) = load_csv(path)?;
        let is  = col(&headers, "Buyers Catalog or Stock Keeping #").ok_or("SKU not found")?;
        let iq  = col(&headers, "Qty Ordered").ok_or("Qty not found")?;
        let iu  = col(&headers, "Unit Price").ok_or("Unit Price not found")?;
        let ipo = col(&headers, "PO Number");
        let idl = col(&headers, "Requested Delivery Date");
        let ist = col(&headers, "Ship To State");
        let mut state_by_po: HashMap<String, String> = HashMap::new();
        for row in &rows {
            if let Some(ip) = ipo {
                let po = row.get(ip).cloned().unwrap_or_default();
                let st = ist.and_then(|i| row.get(i)).cloned().unwrap_or_default();
                if !st.is_empty() && st != "nan" { state_by_po.insert(po, st); }
            }
        }
        for row in &rows {
            let sku = row.get(is).cloned().unwrap_or_default();
            if sku.is_empty() || sku == "nan" { continue; }
            let qty: f64  = row.get(iq).and_then(|v| v.parse().ok()).unwrap_or(0.0);
            let unit: f64 = row.get(iu).and_then(|v| v.parse().ok()).unwrap_or(0.0);
            let po = ipo.and_then(|i| row.get(i)).cloned().unwrap_or_default();
            ih_detail.entry(sku.clone()).or_default().push(PoDetailRow {
                po: po.clone(),
                date:    idl.and_then(|i| row.get(i)).cloned().unwrap_or_default(),
                ship_to: state_by_po.get(&po).cloned().unwrap_or_default(),
                qty: qty / units_per_case(&sku), total: qty * unit,
            });
        }
    }
    let mut sections: Vec<BigTicketSection> = Vec::new();
    let mut used_asins: HashSet<String> = HashSet::new();
    let mut used_skus:  HashSet<String> = HashSet::new();
    for m in &mapping {
        let at: f64 = if !m.asin.is_empty() { amz_detail.get(&m.asin).map(|r| r.iter().map(|x| x.total).sum()).unwrap_or(0.0) } else { 0.0 };
        let it: f64 = ih_detail.get(&m.sku).map(|r| r.iter().map(|x| x.total).sum()).unwrap_or(0.0);
        if at + it >= BIG_TICKET_THRESHOLD {
            sections.push(BigTicketSection {
                id: if !m.asin.is_empty() { m.asin.clone() } else { m.sku.clone() },
                name: m.name.clone(), combined_total: at + it,
                amz_rows:   amz_detail.get(&m.asin).cloned().unwrap_or_default(),
                iherb_rows: ih_detail.get(&m.sku).cloned().unwrap_or_default(),
            });
        }
        if !m.asin.is_empty() { used_asins.insert(m.asin.clone()); }
        used_skus.insert(m.sku.clone());
    }
    for (asin, rows) in &amz_detail {
        if used_asins.contains(asin) { continue; }
        let t: f64 = rows.iter().map(|r| r.total).sum();
        if t >= BIG_TICKET_THRESHOLD { sections.push(BigTicketSection { id: asin.clone(), name: asin.clone(), combined_total: t, amz_rows: rows.clone(), iherb_rows: vec![] }); }
        used_asins.insert(asin.clone());
    }
    for (sku, rows) in &ih_detail {
        if used_skus.contains(sku) { continue; }
        let t: f64 = rows.iter().map(|r| r.total).sum();
        if t >= BIG_TICKET_THRESHOLD { sections.push(BigTicketSection { id: sku.clone(), name: sku.clone(), combined_total: t, amz_rows: vec![], iherb_rows: rows.clone() }); }
        used_skus.insert(sku.clone());
    }
    sections.sort_by(|a, b| b.combined_total.partial_cmp(&a.combined_total).unwrap());
    Ok(sections)
}

fn tnr(size: f64, bold: bool) -> Format {
    let mut f = Format::new().set_font_name("Times New Roman").set_font_size(size);
    if bold { f = f.set_bold(); } f
}
fn money_fmt(bold: bool) -> Format {
    let mut f = Format::new().set_font_name("Times New Roman").set_font_size(14.0).set_num_format("\"$\"#,##0.00");
    if bold { f = f.set_bold(); } f
}
fn num_fmt(bold: bool) -> Format {
    let mut f = Format::new().set_font_name("Times New Roman").set_font_size(14.0).set_num_format("#,##0.00");
    if bold { f = f.set_bold(); } f
}
fn int_fmt(bold: bool) -> Format {
    let mut f = Format::new().set_font_name("Times New Roman").set_font_size(14.0).set_num_format("#,##0");
    if bold { f = f.set_bold(); } f
}

#[tauri::command]
fn save_aggregator_excel(rows: Vec<AggRow>, path: String, platform: String) -> Result<(), String> {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet();
    ws.set_name(&format!("{platform} PO Summary")).map_err(|e| e.to_string())?;
    let is_iherb = platform == "iHerb";
    let headers: &[&str] = if is_iherb { &["SKU","Title","Case Qty","Total Cost"] } else { &["ASIN","Title","Qty Ordered","Total Cost"] };
    for (ci, (h, w)) in headers.iter().zip([18.0f64,44.0,14.0,16.0].iter()).enumerate() {
        ws.write_with_format(0, ci as u16, *h, &tnr(14.0,true)).map_err(|e| e.to_string())?;
        ws.set_column_width(ci as u16, *w).map_err(|e| e.to_string())?;
    }
    for (ri, row) in rows.iter().enumerate() {
        let r = (ri+1) as u32;
        ws.write_with_format(r, 0, &row.id,    &tnr(14.0,false)).map_err(|e| e.to_string())?;
        ws.write_with_format(r, 1, &row.title, &tnr(14.0,false)).map_err(|e| e.to_string())?;
        if is_iherb { ws.write_with_format(r, 2, row.case_qty.unwrap_or(0.0), &num_fmt(false)).map_err(|e| e.to_string())?; }
        else        { ws.write_with_format(r, 2, row.qty_ordered, &int_fmt(false)).map_err(|e| e.to_string())?; }
        ws.write_with_format(r, 3, row.total_cost, &money_fmt(false)).map_err(|e| e.to_string())?;
    }
    let last = (rows.len()+1) as u32;
    ws.write_with_format(last, 0, "GRAND TOTAL", &tnr(14.0,true)).map_err(|e| e.to_string())?;
    let tq: f64 = if is_iherb { rows.iter().filter_map(|r| r.case_qty).sum() } else { rows.iter().map(|r| r.qty_ordered).sum() };
    if is_iherb { ws.write_with_format(last, 2, tq, &num_fmt(true)).map_err(|e| e.to_string())?; }
    else        { ws.write_with_format(last, 2, tq, &int_fmt(true)).map_err(|e| e.to_string())?; }
    ws.write_with_format(last, 3, rows.iter().map(|r| r.total_cost).sum::<f64>(), &money_fmt(true)).map_err(|e| e.to_string())?;
    ws.autofit();
    wb.save(&path).map_err(|e| e.to_string())
}

#[tauri::command]
fn save_combined_excel(rows: Vec<CombinedRow>, path: String) -> Result<(), String> {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet();
    ws.set_name("Combined PO Summary").map_err(|e| e.to_string())?;
    let grey = Format::new().set_font_name("Times New Roman").set_font_size(14.0).set_background_color(Color::RGB(0xD9D9D9));
    let headers = ["ASIN / SKU","Product Title","AMZ Qty (Cases)","Cost AMZ","iHerb Qty (Cases)","Cost iHerb","Total Cases AMZ+iHerb","Total Cost AMZ+iHerb"];
    for (ci,(h,w)) in headers.iter().zip([18.0f64,44.0,16.0,14.0,16.0,14.0,22.0,20.0].iter()).enumerate() {
        ws.write_with_format(0, ci as u16, *h, &tnr(14.0,true)).map_err(|e| e.to_string())?;
        ws.set_column_width(ci as u16, *w).map_err(|e| e.to_string())?;
    }
    for (ri, row) in rows.iter().enumerate() {
        let r = (ri+1) as u32;
        let idf = if row.is_pack { &grey } else { &tnr(14.0,false) };
        ws.write_with_format(r, 0, &row.id,    idf).map_err(|e| e.to_string())?;
        ws.write_with_format(r, 1, &row.title, idf).map_err(|e| e.to_string())?;
        if let Some(v) = row.amz_qty     { ws.write_with_format(r, 2, v, &num_fmt(false)).map_err(|e| e.to_string())?; }
        if let Some(v) = row.amz_cost    { ws.write_with_format(r, 3, v, &money_fmt(false)).map_err(|e| e.to_string())?; }
        if let Some(v) = row.iherb_cases { ws.write_with_format(r, 4, v, &num_fmt(false)).map_err(|e| e.to_string())?; }
        if let Some(v) = row.iherb_cost  { ws.write_with_format(r, 5, v, &money_fmt(false)).map_err(|e| e.to_string())?; }
        if let Some(v) = row.total_cases { ws.write_with_format(r, 6, v, &num_fmt(false)).map_err(|e| e.to_string())?; }
        ws.write_with_format(r, 7, row.total_cost, &money_fmt(false)).map_err(|e| e.to_string())?;
    }
    let last = (rows.len()+1) as u32;
    ws.write_with_format(last, 0, "GRAND TOTAL", &tnr(14.0,true)).map_err(|e| e.to_string())?;
    ws.write_with_format(last, 2, rows.iter().filter_map(|r| r.amz_qty).sum::<f64>(),    &num_fmt(true)).map_err(|e| e.to_string())?;
    ws.write_with_format(last, 3, rows.iter().filter_map(|r| r.amz_cost).sum::<f64>(),   &money_fmt(true)).map_err(|e| e.to_string())?;
    ws.write_with_format(last, 4, rows.iter().filter_map(|r| r.iherb_cases).sum::<f64>(),&num_fmt(true)).map_err(|e| e.to_string())?;
    ws.write_with_format(last, 5, rows.iter().filter_map(|r| r.iherb_cost).sum::<f64>(), &money_fmt(true)).map_err(|e| e.to_string())?;
    ws.write_with_format(last, 6, rows.iter().filter_map(|r| r.total_cases).sum::<f64>(),&num_fmt(true)).map_err(|e| e.to_string())?;
    ws.write_with_format(last, 7, rows.iter().map(|r| r.total_cost).sum::<f64>(),        &money_fmt(true)).map_err(|e| e.to_string())?;
    ws.autofit();
    wb.save(&path).map_err(|e| e.to_string())
}

#[tauri::command]
fn save_bigticket_excel(sections: Vec<BigTicketSection>, path: String) -> Result<(), String> {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet();
    ws.set_name("Big Ticket Breakdown").map_err(|e| e.to_string())?;
    for (ci,w) in [18.0f64,40.0,12.0,18.0,20.0,18.0,14.0,16.0].iter().enumerate() {
        ws.set_column_width(ci as u16, *w).map_err(|e| e.to_string())?;
    }
    let sub = Format::new().set_font_name("Times New Roman").set_font_size(14.0).set_bold().set_underline(FormatUnderline::Single);
    let mut row: u32 = 0;
    for sec in &sections {
        ws.write_with_format(row, 0, format!("{}   {}", sec.id, sec.name), &tnr(14.0,true)).map_err(|e| e.to_string())?;
        ws.write_with_format(row, 6, "Combined Total", &tnr(14.0,true)).map_err(|e| e.to_string())?;
        ws.write_with_format(row, 7, sec.combined_total, &money_fmt(true)).map_err(|e| e.to_string())?;
        row += 1;
        if !sec.amz_rows.is_empty() {
            for (ci,h) in ["","","Amazon","PO #","Expected Date","Ship To","Qty (Cases)","Total Cost"].iter().enumerate() {
                ws.write_with_format(row, ci as u16, *h, &sub).map_err(|e| e.to_string())?;
            }
            row += 1;
            for r in &sec.amz_rows {
                ws.write_with_format(row,2,"Amazon",&tnr(14.0,false)).map_err(|e| e.to_string())?;
                ws.write_with_format(row,3,&r.po,&tnr(14.0,false)).map_err(|e| e.to_string())?;
                ws.write_with_format(row,4,&r.date,&tnr(14.0,false)).map_err(|e| e.to_string())?;
                ws.write_with_format(row,5,&r.ship_to,&tnr(14.0,false)).map_err(|e| e.to_string())?;
                ws.write_with_format(row,6,r.qty,&num_fmt(false)).map_err(|e| e.to_string())?;
                ws.write_with_format(row,7,r.total,&money_fmt(false)).map_err(|e| e.to_string())?;
                row += 1;
            }
        }
        if !sec.iherb_rows.is_empty() {
            for (ci,h) in ["","","iHerb","PO #","Delivery Date","Ship To State","Cases","Total Cost"].iter().enumerate() {
                ws.write_with_format(row, ci as u16, *h, &sub).map_err(|e| e.to_string())?;
            }
            row += 1;
            for r in &sec.iherb_rows {
                ws.write_with_format(row,2,"iHerb",&tnr(14.0,false)).map_err(|e| e.to_string())?;
                ws.write_with_format(row,3,&r.po,&tnr(14.0,false)).map_err(|e| e.to_string())?;
                ws.write_with_format(row,4,&r.date,&tnr(14.0,false)).map_err(|e| e.to_string())?;
                ws.write_with_format(row,5,&r.ship_to,&tnr(14.0,false)).map_err(|e| e.to_string())?;
                ws.write_with_format(row,6,r.qty,&num_fmt(false)).map_err(|e| e.to_string())?;
                ws.write_with_format(row,7,r.total,&money_fmt(false)).map_err(|e| e.to_string())?;
                row += 1;
            }
        }
        row += 1;
    }
    ws.autofit();
    wb.save(&path).map_err(|e| e.to_string())
}

fn config_path() -> Option<std::path::PathBuf> {
    dirs::home_dir().map(|h| h.join(".taoftea_po_config.txt"))
}

#[tauri::command]
fn load_config() -> String {
    config_path()
        .and_then(|p| std::fs::read_to_string(p).ok())
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty() && std::path::Path::new(s).exists())
        .unwrap_or_default()
}

#[tauri::command]
fn save_config(mapping_path: String) -> Result<(), String> {
    if let Some(p) = config_path() {
        std::fs::write(p, &mapping_path).map_err(|e| e.to_string())?;
    }
    Ok(())
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            process_amazon,
            process_iherb,
            build_combined,
            build_bigticket,
            save_aggregator_excel,
            save_combined_excel,
            save_bigticket_excel,
            load_config,
            save_config,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
