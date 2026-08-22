import math
import io
import docx
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from pydantic import BaseModel
from typing import List, Optional

app = FastAPI(title="Qatar MOI CCTV Costing Engine")

# ==========================================
# 1. PARSING ENGINES
# ==========================================

def parse_survey_word(doc_bytes: bytes) -> dict:
    """Parses CCTV Site Survey Word Document and extracts item quantities."""
    doc = docx.Document(io.BytesIO(doc_bytes))
    extracted_data = {
        "dome": 0, "bullet": 0, "ptz": 0,
        "anpr": 0, "kpoi": 0,
        "nvr_count": 0, "workstation_count": 0,
        "poles_3m": 0, "poles_1_5m": 0
    }
    
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip().lower() for c in row.cells]
            if len(cells) >= 4:
                item_desc = cells[1]
                qty_str = cells[3]
                
                try:
                    qty = int(qty_str) if qty_str.isdigit() else 0
                except ValueError:
                    qty = 0
                
                if "dome" in item_desc:
                    extracted_data["dome"] += qty
                elif "bullet" in item_desc:
                    extracted_data["bullet"] += qty
                elif "ptz" in item_desc:
                    extracted_data["ptz"] += qty
                elif "anpr" in item_desc:
                    extracted_data["anpr"] += qty
                elif "kpoi" in item_desc or "k-poi" in item_desc:
                    extracted_data["kpoi"] += qty
                elif "recording" in item_desc or "nvr" in item_desc:
                    extracted_data["nvr_count"] += max(qty, 1)
                elif "monitoring" in item_desc or "workstation" in item_desc:
                    extracted_data["workstation_count"] += max(qty, 1)
                elif "civil" in item_desc:
                    extracted_data["poles_3m"] += qty

    extracted_data["total_cameras"] = (
        extracted_data["dome"] + extracted_data["bullet"] + 
        extracted_data["ptz"] + extracted_data["anpr"] + extracted_data["kpoi"]
    )
    return extracted_data


def parse_price_list(excel_bytes: bytes) -> dict:
    """Parses multi-tab price list Excel file into organized brand lookup tables."""
    xls = pd.ExcelFile(io.BytesIO(excel_bytes))
    price_db = {}
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        df.columns = [str(col).strip().lower() for col in df.columns]
        
        # Standardize column naming
        df = df.rename(columns={'unit price': 'price', 'unit_price': 'price'})
        
        items = []
        for _, row in df.iterrows():
            if pd.notna(row.get('item')) and pd.notna(row.get('price')):
                items.append({
                    "item": str(row['item']).strip(),
                    "model": str(row.get('model', 'TBD')).strip(),
                    "price": float(row['price'])
                })
        price_db[sheet_name.strip()] = items
        
    return price_db

# ==========================================
# 2. CALCULATION ENGINE (POT27813 LOGIC)
# ==========================================

class CalculationRequest(BaseModel):
    survey_data: dict
    selected_brand: str  # e.g., "Dahua", "UNV", "HikVISION"
    selected_hdd_brand: str  # e.g., "HDD"
    avg_cable_length_m: float = 100.0
    material_margin: float = 0.12  # 12%
    service_margin: float = 0.05   # 5%

@app.post("/api/generate-boq")
def generate_boq(
    req: CalculationRequest, 
    price_db: dict
):
    survey = req.survey_data
    brand = req.selected_brand
    tot_cams = survey.get("total_cameras", 0)
    avg_len = req.avg_cable_length_m
    
    brand_prices = {item['item'].lower(): item for item in price_db.get(brand, [])}
    hdd_prices = {item['item'].lower(): item for item in price_db.get(req.selected_hdd_brand, [])}
    
    # ----------------------------------------------------
    # A. STORAGE & RAID-5 CALCULATIONS (120-DAY RETENTION)
    # ----------------------------------------------------
    dome_qty = survey.get("dome", 0)
    bullet_qty = survey.get("bullet", 0)
    kpoi_qty = survey.get("kpoi", 0)
    
    # Storage per camera: 720P Dome = 2.0TB, 1080P = 3.3TB for 120 Days
    req_storage_tb = (dome_qty * 2.0) + ((bullet_qty + kpoi_qty) * 3.3)
    
    # RAID 5 calculation (16TB HDD = 14.55TB Usable)
    data_drives = math.ceil(req_storage_tb / 14.55) if req_storage_tb > 0 else 1
    total_hdds = data_drives + 1 (Parity) + 1 (Hot Spare)  # RAID 5 + HS
    
    # Lookup HDD Price
    hdd_info = hdd_prices.get('16tb', {'model': '16XTB', 'price': 2300.0})
    hdd_unit_cost = hdd_info['price']

    # ----------------------------------------------------
    # B. UPS POWER CALCULATION ENGINE
    # ----------------------------------------------------
    device_watts = {
        "monitors": 125,          # 24" (50W) + 32" (75W)
        "workstation": 250,
        "poe_switch_24p": 300,
        "poe_switch_8p": 100,
        "nvr": 100,
        "kpoi_nvr": 80 if kpoi_qty > 0 else 0
    }
    base_watts = sum(device_watts.values())
    total_watts = base_watts * 1.15  # 15% safety factor
    recommended_ups_kva = math.ceil(total_watts / 900.0)

    # ----------------------------------------------------
    # C. INFRASTRUCTURE & CABLING (DYNAMIC LENGTH)
    # ----------------------------------------------------
    total_cable_meters = tot_cams * avg_len
    cat6_boxes = total_cable_meters / 305.0
    pvc_conduit_m = total_cable_meters * 0.5
    gi_conduit_m = total_cable_meters * 0.5
    
    # ----------------------------------------------------
    # D. BOQ LINE ITEMS GENERATION & MARKUP
    # ----------------------------------------------------
    boq_items = []
    
    def add_line(category, desc, brand_name, model, uom, qty, unit_cost, is_service=False):
        margin = req.service_margin if is_service else req.material_margin
        sell_unit_price = unit_cost * (1.0 + margin)
        total_cost = unit_cost * qty
        total_sell = sell_unit_price * qty
        boq_items.append({
            "category": category,
            "description": desc,
            "brand": brand_name,
            "model": model,
            "uom": uom,
            "qty": round(qty, 2),
            "unit_cost": round(unit_cost, 2),
            "total_cost": round(total_cost, 2),
            "unit_sell": round(sell_unit_price, 2),
            "total_sell": round(total_sell, 2)
        })

    # Cameras
    if dome_qty > 0:
        c_info = brand_prices.get('2mp camera', {'model': 'D2f3', 'price': 520.0})
        add_line("Cameras", "Dome Camera 2MP IR Network", brand, c_info['model'], "EA", dome_qty, c_info['price'])
        
    if bullet_qty > 0:
        c_info = brand_prices.get('4mp camera', {'model': 'D2f2', 'price': 1500.0})
        add_line("Cameras", "Bullet Camera 4MP VF Network", brand, c_info['model'], "EA", bullet_qty, c_info['price'])

    # Storage & NVR
    nvr_info = brand_prices.get('nvr 32ch', {'model': 'Dnvr13', 'price': 4500.0})
    add_line("NVR & Storage", "Network Video Recorder", brand, nvr_info['model'], "EA", survey.get("nvr_count", 1), nvr_info['price'])
    add_line("NVR & Storage", "16TB Surveillance HDD", "Seagate/WDC", hdd_info['model'], "EA", total_hdds, hdd_unit_cost)

    # Cabling & Conduits (Breakdown sheet rollup)
    add_line("Cabling", "Cat6 Cable 305M Box", "Standard", "Cat6", "BOX", cat6_boxes, 460.0)
    add_line("Infrastructure", "PVC Conduit & Accessories", "Standard", "PVC", "MTR", pvc_conduit_m, 2.80)
    add_line("Infrastructure", "GI Conduit & Accessories", "Standard", "GI", "MTR", gi_conduit_m, 4.70)

    # Services
    installation_labor_cost = (tot_cams * 150.0) + (total_cable_meters * 2.0) + 2500.0
    add_line("Services", "Installation, Testing & Commissioning", "Flora", "LABOR", "LOT", 1, installation_labor_cost, is_service=True)

    # Totals
    total_project_cost = sum(i["total_cost"] for i in boq_items)
    total_project_sell = sum(i["total_sell"] for i in boq_items)
    overall_margin = total_project_sell - total_project_cost

    return {
        "summary": {
            "total_cameras": tot_cams,
            "required_storage_tb": round(req_storage_tb, 2),
            "total_hdds_raid5": total_hdds,
            "ups_recommended_rating": f"{recommended_ups_kva} kVA",
            "total_cable_meters": total_cable_meters,
            "total_cost_qar": round(total_project_cost, 2),
            "total_sales_qar": round(total_project_sell, 2),
            "margin_qar": round(overall_margin, 2)
        },
        "boq_lines": boq_items
    }
