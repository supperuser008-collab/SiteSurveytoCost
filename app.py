import streamlit as st
import pandas as pd
import io
from docx import Document

# --- PARSING ENGINE FOR SURVEY SECTIONS 6, 7, and 8 ---

def parse_survey_tables(doc):
    extracted_items = []
    current_section = None
    
    for table in doc.tables:
        if not table.rows:
            continue
        first_row_text = "".join([cell.text.strip() for cell in table.rows[0].cells]).upper()
        
        if "6. NEW SYSTEM REQUIREMENTS" in first_row_text:
            current_section = 6
        elif "7. ANPR & K-POI REQUIREMENTS" in first_row_text:
            current_section = 7
        elif "8. CIVIL / ELECTRICAL" in first_row_text:
            current_section = 8
        elif "9. SITE OBSERVATIONS" in first_row_text or "10. SIGN-OFF" in first_row_text:
            current_section = None
            
        if current_section in [6, 7, 8]:
            for row in table.rows[1:]:
                cells = [cell.text.strip() for cell in row.cells]
                if len(cells) >= 4:
                    item_name = cells[1]
                    description = cells[2]
                    qty_str = cells[3]
                    
                    if item_name and qty_str:
                        try:
                            qty = float(qty_str)
                        except ValueError:
                            qty = 0.0
                            
                        if qty > 0:
                            extracted_items.append({
                                "Section": current_section,
                                "Item": item_name,
                                "Description": description,
                                "Qty": qty,
                                "Unit": cells[4] if len(cells) > 4 else "EA"
                            })
                        
    return extracted_items


# --- MASTER EXCEL BOQ GENERATOR WITH MULTI-SHEET PRICING ---

def generate_master_boq_excel(survey_items, brand_price_dfs, hdd_df, selected_configs):
    output = io.BytesIO()
    
    total_cameras = sum([item['Qty'] for item in survey_items if item['Section'] == 6 and ("camera" in item['Item'].lower() or "cctv" in item['Item'].lower())])
    dome_qty = sum([item['Qty'] for item in survey_items if "dome" in item['Item'].lower()])
    bullet_qty = sum([item['Qty'] for item in survey_items if "bullet" in item['Item'].lower()])
    kpoi_qty = sum([item['Qty'] for item in survey_items if "k-poi" in item['Item'].lower() or "anpr" in item['Item'].lower()])
    
    if total_cameras > 0 and dome_qty == 0 and bullet_qty == 0:
        dome_qty = int(total_cameras)

    # Helper function to get price from selected brand sheet
    def get_brand_item_price(brand, model_name, default_price=0):
        if brand_price_dfs and brand in brand_price_dfs:
            df = brand_price_dfs[brand]
            match = df[df['Model'].astype(str).str.lower() == str(model_name).lower()]
            if not match.empty:
                return float(match.iloc[0]['Unit Price'])
        return default_price

    # Helper function to get HDD price
    def get_hdd_price(model_name, default_price=2300):
        if hdd_df is not None and not hdd_df.empty:
            match = hdd_df[hdd_df['Model'].astype(str).str.lower() == str(model_name).lower()]
            if not match.empty:
                return float(match.iloc[0]['Unit Price'])
        return default_price

    # Retrieve selected prices
    dome_price = get_brand_item_price(selected_configs.get('Dome_Brand'), selected_configs.get('Dome_Model'), 200)
    bullet_price = get_brand_item_price(selected_configs.get('Bullet_Brand'), selected_configs.get('Bullet_Model'), 371)
    nvr_price = get_brand_item_price(selected_configs.get('NVR_Brand'), selected_configs.get('NVR_Model'), 4350)
    hdd_price = get_hdd_price(selected_configs.get('HDD_Model'), 2300)
    switch_price = 985
    ups_price = 3950

    # --- SHEET 1: BOQ ---
    boq_rows = [
        [None, "Costing Summary", None, "Total Camera", None, None, None, "Shipping", 0.15, "Total Sales (QAR)", None, 0.0, 0.16, "MATERIALS"],
        [None, None, None, total_cameras, None, None, None, "Customs", 0.05, "Total Cost (QAR)", None, 0.0, 0.05, "SERVICES"],
        [None, None, None, None, None, None, None, None, None, None, None, None, None, None],
        [None, "SUBJECT :", "POT27605 - SITE SURVEY BOQ (Location - LVQ)", None, None, None, None, None, None, None, None, None, None, None],
        [None, None, None, None, None, "Ex-Works (USD)", None, "DDP - USD", None, "All-in-Cost Price (QAR)", None, "Sell Price (QAR)", None, None],
        [None, "No", "Product Description", "UOM", "Qty", "Unit Price", "Total Price", "Shipping", "Customs", "Unit Price", "Total Price", "Unit Price", "Total Price", "Remarks"],
        # Section A
        [None, "A", "Cameras and accessories", None, None, None, None, None, None, None, None, None, None, None],
        [None, 1, f"Dome Camera\nBrand: {selected_configs.get('Dome_Brand', 'TBD')} | Model: {selected_configs.get('Dome_Model', 'TBD')}", "EA", dome_qty, 0, 0, 0, 0, dome_price, dome_qty * dome_price, round(dome_price * 1.16, 2), round(dome_qty * dome_price * 1.16, 2), None],
        [None, 2, f"Bullet Camera\nBrand: {selected_configs.get('Bullet_Brand', 'TBD')} | Model: {selected_configs.get('Bullet_Model', 'TBD')}", "EA", bullet_qty, 0, 0, 0, 0, bullet_price, bullet_qty * bullet_price, round(bullet_price * 1.16, 2), round(bullet_qty * bullet_price * 1.16, 2), None],
        # Section B
        [None, "B", "VMS, NVR & Storage", None, None, None, None, None, None, None, None, None, None, None],
        [None, 3, "VMS Software", "EA", 1 if total_cameras > 0 else 0, 0, 0, 0, 0, "Included", "Included", "Included", "Included", None],
        [None, 4, f"NVR Recorder\nBrand: {selected_configs.get('NVR_Brand', 'TBD')} | Model: {selected_configs.get('NVR_Model', 'TBD')}", "EA", 1 if total_cameras > 0 else 0, 0, 0, 0, 0, nvr_price, nvr_price if total_cameras > 0 else 0, round(nvr_price * 1.16, 2), round(nvr_price * 1.16, 2) if total_cameras > 0 else 0, None],
        [None, 5, f"HDD Storage\nModel: {selected_configs.get('HDD_Model', 'TBD')}", "EA", max(0, int(total_cameras / 3)), 0, 0, 0, 0, hdd_price, max(0, int(total_cameras / 3)) * hdd_price, round(hdd_price * 1.16, 2), round(max(0, int(total_cameras / 3)) * hdd_price * 1.16, 2), None],
        # Section C
        [None, "C", "Network Switches", None, None, None, None, None, None, None, None, None, None, None],
        [None, 6, "24Port POE Switch", "EA", 2 if total_cameras > 10 else (1 if total_cameras > 0 else 0), 0, 0, 0, 0, switch_price, 0, round(switch_price * 1.16, 2), 0.0, None],
        # Section F
        [None, "F", "UPS System", None, None, None, None, None, None, None, None, None, None, None],
        [None, 7, "UPS System", "EA", 1 if total_cameras > 0 else 0, 0, 0, 0, 0, ups_price, 0, round(ups_price * 1.16, 2), 0.0, None]
    ]
    df_boq = pd.DataFrame(boq_rows)

    # --- SHEET 2: BD (Break Down Details) ---
    bd_rows = [
        [None, "BREAK DOWN DETAILS DO NOT ATTACHED WITH THE FINAL QUOTATION", None, None, None, None, None, None, None, None],
        [None, "BREAK DOWN", None, None, None, None, None, None, None, None],
        [None, None, "Cables and accessories", "Brand", "Model No", "UOM", "Qty", "Unit Price", "Total Price", "Remarks"],
        [None, 1, "Cat6 Cable 305M", "TBD", "TBD", "EA", max(0, int(total_cameras / 3)), 460, 0, None],
        [None, 2, "24 Port Patch panel", "TBD", "TBD", "EA", 2 if total_cameras > 0 else 0, 40, 0, None],
        [None, 3, "Rj45 Keystone", "TBD", "TBD", "EA", total_cameras + 10 if total_cameras > 0 else 0, 9, 0, None],
        [None, 4, "Rj45 Connector - Packet (50 pcs)", "TBD", "TBD", "EA", 1 if total_cameras > 0 else 0, 60, 0, None],
        [None, 5, "Cable Manager", "TBD", "TBD", "EA", 2 if total_cameras > 0 else 0, 35, 0, None],
        [None, 6, "Patch cord - 1M", "TBD", "TBD", "EA", total_cameras + 2 if total_cameras > 0 else 0, 8, 0, None],
        [None, 7, "Patch cord - 3M", "TBD", "TBD", "EA", 3 if total_cameras > 0 else 0, 12, 0, None],
        [None, 8, "CCTV Stickers", "TBD", "TBD", "LOT", 10 if total_cameras > 0 else 0, 20, 0, None],
        [None, 9, "Sundries and miscellaneous (HDMI Cables)", "TBD", "TBD", "LOT", 1 if total_cameras > 0 else 0, 500, 0, None],
        [None, None, "Electrical cables and accessories", None, None, None, None, None, 0, None],
        [None, 1, "Electrical cables & conduits", "TBD", "TBD", "EA", 1 if total_cameras > 0 else 0, 1000, 0, None]
    ]
    df_bd = pd.DataFrame(bd_rows)

    # --- SHEET 3: Storage ---
    required_storage_tb = round(total_cameras * 2.65, 2)
    hdd_count = max(0, int(total_cameras / 3))
    available_storage_tb = round(hdd_count * 11.9, 2)
    storage_rows = [
        [None, "PROJECT", "Site Survey Extraction BOQ", None, None, None, "REQUIRED STORAGE - TB", None, required_storage_tb, None, None, None],
        [None, "STORAGE TYPE", f"NVR Storage ({selected_configs.get('HDD_Model', 'TBD')})", None, None, None, "AVAILABLE STORAGE - TB", None, available_storage_tb, None, None, None],
        [None, "TOTAL HDD", hdd_count, None, None, None, "OVERALL STORAGE LOAD", None, round(required_storage_tb / available_storage_tb, 2) if available_storage_tb > 0 else 0.0, None, None, None],
        [None, "TOTAL CAMERA", total_cameras, None, None, None, None, None, None, None, None, None],
        [None, None, None, None, None, None, None, None, None, None, None, None],
        [None, "CAMERA CONFIGURATION PARAMETERS", None, None, None, None, None, None, None, None, None, None],
        [None, "Camera", "Encoding Mode", "Recording Resolution", "Frame Rate (fps)", "Bitrate (Mbps)", "No. of Cameras", "Estimated (TB) Storage per camera", "Required Storage (TB)", "Remarks", None, None],
        [None, "Dome", "H.264", "1280x720", 15, 1.5, dome_qty, 2.0, round(dome_qty * 2.0, 2), None, None, None],
        [None, "Bullet", "H.264", "1920x1080P", 15, 2.5, bullet_qty, 3.3, round(bullet_qty * 3.3, 2), None, None, None],
        [None, None, None, None, None, None, total_cameras, None, required_storage_tb, None, None, None]
    ]
    df_storage = pd.DataFrame(storage_rows)

    # --- SHEET 4: UPS ---
    ups_rows = [
        [None, "ESTIMATED UPS CALCULATION - MDF", None, None, None, None, None, None, "UPS KVA"],
        [None, "PROJECT", None, None, None, None, None, None, 3],
        [None, "PROPOSED UPS", None, None, None, None, None, None, 3],
        [None, "No", "Item Description", "Brand", "Availability", "Model No", "Quantity", "Watts", "Total Watts"],
        [None, 1, "24\" Monitor", "TBD", None, "TBD", 1, 50, 50],
        [None, 2, "Workstation", "TBD", None, "TBD", 2, 250, 500],
        [None, 3, "24 Port Switch", "TBD", None, "TBD", 2, 370, 740],
        [None, 4, "NVR", selected_configs.get('NVR_Brand', 'TBD'), None, selected_configs.get('NVR_Model', 'TBD'), 1, 100, 100]
    ]
    df_ups = pd.DataFrame(ups_rows)

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_boq.to_excel(writer, sheet_name='BOQ', index=False, header=False)
        df_bd.to_excel(writer, sheet_name='BD', index=False, header=False)
        df_storage.to_excel(writer, sheet_name='Storage', index=False, header=False)
        df_ups.to_excel(writer, sheet_name='UPS', index=False, header=False)
        
    output.seek(0)
    return output


# --- STREAMLIT APP ---

st.set_page_config(page_title="Flora Tech BOQ Generator", layout="wide")

st.title("Flora Technology - MOI CCTV BOQ Generator")
st.markdown("Upload your filled **Site Survey Word Document** and your **Price List Excel File** (containing sheets for Dahua, UNV, HikVISION, and HDD).")

uploaded_survey = st.file_uploader("1. Upload Site Survey Word Document (.docx)", type="docx")
uploaded_pricing = st.file_uploader("2. Upload Price List Excel File (.xlsx)", type="xlsx")

brand_price_dfs = {}
hdd_df = pd.DataFrame()
available_brands = ["Dahua", "UNV", "HikVISION"]
camera_models = ["D2f2", "D2f3", "D2f4", "D2f5"]
nvr_models = ["Dnvr12", "Dnvr13", "Dnvr14"]
hdd_models = ["1XTB", "8XTB", "16XTB"]

if uploaded_pricing is not None:
    try:
        xls = pd.ExcelFile(uploaded_pricing)
        sheet_names = xls.sheet_names
        
        # Load brand sheets
        for brand in ["Dahua", "UNV", "HikVISION"]:
            if brand in sheet_names:
                brand_price_dfs[brand] = pd.read_excel(uploaded_pricing, sheet_name=brand)
                
        # Load HDD sheet
        if "HDD" in sheet_names:
            hdd_df = pd.read_excel(uploaded_pricing, sheet_name="HDD")
            if 'Model' in hdd_df.columns:
                hdd_models = hdd_df['Model'].dropna().unique().tolist()
    except Exception as e:
        st.warning(f"Error reading price list sheets: {e}")

selected_configs = {}

if uploaded_survey is not None:
    doc = Document(uploaded_survey)
    extracted_items = parse_survey_tables(doc)
    
    st.success("Survey file parsed successfully!")
    
    st.subheader("Select Brands & Models (Synced with Price List)")
    col1, col2 = st.columns(2)
    
    with col1:
        selected_configs['Dome_Brand'] = st.selectbox("Dome Camera Brand", available_brands)
        # Filter models for selected brand containing camera
        if selected_configs['Dome_Brand'] in brand_price_dfs:
            df_b = brand_price_dfs[selected_configs['Dome_Brand']]
            cam_models = df_b[df_b['Item'].str.contains('camera|dome', case=False, na=False)]['Model'].tolist()
            if cam_models: camera_models = cam_models
        selected_configs['Dome_Model'] = st.selectbox("Dome Camera Model", camera_models)
        
        selected_configs['Bullet_Brand'] = st.selectbox("Bullet Camera Brand", available_brands)
        selected_configs['Bullet_Model'] = st.selectbox("Bullet Camera Model", camera_models)
        
    with col2:
        selected_configs['NVR_Brand'] = st.selectbox("NVR Recorder Brand", available_brands)
        if selected_configs['NVR_Brand'] in brand_price_dfs:
            df_b = brand_price_dfs[selected_configs['NVR_Brand']]
            n_models = df_b[df_b['Item'].str.contains('nvr', case=False, na=False)]['Model'].tolist()
            if n_models: nvr_models = n_models
        selected_configs['NVR_Model'] = st.selectbox("NVR Recorder Model", nvr_models)
        
        selected_configs['HDD_Model'] = st.selectbox("HDD Storage Model / Size", hdd_models)

    if st.button("Generate Complete Master BOQ Excel"):
        with st.spinner("Extracting pricing from all sheets and generating master workbook..."):
            excel_data = generate_master_boq_excel(extracted_items, brand_price_dfs, hdd_df, selected_configs)
            
            st.success("Master BOQ Excel generated successfully!")
            st.download_button(
                label="📥 Download Complete Master BOQ Excel",
                data=excel_data,
                file_name="POT27605 - Complete Master BOQ.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
