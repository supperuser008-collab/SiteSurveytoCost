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


# --- MASTER EXCEL BOQ GENERATOR (27-ITEM FORMAT) ---

def generate_master_boq_excel(survey_items, brand_price_dfs, hdd_df, selected_configs):
    output = io.BytesIO()
    
    total_cameras = sum([item['Qty'] for item in survey_items if item['Section'] == 6 and ("camera" in item['Item'].lower() or "cctv" in item['Item'].lower())])
    dome_qty = sum([item['Qty'] for item in survey_items if "dome" in item['Item'].lower()])
    bullet_qty = sum([item['Qty'] for item in survey_items if "bullet" in item['Item'].lower()])
    
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

    # --- SHEET 1: BOQ (Full 27-item format) ---
    boq_rows = [
        [None, "Costing Summary", None, "Total Camera", None, None, None, "Shipping", 0.15, "Total Sales (QAR)", None, 141579.49, 0.16, "MATERIALS"],
        [None, None, None, total_cameras if total_cameras > 0 else 36, None, None, None, "Customs", 0.05, "Total Cost (QAR)", None, 123769.56, 0.05, "SERVICES"],
        [None, None, None, None, None, None, None, None, None, None, None, None, None, None],
        [None, "SUBJECT :", "POT27605 - AL SHARQ CONSULTING ENGINEERING - RFQ (Location - LVQ)", None, None, None, None, None, None, None, None, None, None, None],
        [None, None, None, None, None, "Ex-Works (USD)", None, "DDP - USD", None, "All-in-Cost Price (QAR)", None, "Sell Price (QAR)", None, None],
        [None, "No", "Product Description", "UOM", "Qty", "Unit Price", "Total Price", "Shipping", "Customs", "Unit Price", "Total Price", "Unit Price", "Total Price", "Remarks"],
        
        # Section A
        [None, "A", "Cameras and accessories", None, None, None, None, None, None, None, None, None, None, None],
        [None, 1, f"Dome Camera - Fixed\n2MP Dome Camera\nBrand: {selected_configs.get('Dome_Brand', 'TBD')} | Model: {selected_configs.get('Dome_Model', 'TBD')}", "EA", 10, 0, 0, 0, 0, dome_price, 10 * dome_price, round(dome_price * 1.16, 2), round(10 * dome_price * 1.16, 2), None],
        [None, 2, f"Dome Camera - Auto Iris\n2MP WDR LightHunter IR Network Dome Camera", "EA", 12, 0, 0, 0, 0, bullet_price, 12 * bullet_price, round(bullet_price * 1.16, 2), round(12 * bullet_price * 1.16, 2), None],
        [None, 3, f"Bullet Camera\n2MP HD IR VF Bullet Network Camera", "EA", 8, 0, 0, 0, 0, bullet_price, 8 * bullet_price, round(bullet_price * 1.16, 2), round(8 * bullet_price * 1.16, 2), None],
        
        # Section B
        [None, "B", "VMS, NVR & Storage", None, None, None, None, None, None, None, None, None, None, None],
        [None, 4, "VMS Software", "EA", 1, 0, 0, 0, 0, "Included", "Included", "Included", "Included", None],
        [None, 5, f"NVR Recorder\nBrand: {selected_configs.get('NVR_Brand', 'TBD')} | Model: {selected_configs.get('NVR_Model', 'TBD')}", "EA", 1, 0, 0, 0, 0, nvr_price, nvr_price, round(nvr_price * 1.16, 2), round(nvr_price * 1.16, 2), None],
        [None, 6, f"HDD Storage\nModel: {selected_configs.get('HDD_Model', 'TBD')}", "EA", 11, 0, 0, 0, 0, hdd_price, 11 * hdd_price, round(hdd_price * 1.16, 2), round(11 * hdd_price * 1.16, 2), None],
        
        # Section C
        [None, "C", "Network Switches", None, None, None, None, None, None, None, None, None, None, None],
        [None, 7, "24Port POE Switch", "EA", 2, 0, 0, 0, 0, 985, 1970, 1142.60, 2285.20, None],
        
        # Section D
        [None, "D", "Workstation and Monitor", None, None, None, None, None, None, None, None, None, None, None],
        [None, 8, '24" LED FHD Monitor_MOI Approved', "EA", 1, 0, 0, 0, 0, 320, 320, 371.20, 371.20, None],
        [None, 9, '32 Inches Monitor_MOI Approved', "EA", 1, 0, 0, 0, 0, 750, 750, 870.00, 870.00, None],
        [None, 10, 'CCTV Workstation with K/M', "EA", 1, 0, 0, 0, 0, 3000, 3000, 3480.00, 3480.00, None],
        
        # Section E
        [None, "E", "CCTV Racks", None, None, None, None, None, None, None, None, None, None, None],
        [None, 11, '18U Floor Mount Rack with all accessories', "EA", 1, 0, 0, 0, 0, 1300, 1300, 1508.00, 1508.00, None],
        
        # Section F
        [None, "F", "UPS System", None, None, None, None, None, None, None, None, None, None, None],
        [None, 12, '3KVA UPS with Battery Pack for 60 Min. Backup', "EA", 1, 0, 0, 0, 0, 3950, 3950, 4582.00, 4582.00, None],
        
        # Section G
        [None, "G", "ANPR System", None, None, None, None, None, None, None, None, None, None, None],
        [None, 13, '4MP Smart ANPR Camera', "EA", 2, 0, 0, None, None, 1504, 3008, 1744.64, 3489.28, None],
        [None, 14, 'Pole Mount Bracket', "EA", 2, 0, 0, None, None, 65, 130, 75.40, 150.80, None],
        [None, 15, '16CH 8DD 2U Network Video Recorder', "EA", 1, 0, 0, None, None, 2271, 2271, 2634.36, 2634.36, None],
        [None, 16, '8TB HDD', "EA", 4, 0, 0, None, None, 1500, 6000, 1740.00, 6960.00, None],
        [None, 16, 'DSS Professional V8 16 Ch Base', "EA", 1, 0, 0, None, None, 1350, 1350, 1566.00, 1566.00, None],
        [None, 17, 'Single Channel License', "EA", 1, None, None, None, None, 600, 600, 696.00, 696.00, None],
        [None, 17, 'ANPR Workstation', "EA", 1, 0, 0, None, None, 4000, 4000, 4640.00, 4640.00, None],
        [None, 18, '24" LED FHD Monitor_MOI Approved', "EA", 1, 0, 0, None, None, 320, 320, 371.20, 371.20, None],
        
        # Section H
        [None, "H", "K-POI System", None, None, None, None, None, None, None, None, None, None, None],
        [None, 19, '4MP BOX WizMind Network Camera', "EA", 4, 0, 0, None, None, 1285, 5140, 1490.60, 5962.40, None],
        [None, 20, "12 MP 1/1.7' 10.5-42mm vari-focal lens", "EA", 4, 0, 0, None, None, 500, 2000, 580.00, 2320.00, None],
        [None, 21, 'Wall mount Bracket', "EA", 4, 0, 0, None, None, 25, 100, 29.00, 116.00, None],
        [None, 22, '1U Intelligent Video Device, 10 Channel. 1 slot, SATA3.0 Max.20 TB/HDD', "EA", 1, 0, 0, None, None, 3600, 3600, 4176.00, 4176.00, None],
        [None, 23, '20TB HDD', "EA", 1, 0, 0, None, None, 3500, 3500, 4060.00, 4060.00, None],
        
        # Section J
        [None, "J", "ACS System", None, None, None, None, None, None, None, None, None, None, None],
        [None, 24, 'Standalone Access control system', "EA", 1, 0, 0, None, None, 300, 300, 348.00, 348.00, None],
        
        # Section K
        [None, "K", "Cabling and accessories", None, None, None, None, None, None, None, None, None, None, None],
        [None, 25, 'Cabling and accessories', "LOT", 1, None, None, None, None, 7570.56, 7570.56, 8781.85, 8781.85, None],
        [None, 26, 'Conduites and accessories', "LOT", 1, None, None, None, None, 15400, 15400, 17864.00, 17864.00, None],
        
        # Section L
        [None, "L", "Engineering & Services", None, None, None, None, None, None, None, None, None, None, None],
        [None, 27, 'Installation, configuration, Testing & Commissioning', "LOT", 1, None, None, None, None, 18120, 18120, 19026.00, 19026.00, None],
        
        [None, None, None, None, None, None, None, None, None, 'Total Cost', 123769.56, 'Total Sales', 141579.49, None]
    ]
    df_boq = pd.DataFrame(boq_rows)

    # --- SHEET 2: BD (Break Down Details) ---
    bd_rows = [
        [None, "BREAK DOWN DETAILS DO NOT ATTACHED WITH THE FINAL QUOTATION", None, None, None, None, None, None, None, None],
        [None, "BREAK DOWN", None, None, None, None, None, None, None, None],
        [None, None, "Cables and accessories", "Brand", "Model No", "UOM", "Qty", "Unit Price", "Total Price", "Remarks"],
        [None, 1, "Cat6 Cable 305M", "TBD", "TBD", "EA", 10.62, 460, 4886.56, None],
        [None, 2, "24 Port Patch panel", "TBD", "TBD", "EA", 2, 40, 80, None],
        [None, 3, "Rj45 Keystone", "TBD", "TBD", "EA", 50, 9, 450, None],
        [None, 4, "Rj45 Connector - Packet (50 pcs)", "TBD", "TBD", "EA", 1, 60, 60, None],
        [None, 5, "Cable Manager", "TBD", "TBD", "EA", 2, 35, 70, None],
        [None, 6, "Patch cord - 1M", "TBD", "TBD", "EA", 36, 8, 288, None],
        [None, 7, "Patch cord - 3M", "TBD", "TBD", "EA", 3, 12, 36, None],
        [None, 8, "CCTV Stickers", "TBD", "TBD", "LOT", 10, 20, 200, None],
        [None, 9, "Sundries and miscellaneous (HDMI Cables and accessories)", "TBD", "TBD", "LOT", 1, 500, 500, None],
        [None, None, "Electrical cables and accessories", None, None, None, None, None, 6570.56, None],
        [None, 1, "Electrical cables", "TBD", "TBD", "EA", 1, 1000, 1000, None],
        [None, 2, "Sundries and miscellaneous", "TBD", "TBD", "EA", 1, 0, 0, None],
        [None, None, "Conduits and accessories", None, None, None, None, None, 1000, None],
        [None, 1, "PVC Conduit and accessories", "TBD", "TBD", "EA", 18, 300, 5400, None],
        [None, 1, "GI Conduit and accessories", "TBD", "TBD", "EA", 18, 500, 9000, None],
        [None, 2, "Flexible conduit and accessories", "TBD", "TBD", "EA", 1, 500, 500, None],
        [None, 3, "Sundries and miscellaneous", "TBD", "TBD", "LOT", 1, 500, 500, None],
        [None, None, "Civil works and accessories", None, None, None, None, None, 15400, None],
        [None, 1, "Road cutting and backflling", "Flora", "TBD", "MTR", 0, 450, 0, None],
        [None, 2, "3 Meter Pole", "Flora", "TBD", "LOT", 0, 1500, 0, None],
        [None, 3, "1.5 Meter Pole", "Flora", "TBD", "LOT", 0, 1000, 0, None],
        [None, 4, "Pole foundation", "Flora", "TBD", "LOT", 0, 500, 0, None]
    ]
    df_bd = pd.DataFrame(bd_rows)

    # --- SHEET 3: Storage ---
    storage_rows = [
        [None, "PROJECT", "POT27519 - QNBN - RFQ - Supply & Installation", None, None, None, "REQUIRED STORAGE - TB", None, 90.1, None, None, None],
        [None, "STORAGE TYPE", "NVR508-32B - Network Video Recorder", None, None, None, "AVAILABLE STORAGE - TB", None, 130.95, None, None, None],
        [None, "TOTAL HDD - 16TB", 11, None, None, None, "OVERALL STORAGE LOAD", None, 0.69, None, None, None],
        [None, "TOTAL CAMERA", 34, None, None, None, None, None, None, None, None, None],
        [None, None, None, None, None, None, None, None, None, None, None, None],
        [None, "CAMERA CONFIGURATION PARAMETERS", None, None, None, None, None, None, None, None, None, None],
        [None, "Camera", "Encoding Mode", "Recording Resolution", "Frame Rate (fps)", "Bitrate (Mbps)", "No. of Cameras", "Estimated (TB) Storage per camera", "Required Storage (TB)", "Remarks", None, None],
        [None, "Dome", "H.264", "1280x720", 15, 1.5, 17, 2.0, 34.0, None, None, None],
        [None, "Dome - AI", "H.264", "1920x1080P", 15, 2.5, 5, 3.3, 16.5, None, None, None],
        [None, "Bullet", "H.264", "1920x1080P", 15, 2.5, 8, 3.3, 26.4, None, None, None],
        [None, "K-POI", "H.264", "1920x1080P", 15, 2.5, 4, 3.3, 13.2, None, None, None]
    ]
    df_storage = pd.DataFrame(storage_rows)

    # --- SHEET 4: UPS ---
    ups_rows = [
        [None, "ESTIMATED UPS CALCULATION - MDF", None, None, None, None, None, None, "UPS KVA"],
        [None, "PROJECT", None, None, None, None, None, None, 3],
        [None, "PROPOSED UPS", None, None, None, None, None, None, 3],
        [None, "No", "Item Description", "Brand", "Availability", "Model No", "Quantity", "Watts", "Total Watts"],
        [None, 1, "24\" Monitor", "TBD", None, "TBD", 1, 50, 50],
        [None, 2, "32\" Monitor", "TBD", None, "TBD", 1, 75, 75],
        [None, 3, "Workstation", "TBD", None, "TBD", 2, 250, 500],
        [None, 4, "24 Port Switch", "TBD", None, "TBD", 2, 370, 740],
        [None, 5, "NVR", "TBD", None, "TBD", 1, 100, 100],
        [None, 6, "K-POI NVR", "Dahua", None, "TBD", 1, 80, 80]
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
st.markdown("Upload your filled **Site Survey Word Document** and your **Price List Excel File**.")

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
        
        for brand in ["Dahua", "UNV", "HikVISION"]:
            if brand in sheet_names:
                brand_price_dfs[brand] = pd.read_excel(uploaded_pricing, sheet_name=brand)
                
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
        with st.spinner("Extracting pricing and generating master workbook..."):
            excel_data = generate_master_boq_excel(extracted_items, brand_price_dfs, hdd_df, selected_configs)
            
            st.success("Master BOQ Excel generated successfully!")
            st.download_button(
                label="📥 Download Complete Master BOQ Excel",
                data=excel_data,
                file_name="POT27605 - Complete Master BOQ.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
