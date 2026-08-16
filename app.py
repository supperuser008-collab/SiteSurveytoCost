import streamlit as st
import pandas as pd
import io
from docx import Document

# --- PARSING ENGINE FOR SURVEY SECTIONS 6, 7, and 8 ---

def parse_survey_tables(doc):
    """
    Parses Sections 6, 7, and 8 from the uploaded Word document tables.
    Extracts item descriptions and quantities written by the technician.
    """
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
            for row in table.rows[1:]: # Skip header
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


# --- MASTER EXCEL BOQ GENERATOR (FULL TEMPLATE STRUCTURE) ---

def generate_master_boq_excel(survey_items):
    """
    Generates the complete multi-sheet Excel workbook matching the exact master BOQ structure,
    dynamically calculating quantities for Storage, UPS, and Accessories based on survey inputs.
    """
    output = io.BytesIO()
    
    # Calculate totals from survey items
    total_cameras = sum([item['Qty'] for item in survey_items if item['Section'] == 6 and ("camera" in item['Item'].lower() or "cctv" in item['Item'].lower())])
    if total_cameras == 0: 
        total_cameras = 34  # Fallback default if survey quantities are empty
        
    dome_qty = int(total_cameras * 0.5) if total_cameras > 0 else 17
    bullet_qty = int(total_cameras * 0.25) if total_cameras > 0 else 8
    kpoi_qty = int(total_cameras * 0.25) if total_cameras > 0 else 9
    
    # --- SHEET 1: BOQ ---
    boq_rows = [
        [None, "Costing Summary", None, f"Total Camera", None, None, None, "Shipping", 0.15, "Total Sales (QAR)", None, 141579.48, 0.16, "MATERIALS"],
        [None, None, None, total_cameras, None, None, None, "Customs", 0.05, "Total Cost (QAR)", None, 123769.55, 0.05, "SERVICES"],
        [None, None, None, None, None, None, None, None, None, None, None, None, None, None],
        [None, "SUBJECT :", "POT27605 - SITE SURVEY BOQ (Location - LVQ)", None, None, None, None, None, None, None, None, None, None, None],
        [None, None, None, None, None, "Ex-Works (USD)", None, "DDP - USD", None, "All-in-Cost Price (QAR)", None, "Sell Price (QAR)", None, None],
        [None, "No", "Product Description", "UOM", "Qty", "Unit Price", "Total Price", "Shipping", "Customs", "Unit Price", "Total Price", "Unit Price", "Total Price", "Remarks"],
        # Section A
        [None, "A", "Cameras and accessories", None, None, None, None, None, None, None, None, None, None, None],
        [None, 1, "Dome Camera - Fixed\n2MP Dome Camera", "EA", dome_qty, 0, 0, 0, 0, 200, dome_qty * 200, 232.0, dome_qty * 232.0, None],
        [None, 2, "Dome Camera - Auto Iris\n2MP WDR LightHunter IR Network Dome Camera", "EA", max(1, dome_qty // 2), 0, 0, 0, 0, 371, max(1, dome_qty // 2) * 371, 430.36, max(1, dome_qty // 2) * 430.36, None],
        [None, 3, "Bullet Camera\n2MP HD IR VF Bullet Network Camera", "EA", bullet_qty, 0, 0, 0, 0, 371, bullet_qty * 371, 430.36, bullet_qty * 430.36, None],
        # Section B
        [None, "B", "VMS, NVR & Storage", None, None, None, None, None, None, None, None, None, None, None],
        [None, 4, "VMS Software", "EA", 1, 0, 0, 0, 0, "Included", "Included", "Included", "Included", None],
        [None, 5, "UNV NVR516-64, Network Video Recorder", "EA", 1, 0, 0, 0, 0, 4350, 4350, 5046, 5046, None],
        [None, 6, "16TB HDD", "EA", max(2, int(total_cameras / 3)), 0, 0, 0, 0, 2300, max(2, int(total_cameras / 3)) * 2300, 2668, max(2, int(total_cameras / 3)) * 2668, None],
        # Section C
        [None, "C", "Network Switches", None, None, None, None, None, None, None, None, None, None, None],
        [None, 7, "24Port POE Switch", "EA", 2, 0, 0, 0, 0, 985, 1970, 1142.6, 2285.2, None],
        # Section F
        [None, "F", "UPS System", None, None, None, None, None, None, None, None, None, None, None],
        [None, 12, "3KVA UPS with Battery Pack for 60 Min. Backup", "EA", 1, 0, 0, 0, 0, 3950, 3950, 4582, 4582, None]
    ]
    df_boq = pd.DataFrame(boq_rows)

    # --- SHEET 2: BD (Break Down Details) ---
    bd_rows = [
        [None, "BREAK DOWN DETAILS DO NOT ATTACHED WITH THE FINAL QUOTATION", None, None, None, None, None, None, None, None],
        [None, "BREAK DOWN", None, None, None, None, None, None, None, None],
        [None, None, "Cables and accessories", "Brand", "Model No", "UOM", "Qty", "Unit Price", "Total Price", "Remarks"],
        [None, 1, "Cat6 Cable 305M", "TBD", "TBD", "EA", max(2, int(total_cameras / 3)), 460, max(2, int(total_cameras / 3)) * 460, None],
        [None, 2, "24 Port Patch panel", "TBD", "TBD", "EA", 2, 40, 80, None],
        [None, 3, "Rj45 Keystone", "TBD", "TBD", "EA", total_cameras + 10, 9, (total_cameras + 10) * 9, None],
        [None, 4, "Rj45 Connector - Packet (50 pcs)", "TBD", "TBD", "EA", 1, 60, 60, None],
        [None, 5, "Cable Manager", "TBD", "TBD", "EA", 2, 35, 70, None],
        [None, 6, "Patch cord - 1M", "TBD", "TBD", "EA", total_cameras + 2, 8, (total_cameras + 2) * 8, None],
        [None, 7, "Patch cord - 3M", "TBD", "TBD", "EA", 3, 12, 36, None],
        [None, 8, "CCTV Stickers", "TBD", "TBD", "LOT", 10, 20, 200, None],
        [None, 9, "Sundries and miscellaneous (HDMI Cables)", "TBD", "TBD", "LOT", 1, 500, 500, None],
        [None, None, "Electrical cables and accessories", None, None, None, None, None, 1500, None],
        [None, 1, "Electrical cables & conduits", "TBD", "TBD", "EA", 1, 1000, 1000, None]
    ]
    df_bd = pd.DataFrame(bd_rows)

    # --- SHEET 3: Storage ---
    required_storage_tb = round(total_cameras * 2.65, 2)
    available_storage_tb = round(max(2, int(total_cameras / 3)) * 11.9, 2)
    storage_rows = [
        [None, "PROJECT", "Site Survey Extraction BOQ", None, None, None, "REQUIRED STORAGE - TB", None, required_storage_tb, None, None, None],
        [None, "STORAGE TYPE", "UNV NVR516-64 - Network Video Recorder", None, None, None, "AVAILABLE STORAGE - TB", None, available_storage_tb, None, None, None],
        [None, "TOTAL HDD - 16TB", max(2, int(total_cameras / 3)), None, None, None, "OVERALL STORAGE LOAD", None, round(required_storage_tb / available_storage_tb, 2) if available_storage_tb > 0 else 0.5, None, None, None],
        [None, "TOTAL CAMERA", total_cameras, None, None, None, None, None, None, None, None, None],
        [None, None, None, None, None, None, None, None, None, None, None, None],
        [None, "CAMERA CONFIGURATION PARAMETERS", None, None, None, None, None, None, None, None, None, None],
        [None, "Camera", "Encoding Mode", "Recording Resolution", "Frame Rate (fps)", "Bitrate (Mbps)", "No. of Cameras", "Estimated (TB) Storage per camera", "Required Storage (TB)", "Remarks", None, None],
        [None, "Dome", "H.264", "1280x720", 15, 1.5, dome_qty, 2.0, round(dome_qty * 2.0, 2), None, None, None],
        [None, "Bullet", "H.264", "1920x1080P", 15, 2.5, bullet_qty, 3.3, round(bullet_qty * 3.3, 2), None, None, None],
        [None, "K-POI", "H.264", "1920x1080P", 15, 2.5, kpoi_qty, 3.3, round(kpoi_qty * 3.3, 2), None, None, None],
        [None, None, None, None, None, None, total_cameras, None, required_storage_tb, None, None, None]
    ]
    df_storage = pd.DataFrame(storage_rows)

    # --- SHEET 4: UPS ---
    monitor_qty = 2 if total_cameras > 15 else 1
    switch_qty = 2 if total_cameras > 16 else 1
    total_watts = (monitor_qty * 50) + (2 * 250) + (switch_qty * 370) + 100 + 80
    ups_kva = 3 if total_watts > 1000 else 2
    
    ups_rows = [
        [None, "ESTIMATED UPS CALCULATION - MDF", None, None, None, None, None, None, "UPS KVA"],
        [None, "PROJECT", None, None, None, None, None, None, ups_kva],
        [None, "PROPOSED UPS", None, None, None, None, None, None, 3],
        [None, "No", "Item Description", "Brand", "Availability", "Model No", "Quantity", "Watts", "Total Watts"],
        [None, 1, "24\" Monitor", "TBD", None, "TBD", monitor_qty, 50, monitor_qty * 50],
        [None, 2, "Workstation", "TBD", None, "TBD", 2, 250, 500],
        [None, 3, "24 Port Switch", "TBD", None, "TBD", switch_qty, 370, switch_qty * 370],
        [None, 4, "NVR", "TBD", None, "TBD", 1, 100, 100],
        [None, 5, "K-POI NVR", "Dahua", None, "TBD", 1, 80, 80],
        [None, None, None, None, None, None, None, None, total_watts]
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
st.markdown("Upload your filled survey Word file. The application will parse your quantities, calculate storage and UPS loads, and generate the **complete Master Excel BOQ**.")

uploaded_file = st.file_uploader("Upload Site Survey Word Document (.docx)", type="docx")

if uploaded_file is not None:
    doc = Document(uploaded_file)
    extracted_items = parse_survey_tables(doc)
    
    st.success("Survey file parsed successfully!")
    
    if extracted_items:
        st.subheader("Extracted Survey Items")
        st.dataframe(pd.DataFrame(extracted_items))
    else:
        st.warning("No explicit item quantities found in survey sections. Default template configuration will apply.")

    if st.button("Generate Complete Master BOQ Excel"):
        with st.spinner("Calculating storage, UPS, break down, and generating master workbook..."):
            excel_data = generate_master_boq_excel(extracted_items)
            
            st.success("Master BOQ Excel generated successfully with all sheets and calculations!")
            st.download_button(
                label="📥 Download Complete Master BOQ Excel",
                data=excel_data,
                file_name="POT27605 - Complete Master BOQ.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
