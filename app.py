import streamlit as st
import pandas as pd
import io
import math
from docx import Document

# --- CORE MOI-SSD CALCULATION ENGINE ---

def extract_survey_sections(doc):
    """
    Parses Sections 6, 7, and 8 from the CCTV Site Visit & Survey Word Document.
    """
    camera_counts = {"Dome": 0, "Bullet": 0, "AI_Dome": 0, "KPOI": 0, "ANPR": 0}
    hardware_counts = {"24_Port_Switch": 0, "NVR": 0, "Workstation": 0, "Monitor_24": 0, "Monitor_32": 0}
    
    # Iterate through all tables in the uploaded document
    for table in doc.tables:
        for row in table.rows:
            row_text = " ".join([cell.text.strip().lower() for cell in row.cells])
            
            # Extract Camera Counts
            if "dome" in row_text:
                for cell in row.cells:
                    if cell.text.strip().isdigit():
                        camera_counts["Dome"] += int(cell.text.strip())
            elif "bullet" in row_text:
                for cell in row.cells:
                    if cell.text.strip().isdigit():
                        camera_counts["Bullet"] += int(cell.text.strip())
            elif "anpr" in row_text:
                for cell in row.cells:
                    if cell.text.strip().isdigit():
                        camera_counts["ANPR"] += int(cell.text.strip())
            elif "k-poi" in row_text or "kpoi" in row_text:
                for cell in row.cells:
                    if cell.text.strip().isdigit():
                        camera_counts["KPOI"] += int(cell.text.strip())
            
            # Extract Hardware
            if "switch" in row_text:
                for cell in row.cells:
                    if cell.text.strip().isdigit():
                        hardware_counts["24_Port_Switch"] += int(cell.text.strip())
            elif "workstation" in row_text:
                for cell in row.cells:
                    if cell.text.strip().isdigit():
                        hardware_counts["Workstation"] += int(cell.text.strip())
            elif "monitor" in row_text:
                for cell in row.cells:
                    if cell.text.strip().isdigit():
                        hardware_counts["Monitor_24"] += int(cell.text.strip())

    return camera_counts, hardware_counts


def calculate_storage(cameras):
    """Calculates MOI 120-Day Storage & RAID 5 HDD Requirements"""
    dome_tb = cameras.get('Dome', 0) * 2.0
    bullet_tb = cameras.get('Bullet', 0) * 3.3
    ai_dome_tb = cameras.get('AI_Dome', 0) * 3.3
    kpoi_tb = cameras.get('KPOI', 0) * 3.3
    anpr_tb = cameras.get('ANPR', 0) * 3.3
    
    req_storage = dome_tb + bullet_tb + ai_dome_tb + kpoi_tb + anpr_tb
    
    usable_per_drive = 14.9
    drives_for_data = math.ceil(req_storage / usable_per_drive) if req_storage > 0 else 1
    total_hdds = drives_for_data + 2  # RAID 5 + Global Hot Spare
    
    return {
        "required_tb": round(req_storage, 2),
        "total_hdds": total_hdds,
        "usable_tb": round(drives_for_data * usable_per_drive, 2)
    }


def calculate_ups(hardware):
    """Calculates UPS Load for 60-Min Backup + 15% Contingency"""
    wattages = {
        "24_Port_Switch": 370,
        "NVR": 100,
        "Workstation": 250,
        "Monitor_24": 50,
        "Monitor_32": 75
    }
    
    total_watts = sum(hardware.get(item, 0) * watts for item, watts in wattages.items())
    load_with_contingency = total_watts * 1.15
    kva_needed = 1 if load_with_contingency < 800 else (3 if load_with_contingency < 2400 else 6)
    
    return {
        "total_watts": round(total_watts, 2),
        "load_with_contingency": round(load_with_contingency, 2),
        "ups_kva": kva_needed
    }


def generate_excel_boq(camera_data, storage_data, ups_data):
    """Generates the structured Excel workbook in memory"""
    output = io.BytesIO()
    
    boq_data = {
        "Item Description": [
            f"Dome Camera - Standard/AI ({camera_data.get('Dome', 0)} pcs)",
            f"Bullet Camera ({camera_data.get('Bullet', 0)} pcs)",
            f"ANPR Smart Camera ({camera_data.get('ANPR', 0)} pcs)",
            f"K-POI Camera ({camera_data.get('KPOI', 0)} pcs)",
            "UNV NVR516-64 Network Video Recorder",
            f"16TB HDD ({storage_data['total_hdds']} pcs)",
            f"{ups_data['ups_kva']} KVA UPS System with Battery Pack (60 Min Backup)"
        ],
        "Qty": [
            camera_data.get('Dome', 0), 
            camera_data.get('Bullet', 0), 
            camera_data.get('ANPR', 0), 
            camera_data.get('KPOI', 0), 
            1, 
            storage_data['total_hdds'], 
            1
        ],
        "Unit Cost (QAR)": [200, 371, 1504, 1285, 4350, 2300, 3950]
    }
    df_boq = pd.DataFrame(boq_data)
    df_boq['Total Cost (QAR)'] = df_boq['Qty'] * df_boq['Unit Cost (QAR)']
    df_boq['Sell Price (QAR)'] = round(df_boq['Total Cost (QAR)'] * 1.16, 2)
    
    df_storage = pd.DataFrame({
        "Parameter": ["Required Storage (TB)", "Total 16TB HDDs Required (RAID 5)", "Available Usable Storage (TB)"],
        "Value": [storage_data['required_tb'], storage_data['total_hdds'], storage_data['usable_tb']]
    })

    df_ups = pd.DataFrame({
        "Parameter": ["Total Power Load (Watts)", "Load with 15% Safety Margin (Watts)", "Required UPS Sizing (KVA)"],
        "Value": [ups_data['total_watts'], ups_data['load_with_contingency'], ups_data['ups_kva']]
    })

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_boq.to_excel(writer, sheet_name='BOQ', index=False)
        df_storage.to_excel(writer, sheet_name='Storage', index=False)
        df_ups.to_excel(writer, sheet_name='UPS', index=False)
        
    output.seek(0)
    return output


# --- STREAMLIT USER INTERFACE ---

st.set_page_config(page_title="Flora Security - BOQ Generator", layout="wide")

st.title("Flora Security Systems - CCTV BOQ Generator")
st.markdown("Upload a completed **CCTV Site Visit Survey Form (.docx)** to automatically extract Sections 6, 7, and 8 and generate an MOI-SSD compliant Excel BOQ.")

uploaded_file = st.file_uploader("Upload Site Survey Form (.docx)", type="docx")

if uploaded_file is not None:
    doc = Document(uploaded_file)
    extracted_cameras, extracted_hardware = extract_survey_sections(doc)
    
    st.success("File parsed successfully!")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Extracted Cameras (Sections 6 & 7)")
        st.json(extracted_cameras)
    with col2:
        st.subheader("Extracted Hardware (Section 6)")
        st.json(extracted_hardware)
        
    if st.button("Generate MOI-Compliant BOQ"):
        with st.spinner("Calculating RAID 5 storage, UPS power load, and pricing models..."):
            storage_results = calculate_storage(extracted_cameras)
            ups_results = calculate_ups(extracted_hardware)
            excel_file = generate_excel_boq(extracted_cameras, storage_results, ups_results)
            
            st.success("Calculation Complete!")
            st.download_button(
                label="📥 Download Generated Excel BOQ",
                data=excel_file,
                file_name="Flora_Generated_BOQ.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
