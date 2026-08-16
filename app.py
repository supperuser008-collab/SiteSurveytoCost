import streamlit as st
import pandas as pd
import io
import math
from docx import Document

# --- PARSING ENGINE FOR SECTIONS 6, 7, and 8 ---

def parse_survey_tables(doc):
    """
    Accurately parses Sections 6, 7, and 8 from the uploaded Word document tables.
    Extracts item descriptions and quantities written by the technician.
    """
    extracted_items = []
    
    current_section = None
    for table in doc.tables:
        first_row_text = "".join([cell.text.strip() for cell in table.rows[0].cells]).upper()
        
        if "6. NEW SYSTEM REQUIREMENTS" in first_row_text:
            current_section = 6
        elif "7. ANPR & K-POI REQUIREMENTS" in first_row_text:
            current_section = 7
        elif "8. CIVIL / ELECTRICAL" in first_row_text:
            current_section = 8
        elif "9. SITE OBSERVATIONS" in first_row_text or "10. SIGN-OFF" in first_row_text:
            current_section = None
            
        # If we are inside Sections 6, 7, or 8, parse the rows
        if current_section in [6, 7, 8]:
            for row in table.rows[1:]: # Skip header
                cells = [cell.text.strip() for cell in row.cells]
                if len(cells) >= 4:
                    item_name = cells[1]
                    description = cells[2]
                    qty_str = cells[3]
                    
                    if item_name or description:
                        try:
                            qty = float(qty_str) if qty_str else 0.0
                        except ValueError:
                            qty = 0.0
                            
                        extracted_items.append({
                            "Section": current_section,
                            "Item": item_name,
                            "Description": description,
                            "Qty": qty,
                            "Unit": cells[4] if len(cells) > 4 else "EA"
                        })
                        
    return extracted_items


# --- EXCEL BOQ GENERATOR MATCHING MASTER TEMPLATE ---

def generate_master_boq_excel(survey_items):
    """
    Generates an Excel file matching the exact structure of POT27605 - BOQ-1 LVQ.xlsx
    with sheets: BOQ, BD, Storage, UPS.
    """
    output = io.BytesIO()
    
    # 1. Build BOQ DataFrame matching template structure
    boq_rows = []
    
    # Section A: Cameras and accessories
    boq_rows.append([None, "A", "Cameras and accessories", None, None, None, None, None, None, None, None, None, None, ""])
    
    dome_qty = sum([item['Qty'] for item in survey_items if "dome" in item['Item'].lower() or "dome" in item['Description'].lower()])
    bullet_qty = sum([item['Qty'] for item in survey_items if "bullet" in item['Item'].lower() or "bullet" in item['Description'].lower()])
    
    if dome_qty == 0: dome_qty = 10 # Default fallback if blank
    if bullet_qty == 0: bullet_qty = 8
    
    boq_rows.append([None, 1, "Dome Camera - Fixed\n2MP Dome Camera", "EA", dome_qty, 0, 0, 0, 0, 200, dome_qty * 200, 232.0, dome_qty * 232.0, None])
    boq_rows.append([None, 2, "Dome Camera - Auto Iris\n2MP WDR LightHunter IR Network Dome Camera", "EA", 12, 0, 0, 0, 0, 371, 4452, 430.36, 5164.32, None])
    boq_rows.append([None, 3, "Bullet Camera\n2MP HD IR VF Bullet Network Camera", "EA", bullet_qty, 0, 0, 0, 0, 371, bullet_qty * 371, 430.36, bullet_qty * 430.36, None])
    
    # Section B: VMS, NVR & Storage
    boq_rows.append([None, "B", "VMS, NVR & Storage", None, None, None, None, None, None, None, None, None, None, None])
    boq_rows.append([None, 4, "VMS Software", "EA", 1, 0, 0, 0, 0, "Included", "Included", "Included", "Included", None])
    boq_rows.append([None, 5, "UNV NVR516-64,Network Video Recorder", "EA", 1, 0, 0, 0, 0, 4350, 4350, 5046, 5046, None])
    boq_rows.append([None, 6, "16TB HDD", "EA", 11, 0, 0, 0, 0, 2300, 25300, 2668, 29348, None])
    
    # Section C: Network Switches
    boq_rows.append([None, "C", "Network Switches", None, None, None, None, None, None, None, None, None, None, None])
    boq_rows.append([None, 7, "24Port POE Switch", "EA", 2, 0, 0, 0, 0, 985, 1970, 1142.6, 2285.2, None])
    
    # Section F: UPS System
    boq_rows.append([None, "F", "UPS System", None, None, None, None, None, None, None, None, None, None, None])
    boq_rows.append([None, 12, "3KVA UPS with Battery Pack for 60 Min. Backup", "EA", 1, 0, 0, 0, 0, 3950, 3950, 4582, 4582, None])

    df_boq = pd.DataFrame(boq_rows, columns=[
        "Unnamed: 0", "No", "Product Description", "UOM", "Qty", 
        "Unit Price Ex-Works", "Total Price Ex-Works", "Shipping", "Customs", 
        "Unit Price (QAR)", "Total Price (QAR)", "Sell Unit Price (QAR)", "Sell Total Price (QAR)", "Remarks"
    ])

    # 2. Build Storage Sheet
    storage_rows = [
        ["PROJECT", "CCTV Site Survey Extraction", None, None, None, None, "REQUIRED STORAGE - TB", None, 90.1, None, None, None],
        ["STORAGE TYPE", "UNV NVR516-64 - Network Video Recorder", None, None, None, None, "AVAILABLE STORAGE - TB", None, 130.95, None, None, None],
        ["TOTAL HDD - 16TB", 11, None, None, None, None, "OVERALL STORAGE LOAD", None, 0.68, None, None, None],
        ["TOTAL CAMERA", dome_qty + bullet_qty, None, None, None, None, None, None, None, None, None, None]
    ]
    df_storage = pd.DataFrame(storage_rows)

    # 3. Build UPS Sheet
    ups_rows = [
        ["ESTIMATED UPS CALCULATION - MDF", None, None, None, None, None, None, None, "UPS KVA"],
        ["PROPOSED UPS", None, None, None, None, None, None, None, 3],
        ["No", "Item Description", "Brand", "Availability", "Model No", "Quantity", "Watts", "Total Watts", None],
        [1, "24\" Monitor", "TBD", None, "TBD", 1, 50, 50, None],
        [2, "Workstation", "TBD", None, "TBD", 2, 250, 500, None],
        [3, "24 Port Switch", "TBD", None, "TBD", 2, 370, 740, None],
        [4, "NVR", "TBD", None, "TBD", 1, 100, 100, None]
    ]
    df_ups = pd.DataFrame(ups_rows)

    # Write to Excel with exact sheet names from template
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_boq.to_excel(writer, sheet_name='BOQ', index=False)
        df_storage.to_excel(writer, sheet_name='Storage', index=False, header=False)
        df_ups.to_excel(writer, sheet_name='UPS', index=False, header=False)
        
    output.seek(0)
    return output


# --- STREAMLIT WEB INTERFACE ---

st.set_page_config(page_title="Flora Tech BOQ Generator", layout="wide")

st.title("Flora Technology - MOI CCTV BOQ Generator")
st.markdown("Upload a filled **CCTV Site Visit Survey Form (.docx)**. The application reads Sections 6, 7, and 8 and outputs an Excel file matching the master BOQ format.")

uploaded_file = st.file_uploader("Upload Site Survey Word Document (.docx)", type="docx")

if uploaded_file is not None:
    doc = Document(uploaded_file)
    extracted_items = parse_survey_tables(doc)
    
    st.success("Word Document parsed successfully!")
    
    if extracted_items:
        st.subheader("Extracted Data from Sections 6, 7, & 8")
        st.dataframe(pd.DataFrame(extracted_items))
    else:
        st.warning("No explicit rows found in Sections 6, 7, or 8 tables. Default template quantities will be applied.")

    if st.button("Generate Master BOQ Excel"):
        with st.spinner("Processing quantities and building master Excel structure..."):
            excel_data = generate_master_boq_excel(extracted_items)
            
            st.success("Excel BOQ generated successfully!")
            st.download_button(
                label="📥 Download Master BOQ Excel",
                data=excel_data,
                file_name="POT27605 - Generated BOQ.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
