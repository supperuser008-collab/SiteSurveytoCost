import streamlit as st
import pandas as pd
import io
from docx import Document

# --- PARSING ENGINE FOR SECTIONS 6, 7, and 8 ---

def parse_survey_tables(doc):
    """
    Parses Sections 6, 7, and 8 from the uploaded Word document tables.
    Extracts item descriptions and quantities written by the technician strictly without assumptions.
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


# --- EXCEL BOQ GENERATOR MATCHING MASTER TEMPLATE STRICTLY ---

def generate_master_boq_excel(survey_items):
    """
    Generates an Excel file matching the template structure, populated 
    ONLY with items and quantities explicitly extracted from the survey.
    """
    output = io.BytesIO()
    boq_rows = []
    
    # Filter items by category
    camera_items = [item for item in survey_items if item['Section'] == 6 and ("cctv" in item['Item'].lower() or "camera" in item['Item'].lower() or "coverage" in item['Item'].lower())]
    nvr_items = [item for item in survey_items if item['Section'] == 6 and ("recording" in item['Item'].lower() or "nvr" in item['Item'].lower() or "server" in item['Item'].lower())]
    storage_items = [item for item in survey_items if item['Section'] == 6 and ("storage" in item['Item'].lower() or "retention" in item['Item'].lower())]
    network_items = [item for item in survey_items if item['Section'] == 6 and ("network" in item['Item'].lower() or "switch" in item['Item'].lower())]
    workstation_items = [item for item in survey_items if item['Section'] == 6 and ("monitoring" in item['Item'].lower() or "workstation" in item['Item'].lower())]
    ups_items = [item for item in survey_items if item['Section'] == 6 and ("power" in item['Item'].lower() or "ups" in item['Item'].lower())]
    vms_items = [item for item in survey_items if item['Section'] == 6 and ("vms" in item['Item'].lower() or "software" in item['Item'].lower())]
    
    anpr_kpoi_items = [item for item in survey_items if item['Section'] == 7]

    # Section A: Cameras and accessories (Only if input exists)
    if camera_items or anpr_kpoi_items:
        boq_rows.append([None, "A", "Cameras and accessories", None, None, None, None, None, None, None, None, None, None, None])
        
        for idx, cam in enumerate(camera_items, start=1):
            desc = f"{cam['Item']}\n{cam['Description']}" if cam['Description'] else cam['Item']
            boq_rows.append([None, idx, desc, cam['Unit'], cam['Qty'], 0, 0, 0, 0, 200, cam['Qty'] * 200, 232.0, cam['Qty'] * 232.0, None])
            
        for idx, ak in enumerate(anpr_kpoi_items, start=len(camera_items) + 1):
            desc = f"{ak['Item']} System\n{ak['Description']}" if ak['Description'] else f"{ak['Item']} System"
            boq_rows.append([None, idx, desc, ak['Unit'], ak['Qty'], 0, 0, 0, 0, 1500, ak['Qty'] * 1500, 1740.0, ak['Qty'] * 1740.0, None])

    # Section B: VMS, NVR & Storage (Only if input exists)
    if nvr_items or storage_items or vms_items:
        boq_rows.append([None, "B", "VMS, NVR & Storage", None, None, None, None, None, None, None, None, None, None, None])
        row_idx = 1
        for nvr in nvr_items:
            boq_rows.append([None, row_idx, f"NVR/Server: {nvr['Description']}", nvr['Unit'], nvr['Qty'], 0, 0, 0, 0, 4350, nvr['Qty'] * 4350, 5046, nvr['Qty'] * 5046, None])
            row_idx += 1
        for stor in storage_items:
            boq_rows.append([None, row_idx, f"Storage: {stor['Description']}", stor['Unit'], stor['Qty'], 0, 0, 0, 0, 2300, stor['Qty'] * 2300, 2668, stor['Qty'] * 2668, None])
            row_idx += 1

    # Section C: Network Switches (Only if input exists)
    if network_items:
        boq_rows.append([None, "C", "Network Switches", None, None, None, None, None, None, None, None, None, None, None])
        for idx, net in enumerate(network_items, start=1):
            boq_rows.append([None, idx, f"{net['Item']}: {net['Description']}", net['Unit'], net['Qty'], 0, 0, 0, 0, 985, net['Qty'] * 985, 1142.6, net['Qty'] * 1142.6, None])

    # Section F: UPS System (Only if input exists)
    if ups_items:
        boq_rows.append([None, "F", "UPS System", None, None, None, None, None, None, None, None, None, None, None])
        for idx, ups in enumerate(ups_items, start=1):
            boq_rows.append([None, idx, f"{ups['Item']}: {ups['Description']}", ups['Unit'], ups['Qty'], 0, 0, 0, 0, 3950, ups['Qty'] * 3950, 4582, ak['Qty'] * 4582 if 'ak' in locals() else 4582, None])

    df_boq = pd.DataFrame(boq_rows, columns=[
        "Unnamed: 0", "No", "Product Description", "UOM", "Qty", 
        "Unit Price Ex-Works", "Total Price Ex-Works", "Shipping", "Customs", 
        "Unit Price (QAR)", "Total Price (QAR)", "Sell Unit Price (QAR)", "Sell Total Price (QAR)", "Remarks"
    ])

    # Supplementary sheets
    df_storage = pd.DataFrame([["PROJECT", "CCTV Survey Extraction", None], ["STATUS", "Parsed strictly from input", None]])
    df_ups = pd.DataFrame([["ESTIMATED UPS", None], ["STATUS", "Parsed strictly from input", None]])

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_boq.to_excel(writer, sheet_name='BOQ', index=False)
        df_storage.to_excel(writer, sheet_name='Storage', index=False, header=False)
        df_ups.to_excel(writer, sheet_name='UPS', index=False, header=False)
        
    output.seek(0)
    return output


# --- STREAMLIT WEB INTERFACE ---

st.set_page_config(page_title="Flora Tech BOQ Generator", layout="wide")

st.title("Flora Technology - MOI CCTV BOQ Generator")
st.markdown("Upload your filled survey Word file. The app will extract **only** the exact items and quantities entered by the technician.")

uploaded_file = st.file_uploader("Upload Site Survey Word Document (.docx)", type="docx")

if uploaded_file is not None:
    doc = Document(uploaded_file)
    extracted_items = parse_survey_tables(doc)
    
    st.success("File parsed successfully!")
    
    if extracted_items:
        st.subheader("Strictly Extracted Items from Survey")
        st.dataframe(pd.DataFrame(extracted_items))
    else:
        st.warning("No quantities greater than 0 found in Sections 6, 7, or 8. Please ensure your survey table has quantities filled in.")

    if st.button("Generate Clean BOQ Excel"):
        with st.spinner("Generating Excel matching exact input..."):
            excel_data = generate_master_boq_excel(extracted_items)
            
            st.success("Excel BOQ generated successfully without extra unrequested items!")
            st.download_button(
                label="📥 Download Clean BOQ Excel",
                data=excel_data,
                file_name="POT27605 - Clean BOQ.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
