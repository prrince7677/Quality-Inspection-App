import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
from PIL import Image
import zipfile
import os
import datetime

# --- Word Document Generator Function ---
def create_inspection_report(images_data, videos_data, all_sections, details):
    doc = Document()
    
    # 1. Add Lifelong Logo (Center Aligned)
    if os.path.exists("logo.png"):
        doc.add_picture("logo.png", width=Inches(2.0))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph() # Add space after logo
        
    doc.add_heading('Quality Inspection Report', 0)
    
    # 2. Add Inspection Details Table
    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.style = 'Table Grid'
    
    detail_items = [
        ("Inspection Date", str(details['date'])),
        ("Vendor Name", details['vendor']),
        ("SKU ID", details['sku']),
        ("Produced Qty", str(details['produced'])),
        ("Inspected Qty", str(details['inspected'])),
        ("Lot Status", details['status'])
    ]
    
    for i, (key, val) in enumerate(detail_items):
        row_cells = meta_table.rows[i].cells
        row_cells[0].text = key
        row_cells[1].text = val if val else "N/A"
        row_cells[0].paragraphs[0].runs[0].bold = True
        
    doc.add_paragraph() # Add space after table
    
    # 3. Add Evidence Sections
    for section in all_sections:
        images = images_data.get(section, [])
        videos = videos_data.get(section, [])
        
        num_cols = 2 
        num_image_rows = max(1, (len(images) + num_cols - 1) // num_cols) if images else 1
        
        table = doc.add_table(rows=num_image_rows + 1, cols=num_cols)
        table.style = 'Table Grid' 
        
        # Section Header
        header_cell = table.cell(0, 0)
        header_cell.merge(table.cell(0, num_cols - 1))
        header_paragraph = header_cell.paragraphs[0]
        run = header_paragraph.add_run(section)
        run.bold = True 
        
        # Add Images to Table
        if images:
            for i, img_file in enumerate(images):
                row_idx = (i // num_cols) + 1
                col_idx = i % num_cols
                cell = table.cell(row_idx, col_idx)
                paragraph = cell.paragraphs[0]
                
                run = paragraph.add_run()
                run.add_picture(io.BytesIO(img_file.getvalue()), width=Inches(2.5)) 
        else:
            empty_cell = table.cell(1, 0)
            empty_cell.merge(table.cell(1, num_cols - 1))
            empty_cell.text = "No images provided"
        
        doc.add_paragraph()
        
        # Add Video Names below the table if any exist
        if videos:
            vid_p = doc.add_paragraph()
            vid_run = vid_p.add_run(f"🎥 Attached Video Evidence for {section}:")
            vid_run.bold = True
            for vid in videos:
                doc.add_paragraph(f" - {vid.name}", style='List Bullet')
            doc.add_paragraph()
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- ZIP Generator Function (Bundles Doc + Videos) ---
def create_zip_package(docx_data, videos_data, report_name):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        zip_file.writestr(f"{report_name}.docx", docx_data)
        for sec, vids in videos_data.items():
            for vid in vids:
                vid.seek(0)
                zip_file.writestr(f"Videos/{sec}_{vid.name}", vid.read())
    return zip_buffer.getvalue()

# --- UI Design for Mobile/Web ---
st.title("📱 Quality Inspection App")

# State Management
if 'selected_images' not in st.session_state:
    st.session_state.selected_images = {}
if 'selected_videos' not in st.session_state:
    st.session_state.selected_videos = {}
if 'sections_list' not in st.session_state:
    st.session_state.sections_list = ['MRP', 'Barcode (EAN Code)', 'Outer Box', 'Inner Box', 'Drop Test', 'Functional Testing', 'Visual Inspection']

# Step 1: Inspection Details Form
st.subheader("1. Inspection Details")
col1, col2 = st.columns(2)
with col1:
    insp_date = st.date_input("Inspection Date", datetime.date.today())
    vendor_name = st.text_input("Vendor Name")
    sku_id = st.text_input("SKU ID")
with col2:
    produced_qty = st.number_input("Produced Qty", min_value=0)
    inspected_qty = st.number_input("Inspected Qty", min_value=0)
    lot_status = st.selectbox("Lot Status", ["Pending", "Accepted", "Rejected"])

st.divider()

# Step 2: Manage Sections
st.subheader("2. Select or Add Section")
section_choice = st.selectbox("Choose section:", st.session_state.sections_list)
with st.expander("Or Add Custom Section"):
    new_section = st.text_input("Section Name")
    if st.button("Add Section"):
        if new_section and new_section not in st.session_state.sections_list:
            st.session_state.sections_list.append(new_section)
            st.success(f"Added '{new_section}'!")
            st.rerun()

# Step 3: Upload Photos & Videos
st.subheader("3. Upload Evidence")
uploaded_files = st.file_uploader(f"Upload photos & videos for {section_choice}", type=['jpg', 'jpeg', 'png', 'mp4', 'mov'], accept_multiple_files=True)
if st.button("Save Files to Section", type="primary"):
    if uploaded_files:
        if section_choice not in st.session_state.selected_images:
            st.session_state.selected_images[section_choice] = []
        if section_choice not in st.session_state.selected_videos:
            st.session_state.selected_videos[section_choice] = []
            
        for file in uploaded_files:
            if file.name.lower().endswith(('.mp4', '.mov')):
                st.session_state.selected_videos[section_choice].append(file)
            else:
                st.session_state.selected_images[section_choice].append(file)
                
        st.success(f"Successfully added {len(uploaded_files)} file(s) to {section_choice}!")
    else:
        st.error("Please select files first.")

# Step 4: Live Preview
if st.session_state.selected_images or st.session_state.selected_videos:
    st.subheader("4. Evidence Preview")
    for sec in st.session_state.sections_list:
        imgs = st.session_state.selected_images.get(sec, [])
        vids = st.session_state.selected_videos.get(sec, [])
        
        if imgs or vids:
            st.write(f"**{sec}** ({len(imgs)} photos, {len(vids)} videos)")
            if imgs:
                cols = st.columns(2)
                for i, img in enumerate(imgs):
                    with cols[i % 2]:
                        st.image(img, use_container_width=True)
            if vids:
                for vid in vids:
                    st.info(f"🎥 {vid.name}")

# Step 5: Generate Report
st.divider()
st.subheader("5. Finalize Report")
report_name = st.text_input("Report Name", f"{sku_id}_Inspection_Report" if sku_id else "Quality_Inspection_Report")

if st.button("Generate & Download Report", type="primary"):
    if not st.session_state.selected_images and not st.session_state.selected_videos:
        st.warning("Please add at least one image or video before generating the report.")
    else:
        with st.spinner("Generating document and packaging videos..."):
            
            # Pack the details into a dictionary to send to the Word function
            report_details = {
                'date': insp_date,
                'vendor': vendor_name,
                'sku': sku_id,
                'produced': produced_qty,
                'inspected': inspected_qty,
                'status': lot_status
            }
            
            docx_data = create_inspection_report(st.session_state.selected_images, st.session_state.selected_videos, st.session_state.sections_list, report_details)
            
            has_videos = any(len(vids) > 0 for vids in st.session_state.selected_videos.values())
            
            if has_videos:
                final_data = create_zip_package(docx_data, st.session_state.selected_videos, report_name)
                final_filename = f"{report_name}.zip"
                mime_type = "application/zip"
            else:
                final_data = docx_data
                final_filename = f"{report_name}.docx"
                mime_type = "application/octet-stream" 
                
            st.download_button(
                label="📥 Download Report Files",
                data=final_data,
                file_name=final_filename,
                mime=mime_type
            )

if st.button("Reset / Clear All Data"):
    st.session_state.selected_images = {}
    st.session_state.selected_videos = {}
    st.rerun()