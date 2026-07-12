import streamlit as st
from docx import Document
from docx.shared import Inches
import io
from PIL import Image

# --- Word Document Generator Function ---
def create_inspection_report(images_data, all_sections):
    doc = Document()
    doc.add_heading('Quality Inspection Report', 0)

    for section in all_sections:
        images = images_data.get(section, [])
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
        
        if images:
            for i, img_file in enumerate(images):
                row_idx = (i // num_cols) + 1
                col_idx = i % num_cols
                cell = table.cell(row_idx, col_idx)
                paragraph = cell.paragraphs[0]
                
                run = paragraph.add_run()
                # Reading the uploaded file as bytes for python-docx
                run.add_picture(io.BytesIO(img_file.getvalue()), width=Inches(2.5)) 
        else:
            empty_cell = table.cell(1, 0)
            empty_cell.merge(table.cell(1, num_cols - 1))
            empty_cell.text = "No images provided"
            
        doc.add_paragraph() 
        
    # Save to memory instead of local disk so mobile users can download it
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- UI Design for Mobile/Web ---
st.title("📱 Quality Inspection App")

# State Management (Memory for the app)
if 'selected_images' not in st.session_state:
    st.session_state.selected_images = {}
if 'sections_list' not in st.session_state:
    st.session_state.sections_list = ['MRP', 'Barcode (EAN Code)', 'Outer Box', 'Inner Box', 'Drop Test', 'Functional Testing', 'Visual Inspection']

# Step 1: Manage Sections
st.subheader("1. Select or Add Section")
section_choice = st.selectbox("Choose section to add photos:", st.session_state.sections_list)

with st.expander("Or Add Custom Section"):
    new_section = st.text_input("Section Name")
    if st.button("Add Section"):
        if new_section and new_section not in st.session_state.sections_list:
            st.session_state.sections_list.append(new_section)
            st.success(f"Added '{new_section}'!")
            st.rerun()

# Step 2: Upload Photos (Uses phone camera or gallery on iPhone)
st.subheader("2. Upload Photos")
uploaded_files = st.file_uploader(f"Upload photos for {section_choice}", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

if st.button("Save Photos to Section", type="primary"):
    if uploaded_files:
        if section_choice not in st.session_state.selected_images:
            st.session_state.selected_images[section_choice] = []
        st.session_state.selected_images[section_choice].extend(uploaded_files)
        st.success(f"Successfully added {len(uploaded_files)} photos to {section_choice}!")
    else:
        st.error("Please select photos first.")

# Step 3: Live Preview
if st.session_state.selected_images:
    st.subheader("3. Photo Preview")
    for sec, imgs in st.session_state.selected_images.items():
        st.write(f"**{sec}** ({len(imgs)} photos)")
        # Show images in a 2-column layout suited for mobile
        cols = st.columns(2)
        for i, img in enumerate(imgs):
            with cols[i % 2]:
                st.image(img, use_container_width=True)

# Step 4: Generate Report
st.divider()
st.subheader("4. Finalize Report")
report_name = st.text_input("Report Name", "Quality_Inspection_Report")

if st.button("Generate Word Report", type="primary"):
    if not st.session_state.selected_images:
        st.warning("Please add at least one image before generating the report.")
    else:
        with st.spinner("Generating document..."):
            docx_data = create_inspection_report(st.session_state.selected_images, st.session_state.sections_list)
            
            st.download_button(
                label="📥 Download .docx Report",
                data=docx_data,
                file_name=f"{report_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

if st.button("Reset / Clear All Data"):
    st.session_state.selected_images = {}
    st.rerun()