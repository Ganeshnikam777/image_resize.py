import os
import shutil
from PIL import Image
from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
import streamlit as st

# Set up Streamlit page
st.set_page_config(page_title="Passport Photo Generator", layout="centered")
st.title("🛂 Passport Photo Table Generator")
st.write("Upload JPEG images and download a Word file with 2×2 inch passport-style photos arranged in a table.")

# Upload images
uploaded_files = st.file_uploader("Upload JPEG images", type=["jpg", "jpeg"], accept_multiple_files=True)

if uploaded_files:
    temp_dir = "temp_images"
    os.makedirs(temp_dir, exist_ok=True)

    # Create Word document
    doc = Document()
    doc.add_heading("Passport Photo Table", level=1)

    images_per_row = 4  # Four passport photos per row
    table = doc.add_table(rows=0, cols=images_per_row)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, uploaded_file in enumerate(uploaded_files):
        # Save uploaded image temporarily
        image_path = os.path.join(temp_dir, uploaded_file.name)
        with open(image_path, "wb") as f:
            f.write(uploaded_file.getvalue())

        # Add a new row after every 4 images
        if i % images_per_row == 0:
            row_cells = table.add_row().cells

        # Insert image at display size 2x2 inches
        row_cells[i % images_per_row].paragraphs[0].add_run().add_picture(
            image_path, width=Inches(2), height=Inches(2)
        )

    # Save Word document
    output_docx_path = "passport_photos_table.docx"
    doc.save(output_docx_path)

    # Offer download button
    with open(output_docx_path, "rb") as file:
        st.success("✅ Your Word document is ready!")
        st.download_button(
            "Download Word File",
            file,
            file_name="passport_photos.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    # Clean up temporary folder
    shutil.rmtree(temp_dir)
