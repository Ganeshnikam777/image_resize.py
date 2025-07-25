import os
import shutil
from PIL import Image
from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
import streamlit as st

# Streamlit setup
st.set_page_config(page_title="Image to Word Table Generator", layout="centered")
st.title("📸 Image to Word Table Generator")
st.write("Upload JPEG images and get a Word document with them arranged in a 3×3 inch table.")

# Upload images
uploaded_files = st.file_uploader("Upload JPEG images", type=["jpg", "jpeg"], accept_multiple_files=True)

if uploaded_files:
    temp_dir = "temp_images"
    os.makedirs(temp_dir, exist_ok=True)

    doc = Document()
    doc.add_heading("Photo Table", level=1)
    images_per_row = 3
    table = doc.add_table(rows=0, cols=images_per_row)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, uploaded_file in enumerate(uploaded_files):
        # Save uploaded file temporarily
        image_path = os.path.join(temp_dir, uploaded_file.name)
        with open(image_path, "wb") as f:
            f.write(uploaded_file.getvalue())

        # Resize and save image
        resized_path = os.path.join(temp_dir, f"resized_{uploaded_file.name}")
        with Image.open(image_path) as img:
            img = img.resize((300, 300))  # Resize to approx 3x3 inches at 100 DPI
            img.save(resized_path)

        if i % images_per_row == 0:
            row_cells = table.add_row().cells

        row_cells[i % images_per_row].paragraphs[0].add_run().add_picture(resized_path, width=Inches(3), height=Inches(3))

    # Save Word document
    output_docx_path = "output_images_table.docx"
    doc.save(output_docx_path)

    # Offer download
    with open(output_docx_path, "rb") as file:
        st.success("✅ Word document generated!")
        st.download_button("Download Word File", file, file_name="photos_table.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    # Cleanup
    shutil.rmtree(temp_dir)
