import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
import io

# Function to read and split PDF pages into images
def read_pdf(file):
    pdf_document = fitz.open(stream=file.read(), filetype="pdf")
    images = []
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    return images

# Streamlit UI
st.title("PDF Page Splitter")
uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")

if uploaded_file is not None:
    pdf_images = read_pdf(uploaded_file)
    
    st.write(f"Total Pages: {len(pdf_images)}")

    # Display each page as an image
    for page_num, img in enumerate(pdf_images):
        st.write(f"Page {page_num + 1}")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        st.image(buf, caption=f"Page {page_num + 1}", use_column_width=True)
