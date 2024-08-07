import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
import io
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# Function to read and split PDF pages into images
def read_pdf(file):
    pdf_document = fitz.open(stream=file.read(), filetype="pdf")
    images = []
    text = ""
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
        text += page.get_text()
    return images, text

# Streamlit UI
st.title("口腔护理数据分析")
uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")

if uploaded_file is not None:
    pdf_images, pdf_text = read_pdf(uploaded_file)
    
    st.write(f"Total Pages: {len(pdf_images)}")

    # Display each page as an image
    for page_num, img in enumerate(pdf_images):
        st.write(f"Page {page_num + 1}")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        st.image(buf, caption=f"Page {page_num + 1}", use_column_width=True)

    # Generate and display word cloud from text
    if pdf_text:
        st.write("Generating word cloud...")
        wordcloud = WordCloud(width=800, height=400, background_color='white').generate(pdf_text)
        
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis("off")
        
        # Convert matplotlib figure to image
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        st.image(buf, caption="Word Cloud", use_column_width=True)
    else:
        st.write("No text found in PDF.")
