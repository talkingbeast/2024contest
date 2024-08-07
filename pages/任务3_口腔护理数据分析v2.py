import streamlit as st
from docx import Document
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import re
import io  # 导入 io 模块

# Function to read Word file and extract text
def read_word(file):
    document = Document(file)
    text = ""
    for para in document.paragraphs:
        text += para.text
    return text

# Function to filter and retain only Chinese characters from the text
def filter_chinese(text):
    return ''.join(re.findall(r'[\u4e00-\u9fff]+', text))

# Streamlit UI
st.title("口腔护理数据分析")
uploaded_file = st.file_uploader("Upload a Word file", type="docx")

if uploaded_file is not None:
    word_text = read_word(uploaded_file)

    # Filter and retain only Chinese text
    chinese_text = filter_chinese(word_text)

    # Generate and display word cloud from text
    if chinese_text:
        st.write("Generating word cloud...")
        wordcloud = WordCloud(width=800, height=400, background_color='white', font_path='simsun.ttc').generate(chinese_text)
        
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis("off")
        
        # Convert matplotlib figure to image
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        st.image(buf, caption="Word Cloud", use_column_width=True)
    else:
        st.write("No valid text found in Word file.")
