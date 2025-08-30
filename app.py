import streamlit as st
import pdfplumber
import pytesseract
from PIL import Image
import io
import spacy
import pandas as pd
import re
from PyPDF2 import PdfReader

# Load SpaCy model
nlp = spacy.load("en_core_web_sm")

# Function: Clean text
def clean_text(text):
    text = re.sub(r'\s+', ' ', text)  # Remove extra spaces
    text = re.sub(r'[^a-zA-Z0-9.,;:!?()\-\n ]', '', text)  # Remove weird chars
    return text.strip()

# Function: Extract text from normal PDFs
def extract_text_pdfplumber(file):
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    return text

# Function: Extract text from scanned PDFs (OCR)
def extract_text_ocr(file):
    pdf = PdfReader(file)
    text = ""
    for page_num in range(len(pdf.pages)):
        page = pdf.pages[page_num]
        try:
            x_object = page['/Resources']['/XObject'].get_object()
            for obj in x_object:
                if x_object[obj]['/Subtype'] == '/Image':
                    size = (x_object[obj]['/Width'], x_object[obj]['/Height'])
                    data = x_object[obj]._data
                    img = Image.open(io.BytesIO(data))
                    text += pytesseract.image_to_string(img)
        except:
            continue
    return text

# Function: Extract relevant paragraphs
def extract_relevant_paragraphs(text, keywords, company_name):
    paragraphs = text.split("\n")
    extracted = []
    for para in paragraphs:
        para_clean = clean_text(para)
        if any(kw.lower() in para_clean.lower() for kw in keywords):
            doc = nlp(para_clean)
            extracted.append({
                "Company": company_name,
                "Paragraph": para_clean
            })
    return pd.DataFrame(extracted)

# Streamlit UI
st.set_page_config(page_title="📊 ESG Insights Extractor", layout="wide")
st.title("📊 ESG Insights Extractor")
st.markdown("Upload a PDF report, enter ESG keywords, and extract relevant paragraphs.")

# Inputs
company_name = st.text_input("Enter Company Name", placeholder="e.g., Infosys")

# Upload PDF
uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

# Enter keywords
keywords_input = st.text_area(
    "Enter ESG keywords (comma separated)",
    "environment, sustainability, diversity, carbon, emissions, renewable, governance, ethics, inclusion, recycling, waste, energy efficiency"
)
keywords = [k.strip() for k in keywords_input.split(",") if k.strip()]

# Extract button
if uploaded_file and company_name and st.button("🔍 Extract ESG Paragraphs"):
    st.info("Processing PDF...")

    # Try normal PDF extraction
    text = extract_text_pdfplumber(uploaded_file)

    # If text seems empty, try OCR
    if len(text.strip()) < 50:
        st.warning("No text detected. Trying OCR for scanned PDF...")
        uploaded_file.seek(0)  # Reset file pointer
        text = extract_text_ocr(uploaded_file)

    if text:
        st.success("Text extraction complete ✅")
        df = extract_relevant_paragraphs(text, keywords, company_name)

        if not df.empty:
            st.subheader("✅ Extracted Relevant Paragraphs")
            st.dataframe(df, use_container_width=True)

            # Save to Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="ESG Extracts")
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Download Extracted Paragraphs (Excel)",
                data=excel_data,
                file_name=f"{company_name}_ESG_extracts.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No relevant paragraphs found for given keywords.")
    else:
        st.error("Failed to extract text from PDF.")
