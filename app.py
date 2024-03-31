import streamlit as st
import docx2pdf
from PyPDF2 import PdfReader, PdfWriter
import io
import os
from multiprocessing import Process, Queue

# Function to convert Word to PDF
def convert_word_to_pdf(file_path):
    try:
        docx2pdf.convert(file_path)
        st.success("Conversion successful!")
    except Exception as e:
        st.error(f"Error converting file: {e}")

# Function to convert PDF to Word
def convert_pdf_to_word(file_path, output_queue):
    try:
        pdf_reader = PdfReader(file_path)
        pdf_writer = PdfWriter()

        for page_num in range(len(pdf_reader.pages)):
            page_obj = pdf_reader.pages[page_num]
            pdf_writer.add_page(page_obj)

        word_output = io.BytesIO()
        pdf_writer.write(word_output)
        word_output.seek(0)

        output_queue.put(word_output)
    except Exception as e:
        output_queue.put(e)

# Streamlit UI
st.title("File Converter")
option = st.selectbox("Select conversion type:", ("Word to PDF", "PDF to Word"))

if option == "Word to PDF":
    st.subheader("Convert Word to PDF")
    word_file = st.file_uploader("Upload a Word file (.docx)", type=["docx"])
    if st.button("Convert"):
        if word_file is not None:
            file_path = os.path.join(os.getcwd(), word_file.name)
            with open(file_path, "wb") as f:
                f.write(word_file.getbuffer())
            convert_word_to_pdf(file_path)
            os.remove(file_path)  # Remove the temporary Word file after conversion
        else:
            st.warning("Please upload a Word file.")

elif option == "PDF to Word":
    st.subheader("Convert PDF to Word")
    pdf_file = st.file_uploader("Upload a PDF file", type=["pdf"])
    if st.button("Convert"):
        if pdf_file is not None:
            file_path = os.path.join(os.getcwd(), pdf_file.name)
            with open(file_path, "wb") as f:
                f.write(pdf_file.getbuffer())
            
            output_queue = Queue()
            process = Process(target=convert_pdf_to_word, args=(file_path, output_queue))
            process.start()
            process.join()

            result = output_queue.get()

            if isinstance(result, io.BytesIO):
                st.success("Conversion successful!")
                st.download_button("Download Word file", result, file_name="converted_word_file.docx", mime="application/octet-stream")
            else:
                st.error(f"Error converting file: {result}")

            os.remove(file_path)  # Remove the temporary PDF file after conversion
        else:
            st.warning("Please upload a PDF file.")
