import io
import zipfile
import streamlit as st
import os
import tempfile
import time
import subprocess

# Configure Streamlit page
st.set_page_config(
    page_title='DOC2PDF',
    layout="centered",
    initial_sidebar_state="auto",
)

# Hide footer style
hide_streamlit_style = """
<style>
footer {visibility: hidden;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# Function to save uploaded file
def save_uploadedfile(uploadedfile):
    temp_dir = os.path.join(os.getcwd(), "temp_files")
    os.makedirs(temp_dir, exist_ok=True)
    file_path = os.path.join(temp_dir, uploadedfile.name)
    with open(file_path, "wb") as f:
        f.write(uploadedfile.getbuffer())
    return file_path

# Function to convert DOCX to PDF using libreoffice
def convert_to_pdf_stream(docx_file):
    pdf_file = docx_file.replace('.docx', '.pdf').replace('.doc', '.pdf')
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(docx_file), docx_file])
    
    with open(pdf_file, "rb") as pdf:
        pdf_stream = io.BytesIO(pdf.read())
    
    os.remove(pdf_file)
    return pdf_stream.getvalue()

# Function to handle conversion based on input type
def convert_to_pdf(Files, InputType):
    Flag = None
    temp_zip_path = os.path.join(tempfile.gettempdir(), "documents.zip")

    with zipfile.ZipFile(temp_zip_path, "w", zipfile.ZIP_DEFLATED, False) as zip_file:
        try:
            if InputType == 'Path':
                File_path = os.path.abspath(Files)
                for filename in os.listdir(File_path):
                    if filename.endswith(".docx") or filename.endswith(".doc"):
                        file = os.path.join(File_path, filename)
                        pdf_content = convert_to_pdf_stream(file)
                        zip_file.writestr(os.path.splitext(filename)[0] + ".pdf", pdf_content)
            elif InputType == 'Multiple':
                for uploaded_file in Files:
                    file_path = save_uploadedfile(uploaded_file)
                    pdf_content = convert_to_pdf_stream(file_path)
                    zip_file.writestr(os.path.splitext(uploaded_file.name)[0] + ".pdf", pdf_content)
                    os.remove(file_path)
            elif InputType == 'Single':
                file_path = save_uploadedfile(Files)
                pdf_content = convert_to_pdf_stream(file_path)
                zip_file.writestr(os.path.splitext(Files.name)[0] + ".pdf", pdf_content)
                os.remove(file_path)
            Flag = True
        except Exception as e:
            st.write(f'ERROR!!: {str(e)}')
            Flag = False

    return temp_zip_path, Flag

# Main function to render the Streamlit app
def main():
    st.markdown("""<h1 style='background-color: #B8D1D8;font-family: Papyrus, fantasy;text-align: center;'>DOC 2 PDF Converter</h1><br><br>""", unsafe_allow_html=True)

    genre = st.radio(
        "Please select a file option",
        ["Want to upload only one file?", "Want to upload multiple files manually?", "Want to read from a folder path?"],
    )

    st.markdown('<br>', unsafe_allow_html=True)

    if genre == 'Want to upload only one file?':
        input_files = st.sidebar.file_uploader('Upload a docx', type=['docx', 'doc'])
        input_type = 'Single'
    elif genre == "Want to upload multiple files manually?":
        input_files = st.sidebar.file_uploader('Upload multiple docx', accept_multiple_files=True, type=['docx', 'doc'])
        input_type = 'Multiple'
    elif genre == "Want to read from a folder path?":
        input_files = st.text_input('Folder Path', placeholder="Please enter your folder path")
        input_type = 'Path'

    st.markdown('<br><br>', unsafe_allow_html=True)

    pdf_file_path = None
    Flag = None
    col1, col2, col3 = st.columns([0.01, 0.3, 0.1])
    col11, col22, col33 = st.columns([0.1, 0.3, 0.1])

    if input_files:
        with col2:
            if st.button('Generate PDF', type='primary'):
                with col22:
                    st.markdown('<br>', unsafe_allow_html=True)
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    for i in range(100):
                        time.sleep(0.05)
                        progress_bar.progress(i + 1)
                    if pdf_file_path is None:
                        pdf_file_path, Flag = convert_to_pdf(input_files, input_type)
                    if Flag:
                        progress_bar.progress(100)
                        st.snow()
                        status_text.text("Process completed. Please download your files.")

        with col3:
            if Flag:
                with open(pdf_file_path, "rb") as temp_zip_file:
                    if st.download_button('Download Files', data=temp_zip_file, file_name='Documents.zip'):
                        pass
                os.remove(pdf_file_path)
                pdf_file_path = None

if __name__ == "__main__":
    main()
