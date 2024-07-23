import io
import zipfile
import streamlit as st
import subprocess
import os
import tempfile
import time

st.set_page_config(
    page_title='DOC2PDF', 
    layout="centered", 
    initial_sidebar_state="auto", 
    menu_items=None
)

hide_streamlit_style = """
            <style>
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

def save_uploadedfile(uploadedfile):
    temp_dir = os.path.join(os.getcwd(), "temp_files")
    os.makedirs(temp_dir, exist_ok=True)
    file_path = os.path.join(temp_dir, uploadedfile.name)
    with open(file_path, "wb") as f:
        f.write(uploadedfile.getbuffer())
    return file_path

def convert_to_pdf_stream(docx_file):
    output_pdf = os.path.splitext(docx_file)[0] + '.pdf'
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', docx_file, '--outdir', os.path.dirname(output_pdf)], check=True)
    
    with open(output_pdf, "rb") as pdf_file:
        pdf_stream = io.BytesIO(pdf_file.read())
    
    os.remove(output_pdf)
    return pdf_stream.getvalue()

def convertToPdf(Files, InputType):
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

def main():
    st.markdown("<h1 style='background-color: #B8D1D8;font-family: Papyrus, fantasy;text-align: center;box-shadow: rgba(50, 50, 93, 0.25) 0px 50px 100px -20px, rgba(0, 0, 0, 0.3) 0px 30px 60px -30px, rgba(10, 37, 64, 0.35) 0px -2px 6px 0px inset;'>DOC 2 PDF Converter</h1><br/><br/>", unsafe_allow_html=True)

    st.markdown("""
        <style>
        
        [class="main css-uf99v8 ea3mdgi5"],[class="css-qg4qf ezrtsby2"]{
            background-color: rgb(133, 157, 216);
        }
        
        .stRadio > div {
            background-color: #4D898C;
            padding: 10px;
            border-radius: 10px;
            box-shadow: rgba(50, 50, 93, 0.25) 0px 50px 100px -20px, rgba(0, 0, 0, 0.3) 0px 30px 60px -30px, rgba(10, 37, 64, 0.35) 0px -2px 6px 0px inset;
        }
        
        .stRadio > label {
            font-size:50%;
        }
        
        button[kind="primary"] {
        background-color:#C1C8E4;
        font-weight:bold;
        box-shadow: rgba(240, 46, 170, 0.4) -5px 5px, rgba(240, 46, 170, 0.3) -10px 10px, rgba(240, 46, 170, 0.2) -15px 15px, rgba(240, 46, 170, 0.1) -20px 20px, rgba(240, 46, 170, 0.05) -25px 25px;
        
        }
        
        button[kind="primary"]:hover {
        background-color: #DEB89B;
        }
    
        [class="css-mrc9ir ef3psqc11"] {
            background-color: #7FAD9C;
            font-weight:bold;
            box-shadow: rgba(240, 46, 170, 0.4) 5px 5px, rgba(240, 46, 170, 0.3) 10px 10px, rgba(240, 46, 170, 0.2) 15px 15px, rgba(240, 46, 170, 0.1) 20px 20px, rgba(240, 46, 170, 0.05) 25px 25px;
        }
        
        [class="css-mrc9ir ef3psqc11"]:hover {
            background-color: #DEB89B;
        }
        
        [class="st-bo st-bp st-br st-bq st-bt st-dy st-b8 st-e1 st-e0"]{
            background-color: #4D898C;
        }
        
        </style>
    """, unsafe_allow_html=True)

    genre = st.radio(
        "Please select a file option",
        ["Want to upload only one file?", "Want to upload multiple files manually?", "Want to read from a folder path?"],
    )

    st.markdown('<br/>', unsafe_allow_html=True)

    if genre == 'Want to upload only one file?':
        input_files = st.sidebar.file_uploader('Upload a docx', type=['docx', 'doc'])
        input_type = 'Single'

    elif genre == "Want to upload multiple files manually?":
        input_files = st.sidebar.file_uploader('Upload multiple docx', accept_multiple_files=True, type=['docx', 'doc'])
        input_type = 'Multiple'
        
    elif genre == "Want to read from a folder path?":
        input_files = st.text_input('Folder Path', placeholder="Please enter your folder path")
        input_type = 'Path'

    st.markdown('<br/><br/>', unsafe_allow_html=True)
        
    pdf_file_path = None
    Flag = None
    col1, col2, col3 = st.columns([0.01, 0.3, 0.1])
    col11, col22, col33 = st.columns([0.1, 0.3, 0.1])

    if input_files:
        with col2:
            if st.button('Generate PDF', type='primary'):
                with col22:
                    st.markdown('<br/>', unsafe_allow_html=True)
                    progress_bar = st.progress(0)
                    status_text = st.empty()  # Initialize the progress bar
                    for i in range(100):  # Simulate progress
                        time.sleep(0.05)  # Simulate work being done
                        progress_bar.progress(i + 1)  # Update progress bar
                    if pdf_file_path is None:
                        pdf_file_path, Flag = convertToPdf(input_files, input_type)
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

main()
