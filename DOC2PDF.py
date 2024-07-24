import io
import zipfile
import streamlit as st
import comtypes.client
import pythoncom
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

# footer = """
#     <style>
#         .footer {
#             position: fixed;
#             font-weight: 300; 
#             bottom: 0;
#             width: 100%;
#             color: black;
#             text-align: center;
#             padding: 10px;
#         }
#     </style>
#     <div class="footer">
#         <h2>Copyright at Gadde</h2>
#     </div>
# """
# st.markdown(footer, unsafe_allow_html=True)



def save_uploadedfile(uploadedfile):
    temp_dir = os.path.join(os.getcwd(), "temp_files")
    os.makedirs(temp_dir, exist_ok=True)
    file_path = os.path.join(temp_dir, uploadedfile.name)
    with open(file_path, "wb") as f:
        f.write(uploadedfile.getbuffer())
    return file_path

def convert_to_pdf_stream(docx_file):
    pythoncom.CoInitialize()
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(docx_file)
    temp_pdf_path = os.path.abspath("temp.pdf")
    doc.SaveAs(temp_pdf_path, FileFormat=17)  
    doc.Close()
    word.Quit()
    
    with open(temp_pdf_path, "rb") as pdf_file:
        pdf_stream = io.BytesIO(pdf_file.read())
    
    os.remove(temp_pdf_path)
    return pdf_stream.getvalue()

def converToPdf(Files, InputType):
    Flag=None
    temp_zip_path = os.path.join(tempfile.gettempdir(), "documents.zip")

    with zipfile.ZipFile(temp_zip_path, "w", zipfile.ZIP_DEFLATED, False) as zip_file:
        try:
            if InputType == 'Multiple':
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
            Flag=True
        except Exception as e:
            st.write(f'ERROR!!: {str(e)}')
            Flag=False

    return temp_zip_path,Flag


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
        ["Want to upload single file?", "Want to upload multiple files?"],
    )
    
    st.markdown('<br/>', unsafe_allow_html=True)

    if genre == 'Want to upload single file?':
        input_files = st.sidebar.file_uploader('Upload a docx', type=['docx', 'doc'])
        input_type = 'Single'

    elif genre == "Want to upload multiple files?":
        st.sidebar.info('You can use Ctrl+A', icon="ℹ️")
        input_files = st.sidebar.file_uploader('Upload multiple docx', accept_multiple_files=True, type=['docx', 'doc'])
        input_type = 'Multiple'

    st.markdown('<br/><br/>', unsafe_allow_html=True)
        
    pdf_file_path=None
    Flag=None
    row1Col1, row1Col2,row1Col3 = st.columns([0.01, 0.3,0.1])
    row2Col1, row2Col2,row2Col3 = st.columns([0.1, 0.3,0.1])
    row3Col1, row3Col2,row3Col3 = st.columns([0.1, 0.21,0.1])

    if input_files:
        with row1Col2:
            if st.button('Generate PDF', type='primary'):
                with row2Col2:
                    st.markdown('<br/>', unsafe_allow_html=True)
                    progress_bar = st.progress(0) 
            
                    for i in range(100): 
                        time.sleep(0.05)  
                        progress_bar.progress(i + 1)  
                    if pdf_file_path is None:
                        pdf_file_path, Flag = converToPdf(input_files, input_type)
                    if Flag:
                        progress_bar.progress(100) 
                        st.snow() 
                        st.markdown('<br/>', unsafe_allow_html=True)
                        # st.success('Process completed. Please download your files!.', icon="✅")
                        with row3Col2:
                            st.success('Woohoo! Your files are ready to download!.', icon="✅")


        with row1Col3:    
            if Flag:
                with open(pdf_file_path, "rb") as temp_zip_file:
                    if st.download_button('Download Files', data=temp_zip_file, file_name='Documents.zip'):
                        pass
                
                os.remove(pdf_file_path)
                pdf_file_path = None  
            

main()
