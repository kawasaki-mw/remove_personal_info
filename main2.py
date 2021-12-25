import streamlit as st
import os
import base64
from PIL import Image
from PIL import ImageDraw


def file_downloader(filename, file_label='File'):
    with open(filename, 'rb') as f:
        data = f.read()


st.title('著者情報削除')
uploaded_file = st.file_uploader('Wordファイルを選んでください…', type='docx')

if uploaded_file is not None:
    pdf = pikepdf.open(filename)
    pdf.docinfo["/Author"] = ""
    pdf.save('output2.pdf')

    b64 = base64.b64encode(pdf.encode()).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="result.pdf">download</a>'
    st.markdown(f"ダウンロードする {href}", unsafe_allow_html=True)


