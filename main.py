import streamlit as st
import os
import base64
import pikepdf


st.title('著者情報削除')
uploaded_file = st.file_uploader('Wordファイルを選んでください…', type='pdf')

if uploaded_file is not None:
    pdf = pikepdf.open(uploaded_file)
    pdf.docinfo["/Author"] = ""
    pdf.save('tmp.pdf')

    with open("tmp.pdf", "rb") as pdf_file:
        encoded_string = base64.b64encode(pdf_file.read())
        href = f'<a href="data:application/pdf;base64,{encoded_string}" download="result.pdf">download</a>'
        #href = f'<a download="result.pdf" href="data:application/pdf;base64,{encoded_string}">download</a>'
        st.markdown(f"ダウンロードする {href}", unsafe_allow_html=True)

    os.remove('tmp.pdf')

