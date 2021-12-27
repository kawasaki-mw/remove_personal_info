import streamlit as st
import os
import base64
import pikepdf
import hashlib
from docx import Document


st.title('著者情報削除')
st.markdown(f"ファイル（docxもしくはpdf）をアップロートしてください。<br>処理終了後、ダウンロードリンクが表示されます。", unsafe_allow_html=True)
uploaded_file = st.file_uploader('下欄にドラッグ＆ドロップできます。', type=['docx', 'pdf'])

if uploaded_file is not None:

    try:
        root, ext = os.path.splitext(uploaded_file.name)

        if ext == '.docx':
            doc = Document(uploaded_file)

            core_properties = doc.core_properties
            meta_fields= ["author", "category", "last_modified_by", "comments", "content_status", "identifier", "keywords", "language", "subject", "title", "version"]
            for meta_field in meta_fields:
                setattr(core_properties, meta_field, "")

            download_filename = uploaded_file.name
            output_filename = hashlib.sha224(uploaded_file.name.encode()).hexdigest()#'tmp.docx'
            doc.save(output_filename)
            with open(output_filename, mode="rb") as f:
                content = f.read()
                encoded_string = base64.b64encode(content)
                encoded_string = encoded_string.decode()
                href = f'<a href="data:application/docx;base64,{encoded_string}" download="{download_filename}">download</a>'
                st.markdown(f"ダウンロードする {href}", unsafe_allow_html=True)

            os.remove(output_filename)


        elif ext == '.pdf':
            pdf = pikepdf.open(uploaded_file)
            try:
                del(pdf.docinfo)
            except:
                pass
            try:
                del(pdf.Root.Metadata)
            except:
                pass

            download_filename = uploaded_file.name
            output_filename = hashlib.sha224(uploaded_file.name.encode()).hexdigest()#'tmp.pdf'
            pdf.save(output_filename)

            with open(output_filename, "rb") as pdf_file:
                content = pdf_file.read()
                encoded_string = base64.b64encode(content)
                encoded_string = encoded_string.decode()
                href = f'<a href="data:application/pdf;base64,{encoded_string}" download="{download_filename}">download</a>'
                st.markdown(f"ダウンロードする {href}", unsafe_allow_html=True)

            os.remove(output_filename)

    except Exception as e:
        st.markdown(f"<b>エラーが発生しました: {str(e)}</b>", unsafe_allow_html=True)
