import streamlit as st
import re
import os
import base64
import hashlib
from zipfile import ZipFile

def docx_blind_comment(input_file, output_file):

    docx_file_name = input_file

    files = dict()

    # We read all of the files and store them in "files" dictionary.
    document_as_zip = ZipFile(docx_file_name, 'r')
    for internal_file in document_as_zip.infolist():
        file_reader = document_as_zip.open(internal_file.filename, "r")
        files[internal_file.filename] = file_reader.readlines()
        file_reader.close()
       
    # We don't need to read anything more, so we close the file.
    document_as_zip.close() 

    # If there are any comments.
    if "word/comments.xml" in files.keys():
        # We will be working on comments file...
        comments = files["word/comments.xml"]

        comments_new = str()

        # Files contents have been read as list of byte strings.
        for comment in comments:
            if isinstance(comment, bytes):
                # Change every author to "Unknown Author".
                comments_new += re.sub(r'w:author="[^"]*"', "w:author=\"Unknown Author\"", comment.decode())

        files["word/comments.xml"] = comments_new
        
        
    docx_outputfile_name = output_file

    # Now we want to save old files to the new archive.
    document_as_zip = ZipFile(docx_outputfile_name, 'w')
    for internal_file_name in files.keys():
        # Those are lists of byte strings, so we merge them...
        merged_binary_data = str()
        for binary_data in files[internal_file_name]:
            # If the file was not edited (therefore is not the comments.xml file).
            try:
                if not isinstance(binary_data, str):
                    binary_data = binary_data.decode()
                # Merge file contents.
                merged_binary_data += binary_data
            except:
                #print(internal_file_name)
                pass

        # We write old file contents to new file in new .docx.
        document_as_zip.writestr(internal_file_name, merged_binary_data)

    # Close file for writing.
    document_as_zip.close()




st.title('Wordコメント著者情報削除')
st.markdown(f"Wordファイル（docx）をアップロートしてください。<br>処理終了後、ダウンロードリンクが表示されます。", unsafe_allow_html=True)
uploaded_file = st.file_uploader('下欄にドラッグ＆ドロップできます。', type='docx')


if uploaded_file is not None:

    try:

        output_filename = hashlib.sha224(uploaded_file.name.encode()).hexdigest()#'tmp.docx'
        docx_blind_comment(uploaded_file, output_filename)

        download_filename = uploaded_file.name
        with open(output_filename, mode="rb") as f:
            content = f.read()
            encoded_string = base64.b64encode(content)
            encoded_string = encoded_string.decode()
            href = f'<a href="data:application/docx;base64,{encoded_string}" download="{download_filename}">download</a>'
            st.markdown(f"ダウンロードする {href}", unsafe_allow_html=True)

        os.remove(output_filename)

    except Exception as e:
        st.markdown(f"<b>エラーが発生しました: {str(e)}</b>", unsafe_allow_html=True)

