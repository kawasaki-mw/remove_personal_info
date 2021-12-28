import os
import re
import io
import zipfile
import tempfile
import hashlib
import base64
from zipfile import ZipFile


def doc_blind_comment(input_file, output_file):

    # generate a temp file
    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(input_file))
    os.close(tmpfd)

    # filename
    srcfile = input_file  # docx file
    dstfile = tmpname

    with zipfile.ZipFile(srcfile) as inzip, zipfile.ZipFile(dstfile, "w") as outzip:
        # Iterate the input files
        for inzipinfo in inzip.infolist():
            # Read input file
            with inzip.open(inzipinfo) as infile:

                if inzipinfo.filename.startswith("word/comments.xml"):

                    comments = infile.read()
                    comments_new = str()
                    comments_new += re.sub(r'w:author="[^"]*"', "w:author=\"Anonymous Author\"", comments.decode())
                    outzip.writestr(inzipinfo.filename, comments_new)

                else: # Other file, dont want to modify => just copy it

                    outzip.writestr(inzipinfo.filename, infile.read())


    # replace with the temp archive
    if os.path.exists(output_file):
        os.remove(output_file)
    os.rename(tmpname, output_file)
    if os.path.exists(tmpname):
        os.remove(tmpname)



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

