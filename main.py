import streamlit as st
from PIL import Image
from PIL import ImageDraw

st.title('写真表示テスト')
uploaded_file = st.file_uploader('Choose an image…', type='jpg')

if uploaded_file is not None:
    img = Image.open(uploaded_file)
    draw = ImageDraw.Draw(img)
    draw.line([(0, 50), (200, 50), (0, 150), (200, 150)], fill='red', width=5)
    st.image(img, caption='Uploaded Image', use_column_width=True)


