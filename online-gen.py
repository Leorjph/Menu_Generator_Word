import streamlit as st
from generate import parse_text, create_doc
from io import BytesIO, StringIO

st.title("Text to Docx Menu Generator")
output_file_name = "Menu-Iron_Skillet.docx"

uploaded = st.file_uploader(
    "Upload text file",
    type=["txt"]
)

stations = {
    'Iron Skillet' : 'iron_skillet',
    'True Balance' : 'true_balance'
}

station_name = st.selectbox(
    "Select station: ",
    list(stations.keys()),
    index=0
)

scale = st.number_input(
    "Scale",
    min_value=0.1,
    max_value=10.0,
    value=1.0,
    step=0.1,
)

if st.button("Generate Menu"):
    if uploaded:
        text = uploaded.read().decode("utf-8")
        station_name = stations[station_name]
            
        items = parse_text(StringIO(text))
        output_doc = create_doc(station_name, items, scale)
        print(f"\nSuccessfully generated {output_file_name}\n")
        
        output = BytesIO()
        output_doc.save(output)
        output.seek(0)
        
        st.download_button(
            "Download Menu",
            output,
            file_name=output_file_name,
            mime = "application/vnd.openxmlformats-officedocuments.wordprocessingxml.document"
        )
    else:
        st.error("Please upload a text file first")
