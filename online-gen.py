import streamlit as st
from generate import parse_text, create_doc, image_path, TAGS
from io import BytesIO, StringIO

st.title("Text to Docx Menu Generator")

uploaded = st.file_uploader(
    "Upload text file",
    type=["txt"]
)

stations = {
    'Iron Skillet' : 'iron_skillet',
    'True Balance' : 'true_balance',
    'Menu of the Day' : 'motd'
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
        
        output = BytesIO()
        output_doc.save(output)
        output.seek(0)
        
        st.download_button(
            "Download Menu",
            output,
            file_name="Menu_" + station_name + ".docx",
            mime = "application/vnd.openxmlformats-officedocuments.wordprocessingxml.document"
        )
    else:
        st.error("Please upload a text file first")


with st.sidebar.expander("? - Help"):
    st.subheader("Text file format")
    st.write("Each item consists for 4 lines-\n")
    st.write("1. Item name")
    st.write("Ingredients")
    st.write("Tags (Comma-separated)")
    st.write("Allergen note")
    st.write("\n Use '-' in any column other than name to leave it empty\n\n")
    
    st.subheader("Tags List")
    tags, icons = st.columns(2)
    
    with col1:
        st.subheader("Tag")
        st.write('Vegan / v')
        st.write('Vegetarian / veg')
        st.write('Eat well / EW / MB')
        st.write('Halal / H')
        st.write('Low carbon / LC / FE')
        
    with col2:
        st.subheader("Icon")
        st.image(image_path + TAGS['v'])
        st.image(image_path + TAGS['veg'])
        st.image(image_path + TAGS['ew'])
        st.image(image_path + TAGS['h'])
        st.image(image_path + TAGS['lc'])
