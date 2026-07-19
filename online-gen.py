import streamlit as st
from generate import parse_text, create_doc, image_path, TAGS
from io import BytesIO, StringIO

st.title("Text to Docx Menu Generator")
filename = None

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
        filename = uploaded.name
            
        items = parse_text(StringIO(text))
        output_doc = create_doc(station_name, items, scale)
        
        output = BytesIO()
        output_doc.save(output)
        output.seek(0)
        
        st.download_button(
            "Download Menu",
            output,
            file_name = filename + " - " + station_name + ".docx",
            mime = "application/vnd.openxmlformats-officedocuments.wordprocessingxml.document"
        )
    else:
        st.error("Please upload a text file first")


with st.sidebar.expander("? - Help"):
    st.subheader("Text file format")
    st.markdown('''
    Each item consists of 4 lines-
    
    Line 1. **Item no. Item name**  
    Line 2. **Ingredients**  
    Line 3. **Tags (Comma-separated)**  
    Line 4. **Allergen notes**
    
    Use '-' in any column other than name to leave it empty  
    Can be written in any language  
    Review final menu because misspellings can result in poor translation  
    Note: 'Menu of the Day' does not support tags
    ''')

    with st.expander("Example item"):
        st.markdown('''
        1. Bean Stew  
        Beans and stuff  
        vegan, lc  
        Contains: Lactose  
        ''')
    
    with st.expander("Tags List"):
        tags, icons = st.columns(2)
        
        with tags:
            st.subheader("Tag")
            st.write('Vegan / v\n')
            st.write('Vegetarian / veg\n')
            st.write('Eat well / EW / MB\n')
            st.write('Halal / H\n')
            st.write('Low carbon / LC / FE\n')
            
        with icons:
            st.subheader("Icon")
            st.image(image_path + TAGS['v'], width=25)
            st.image(image_path + TAGS['veg'], width=25)
            st.image(image_path + TAGS['ew'], width=25)
            st.image(image_path + TAGS['h'], width=25)
            st.image(image_path + TAGS['lc'], width=25)
