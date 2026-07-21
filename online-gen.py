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
        filename = uploaded.name.split('.')[0]
            
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


with st.sidebar.expander("? - Help / Aide"):
    language = st.segmented_control(
        "Language",
        ["English", "Français"],
        default = "English"
    )

    if language == "English":
        st.subheader("Text file format")
        st.markdown('''
        Each item consists of 4 lines-
        
        Line 1. **Item no. Item name**  
        Line 2. **Ingredients**  
        Line 3. **Tags (Comma-separated)**  
        Line 4. **Allergen notes**
        
        Use '-' in any column other than name to leave it empty.  
        Can be written in any language.  
        Review final menu because misspellings can result in poor translation.    
        View all tags by clicking on the "Tags List" section below.    
        Note: 'Menu of the Day' does not support tags.
        ''')
    
        with st.expander("Example item"):
            st.markdown('''
            1. Bean Stew  
            Beans and stuff  
            vegan, lc  
            Contains: Lactose  
            ''')
    elif language == "Français":
        st.subheader("Format de fichier texte")
        st.markdown('''
        Chaque élément comporte 4 lignes:
        
        Ligne 1. **Numéro de l’élément. Nom de l’élément**  
        Ligne 2. **Ingrédients**  
        Ligne 3. **Attributs (séparées par des virgules)**  
        Ligne 4. **Notes sur les allergènes**
        
        Utilisez «-» dans toute colonne autre que le nom pour la laisser vide.  
        Le texte peut être dans n'importe quelle langue.  
        Vérifiez le menu final, car les fautes d’orthographe peuvent entraîner une mauvaise traduction.    
        Consultez tous les attributs en cliquant sur la section «Tags List» ci-dessous.    
        Note: Le «Menu du jour» ne prend pas en charge les attributs.
        ''')
    
        with st.expander("Exemple d'élément"):
            st.markdown('''
            1. Ragoût de fèves  
            Des fèves et tout ça  
            vegan, lc  
            Contient: lactose  
            ''')
    
    with st.expander("Tags List"):
        tags, icons = st.columns(2)
        
        with tags:
            st.subheader("Tag")
            st.write('Vegan / v')
            st.empty()
            st.empty()
            st.empty()
            st.write('Vegetarian / veg')
            st.empty()
            st.empty()
            st.empty()
            st.write('Eat well / EW / MB')
            st.empty()
            st.empty()
            st.empty()
            st.write('Halal / H')
            st.empty()
            st.empty()
            st.empty()
            st.write('Low carbon / LC / FE')
            
        with icons:
            st.subheader("Icon")
            st.image(image_path + TAGS['v'], width=75)
            st.image(image_path + TAGS['veg'], width=75)
            st.image(image_path + TAGS['ew'], width=75)
            st.image(image_path + TAGS['h'], width=75)
            st.image(image_path + TAGS['lc'], width=75)
