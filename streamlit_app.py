import streamlit as st
import pandas as pd
import math
import zipfile
import io
import re
import unicodedata

st.set_page_config(page_title="La Boîte à Outils", page_icon="🛠️", layout="wide")

# Dictionnaires globaux
REGION_DICTIONARY = {
    "Montreal + Environ": ['H1', 'H2', 'H3', 'H4', 'H5', 'H8', 'H9'], 
    "Rive Sud": ['J2G','J2H','J2S','J2T','J2W','J2X','J2Y','J3','J4','J5A','J5B','J5C','J5R','J6J','J6K','J6N','J6R','J6S','J6T','J0L','J0J','J7V','J7T'],
    "Rive Nord": ['H7','J5J','J5K','J5L','J5M','J5N','J5T','J5W','J5X','J5Y','J5Z','J6A','J6E','J6V','J6X','J6Y','J6Z','J7A','J7B','J7C','J7E','J7G','J7H','J7J','J7K','J7L','J7M','J7N','J7P','J7R','J7Z']
}

CITY_KEYWORDS = {
    "Montreal + Environ": ['montreal', 'mtl', 'outremont', 'westmount', 'anjou', 'lachine', 'lasalle', 'pierrefonds', 'roxboro', 'dorval', 'pointe claire', 'kirkland', 'beaconsfield', 'baie d urfe', 'ste anne de bellevue', 'senneville', 'dollard', 'ddo', 'hampstead', 'cote st luc', 'csl', 'mont royal'],
    "Rive Sud": ['longueuil', 'brossard', 'boucherville', 'st hubert', 'st lambert', 'greenfield', 'st bruno', 'ste julie', 'beloeil', 'chambly', 'la prairie', 'laprairie', 'candiac', 'delson', 'st constant', 'ste catherine', 'chateauguay', 'mercier', 'varennes', 'vercheres', 'st jean', 'st luc', 'st mathieu', 'st philippe', 'carignan', 'st basile', 'st mathias', 'richelieu', 'sorel', 'valleyfield'],
    "Rive Nord": ['laval', 'vimont', 'chomedey', 'terrebonne', 'repentigny', 'blainville', 'boisbriand', 'ste therese', 'st eustache', 'rosemere', 'mascouche', 'mirabel', 'deux montagnes', 'st jerome', 'bois des filion', 'lorraine', 'ste marthe', 'oka', 'pointe calumet', 'st lin', 'lassomption']
}

column_synonyms = {
    'Phone Number': ['phone number', 'phone', 'telephone', 'tel', 'mobile', 'cell', 'numéro de téléphone', 'numero', 'téléphone', 'phone_home','Num Tel','phoneNumber','NUM TEL ','NUM TEL','phonenumber'],
    'First Name': ['first name', 'first_name', 'prenom', 'prénom', 'fname', 'given name','firstName'],
    'Last Name': ['last name', 'last_name', 'name', 'nom', 'surname', 'lname', 'nom de famille', 'family name','last_name','lastName'],
    'Address': ['adress', 'address_street', 'address_full', 'address', 'adresse', 'street', 'rue', 'adresse postale','Street Address','ADRESSE'],
    'Postal Code': ['code postale', 'address_postal_code', 'postalcode', 'code postal', 'postal code', 'zip code', 'zip', 'cp'],
    'Mail': ['courriel', 'mail', 'email', 'adresse mail'],
    'City': ['city', '_requestedcity', 'address_city', 'ville','Ville originale'],
    'Date Naissance': ['birth_date', 'date', 'date naissance', 'age','birth date','COMMENT'],
    'IBAN': ['iban', 'iban number', 'bank account', 'bank account number','IBAN', 'iban rib','nir_pretty','IBAN/RIB','IBANRIB'],
}

# Fonctions utilitaires
def normalize_text(text):
    if pd.isna(text): return ""
    text = str(text).lower()
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('utf-8')
    text = re.sub(r'[^a-z0-9]', ' ', text)
    return re.sub(r'\s+', ' ', text).strip()

def lire_fichier(file_obj):
    file_obj.seek(0)
    ext = file_obj.name.rsplit('.', 1)[1].lower()
    
    if ext == 'csv': 
        try:
            return pd.read_csv(file_obj, sep=';', dtype=str, encoding='utf-8')
        except UnicodeDecodeError:
            file_obj.seek(0)
            return pd.read_csv(file_obj, sep=';', dtype=str, encoding='latin-1')
    else: 
        return pd.read_excel(file_obj, dtype=str)

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ==========================================
# 1. MENU LATÉRAL (SIDEBAR)
# ==========================================
st.sidebar.title("🛠️ La Boîte à Outils")
choix = st.sidebar.radio("Choisissez votre application :", [
    "1. ✂️ Découpeur",
    "2. 🔗 Mélangeur de Fichiers",
    "3. 🌍 Filtre Régional",
    "4. 🎂 Filtre par Âge",
    "5. 📱 Formateur Téléphone(+1 ou +33)",
    "6. 🇫🇷 Séparateur Départements (FR)"
])

st.sidebar.markdown("---")
st.sidebar.info("Développé pour un traitement de données rapide et indestructible.")

# ==========================================
# 2. APPLICATION : DÉCOUPEUR (9asem)
# ==========================================
if choix.startswith("1"):
    st.title("✂️ Découpeur")
    lignes_par_fichier = st.number_input("Nombre de contacts par fichier :", min_value=1, value=19999)
    uploaded_file = st.file_uploader("Chargez votre fichier (CSV/Excel)", type=["csv", "xlsx", "xls"], key="split")

    if uploaded_file and st.button("Lancer le découpage 🚀"):
        with st.spinner("Découpage en cours..."):
            df = lire_fichier(uploaded_file)
            ext = uploaded_file.name.rsplit('.', 1)[1].lower()
            nom_base = uploaded_file.name.rsplit('.', 1)[0]
            nb_fichiers = math.ceil(len(df) / lignes_par_fichier)
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for i in range(nb_fichiers):
                    df_morceau = df.iloc[i*lignes_par_fichier : (i+1)*lignes_par_fichier]
                    file_buffer = io.BytesIO()
                    if ext == 'csv':
                        df_morceau.to_csv(file_buffer, index=False, sep=';')
                    else:
                        df_morceau.to_excel(file_buffer, index=False, engine='xlsxwriter')
                    zip_file.writestr(f"{nom_base}_partie_{i+1}.{ext}", file_buffer.getvalue())
            
            st.success(f"✅ {len(df)} contacts découpés en {nb_fichiers} fichiers !")
            st.download_button("📥 Télécharger le ZIP", data=zip_buffer.getvalue(), file_name=f"{nom_base}_decoupe.zip", mime="application/zip")


# ==========================================
# 3. APPLICATION : MÉLANGEUR
# ==========================================
elif choix.startswith("2"):
    st.title("🔗 Mélangeur de Fichiers Excel")
    st.write("Glissez tous les fichiers que vous voulez fusionner. Le programme standardisera les colonnes.")
    
    # ---> Le sélecteur de pays
    pays_fusion = st.selectbox("Sélectionnez le pays cible :", ["Canada", "France"], key="merge_country")
    
    uploaded_files = st.file_uploader("Chargez plusieurs fichiers Excel", type=["xlsx", "xls"], accept_multiple_files=True, key="merge")
    
    if uploaded_files and st.button("Fusionner les fichiers 🚀"):
        with st.spinner("Analyse et fusion..."):
            all_dataframes = []
            standard_columns = list(column_synonyms.keys())
            
            # 1. On met uniquement les colonnes strictement obligatoires ici
            mandatory_columns = ['Phone Number', 'Address']
            if pays_fusion == "France":
                mandatory_columns.append("IBAN")
                mandatory_columns.append("Date Naissance")
            
            erreurs = []

           # UNE SEULE BOUCLE POUR TOUT FAIRE
            for file in uploaded_files:
                try:
                    # LECTURE ET NETTOYAGE SUPER PUISSANT
                    df = pd.read_excel(file)
                    rename_mapping = {}
                    
                    for original_col in df.columns:
                        # 1. On "écrase" la colonne du fichier (ex: " Last_Name " devient "lastname")
                        clean_col = re.sub(r'[^a-z0-9]', '', str(original_col).lower())
                        
                        for std_name, synonyms in column_synonyms.items():
                            # 2. On "écrase" aussi les mots de notre dictionnaire de la même façon
                            clean_synonyms = [re.sub(r'[^a-z0-9]', '', syn.lower()) for syn in synonyms]
                            
                            if clean_col in clean_synonyms:
                                rename_mapping[original_col] = std_name
                                break
                                
                    df = df.rename(columns=rename_mapping)
                    
                    
                    # LOGIQUE DE TRI INTELLIGENT (Garde la colonne avec le plus de données)
                    meilleures_colonnes = {}
                    for col in df.columns.unique():
                        homonymes = df.loc[:, df.columns == col]
                        if isinstance(homonymes, pd.DataFrame) and homonymes.shape[1] > 1:
                            meilleur_index = homonymes.notna().sum().argmax()
                            meilleures_colonnes[col] = homonymes.iloc[:, meilleur_index]
                        else:
                            meilleures_colonnes[col] = homonymes.iloc[:, 0] if isinstance(homonymes, pd.DataFrame) else homonymes
                            
                    df = pd.DataFrame(meilleures_colonnes)
                    
                    # 2. Vérification des colonnes de base (Téléphone, Adresse, IBAN)
                    missing = [col for col in mandatory_columns if col not in df.columns]
                    
                    # 3. LA CONDITION INTELLIGENTE : Prénom OU Nom
                    if 'First Name' not in df.columns and 'Last Name' not in df.columns:
                        missing.append("First Name OU Last Name")
                    
                    # 4. S'il manque quelque chose, on rejette le fichier
                    if missing:
                        erreurs.append(f"❌ {file.name} : Manque {', '.join(missing)}")
                        continue 
                        
                    # 5. Si tout est bon, on l'ajoute à la liste
                    filtered_df = df[[c for c in standard_columns if c in df.columns]].copy()
                    for col in standard_columns:
                        if col not in filtered_df.columns: 
                            filtered_df[col] = pd.NA
                    all_dataframes.append(filtered_df[standard_columns])
                
                except Exception as e:
                    erreurs.append(f"⚠️ {file.name} : Erreur -> {e}")

            # FIN DE LA BOUCLE - FUSION FINALE
            if all_dataframes:
                merged_df = pd.concat(all_dataframes, ignore_index=True)
                st.success(f"✅ {len(all_dataframes)} fichiers fusionnés avec succès ! ({len(merged_df)} lignes au total)")
                st.download_button("📥 Télécharger le fichier fusionné", data=to_excel_bytes(merged_df), file_name="Merged_Final_Output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.error("Aucun fichier valide n'a pu être fusionné.")
            
            if erreurs:
                st.warning("Certains fichiers ont été ignorés car il manquait des colonnes obligatoires :")
                for err in erreurs: 
                    st.write(err)
# ==========================================
# 4. APPLICATION : FILTRE RÉGIONAL
# ==========================================
elif choix.startswith("3"):
    st.title("🌍 Filtre Régional (❌PROTOTYPE❌)")
    
    col1, col2 = st.columns(2)
    with col1: city_col = st.text_input("Nom de la colonne Ville:", value="Ville")
    with col2: zip_col = st.text_input("Nom de la colonne Code Postal:", value="Code Postal")
    
    mode = st.radio("Choisissez la région à garder :", ["Villes Spécifiques (Sélection manuelle)", "Montreal + Environ", "Rive Sud", "Rive Nord"])
    
    uploaded_file = st.file_uploader("Chargez votre fichier de données", type=["csv", "xlsx", "xls"], key="region")
    
    if uploaded_file is not None:
        uploaded_file.seek(0) 
        
        df = lire_fichier(uploaded_file)
        cols_lower = {str(c).strip().lower(): c for c in df.columns}
        
        if city_col.lower() not in cols_lower or zip_col.lower() not in cols_lower:
            st.error("❌ Les colonnes Ville ou Code Postal sont introuvables. Vérifiez l'orthographe.")
        else:
            actual_city = cols_lower[city_col.lower()]
            actual_zip = cols_lower[zip_col.lower()]
            
            clean_zips = df[actual_zip].astype(str).str.upper().replace(r'\s+|-', '', regex=True)
            clean_cities = df[actual_city].apply(normalize_text)

            filtered_df = pd.DataFrame()
            villes_choisies = []

            # --- SI MODE MANUEL ---
            if mode == "Villes Spécifiques (Sélection manuelle)":
                
                # LA CORRECTION EST ICI : On extrait UNIQUEMENT les villes uniques pour ne pas faire crasher le web
                villes_uniques = df[actual_city].dropna().astype(str).str.strip().unique()
                liste_villes = sorted(list(villes_uniques))
                
                villes_choisies = st.multiselect("🔎 Recherchez et sélectionnez les VILLES à garder :", options=liste_villes)
                
                if villes_choisies:
                    # On garde uniquement les lignes où la ville correspond à votre sélection
                    filtered_df = df[df[actual_city].astype(str).str.strip().isin(villes_choisies)]
                else:
                    st.info("👆 La liste des villes uniques est chargée ! Cliquez ci-dessus pour faire votre choix.")
                    
            # --- SI MODE REGION (Nord, Sud, Mtl) ---
            else:
                valid_prefixes = REGION_DICTIONARY[mode]
                mask_zip = clean_zips.apply(lambda z: any(str(z).startswith(prefix) for prefix in valid_prefixes))
                mask_city = clean_cities.apply(lambda c: any(kw in c for kw in CITY_KEYWORDS[mode]))
                filtered_df = df[mask_zip | mask_city]
            
            # --- AFFICHAGE ET TÉLÉCHARGEMENT ---
            if not filtered_df.empty:
                st.success(f"✅ {len(filtered_df)} contacts trouvés pour cette sélection !")
                ext = uploaded_file.name.rsplit('.', 1)[1].lower()
                safe_name = mode.replace(' + ', '_').replace(' ', '_').replace('(', '').replace(')', '')
                
                if ext == 'csv':
                    csv = filtered_df.to_csv(index=False, sep=';').encode('utf-8')
                    st.download_button("📥 Télécharger le résultat (CSV)", data=csv, file_name=f"Filtre_{safe_name}.csv", mime="text/csv")
                else:
                    st.download_button("📥 Télécharger le résultat (Excel)", data=to_excel_bytes(filtered_df), file_name=f"Filtre_{safe_name}.xlsx")
            elif mode != "Villes Spécifiques (Sélection manuelle)" or villes_choisies:
                st.warning("Aucun contact trouvé avec ce filtre.")
# ==========================================
# 5. APPLICATION : FILTRE ÂGE
# ==========================================
elif choix.startswith("4"):
    st.title("🎂 Filtre par Âge")
    
    col_name = st.text_input("Nom de la colonne Date de Naissance :", value="Date de naissance")
    target_year = st.number_input("Garder les contacts nés AVANT l'année :", value=1980, step=1)
    uploaded_file = st.file_uploader("Chargez votre fichier", type=["csv", "xlsx", "xls"], key="age")
    
    if uploaded_file and st.button("Filtrer l'Âge 🚀"):
        df = lire_fichier(uploaded_file)
        cols_lower = {str(c).strip().lower(): c for c in df.columns}
        
        if col_name.lower() not in cols_lower:
            st.error("❌ Colonne introuvable.")
        else:
            actual_col = cols_lower[col_name.lower()]
            years = df[actual_col].astype(str).str.extract(r'((?:19|20)\d{2})')[0].astype(float)
            filtered_df = df[years < target_year]
            
            st.success(f"✅ {len(filtered_df)} contacts nés avant {target_year} conservés.")
            st.download_button("📥 Télécharger le résultat", data=to_excel_bytes(filtered_df), file_name=f"Filtre_Avant_{target_year}.xlsx")

# ==========================================
# 6. APPLICATION : FILTRE DÉPARTEMENTS FRANCE
# ==========================================
elif choix.startswith("6"):
    st.title("🇫🇷 Séparateur de Départements")
    st.write("Ce programme ajoute un 0 devant les codes postaux à 4 chiffres, puis sépare les contacts en deux fichiers cibles. (ABC/IDF)")
    
    col_cp = st.text_input("Nom de la colonne des codes postaux :", value="Code Postal")
    uploaded_file = st.file_uploader("Chargez votre fichier (Excel ou CSV)", type=["csv", "xlsx", "xls"], key="dept")
    
    if uploaded_file and st.button("Filtrer et Séparer "):
        with st.spinner("Traitement en cours..."):
            uploaded_file.seek(0)
            df = lire_fichier(uploaded_file)
            
            cols_lower = {str(c).strip().lower(): c for c in df.columns}
            
            if col_cp.lower() not in cols_lower:
                st.error(f"❌ La colonne '{col_cp}' est introuvable. Vérifiez l'orthographe.")
            else:
                actual_cp_col = cols_lower[col_cp.lower()]
                
                # 1. LA CORRECTION MAGIQUE (Ajout du 0)
                def corriger_cp(cp):
                    cp_str = str(cp).strip()
                    # Si Excel a transformé 7000 en 7000.0, on enlève le .0
                    if cp_str.endswith('.0'): 
                        cp_str = cp_str[:-2]
                    # Si la case ne contient que des chiffres et fait exactement 4 de long
                    if len(cp_str) == 4 and cp_str.isdigit():
                        return '0' + cp_str
                    return cp_str
                    
                df[actual_cp_col] = df[actual_cp_col].apply(corriger_cp)
                
                # 2. Définition des cibles
                liste_1 = ('07', '26', '42', '38', '69', '73', '74')
                liste_2 = ('91', '92', '93', '94', '95', '77', '78')
                
                # 3. Le Filtrage (startswith permet de vérifier les 2 premiers caractères)
                df_liste1 = df[df[actual_cp_col].astype(str).str.startswith(liste_1, na=False)]
                df_liste2 = df[df[actual_cp_col].astype(str).str.startswith(liste_2, na=False)]
                
                # 4. Création du fichier ZIP pour tout télécharger d'un coup
                nom_base = uploaded_file.name.rsplit('.', 1)[0]
                ext = uploaded_file.name.rsplit('.', 1)[1].lower()
                
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    
                    # On crée le Fichier 1 s'il n'est pas vide
                    if not df_liste1.empty:
                        file_buffer1 = io.BytesIO()
                        if ext == 'csv':
                            df_liste1.to_csv(file_buffer1, index=False, sep=';')
                        else:
                            df_liste1.to_excel(file_buffer1, index=False, engine='xlsxwriter')
                        zip_file.writestr(f"{nom_base}_07_26_42_38_69_73_74.{ext}", file_buffer1.getvalue())
                        
                    # On crée le Fichier 2 s'il n'est pas vide
                    if not df_liste2.empty:
                        file_buffer2 = io.BytesIO()
                        if ext == 'csv':
                            df_liste2.to_csv(file_buffer2, index=False, sep=';')
                        else:
                            df_liste2.to_excel(file_buffer2, index=False, engine='xlsxwriter')
                        zip_file.writestr(f"{nom_base}_91_a_95_77_78.{ext}", file_buffer2.getvalue())
                
                # 5. Affichage des résultats
                st.success("✅ Traitement terminé !")
                st.info(f"📁 Fichier 1 (07, 26, 42, 38, 69, 73, 74) : {len(df_liste1)} contacts.")
                st.info(f"📁 Fichier 2 (91, 92, 93, 94, 95, 77, 78) : {len(df_liste2)} contacts.")
                
                if not df_liste1.empty or not df_liste2.empty:
                    st.download_button(
                        label="📥 Télécharger les 2 fichiers (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"{nom_base}_Filtres_Departements.zip",
                        mime="application/zip"
                    )
                else:
                    st.warning("Aucun contact n'a été trouvé pour ces départements.")


# ==========================================
# 6. APPLICATION : FORMATEUR TÉLÉPHONE
# ==========================================
elif choix.startswith("5"):
    st.title("📱 Formateur de Téléphone")
    
    country = st.selectbox("Pays :", ["Canada", "France"])
    phone_col = st.text_input("Nom de la colonne Téléphone :", value="Téléphone")
    uploaded_files = st.file_uploader("Chargez les fichiers à corriger", type=["xlsx", "xls"], accept_multiple_files=True, key="tel")
    
    if uploaded_files and st.button("Corriger les numéros 🚀"):
        with st.spinner("Correction en cours..."):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for file in uploaded_files:
                    df = pd.read_excel(file, dtype=str)
                    if phone_col in df.columns:
                        def clean_phone(phone):
                            digits = re.sub(r'\D', '', str(phone))
                            if not digits: return phone
                            if country == 'Canada' and len(digits) == 10: return '1' + digits
                            if country == 'France':
                                if digits.startswith('0'): digits = digits[1:]
                                if len(digits) == 9: return '33' + digits
                            return digits
                        df[phone_col] = df[phone_col].apply(clean_phone)
                        
                        file_buffer = io.BytesIO()
                        df.to_csv(file_buffer, index=False, sep=';')
                        zip_file.writestr(f"{file.name.rsplit('.', 1)[0]}_clean.csv", file_buffer.getvalue())
                        
            st.success(f"✅ {len(uploaded_files)} fichiers corrigés et convertis en CSV !")
            st.download_button("📥 Télécharger les fichiers (ZIP)", data=zip_buffer.getvalue(), file_name="Telephones_Corriges.zip", mime="application/zip")