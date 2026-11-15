import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from deep_translator import MyMemoryTranslator
import io
import time
import re # Ajouté pour vérifier la ponctuation

st.set_page_config(page_title="Traducteur Scolaire", page_icon="🏫")
st.title("🌐 Traducteur de Cours & Examens (Correction Arabe)")
st.warning("⚠️ Outil gratuit utilisant MyMemory. Amélioré pour l'arabe et les tableaux.")

langues_cibles = {
    "Allemand": "de-DE",
    "Anglais": "en-US",
    "Arabe": "ar-SA",
    "Espagnol": "es-ES"
}
langue_choisie = st.selectbox("Langue de destination :", list(langues_cibles.keys()))

fichier_upload = st.file_uploader("Uploadez votre fichier", type=["docx"])

# Fonction pour vérifier si un texte ne contient que des symboles/ponctuation
def contient_seulement_ponctuation(texte):
    # Enlève tous les symboles, espaces et points. Si le résultat est vide, c'est qu'il n'y avait pas de mots.
    texte_nettoye = re.sub(r'[^\w\s]', '', texte).strip()
    return len(texte_nettoye) == 0

if st.button("Traduire le document") and fichier_upload is not None:
    traducteur = MyMemoryTranslator(source='fr-FR', target=langues_cibles[langue_choisie])
    est_arabe = (langue_choisie == "Arabe")
    
    with st.spinner(f'Traduction en {langue_choisie} en cours...'):
        try:
            doc = docx.Document(fichier_upload)
            
            # --- 1. Traduction des paragraphes normaux ---
            for para in doc.paragraphs:
                texte_original = para.text.strip()
                if texte_original:
                    # On ignore les lignes pointillées (ex: "........")
                    if contient_seulement_ponctuation(texte_original):
                        continue 
                        
                    try:
                        traduction = traducteur.translate(texte_original, return_all=False)
                        if isinstance(traduction, list):
                            para.text = traduction[0]
                        else:
                            para.text = traduction
                        
                        # Si c'est de l'arabe, on force l'alignement à droite (RTL)
                        if est_arabe:
                            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            
                        time.sleep(0.3)
                    except:
                        pass
            
            # --- 2. Traduction des Tableaux (Très important pour les examens) ---
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            texte_original = para.text.strip()
                            if texte_original:
                                if contient_seulement_ponctuation(texte_original):
                                    continue
                                try:
                                    traduction = traducteur.translate(texte_original, return_all=False)
                                    if isinstance(traduction, list):
                                        para.text = traduction[0]
                                    else:
                                        para.text = traduction
                                        
                                    if est_arabe:
                                        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                                        
                                    time.sleep(0.3)
                                except:
                                    pass

            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            
            st.success("✅ Fichier Word traduit avec succès (Mise en page corrigée) !")
            st.download_button(
                label="⬇️ Télécharger le Word", 
                data=doc_io, 
                file_name=f"Traduit_{fichier_upload.name}", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"Erreur technique de lecture du fichier : {e}")
