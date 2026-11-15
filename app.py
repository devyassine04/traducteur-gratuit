import streamlit as st
import docx
from deep_translator import MyMemoryTranslator
import io
import time

st.set_page_config(page_title="Traducteur Scolaire", page_icon="🏫")
st.title("🌐 Traducteur de Cours & Examens (Version Gratuite)")
st.warning("⚠️ Outil gratuit utilisant MyMemory Translator.")
st.write("Format accepté : Uniquement Word (.docx) pour conserver la mise en page.")

# NOUVEAUX CODES DE LANGUE COMPATIBLES AVEC MYMEMORY
langues_cibles = {
    "Allemand": "de-DE",
    "Anglais": "en-US",
    "Arabe": "ar-SA",
    "Espagnol": "es-ES"
}
langue_choisie = st.selectbox("Langue de destination :", list(langues_cibles.keys()))

fichier_upload = st.file_uploader("Uploadez votre fichier", type=["docx"])

if st.button("Traduire le document") and fichier_upload is not None:
    # Utilisation des nouveaux codes de langue
    traducteur = MyMemoryTranslator(source='fr-FR', target=langues_cibles[langue_choisie])
    erreur_affichee = False
    
    with st.spinner(f'Traduction en {langue_choisie} en cours...'):
        try:
            doc = docx.Document(fichier_upload)
            
            # Traduction des paragraphes
            for para in doc.paragraphs:
                if para.text.strip():
                    try:
                        # MyMemory renvoie parfois une liste si return_all n'est pas spécifié, on s'assure d'avoir du texte
                        traduction = traducteur.translate(para.text, return_all=False)
                        if isinstance(traduction, list):
                            para.text = traduction[0]
                        else:
                            para.text = traduction
                        time.sleep(0.3) # Pause légère
                    except Exception as e:
                        if not erreur_affichee:
                            st.error(f"Avertissement de traduction sur un bloc : {e}")
                            erreur_affichee = True
            
            # Traduction des tableaux
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if para.text.strip():
                                try:
                                    traduction = traducteur.translate(para.text, return_all=False)
                                    if isinstance(traduction, list):
                                        para.text = traduction[0]
                                    else:
                                        para.text = traduction
                                    time.sleep(0.3)
                                except:
                                    pass

            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            
            st.success("✅ Fichier Word traduit avec succès !")
            st.download_button(
                label="⬇️ Télécharger le Word", 
                data=doc_io, 
                file_name=f"Traduit_{fichier_upload.name}", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"Erreur technique de lecture du fichier : {e}")
