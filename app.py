import streamlit as st
import docx
from deep_translator import GoogleTranslator
import io
import time # Nécessaire pour la temporisation anti-blocage

st.set_page_config(page_title="Traducteur Scolaire", page_icon="🏫")
st.title("🌐 Traducteur de Cours & Examens (Version Gratuite)")
st.warning("⚠️ Pour éviter les blocages de traduction, ce système traite le document lentement (environ 2 secondes par paragraphe). Merci de patienter pendant le chargement.")
st.write("Format accepté : Uniquement Word (.docx) pour conserver la mise en page.")

langues_cibles = {
    "Allemand": "de",
    "Anglais": "en",
    "Arabe": "ar",
    "Espagnol": "es"
}
langue_choisie = st.selectbox("Langue de destination :", list(langues_cibles.keys()))

fichier_upload = st.file_uploader("Uploadez votre fichier", type=["docx"])

if st.button("Traduire le document") and fichier_upload is not None:
    traducteur = GoogleTranslator(source='auto', target=langues_cibles[langue_choisie])
    erreur_affichee = False
    
    with st.spinner(f'Traduction en {langue_choisie} en cours... (Processus ralenti volontairement)'):
        try:
            doc = docx.Document(fichier_upload)
            
            # Traduction des paragraphes avec pause de 2 secondes
            for para in doc.paragraphs:
                if para.text.strip():
                    try:
                        para.text = traducteur.translate(para.text)
                        time.sleep(2) 
                    except Exception as e:
                        if not erreur_affichee:
                            st.error(f"Erreur de connexion (Traduction ignorée sur un bloc) : {e}")
                            erreur_affichee = True
            
            # Traduction des tableaux avec pause de 2 secondes
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if para.text.strip():
                                try:
                                    para.text = traducteur.translate(para.text)
                                    time.sleep(2)
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
