import os

import streamlit as st

from converter import pdf_to_docx

st.set_page_config(page_title="PDF Slide -> Word", layout="centered")

st.title("Convertitore PDF Slide -> Word")
st.write(
    "Carica un file PDF con slide orizzontali (testo e immagini). "
    "Il testo viene riorganizzato in titoli, paragrafi ed elenchi puntati "
    "e le immagini vengono inserite nel documento Word risultante."
)

uploaded_file = st.file_uploader("Carica il file PDF", type=["pdf"])

if uploaded_file is not None:
    if st.button("Avvia conversione", type="primary"):
        with st.spinner("Conversione in corso..."):
            try:
                pdf_bytes = uploaded_file.read()
                base_name = os.path.splitext(uploaded_file.name)[0]
                docx_bytes = pdf_to_docx(pdf_bytes, title=base_name)
            except Exception as exc:
                st.error(f"Errore durante la conversione: {exc}")
            else:
                st.session_state["docx_bytes"] = docx_bytes
                st.session_state["docx_name"] = f"{base_name}.docx"
                st.success("Conversione completata.")

if "docx_bytes" in st.session_state:
    st.download_button(
        label="Scarica il file Word",
        data=st.session_state["docx_bytes"],
        file_name=st.session_state["docx_name"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
