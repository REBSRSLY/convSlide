import os

import streamlit as st

from converter import build_document, extract_content, translate_items

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

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
                items = extract_content(pdf_bytes, title=base_name)
                docx_bytes = build_document(items)
            except Exception as exc:
                st.error(f"Errore durante la conversione: {exc}")
            else:
                st.session_state["content_items"] = items
                st.session_state["base_name"] = base_name
                st.session_state["docx_bytes"] = docx_bytes
                st.session_state.pop("docx_it_bytes", None)
                st.success("Conversione completata.")

if "docx_bytes" in st.session_state:
    st.download_button(
        label="Scarica il file Word",
        data=st.session_state["docx_bytes"],
        file_name=f"{st.session_state['base_name']}.docx",
        mime=DOCX_MIME,
    )

    if st.button("Traduci in italiano"):
        with st.spinner("Traduzione in corso..."):
            try:
                translated_items = translate_items(st.session_state["content_items"], target_lang="it")
                docx_it_bytes = build_document(translated_items)
            except Exception as exc:
                st.error(f"Errore durante la traduzione: {exc}")
            else:
                st.session_state["docx_it_bytes"] = docx_it_bytes
                st.success("Traduzione completata.")

if "docx_it_bytes" in st.session_state:
    st.download_button(
        label="Scarica la traduzione in italiano",
        data=st.session_state["docx_it_bytes"],
        file_name=f"{st.session_state['base_name']}_IT.docx",
        mime=DOCX_MIME,
    )
