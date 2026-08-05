import io
import os
import zipfile

import streamlit as st

from converter import build_document, extract_content, translate_items

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
ZIP_MIME = "application/zip"

st.set_page_config(page_title="PDF Slide -> Word", layout="centered")

st.title("Convertitore PDF Slide -> Word")
st.write(
    "Carica uno o piu file PDF con slide orizzontali (testo e immagini). "
    "Il testo viene riorganizzato in titoli, paragrafi ed elenchi puntati "
    "e le immagini vengono inserite nel documento Word risultante. "
    "Con piu file caricati insieme, il download sara uno zip con tutti i risultati."
)

uploaded_files = st.file_uploader("Carica i file PDF", type=["pdf"], accept_multiple_files=True)


def _zip_bytes(named_docs):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_name, data in named_docs:
            zf.writestr(file_name, data)
    return buffer.getvalue()


if uploaded_files:
    if st.button("Avvia conversione", type="primary"):
        results = []
        progress = st.progress(0.0)
        status = st.empty()
        for i, uploaded_file in enumerate(uploaded_files):
            base_name = os.path.splitext(uploaded_file.name)[0]
            status.write(f"Conversione di **{uploaded_file.name}**...")
            try:
                pdf_bytes = uploaded_file.read()
                items = extract_content(pdf_bytes, title=base_name)
                docx_bytes = build_document(items)
            except Exception as exc:
                results.append({"base_name": base_name, "error": str(exc)})
            else:
                results.append({"base_name": base_name, "items": items, "docx_bytes": docx_bytes})
            progress.progress((i + 1) / len(uploaded_files))
        status.empty()
        progress.empty()

        st.session_state["results"] = results
        st.session_state.pop("results_it", None)

        failed = [r for r in results if "error" in r]
        ok = [r for r in results if "error" not in r]
        if ok:
            st.success(f"Conversione completata: {len(ok)} su {len(results)} file.")
        for r in failed:
            st.error(f"Errore su {r['base_name']}: {r['error']}")

if st.session_state.get("results"):
    ok_results = [r for r in st.session_state["results"] if "error" not in r]
    if ok_results:
        if len(ok_results) == 1:
            r = ok_results[0]
            st.download_button(
                label="Scarica il file Word",
                data=r["docx_bytes"],
                file_name=f"{r['base_name']}.docx",
                mime=DOCX_MIME,
            )
        else:
            zip_data = _zip_bytes([(f"{r['base_name']}.docx", r["docx_bytes"]) for r in ok_results])
            st.download_button(
                label=f"Scarica tutti i {len(ok_results)} file Word (zip)",
                data=zip_data,
                file_name="convslide_output.zip",
                mime=ZIP_MIME,
            )

        if st.button("Traduci in italiano"):
            results_it = []
            progress = st.progress(0.0)
            status = st.empty()
            for i, r in enumerate(ok_results):
                status.write(f"Traduzione di **{r['base_name']}**...")
                try:
                    translated_items = translate_items(r["items"], target_lang="it")
                    docx_it_bytes = build_document(translated_items)
                except Exception as exc:
                    results_it.append({"base_name": r["base_name"], "error": str(exc)})
                else:
                    results_it.append({"base_name": r["base_name"], "docx_bytes": docx_it_bytes})
                progress.progress((i + 1) / len(ok_results))
            status.empty()
            progress.empty()

            st.session_state["results_it"] = results_it
            failed_it = [r for r in results_it if "error" in r]
            ok_it = [r for r in results_it if "error" not in r]
            if ok_it:
                st.success(f"Traduzione completata: {len(ok_it)} su {len(results_it)} file.")
            for r in failed_it:
                st.error(f"Errore su {r['base_name']}: {r['error']}")

if st.session_state.get("results_it"):
    ok_it = [r for r in st.session_state["results_it"] if "error" not in r]
    if ok_it:
        if len(ok_it) == 1:
            r = ok_it[0]
            st.download_button(
                label="Scarica la traduzione in italiano",
                data=r["docx_bytes"],
                file_name=f"{r['base_name']}_IT.docx",
                mime=DOCX_MIME,
            )
        else:
            zip_data = _zip_bytes([(f"{r['base_name']}_IT.docx", r["docx_bytes"]) for r in ok_it])
            st.download_button(
                label=f"Scarica tutte le {len(ok_it)} traduzioni (zip)",
                data=zip_data,
                file_name="convslide_output_IT.zip",
                mime=ZIP_MIME,
            )
