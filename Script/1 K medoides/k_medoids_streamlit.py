# streamlit_app.py
# App Streamlit: input = texto, output = archivo madre Excel etiquetado.
from __future__ import annotations

import io
import re
from typing import Dict, List

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from spanish_kmedoids_10 import ICOM_KEYWORDS, LEY_19_2013_KEYWORDS, keyword_labels

# --------------------------
# Utilidades
# --------------------------
def split_sentences_or_paragraphs(text: str) -> List[str]:
    """
    Divide el texto en oraciones o párrafos.
    Usa saltos de línea si existen; si no, divide por signos de puntuación.
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if len(lines) > 1:
        return lines
    sentences = [s.strip() for s in re.split(r"[.!?]\s+", text) if len(s.strip()) > 0]
    return sentences

def make_excel_bytes(rows: List[List], header: List[str]) -> bytes:
    buf = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "archivo_madre"
    ws.append(header)
    for row in rows:
        ws.append(row)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 90
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["D"].width = 25
    wb.save(buf)
    return buf.getvalue()

def taxonomy_to_rows(taxonomy: Dict[str, List[str]]) -> List[Dict[str, str]]:
    return [
        {"etiqueta": label, "palabras_clave": ", ".join(keywords)}
        for label, keywords in taxonomy.items()
    ]

def rows_to_taxonomy(rows) -> Dict[str, List[str]]:
    taxonomy = {}
    for row in rows:
        label = str(row.get("etiqueta", "")).strip()
        raw_keywords = str(row.get("palabras_clave", "")).strip()
        if not label or not raw_keywords:
            continue
        keywords = [kw.strip() for kw in raw_keywords.split(",") if kw.strip()]
        if keywords:
            taxonomy[label] = keywords
    return taxonomy

def normative_columns_custom(text: str, ley_taxonomy: Dict[str, List[str]], icom_taxonomy: Dict[str, List[str]]) -> Dict[str, str]:
    ley_labels = keyword_labels(text, ley_taxonomy)
    icom_labels = keyword_labels(text, icom_taxonomy)
    return {
        "ley 19/2013": "; ".join(ley_labels),
        "codigo deontologico": "; ".join(icom_labels),
        "otros": "" if ley_labels or icom_labels else "otros",
    }

# --------------------------
# UI
# --------------------------
st.set_page_config(page_title="Archivo madre etiquetado", page_icon="🧩", layout="centered")
st.title("Archivo madre etiquetado")
st.write("Esta herramienta lee una memoria, separa sus parrafos y genera un solo Excel con las etiquetas normativas.")

with st.expander("Que significa cada parte", expanded=True):
    st.markdown(
        """
        - **Pegar texto**: pega aqui el contenido de una memoria si quieres probar sin subir archivo.
        - **Subir archivo .txt/.md**: carga una memoria completa desde tu ordenador.
        - **Ley 19/2013**: etiquetas de transparencia. La herramienta mira si cada parrafo contiene palabras clave de esas etiquetas.
        - **Codigo deontologico ICOM**: etiquetas de buenas practicas museologicas. Funciona igual: busca palabras clave en cada parrafo.
        - **Palabras clave**: terminos que activan una etiqueta. Puedes editarlas separandolas con comas.
        - **Otros temas**: se rellena cuando el parrafo no encaja ni en Ley 19/2013 ni en Codigo deontologico ICOM.
        - **Restaurar Ley / Restaurar ICOM**: vuelve a poner las etiquetas originales si cambiaste algo y quieres empezar de nuevo.
        - **Procesar**: crea el Excel final.
        - **Archivo madre etiquetado**: el unico archivo de salida. Contiene una fila por parrafo y cuatro columnas: parrafo, ley 19/2013, codigo deontologico icom y otros temas.
        """
    )

with st.sidebar:
    st.header("Etiquetas")
    st.caption("Edita las palabras clave separadas por comas. Puedes agregar o borrar filas.")

    if "ley_taxonomy_rows" not in st.session_state:
        st.session_state.ley_taxonomy_rows = taxonomy_to_rows(LEY_19_2013_KEYWORDS)
    if "icom_taxonomy_rows" not in st.session_state:
        st.session_state.icom_taxonomy_rows = taxonomy_to_rows(ICOM_KEYWORDS)

    with st.expander("Ley 19/2013", expanded=False):
        ley_taxonomy_rows = st.data_editor(
            st.session_state.ley_taxonomy_rows,
            num_rows="dynamic",
            use_container_width=True,
            key="ley_taxonomy_editor",
            column_config={
                "etiqueta": st.column_config.TextColumn("Etiqueta"),
                "palabras_clave": st.column_config.TextColumn("Palabras clave"),
            },
        )

    with st.expander("Codigo deontologico ICOM", expanded=False):
        icom_taxonomy_rows = st.data_editor(
            st.session_state.icom_taxonomy_rows,
            num_rows="dynamic",
            use_container_width=True,
            key="icom_taxonomy_editor",
            column_config={
                "etiqueta": st.column_config.TextColumn("Etiqueta"),
                "palabras_clave": st.column_config.TextColumn("Palabras clave"),
            },
        )

    col_reset_ley, col_reset_icom = st.columns(2)
    with col_reset_ley:
        if st.button("Restaurar Ley"):
            st.session_state.ley_taxonomy_rows = taxonomy_to_rows(LEY_19_2013_KEYWORDS)
            st.rerun()
    with col_reset_icom:
        if st.button("Restaurar ICOM"):
            st.session_state.icom_taxonomy_rows = taxonomy_to_rows(ICOM_KEYWORDS)
            st.rerun()

ley_taxonomy = rows_to_taxonomy(ley_taxonomy_rows)
icom_taxonomy = rows_to_taxonomy(icom_taxonomy_rows)

st.write("**Entrada de texto** (pega tu texto o sube un .txt/.md). Si hay saltos de línea, se usan como párrafos; si no, se segmenta por oraciones.")
col1, col2 = st.columns(2)
with col1:
    texto = st.text_area("Pegar texto", height=220, placeholder="Pega aquí tu texto en español…")
with col2:
    up = st.file_uploader("…o subir archivo .txt/.md", type=["txt", "md"])

if up and not texto.strip():
    texto = up.read().decode("utf-8", errors="replace")

if st.button("Procesar"):
    if not texto.strip():
        st.error("Por favor, introduce texto o sube un .txt.")
        st.stop()

    items = split_sentences_or_paragraphs(texto)
    if len(items) == 0:
        st.error("No se detectaron oraciones/párrafos en el texto.")
        st.stop()

    master_rows = []
    preview_rows = []
    for paragraph in items:
        norm_labels = normative_columns_custom(paragraph, ley_taxonomy, icom_taxonomy)
        master_rows.append([
            paragraph,
            norm_labels["ley 19/2013"],
            norm_labels["codigo deontologico"],
            norm_labels["otros"],
        ])
        preview_rows.append({
            "parrafo": paragraph,
            "ley 19/2013": norm_labels["ley 19/2013"],
            "codigo deontologico icom": norm_labels["codigo deontologico"],
            "otros temas": norm_labels["otros"],
        })
    master_excel_bytes = make_excel_bytes(
        master_rows,
        header=["parrafo", "ley 19/2013", "codigo deontologico icom", "otros temas"],
    )

    excel_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    st.success("Archivo madre listo.")
    st.download_button("Descargar archivo_madre_etiquetado.xlsx", data=master_excel_bytes, file_name="archivo_madre_etiquetado.xlsx", mime=excel_mime)

    # Opcional: previews
    st.subheader("Vista previa (primeras filas)")
    st.write("**Archivo madre**")
    st.dataframe(preview_rows[:min(10, len(preview_rows))])
