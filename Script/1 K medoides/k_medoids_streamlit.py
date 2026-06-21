# streamlit_app.py
# App Streamlit: input = texto, output = archivo madre Excel etiquetado.
from __future__ import annotations

import io
import re
from typing import Dict, List

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from spanish_kmedoids_10 import ICOM_KEYWORDS, LEY_19_2013_KEYWORDS, keyword_labels, select_k_medoids_paragraphs

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

def taxonomy_to_text(taxonomy: Dict[str, List[str]]) -> str:
    return "\n".join(
        f"{label} = {', '.join(keywords)}"
        for label, keywords in taxonomy.items()
    )

def text_to_taxonomy(text: str) -> Dict[str, List[str]]:
    taxonomy = {}
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" in line:
            label, raw_keywords = line.split("=", 1)
        elif ":" in line:
            label, raw_keywords = line.split(":", 1)
        else:
            continue
        label = label.strip()
        raw_keywords = raw_keywords.strip()
        if not label or not raw_keywords:
            continue
        keywords = [kw.strip() for kw in raw_keywords.split(",") if kw.strip()]
        if keywords:
            taxonomy[label] = keywords
    return taxonomy

def normative_columns_custom(text: str, law_taxonomy: Dict[str, List[str]], icom_taxonomy: Dict[str, List[str]]) -> Dict[str, str]:
    law_labels = keyword_labels(text, law_taxonomy)
    icom_labels = keyword_labels(text, icom_taxonomy)
    return {
        "ley": "; ".join(law_labels),
        "codigo deontologico": "; ".join(icom_labels),
        "otros": "" if law_labels or icom_labels else "otros",
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
        - **Numero de K medoides**: cantidad de parrafos representativos que quieres obtener. Si pones 24, el Excel tendra 24 parrafos.
        - **Nombre de la ley o marco 1**: cambia aqui Ley 19/2013 por cualquier otra ley o marco que quieras analizar.
        - **Nombre del marco 2**: puedes dejar Codigo deontologico ICOM o cambiarlo por otro marco.
        - **Etiquetas y palabras clave**: escribe una etiqueta por linea con este formato: etiqueta = palabra1, palabra2, palabra3.
        - **Otros temas**: se rellena cuando el parrafo no encaja en ninguno de los dos marcos.
        - **Restaurar marco 1 / Restaurar marco 2**: vuelve a poner las etiquetas originales si cambiaste algo y quieres empezar de nuevo.
        - **Procesar**: crea el Excel final.
        - **Archivo madre etiquetado**: el unico archivo de salida. Contiene una fila por parrafo y cuatro columnas: parrafo, marco 1, marco 2 y otros temas.
        """
    )

with st.container():
    st.subheader("Editar leyes y etiquetas")
    st.caption("Formato: una etiqueta por linea. Ejemplo: presupuesto_aprobado = presupuesto, cuentas, gasto")

    if "law_name" not in st.session_state:
        st.session_state.law_name = "ley 19/2013"
    if "icom_name" not in st.session_state:
        st.session_state.icom_name = "codigo deontologico icom"
    if "law_taxonomy_text" not in st.session_state:
        st.session_state.law_taxonomy_text = taxonomy_to_text(LEY_19_2013_KEYWORDS)
    if "icom_taxonomy_text" not in st.session_state:
        st.session_state.icom_taxonomy_text = taxonomy_to_text(ICOM_KEYWORDS)

    law_name = st.text_input(
        "Nombre de la ley o marco 1",
        value=st.session_state.law_name,
        help="Este nombre sera el titulo de la columna en el Excel. Puedes cambiarlo por otra ley.",
    )
    icom_name = st.text_input(
        "Nombre del marco 2",
        value=st.session_state.icom_name,
        help="Este nombre sera el titulo de la segunda columna de etiquetas.",
    )

    with st.expander(law_name, expanded=True):
        law_taxonomy_text = st.text_area(
            "Etiquetas y palabras clave del marco 1",
            value=st.session_state.law_taxonomy_text,
            height=320,
            help="Una etiqueta por linea. Ejemplo: contratos_publicos = contrato, licitacion, adjudicacion",
        )

    with st.expander(icom_name, expanded=True):
        icom_taxonomy_text = st.text_area(
            "Etiquetas y palabras clave del marco 2",
            value=st.session_state.icom_taxonomy_text,
            height=320,
            help="Una etiqueta por linea. Ejemplo: accesibilidad = accesible, discapacidad, inclusion",
        )

    col_reset_ley, col_reset_icom = st.columns(2)
    with col_reset_ley:
        if st.button("Restaurar marco 1"):
            st.session_state.law_name = "ley 19/2013"
            st.session_state.law_taxonomy_text = taxonomy_to_text(LEY_19_2013_KEYWORDS)
            st.rerun()
    with col_reset_icom:
        if st.button("Restaurar marco 2"):
            st.session_state.icom_name = "codigo deontologico icom"
            st.session_state.icom_taxonomy_text = taxonomy_to_text(ICOM_KEYWORDS)
            st.rerun()

law_taxonomy = text_to_taxonomy(law_taxonomy_text)
icom_taxonomy = text_to_taxonomy(icom_taxonomy_text)

if not law_taxonomy:
    st.warning("El marco 1 no tiene etiquetas validas. Usa el formato: etiqueta = palabra1, palabra2")
if not icom_taxonomy:
    st.warning("El marco 2 no tiene etiquetas validas. Usa el formato: etiqueta = palabra1, palabra2")

st.subheader("Cantidad de parrafos representativos")
k_medoids = st.number_input(
    "Numero de K medoides",
    min_value=1,
    max_value=200,
    value=24,
    step=1,
    help="Controla cuantos parrafos representativos apareceran en el Excel final. Por ejemplo: si pones 24, el Excel tendra 24 parrafos.",
)

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

    selected_items = select_k_medoids_paragraphs(items, k=int(k_medoids), seed=42)
    master_rows = []
    preview_rows = []
    for paragraph in selected_items:
        norm_labels = normative_columns_custom(paragraph, law_taxonomy, icom_taxonomy)
        master_rows.append([
            paragraph,
            norm_labels["ley"],
            norm_labels["codigo deontologico"],
            norm_labels["otros"],
        ])
        preview_rows.append({
            "parrafo": paragraph,
            law_name: norm_labels["ley"],
            icom_name: norm_labels["codigo deontologico"],
            "otros temas": norm_labels["otros"],
        })
    master_excel_bytes = make_excel_bytes(
        master_rows,
        header=["parrafo", law_name, icom_name, "otros temas"],
    )

    excel_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    st.success(f"Archivo madre listo con {len(selected_items)} parrafos representativos.")
    st.download_button("Descargar archivo_madre_etiquetado.xlsx", data=master_excel_bytes, file_name="archivo_madre_etiquetado.xlsx", mime=excel_mime)

    # Opcional: previews
    st.subheader("Vista previa (primeras filas)")
    st.write("**Archivo madre**")
    st.dataframe(preview_rows[:min(10, len(preview_rows))])
