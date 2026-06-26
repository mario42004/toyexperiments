# streamlit_app.py
# App Streamlit: input = texto, output = archivo madre Excel etiquetado.
from __future__ import annotations

import io
import re
from typing import Dict, List

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from spanish_kmedoids_10 import (
    ICOM_KEYWORDS,
    LEY_19_2013_KEYWORDS,
    keyword_labels,
    normalize_for_keywords,
    select_k_medoids_paragraphs,
)


# --------------------------
# Utilidades
# --------------------------
STOPWORDS = {
    "para", "por", "con", "sin", "del", "las", "los", "una", "uno", "unos",
    "unas", "que", "como", "sobre", "entre", "desde", "hasta", "esta", "este",
    "estos", "estas", "tambien", "donde", "cuando", "cada", "otros", "otras",
    "informacion", "transparencia", "dimension", "codigo", "marco", "temas",
}


def is_legible_paragraph(text: str) -> bool:
    letters = re.findall(r"[^\W\d_]", text)
    visible_chars = re.findall(r"\S", text)
    words = re.findall(r"[^\W\d_]{3,}", text.lower())
    short_tokens = re.findall(r"\b[^\W\d_]{1,2}\b", text.lower())
    numeric_tokens = re.findall(r"\b\d+(?:[.,]\d+)?\b", text)
    table_tokens = re.findall(r"\b(?:td|tr|lt|gt)\b", text.lower())
    all_tokens = re.findall(r"\b\w+\b", text.lower())

    if len(letters) < 12:
        return False
    if len(words) < 4:
        return False
    if len(table_tokens) >= 3:
        return False
    if len(numeric_tokens) > len(words):
        return False
    if len(short_tokens) / max(len(all_tokens), 1) > 0.35:
        return False
    return len(letters) / max(len(visible_chars), 1) >= 0.45


def split_sentences_or_paragraphs(text: str) -> List[str]:
    """
    Divide el texto en oraciones o parrafos.
    Usa saltos de linea si existen; si no, divide por signos de puntuacion.
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if len(lines) > 1:
        return [line for line in lines if is_legible_paragraph(line)]

    sentences = [s.strip() for s in re.split(r"[.!?]\s+", text) if s.strip()]
    return [sentence for sentence in sentences if is_legible_paragraph(sentence)]


def semantic_tokens(text: str) -> set[str]:
    normalized = normalize_for_keywords(text.replace("_", " "))
    tokens = re.findall(r"\b[a-z0-9]{3,}\b", normalized)
    return {token for token in tokens if token not in STOPWORDS}


def semantic_label_suggestions(
    text: str,
    taxonomy: Dict[str, List[str]],
    threshold: float,
    max_labels: int = 2,
) -> List[str]:
    text_tokens = semantic_tokens(text)
    if not text_tokens:
        return []

    scored_labels = []
    for label, keywords in taxonomy.items():
        label_text = label.replace("_", " ")
        prototype_tokens = semantic_tokens(" ".join([label_text, *keywords]))
        if not prototype_tokens:
            continue

        overlap = text_tokens.intersection(prototype_tokens)
        if not overlap:
            continue

        score = len(overlap) / ((len(text_tokens) * len(prototype_tokens)) ** 0.5)
        if score >= threshold:
            scored_labels.append((score, label))

    scored_labels.sort(reverse=True)
    return [label for _, label in scored_labels[:max_labels]]


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

    widths = [90, 35, 35, 35, 35, 45, 20]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + index)].width = width
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


def assisted_labeling_columns(
    text: str,
    law_taxonomy: Dict[str, List[str]],
    icom_taxonomy: Dict[str, List[str]],
    semantic_threshold: float,
) -> Dict[str, str]:
    law_labels = keyword_labels(text, law_taxonomy)
    icom_labels = keyword_labels(text, icom_taxonomy)
    law_semantic = [] if law_labels else semantic_label_suggestions(text, law_taxonomy, semantic_threshold)
    icom_semantic = [] if icom_labels else semantic_label_suggestions(text, icom_taxonomy, semantic_threshold)
    final_labels = law_labels + icom_labels + law_semantic + icom_semantic
    used_semantic = bool(law_semantic or icom_semantic)
    return {
        "ley_diccionario": "; ".join(law_labels),
        "codigo_diccionario": "; ".join(icom_labels),
        "ley_semantica": "; ".join(law_semantic),
        "codigo_semantica": "; ".join(icom_semantic),
        "etiqueta_final": "; ".join(final_labels) if final_labels else "otros temas",
        "requiere_revision": "si" if used_semantic or not final_labels else "no",
    }


# --------------------------
# UI
# --------------------------
st.set_page_config(
    page_title="Archivo madre etiquetado",
    page_icon=":material/table_chart:",
    layout="centered",
)

st.markdown(
    """
    <style>
    div.stButton > button[kind="primary"] {
        background-color: #16a34a;
        border: 1px solid #15803d;
        border-radius: 8px;
        color: white;
        font-size: 1.35rem;
        font-weight: 700;
        min-height: 4rem;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #15803d;
        border-color: #166534;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Archivo madre etiquetado")
st.write(
    "Esta herramienta lee una memoria, separa sus parrafos y genera un solo "
    "Excel con las etiquetas normativas."
)

with st.expander("Que significa cada parte", expanded=True):
    st.markdown(
        """
        - **Pegar texto**: pega aqui el contenido de una memoria si quieres probar sin subir archivo.
        - **Subir archivo .txt/.md**: carga una memoria completa desde tu ordenador.
        - **Numero de K medoides**: cantidad de parrafos representativos que quieres obtener. Si pones 24, el Excel tendra 24 parrafos.
        - **Nombre de la ley o marco 1**: cambia aqui Ley 19/2013 por cualquier otra ley o marco que quieras analizar.
        - **Nombre del marco 2**: puedes dejar Codigo deontologico ICOM o cambiarlo por otro marco.
        - **Etiquetas y palabras clave**: escribe una etiqueta por linea con este formato: etiqueta = palabra1, palabra2, palabra3.
        - **Etiqueta final**: combina etiquetas por diccionario y, si falta coincidencia literal, sugerencias semanticas marcadas para revision.
        - **Restaurar marco 1 / Restaurar marco 2**: vuelve a poner las etiquetas originales si cambiaste algo y quieres empezar de nuevo.
        - **Procesar**: crea el Excel final.
        - **Archivo madre etiquetado**: el unico archivo de salida. Contiene una fila por parrafo y columnas separadas para diccionario, sugerencia semantica, etiqueta final y revision.
        """
    )

with st.expander("Nota metodologica", expanded=False):
    st.markdown(
        """
        La version actual usa una clasificacion asistida por diccionario normativo:
        es transparente, reproducible y facil de revisar. La capa semantica
        funciona como apoyo para parrafos sin coincidencias literales y sus
        resultados quedan marcados para revision.
        """
    )

with st.container():
    st.subheader("Editar leyes y etiquetas")
    st.caption(
        "Formato: una etiqueta por linea. Ejemplo: "
        "presupuesto_aprobado = presupuesto, cuentas, gasto"
    )

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
    help=(
        "Controla cuantos parrafos representativos apareceran en el Excel final. "
        "Por ejemplo: si pones 24, el Excel tendra 24 parrafos."
    ),
)

semantic_threshold = st.slider(
    "Sensibilidad de la sugerencia semantica",
    min_value=0.03,
    max_value=0.30,
    value=0.08,
    step=0.01,
    help=(
        "Valores mas bajos sugieren mas etiquetas, pero pueden generar mas falsos positivos. "
        "Valores mas altos son mas conservadores."
    ),
)

st.write(
    "**Entrada de texto** (pega tu texto o sube un .txt/.md). Si hay saltos "
    "de linea, se usan como parrafos; si no, se segmenta por oraciones."
)
col1, col2 = st.columns(2)
with col1:
    texto = st.text_area(
        "Pegar texto",
        height=220,
        placeholder="Pega aqui tu texto en espanol...",
    )
with col2:
    up = st.file_uploader("...o subir archivo .txt/.md", type=["txt", "md"])

if up and not texto.strip():
    texto = up.read().decode("utf-8", errors="replace")

_, process_col, _ = st.columns([1, 2, 1])
with process_col:
    process_clicked = st.button("Procesar", type="primary", use_container_width=True)

if process_clicked:
    if not texto.strip():
        st.error("Por favor, introduce texto o sube un .txt.")
        st.stop()

    items = split_sentences_or_paragraphs(texto)
    if len(items) == 0:
        st.error(
            "No se detectaron parrafos con suficiente texto legible. "
            "Se descartan fragmentos compuestos solo por numeros, codigos o signos."
        )
        st.stop()

    selected_items = select_k_medoids_paragraphs(items, k=int(k_medoids), seed=42)
    master_rows = []
    preview_rows = []

    for paragraph in selected_items:
        norm_labels = assisted_labeling_columns(
            paragraph,
            law_taxonomy,
            icom_taxonomy,
            semantic_threshold,
        )
        master_rows.append(
            [
                paragraph,
                norm_labels["ley_diccionario"],
                norm_labels["codigo_diccionario"],
                norm_labels["ley_semantica"],
                norm_labels["codigo_semantica"],
                norm_labels["etiqueta_final"],
                norm_labels["requiere_revision"],
            ]
        )
        preview_rows.append(
            {
                "parrafo": paragraph,
                f"{law_name}_diccionario": norm_labels["ley_diccionario"],
                f"{icom_name}_diccionario": norm_labels["codigo_diccionario"],
                f"{law_name}_semantica": norm_labels["ley_semantica"],
                f"{icom_name}_semantica": norm_labels["codigo_semantica"],
                "etiqueta_final": norm_labels["etiqueta_final"],
                "requiere_revision": norm_labels["requiere_revision"],
            }
        )

    master_excel_bytes = make_excel_bytes(
        master_rows,
        header=[
            "parrafo",
            f"{law_name}_diccionario",
            f"{icom_name}_diccionario",
            f"{law_name}_semantica",
            f"{icom_name}_semantica",
            "etiqueta_final",
            "requiere_revision",
        ],
    )

    excel_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    st.success(f"Archivo madre listo con {len(selected_items)} parrafos representativos.")
    st.download_button(
        "Descargar archivo_madre_etiquetado.xlsx",
        data=master_excel_bytes,
        file_name="archivo_madre_etiquetado.xlsx",
        mime=excel_mime,
    )

    st.subheader("Vista previa")
    st.write("**Archivo madre**")
    st.dataframe(preview_rows[: min(10, len(preview_rows))])
