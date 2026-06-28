# streamlit_app.py
# App Streamlit: input = texto, output = archivo madre Excel etiquetado.
from __future__ import annotations

import io
import json
import re
import urllib.error
import urllib.request
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
    "institucion", "institucional", "museo", "museos",
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

    widths = [90, 35, 35, 35, 18, 55, 35, 35]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + index)].width = width
    wb.save(buf)
    return buf.getvalue()


def label_to_display(label: str) -> str:
    icom_labels = {
        "dimension_01_transparencia_institucional": "transparencia institucional",
        "dimension_02_transparencia_economica": "transparencia economica",
        "dimension_03_transparencia_colecciones": "transparencia sobre las colecciones",
        "dimension_04_transparencia_procedencia_legalidad": "transparencia sobre procedencia y legalidad",
        "dimension_05_transparencia_cientifica": "transparencia cientifica",
        "dimension_06_transparencia_educativa": "transparencia educativa",
        "dimension_07_transparencia_acceso": "transparencia sobre el acceso",
        "dimension_08_transparencia_servicio_publico": "transparencia del servicio publico",
        "dimension_09_transparencia_social": "transparencia social",
        "dimension_10_transparencia_cooperacion_institucional": "transparencia en cooperacion institucional",
        "dimension_11_transparencia_juridica": "transparencia juridica",
        "dimension_12_transparencia_etica": "transparencia etica",
    }
    if label in icom_labels:
        return icom_labels[label]
    label = re.sub(r"^dimension_\d+_", "", label)
    prefixes = [
        ("informacion_institucional_", "informacion institucional"),
        ("informacion_organizativa_", "informacion organizativa"),
        ("planificacion_", "informacion sobre planificacion"),
        ("evaluacion_", "evaluacion del grado de cumplimiento"),
        ("normativa_", "normativa que las rige"),
        ("funciones_", "funciones que desarrolla"),
        ("informacion_juridica_", "informacion juridica relevante"),
        ("contratos_", "contratos"),
        ("convenios_", "convenios"),
        ("subvenciones_", "subvenciones y ayudas publicas"),
        ("presupuestos_", "presupuestos"),
        ("cuentas_anuales_", "cuentas anuales"),
        ("estadisticas_", "estadisticas"),
        ("acceso_", "acceso a la informacion"),
        ("buen_gobierno_", "buen gobierno"),
    ]
    for prefix, group in prefixes:
        if label.startswith(prefix):
            topic = label.removeprefix(prefix).replace("_", " ")
            return f"{group}, {topic}"
    return label.replace("_", " ")


def join_display_labels(labels: List[str]) -> str:
    return ", ".join(label_to_display(label) for label in labels)


def display_taxonomy(taxonomy: Dict[str, List[str]]) -> Dict[str, List[str]]:
    return {label_to_display(label): keywords for label, keywords in taxonomy.items()}


def tokens_for_candidate_selection(text: str) -> set[str]:
    normalized = normalize_for_keywords(text.replace("_", " "))
    tokens = re.findall(r"\b[a-z0-9]{3,}\b", normalized)
    return {token for token in tokens if token not in STOPWORDS}


def select_candidate_taxonomies(
    paragraph: str,
    law_taxonomy: Dict[str, List[str]],
    icom_taxonomy: Dict[str, List[str]],
    max_candidates: int = 18,
) -> tuple[Dict[str, List[str]], Dict[str, List[str]]]:
    paragraph_tokens = tokens_for_candidate_selection(paragraph)
    scored = []
    for marco, taxonomy in [("ley", law_taxonomy), ("icom", icom_taxonomy)]:
        for label, keywords in taxonomy.items():
            display_label = label_to_display(label)
            candidate_tokens = tokens_for_candidate_selection(" ".join([display_label, *keywords]))
            overlap = paragraph_tokens.intersection(candidate_tokens)
            if overlap:
                score = len(overlap)
                scored.append((score, marco, display_label, keywords))

    if not scored:
        return {}, display_taxonomy(icom_taxonomy)

    scored.sort(reverse=True)
    law_candidates = {}
    icom_candidates = {}
    for _, marco, label, keywords in scored[:max_candidates]:
        if marco == "ley":
            law_candidates[label] = keywords
        else:
            icom_candidates[label] = keywords
    return law_candidates, icom_candidates


def taxonomy_to_text(taxonomy: Dict[str, List[str]]) -> str:
    return "\n".join(
        f"{label_to_display(label)} = {', '.join(keywords)}"
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


def taxonomy_options_text(
    law_taxonomy: Dict[str, List[str]],
    icom_taxonomy: Dict[str, List[str]],
) -> str:
    lines = ["LEY_19_2013:"]
    for label, keywords in law_taxonomy.items():
        lines.append(f"- {label}: {', '.join(keywords)}")
    lines.append("CODIGO_MUSEOS:")
    for label, keywords in icom_taxonomy.items():
        lines.append(f"- {label}: {', '.join(keywords)}")
    return "\n".join(lines)


def keyword_hits_for_label(text: str, taxonomy: Dict[str, List[str]], label: str) -> List[str]:
    if label == "indeterminado" or label not in taxonomy:
        return []
    normalized_text = normalize_for_keywords(text)
    return [
        keyword
        for keyword in taxonomy[label]
        if normalize_for_keywords(keyword) in normalized_text
    ]


def quantitative_ia_justification(
    paragraph: str,
    law_taxonomy: Dict[str, List[str]],
    icom_taxonomy: Dict[str, List[str]],
    law_label: str,
    icom_label: str,
) -> str:
    law_hits = keyword_hits_for_label(paragraph, law_taxonomy, law_label)
    icom_hits = keyword_hits_for_label(paragraph, icom_taxonomy, icom_label)
    law_total = len(law_taxonomy.get(law_label, []))
    icom_total = len(icom_taxonomy.get(icom_label, []))

    law_summary = (
        f"Ley 19/2013: {law_label} ({len(law_hits)}/{law_total} indicios lexicos)"
        if law_label != "indeterminado"
        else "Ley 19/2013: indeterminado (0 indicios suficientes)"
    )
    icom_summary = (
        f"Codigo deontologico: {icom_label} ({len(icom_hits)}/{icom_total} indicios lexicos)"
        if icom_label != "indeterminado"
        else "Codigo deontologico: indeterminado (0 indicios suficientes)"
    )
    return (
        "Codificacion manual simulada: "
        f"{law_summary}; {icom_summary}. "
        "La asignacion se basa en la concentracion relativa de evidencias textuales del parrafo."
    )


def extract_json_object(text: str) -> Dict[str, object]:
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", text, flags=re.DOTALL)
        if not match:
            raise
        return json.loads(match.group(0))


def classify_with_ollama(
    paragraph: str,
    law_taxonomy: Dict[str, List[str]],
    icom_taxonomy: Dict[str, List[str]],
    model: str,
    ollama_url: str,
) -> Dict[str, str]:
    law_display_taxonomy = display_taxonomy(law_taxonomy)
    icom_display_taxonomy = display_taxonomy(icom_taxonomy)
    law_prompt_taxonomy, icom_prompt_taxonomy = select_candidate_taxonomies(
        paragraph,
        law_taxonomy,
        icom_taxonomy,
    )
    valid_law_labels = set(law_display_taxonomy)
    valid_icom_labels = set(icom_display_taxonomy)
    prompt = f"""
Eres un clasificador academico para una tesis doctoral.
Debes leer el parrafo y asignar DOS etiquetas independientes dentro de taxonomias cerradas:
1. Una etiqueta segun Ley 19/2013.
2. Una etiqueta segun Codigo deontologico de museos.

No inventes etiquetas. Si ninguna etiqueta corresponde razonablemente en una taxonomia, usa "indeterminado" para esa taxonomia.
Actua como una persona codificadora que revisa evidencias textuales y asigna categorias normativas.

Taxonomia cerrada:
{taxonomy_options_text(law_prompt_taxonomy, icom_prompt_taxonomy)}

Parrafo:
\"\"\"{paragraph}\"\"\"

Responde solo JSON valido con este esquema:
{{
  "etiqueta_ia_ley_19_2013": "una etiqueta exacta de LEY_19_2013 o indeterminado",
  "etiqueta_ia_codigo_deontologico": "una etiqueta exacta de CODIGO_MUSEOS o indeterminado"
}}
"""
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "format": "json",
        "options": {"temperature": 0.0},
    }
    request = urllib.request.Request(
        ollama_url.rstrip("/") + "/api/generate",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            raw = json.loads(response.read().decode("utf-8", errors="replace"))
        parsed = extract_json_object(raw.get("response", "{}"))
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, KeyError) as exc:
        return {
            "etiqueta_ia_ley_19_2013": "indeterminado",
            "etiqueta_ia_codigo_deontologico": "indeterminado",
            "justificacion_ia": f"Error al consultar Ollama: {exc}",
        }

    law_label = str(parsed.get("etiqueta_ia_ley_19_2013", "indeterminado")).strip()
    icom_label = str(parsed.get("etiqueta_ia_codigo_deontologico", "indeterminado")).strip()
    if law_label not in valid_law_labels:
        law_label = "indeterminado"
    if icom_label not in valid_icom_labels:
        icom_label = "indeterminado"

    return {
        "etiqueta_ia_ley_19_2013": law_label,
        "etiqueta_ia_codigo_deontologico": icom_label,
        "justificacion_ia": quantitative_ia_justification(
            paragraph,
            law_taxonomy,
            icom_taxonomy,
            law_label,
            icom_label,
        ),
    }


def automated_labeling_columns(
    text: str,
    law_taxonomy: Dict[str, List[str]],
    icom_taxonomy: Dict[str, List[str]],
    model: str,
    ollama_url: str,
) -> Dict[str, str]:
    law_labels = keyword_labels(text, law_taxonomy)
    icom_labels = keyword_labels(text, icom_taxonomy)
    ia_labels = classify_with_ollama(text, law_taxonomy, icom_taxonomy, model, ollama_url)
    final_candidates = [
        ia_labels["etiqueta_ia_ley_19_2013"],
        ia_labels["etiqueta_ia_codigo_deontologico"],
    ]
    final_label = ", ".join(label for label in final_candidates if label != "indeterminado")
    if not final_label:
        dict_labels = [label_to_display(label) for label in law_labels + icom_labels]
        final_label = ", ".join(dict_labels) if dict_labels else "indeterminado"

    return {
        "ley_diccionario": join_display_labels(law_labels),
        "codigo_diccionario": join_display_labels(icom_labels),
        **ia_labels,
        "etiqueta_final": final_label,
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
        - **Etiqueta final**: resume las etiquetas IA asignadas para los dos marcos normativos.
        - **Restaurar marco 1 / Restaurar marco 2**: vuelve a poner las etiquetas originales si cambiaste algo y quieres empezar de nuevo.
        - **Procesar**: crea el Excel final.
        - **Archivo madre etiquetado**: el unico archivo de salida. Contiene una fila por parrafo, evidencia por diccionario, clasificacion IA por marco, justificacion cuantitativa y etiqueta final.
        """
    )

with st.expander("Nota metodologica", expanded=False):
    st.markdown(
        """
        La version actual usa una clasificacion automatizada con taxonomia
        cerrada: el diccionario conserva evidencia textual y la IA asigna
        una etiqueta interpretativa para Ley 19/2013 y otra para el Codigo
        deontologico, sin poder inventar categorias nuevas.
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

st.subheader("Clasificacion automatizada con IA")
ollama_model = st.text_input(
    "Modelo Ollama",
    value="mistral:latest",
    help="Modelo local usado para clasificar cada parrafo dentro de la taxonomia cerrada.",
)
ollama_url = st.text_input(
    "URL de Ollama",
    value="http://127.0.0.1:11434",
    help="Endpoint del servidor Ollama. En el remoto normalmente se deja como 127.0.0.1.",
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
        norm_labels = automated_labeling_columns(
            paragraph,
            law_taxonomy,
            icom_taxonomy,
            ollama_model,
            ollama_url,
        )
        master_rows.append(
            [
                paragraph,
                norm_labels["ley_diccionario"],
                norm_labels["codigo_diccionario"],
                norm_labels["etiqueta_ia_ley_19_2013"],
                norm_labels["etiqueta_ia_codigo_deontologico"],
                norm_labels["justificacion_ia"],
                norm_labels["etiqueta_final"],
            ]
        )
        preview_rows.append(
            {
                "parrafo": paragraph,
                f"{law_name}_diccionario": norm_labels["ley_diccionario"],
                f"{icom_name}_diccionario": norm_labels["codigo_diccionario"],
                "etiqueta_ia_ley_19_2013": norm_labels["etiqueta_ia_ley_19_2013"],
                "etiqueta_ia_codigo_deontologico": norm_labels["etiqueta_ia_codigo_deontologico"],
                "justificacion_ia": norm_labels["justificacion_ia"],
                "etiqueta_final": norm_labels["etiqueta_final"],
            }
        )

    master_excel_bytes = make_excel_bytes(
        master_rows,
        header=[
            "parrafo",
            f"{law_name}_diccionario",
            f"{icom_name}_diccionario",
            "etiqueta_ia_ley_19_2013",
            "etiqueta_ia_codigo_deontologico",
            "justificacion_ia",
            "etiqueta_final",
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
