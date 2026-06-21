# streamlit_app.py
# App Streamlit: input = texto, output = CSVs (elbow, medoids, assignments)
# Basado en la lógica de spanish_kmedoids_10.py (segmentación, embeddings, elbow, K-Medoids)

import io
import csv
import os
import re
import random
from pathlib import Path
from typing import List, Dict, Tuple

import numpy as np
import streamlit as st
import torch
from transformers import AutoTokenizer, AutoModel
from sklearn.metrics import pairwise_distances
from sklearn_extra.cluster import KMedoids

from spanish_kmedoids_10 import ICOM_KEYWORDS, LEY_19_2013_KEYWORDS, keyword_labels

DEFAULT_MODEL = "dccuchile/bert-base-spanish-wwm-cased"

# --------------------------
# Utilidades (adaptadas del script)
# --------------------------
def set_seed(seed: int = 42):
    random.seed(seed)
    np.random.seed(seed)
    torch.manual_seed(seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed_all(seed)

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

def mean_pooling(last_hidden_state: torch.Tensor, attention_mask: torch.Tensor) -> torch.Tensor:
    mask = attention_mask.unsqueeze(-1).expand(last_hidden_state.size()).float()
    summed = torch.sum(last_hidden_state * mask, dim=1)
    counts = torch.clamp(mask.sum(dim=1), min=1e-9)
    return summed / counts

@st.cache_resource(show_spinner=False)
def load_model(model_name: str):
    tokenizer = AutoTokenizer.from_pretrained(model_name)
    model = AutoModel.from_pretrained(model_name)
    device = "cuda" if torch.cuda.is_available() else "cpu"
    model.to(device)
    model.eval()
    return tokenizer, model, device

def get_bert_embeddings(texts: List[str], model_name: str = DEFAULT_MODEL, batch_size: int = 16) -> np.ndarray:
    tokenizer, model, device = load_model(model_name)
    embs = []
    with torch.no_grad():
        for i in range(0, len(texts), batch_size):
            batch = texts[i:i+batch_size]
            enc = tokenizer(batch, padding=True, truncation=True, return_tensors="pt", max_length=256)
            enc = {k: v.to(device) for k, v in enc.items()}
            out = model(**enc)
            pooled = mean_pooling(out.last_hidden_state, enc["attention_mask"])
            embs.append(pooled.detach().cpu().numpy())
    return np.vstack(embs)

def elbow_curve(D: np.ndarray, max_k: int = 12, seed: int = 42) -> Dict[int, float]:
    max_k = min(max_k, D.shape[0])
    results = {}
    for k in range(1, max_k + 1):
        model = KMedoids(n_clusters=k, metric="precomputed", random_state=seed)
        model.fit(D)
        labels_k = model.labels_
        medoids_k = model.medoid_indices_
        cost = float(sum(D[i, medoids_k[labels_k[i]]] for i in range(D.shape[0])))
        results[k] = cost
    return results

def make_csv_bytes(rows: List[List], header: List[str]) -> bytes:
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(header)
    for r in rows:
        w.writerow(r)
    return buf.getvalue().encode("utf-8")

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
st.set_page_config(page_title="Clustering de texto (K-Medoids)", page_icon="🧩", layout="centered")
st.title("🧩 Clustering de texto con BERT + K-Medoids (CSV outputs)")

with st.sidebar:
    st.header("Parámetros")
    modelo = st.text_input("Modelo HF", value=DEFAULT_MODEL, help="Nombre del modelo en Hugging Face")
    k = st.number_input("k (número de clusters)", min_value=1, max_value=100, value=6, step=1)
    max_k = st.number_input("max_k (para curva elbow)", min_value=1, max_value=200, value=12, step=1)
    metric = st.selectbox("Métrica de distancia", options=["cosine", "euclidean"], index=0)
    seed = st.number_input("Semilla", min_value=0, max_value=2**32-1, value=42, step=1)
    batch_size = st.number_input("Batch size (embeddings)", min_value=1, max_value=128, value=16, step=1)

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

    set_seed(int(seed))
    items = split_sentences_or_paragraphs(texto)
    if len(items) == 0:
        st.error("No se detectaron oraciones/párrafos en el texto.")
        st.stop()

    if k > len(items):
        st.warning(f"k={k} es mayor que el número de oraciones/párrafos ({len(items)}). Se ajusta a {len(items)}.")
        k = len(items)

    with st.status("Calculando embeddings BERT…", expanded=False) as s:
        X = get_bert_embeddings(items, model_name=modelo, batch_size=int(batch_size))
        s.update(label="Calculando matriz de distancias…")
        D = pairwise_distances(X, metric=metric)
        s.update(label="Calculando curva elbow…")
        elbow = elbow_curve(D, max_k=int(max_k), seed=int(seed))
        s.update(label=f"Ejecutando K-Medoids con k={k}…")
        km = KMedoids(n_clusters=int(k), metric="precomputed", random_state=int(seed))
        km.fit(D)
        labels = km.labels_
        medoids_idx = km.medoid_indices_
        s.update(label="Listo ✅", state="complete")

    # ---- Construir CSVs en memoria ----
    # 1) elbow
    elbow_rows = [[int(kk), float(cost)] for kk, cost in sorted(elbow.items(), key=lambda x: x[0])]
    elbow_csv_bytes = make_csv_bytes(elbow_rows, header=["k", "cost"])

    # 2) medoids con etiquetas normativas
    medoids_rows = []
    for cid, idx in enumerate(medoids_idx):
        representative_sentence = items[idx]
        norm_labels = normative_columns_custom(representative_sentence, ley_taxonomy, icom_taxonomy)
        medoids_rows.append([
            cid,
            representative_sentence,
            norm_labels["ley 19/2013"],
            norm_labels["codigo deontologico"],
            norm_labels["otros"],
        ])
    medoids_csv_bytes = make_csv_bytes(
        medoids_rows,
        header=["cluster_id", "representative_sentence", "ley 19/2013", "codigo deontologico", "otros"],
    )

    # 3) assignments
    assign_rows = [[sent, int(lab)] for sent, lab in zip(items, labels)]
    assign_csv_bytes = make_csv_bytes(assign_rows, header=["sentence", "cluster_id"])

    st.success("Procesamiento completado. Descarga tus CSVs:")
    st.download_button("⬇️ Descargar elbow_k_vs_cost.csv", data=elbow_csv_bytes, file_name="elbow_k_vs_cost.csv", mime="text/csv")
    st.download_button("⬇️ Descargar kmedoids_medoids.csv", data=medoids_csv_bytes, file_name="kmedoids_medoids.csv", mime="text/csv")
    st.download_button("⬇️ Descargar cluster_assignments.csv", data=assign_csv_bytes, file_name="cluster_assignments.csv", mime="text/csv")

    # Opcional: previews
    st.subheader("Vista previa (primeras filas)")
    st.write("**Elbow (k vs cost)**")
    st.dataframe(elbow_rows[:10])
    st.write("**Medoids (frases representativas)**")
    st.dataframe(medoids_rows[:min(10, len(medoids_rows))])
    st.write("**Assignments (oración → cluster)**")
    st.dataframe(assign_rows[:min(10, len(assign_rows))])
