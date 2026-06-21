#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import csv
import random
import argparse
import unicodedata
from pathlib import Path
from typing import List, Dict

import numpy as np
import matplotlib.pyplot as plt
from sklearn.manifold import TSNE
from sklearn_extra.cluster import KMedoids
from sklearn.metrics import pairwise_distances

import torch
from transformers import AutoTokenizer, AutoModel

DEFAULT_MODEL = "dccuchile/bert-base-spanish-wwm-cased"


LEY_19_2013_KEYWORDS = {
    "identidad_institucional": ["museo", "institucion", "identidad", "mision", "vision"],
    "naturaleza_juridica": ["naturaleza juridica", "fundacion", "consorcio", "organismo", "entidad publica", "entidad"],
    "ambito_actuacion": ["ambito", "territorio", "actuacion", "competencia territorial"],
    "marco_normativo_aplicable": ["normativa", "ley", "real decreto", "reglamento", "estatutos", "marco normativo"],
    "funciones_institucionales": ["funciones", "fines", "cometido", "responsabilidades"],
    "competencias_asignadas": ["competencias", "atribuciones", "asignadas"],
    "servicios_prestados": ["servicios", "prestacion", "atencion al publico"],
    "estructura_organizativa": ["estructura organizativa", "estructura", "areas", "departamentos"],
    "organigrama": ["organigrama"],
    "equipo_directivo": ["director", "directora", "equipo directivo", "direccion", "gerencia"],
    "responsables_areas": ["responsable", "jefe de area", "jefa de area", "coordinador", "coordinadora"],
    "modelo_gobernanza": ["gobernanza", "patronato", "consejo rector", "comision ejecutiva", "junta"],
    "objetivos_anuales": ["objetivos", "objetivo anual", "metas"],
    "planes_y_programas": ["plan", "programa", "planes", "programas"],
    "lineas_estrategicas": ["lineas estrategicas", "estrategia", "estrategico"],
    "planificacion_actividad": ["planificacion", "calendario", "programacion prevista"],
    "actividades_realizadas": ["actividades", "realizadas", "desarrolladas", "celebradas"],
    "programacion_anual": ["programacion anual", "programacion"],
    "servicios_culturales": ["servicios culturales", "visitas", "talleres", "exposiciones"],
    "acciones_publicas": ["acciones publicas", "acto publico", "campana", "jornada"],
    "presupuesto_aprobado": ["presupuesto aprobado", "presupuesto"],
    "ejecucion_presupuestaria": ["ejecucion presupuestaria", "ejecutado", "liquidacion"],
    "fuentes_financiacion": ["financiacion", "fuentes de financiacion", "aportaciones"],
    "ingresos_propios": ["ingresos propios", "taquilla", "venta de entradas", "tienda"],
    "subvenciones_y_ayudas": ["subvencion", "subvenciones", "ayuda", "ayudas"],
    "gastos_desglosados": ["gastos", "costes", "desglose", "partidas"],
    "contratos_publicos": ["contrato", "contratos", "licitacion", "adjudicacion"],
    "convenios_colaboracion": ["convenio", "convenios"],
    "acuerdos_institucionales": ["acuerdo", "acuerdos institucionales"],
    "coproducciones": ["coproduccion", "coproducciones"],
    "personal_museo": ["personal", "plantilla", "empleados", "trabajadores"],
    "distribucion_funcional_personal": ["distribucion funcional", "puestos", "areas de personal"],
    "perfiles_profesionales": ["perfil profesional", "perfiles profesionales", "tecnicos", "conservadores"],
    "indicadores_resultado": ["indicadores", "resultados", "metricas"],
    "grado_cumplimiento_objetivos": ["cumplimiento", "grado de cumplimiento", "objetivos alcanzados"],
    "evaluacion_interna": ["evaluacion interna", "autoevaluacion"],
    "evaluacion_externa": ["evaluacion externa", "auditoria", "informe externo"],
    "publicidad_activa": ["publicidad activa", "portal de transparencia", "transparencia"],
    "canales_transparencia": ["canales de transparencia", "sede electronica", "pagina web", "web"],
    "derecho_acceso_informacion": ["derecho de acceso", "acceso a la informacion", "solicitud de informacion"],
}


ICOM_KEYWORDS = {
    "institucion_permanente": ["institucion permanente", "permanente"],
    "sin_animo_lucro": ["sin animo de lucro", "no lucrativa"],
    "servicio_sociedad": ["servicio de la sociedad", "servicio sociedad", "sociedad"],
    "coleccionar": ["coleccion", "coleccionar", "adquisicion", "adquisiciones"],
    "conservar": ["conservar", "conservacion", "preservacion"],
    "investigar": ["investigar", "investigacion", "estudio"],
    "interpretar": ["interpretar", "interpretacion"],
    "exhibir": ["exhibir", "exposicion", "exposiciones", "muestra"],
    "politica_colecciones": ["politica de colecciones", "gestion de colecciones"],
    "conservacion_preventiva": ["conservacion preventiva", "prevencion", "control ambiental"],
    "restauracion": ["restauracion", "restaurar", "restauradas", "restaurados"],
    "documentacion_colecciones": ["documentacion", "inventario", "catalogacion", "registro"],
    "prestamos_responsables": ["prestamo", "prestamos", "cesion", "deposito"],
    "proyectos_investigacion": ["proyecto de investigacion", "proyectos de investigacion"],
    "publicaciones": ["publicacion", "publicaciones", "articulo", "catalogo"],
    "produccion_editorial": ["produccion editorial", "edicion", "editorial"],
    "archivos_documentales": ["archivo", "archivos", "documental", "documentales"],
    "programas_educativos": ["programa educativo", "programas educativos", "educativo", "escolares"],
    "mediacion_cultural": ["mediacion", "mediacion cultural", "mediadores"],
    "educacion_no_reglada": ["educacion no reglada", "educacion informal"],
    "acceso_publico": ["acceso publico", "publico", "visitantes"],
    "accesibilidad": ["accesibilidad", "accesible", "discapacidad", "adaptado"],
    "inclusion": ["inclusion", "inclusivo", "diversidad", "vulnerabilidad"],
    "participacion_ciudadana": ["participacion ciudadana", "participacion", "comunidad"],
    "desarrollo_publicos": ["desarrollo de publicos", "publicos", "audiencias"],
    "principios_eticos": ["etica", "principios eticos", "codigo deontologico"],
    "buenas_practicas": ["buenas practicas", "calidad", "responsabilidad"],
    "profesionalidad": ["profesionalidad", "profesional", "formacion", "capacitacion"],
    "colaboraciones_institucionales": ["colaboracion", "colaboraciones", "alianza", "instituciones"],
    "trabajo_en_red": ["trabajo en red", "redes", "red de museos"],
}


# --------------------------
# Utilidades
# --------------------------
def set_seed(seed: int = 42):
    random.seed(seed)
    np.random.seed(seed)
    torch.manual_seed(seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed_all(seed)


def read_txt_file(path: str) -> str:
    """Lee texto con tolerancia a codificaciones."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except UnicodeDecodeError:
        try:
            with open(path, "r", encoding="latin-1") as f:
                return f.read()
        except UnicodeDecodeError:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                return f.read()


def split_sentences_or_paragraphs(text: str) -> List[str]:
    """
    Divide el texto en oraciones o párrafos.
    Usa saltos de línea si existen; si no, divide por puntos.
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if len(lines) > 1:
        return lines
    # si no hay saltos, segmentar por punto
    sentences = [s.strip() for s in re.split(r"[.!?]\s+", text) if len(s.strip()) > 0]
    return sentences


def mean_pooling(last_hidden_state: torch.Tensor, attention_mask: torch.Tensor) -> torch.Tensor:
    mask = attention_mask.unsqueeze(-1).expand(last_hidden_state.size()).float()
    summed = torch.sum(last_hidden_state * mask, dim=1)
    counts = torch.clamp(mask.sum(dim=1), min=1e-9)
    return summed / counts


def get_bert_embeddings(texts: List[str], model_name: str = DEFAULT_MODEL, device: str = None, batch_size: int = 16) -> np.ndarray:
    """Devuelve embeddings promedio para cada oración/párrafo."""
    if device is None:
        device = "cuda" if torch.cuda.is_available() else "cpu"

    tokenizer = AutoTokenizer.from_pretrained(model_name)
    model = AutoModel.from_pretrained(model_name)
    model.to(device)
    model.eval()

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


def pairwise_distance_matrix(X: np.ndarray, metric: str = "cosine") -> np.ndarray:
    """Devuelve matriz de distancias (por defecto: coseno)."""
    return pairwise_distances(X, metric=metric)


def normalize_for_keywords(text: str) -> str:
    """Normaliza texto para detectar etiquetas por palabras clave."""
    normalized = unicodedata.normalize("NFKD", text)
    without_accents = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    return without_accents.lower()


def keyword_labels(text: str, taxonomy: Dict[str, List[str]]) -> List[str]:
    """Devuelve las etiquetas cuya lista de palabras clave aparece en el texto."""
    normalized_text = normalize_for_keywords(text)
    labels = []
    for label, keywords in taxonomy.items():
        for keyword in keywords:
            if normalize_for_keywords(keyword) in normalized_text:
                labels.append(label)
                break
    return labels


def normative_columns(text: str) -> Dict[str, str]:
    """Clasifica una frase representativa en Ley 19/2013, ICOM u otros."""
    ley_labels = keyword_labels(text, LEY_19_2013_KEYWORDS)
    icom_labels = keyword_labels(text, ICOM_KEYWORDS)
    return {
        "ley 19/2013": "; ".join(ley_labels),
        "codigo deontologico": "; ".join(icom_labels),
        "otros": "" if ley_labels or icom_labels else "otros",
    }


def elbow_curve(D: np.ndarray, max_k: int = 12, seed: int = 42) -> Dict[int, float]:
    """Calcula curva elbow para distintos K."""
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


# --------------------------
# MAIN
# --------------------------
def main():
    parser = argparse.ArgumentParser(description="Clustering de oraciones/párrafos con BERT + K-Medoids")
    parser.add_argument("--file", type=str, required=True, help="Ruta al archivo .txt")
    parser.add_argument("--k", type=int, default=25, help="Número de clusters (K-Medoids)")
    parser.add_argument("--max_k", type=int, default=35, help="Máximo K para el elbow")
    parser.add_argument("--model", type=str, default=DEFAULT_MODEL, help="Modelo BERT de Transformers")
    parser.add_argument("--seed", type=int, default=42, help="Semilla aleatoria")
    parser.add_argument("--save_dir", type=str, default="./resultados_kmedoids", help="Carpeta de salida")
    parser.add_argument("--distance_metric", type=str, default="cosine", choices=["cosine", "euclidean"], help="Métrica de distancia")
    args = parser.parse_args()

    set_seed(args.seed)
    os.makedirs(args.save_dir, exist_ok=True)

    base_name = Path(args.file).stem
    print(f"\n📘 Procesando archivo: {args.file}")
    print(f"📂 Guardando en: {args.save_dir}/")

    # --- Lectura y segmentación ---
    text = read_txt_file(args.file)
    items = split_sentences_or_paragraphs(text)
    print(f"🧩 Total de oraciones/párrafos detectados: {len(items)}")

    if len(items) < args.k:
        print(f"[WARN] K ({args.k}) mayor que número de oraciones ({len(items)}). Se ajusta a {len(items)}.")
        args.k = len(items)

    # --- Embeddings + distancias ---
    device = "cuda" if torch.cuda.is_available() else "cpu"
    X = get_bert_embeddings(items, model_name=args.model, device=device)
    D = pairwise_distance_matrix(X, metric=args.distance_metric)

    # --- Elbow ---
    elbow = elbow_curve(D, max_k=args.max_k, seed=args.seed)
    elbow_csv = os.path.join(args.save_dir, f"{base_name}_elbow_k_vs_cost.csv")
    with open(elbow_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["k", "cost"])
        for k in sorted(elbow.keys()):
            w.writerow([k, elbow[k]])
    print(f"[OK] Guardado elbow: {elbow_csv}")

    # --- K-Medoids ---
    print(f"⚙️  Ejecutando K-Medoids con k={args.k} ...")
    model = KMedoids(n_clusters=args.k, metric="precomputed", random_state=args.seed)
    model.fit(D)
    labels = model.labels_
    medoids = model.medoid_indices_
    medoid_texts = [items[i] for i in medoids]

    # --- CSV 1: Medoids (frases representativas) ---
    medoids_csv = os.path.join(args.save_dir, f"{base_name}_kmedoids_medoids.csv")
    with open(medoids_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["cluster_id", "representative_sentence", "ley 19/2013", "codigo deontologico", "otros"])
        for cid, mtxt in enumerate(medoid_texts):
            labels = normative_columns(mtxt)
            w.writerow([cid, mtxt, labels["ley 19/2013"], labels["codigo deontologico"], labels["otros"]])
    print(f"[OK] Guardado medoides: {medoids_csv}")

    # --- CSV 2: Asignaciones (frase → cluster) ---
    assign_csv = os.path.join(args.save_dir, f"{base_name}_cluster_assignments.csv")
    with open(assign_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["sentence", "cluster_id"])
        for sent, lab in zip(items, labels):
            w.writerow([sent, int(lab)])
    print(f"[OK] Guardado asignaciones: {assign_csv}")

    print(f"\n✅ Completado: {base_name}")
    print(f"📄 Archivos generados:")
    print(f"   ├─ {os.path.basename(elbow_csv)}")
    print(f"   ├─ {os.path.basename(medoids_csv)}")
    print(f"   └─ {os.path.basename(assign_csv)}")


if __name__ == "__main__":
    main()
