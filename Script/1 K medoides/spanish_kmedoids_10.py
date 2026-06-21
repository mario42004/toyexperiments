#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
import random
import argparse
import math
import unicodedata
from collections import Counter
from pathlib import Path
from typing import List, Dict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

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


def select_k_medoids_paragraphs(paragraphs: List[str], k: int = 24, seed: int = 42) -> List[str]:
    """Selecciona K parrafos representativos con una aproximacion ligera a K-Medoids."""
    if not paragraphs:
        return []
    if k <= 0:
        k = 1
    if len(paragraphs) <= k:
        return paragraphs

    vectors = tfidf_vectors(paragraphs)
    distances = cosine_distance_matrix(vectors)
    medoid_indices = initialize_medoids(distances, k)
    medoid_indices = improve_medoids(distances, medoid_indices, max_iter=5)
    return [paragraphs[idx] for idx in medoid_indices]


def tokenize_for_tfidf(text: str) -> List[str]:
    normalized = normalize_for_keywords(text)
    return re.findall(r"\b[a-z0-9]{3,}\b", normalized)


def tfidf_vectors(texts: List[str]) -> List[Dict[str, float]]:
    tokenized = [tokenize_for_tfidf(text) for text in texts]
    doc_freq = Counter()
    for tokens in tokenized:
        doc_freq.update(set(tokens))

    total_docs = len(texts)
    vectors = []
    for tokens in tokenized:
        counts = Counter(tokens)
        total_terms = sum(counts.values()) or 1
        vector = {}
        for token, count in counts.items():
            tf = count / total_terms
            idf = math.log((1 + total_docs) / (1 + doc_freq[token])) + 1
            vector[token] = tf * idf
        norm = math.sqrt(sum(value * value for value in vector.values())) or 1
        vectors.append({token: value / norm for token, value in vector.items()})
    return vectors


def cosine_distance(vec_a: Dict[str, float], vec_b: Dict[str, float]) -> float:
    if len(vec_a) > len(vec_b):
        vec_a, vec_b = vec_b, vec_a
    similarity = sum(value * vec_b.get(token, 0.0) for token, value in vec_a.items())
    return 1.0 - similarity


def cosine_distance_matrix(vectors: List[Dict[str, float]]) -> List[List[float]]:
    size = len(vectors)
    distances = [[0.0] * size for _ in range(size)]
    for i in range(size):
        for j in range(i + 1, size):
            distance = cosine_distance(vectors[i], vectors[j])
            distances[i][j] = distance
            distances[j][i] = distance
    return distances


def initialize_medoids(distances: List[List[float]], k: int) -> List[int]:
    total_distances = [sum(row) for row in distances]
    medoids = [min(range(len(distances)), key=lambda idx: total_distances[idx])]
    while len(medoids) < k:
        candidates = [idx for idx in range(len(distances)) if idx not in medoids]
        next_medoid = max(candidates, key=lambda idx: min(distances[idx][m] for m in medoids))
        medoids.append(next_medoid)
    return medoids


def medoid_cost(distances: List[List[float]], medoids: List[int]) -> float:
    return sum(min(row[medoid] for medoid in medoids) for row in distances)


def improve_medoids(distances: List[List[float]], medoids: List[int], max_iter: int = 5) -> List[int]:
    medoids = list(medoids)
    best_cost = medoid_cost(distances, medoids)
    all_indices = set(range(len(distances)))

    for _ in range(max_iter):
        improved = False
        non_medoids = sorted(all_indices.difference(medoids))
        for medoid in list(medoids):
            medoid_position = medoids.index(medoid)
            for candidate in non_medoids:
                trial = list(medoids)
                trial[medoid_position] = candidate
                cost = medoid_cost(distances, trial)
                if cost < best_cost:
                    medoids = trial
                    best_cost = cost
                    improved = True
                    break
            if improved:
                break
        if not improved:
            break

    return sorted(medoids)


def write_xlsx(path: str, header: List[str], rows: List[List]):
    """Escribe una hoja Excel simple con cabecera y filas."""
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

    wb.save(path)


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
    parser = argparse.ArgumentParser(description="Seleccion de parrafos representativos con K-Medoids y etiquetado normativo")
    parser.add_argument("--file", type=str, required=True, help="Ruta al archivo .txt")
    parser.add_argument("--k", type=int, default=24, help="Numero de parrafos representativos que quieres obtener")
    parser.add_argument("--max_k", type=int, default=35, help="Máximo K para el elbow")
    parser.add_argument("--model", type=str, default=DEFAULT_MODEL, help="Modelo BERT de Transformers")
    parser.add_argument("--seed", type=int, default=42, help="Semilla aleatoria")
    parser.add_argument("--save_dir", type=str, default="./resultados_kmedoids", help="Carpeta de salida")
    parser.add_argument("--distance_metric", type=str, default="cosine", choices=["cosine", "euclidean"], help="Métrica de distancia")
    args = parser.parse_args()

    set_seed(args.seed)
    os.makedirs(args.save_dir, exist_ok=True)

    base_name = Path(args.file).stem
    print(f"\nProcesando archivo: {args.file}")
    print(f"Guardando en: {args.save_dir}/")

    # --- Lectura y segmentación ---
    text = read_txt_file(args.file)
    items = split_sentences_or_paragraphs(text)
    selected_items = select_k_medoids_paragraphs(items, k=args.k, seed=args.seed)
    print(f"Total de oraciones/parrafos detectados: {len(items)}")
    print(f"Parrafos representativos seleccionados (K): {len(selected_items)}")

    # --- Archivo madre: parrafos completos + etiquetas normativas ---
    master_xlsx = os.path.join(args.save_dir, f"{base_name}_archivo_madre_etiquetado.xlsx")
    master_rows = []
    for paragraph in selected_items:
        norm_labels = normative_columns(paragraph)
        master_rows.append([
            paragraph,
            norm_labels["ley 19/2013"],
            norm_labels["codigo deontologico"],
            norm_labels["otros"],
        ])
    write_xlsx(
        master_xlsx,
        ["parrafo", "ley 19/2013", "codigo deontologico icom", "otros temas"],
        master_rows,
    )
    print(f"[OK] Guardado archivo madre: {master_xlsx}")

    print(f"\nCompletado: {base_name}")
    print("Archivo generado:")
    print(f"   - {os.path.basename(master_xlsx)}")


if __name__ == "__main__":
    main()
