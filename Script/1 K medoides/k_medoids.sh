#!/bin/bash
# ===============================================================
# Ejecuta el etiquetado de parrafos sobre todos los .md y guarda
# un unico archivo madre Excel por memoria en ./resultados_kmedoids
# ===============================================================

set -euo pipefail

# --- Configuración ---
SCRIPT="spanish_kmedoids_10.py"   # <-- pon aquí el nombre de tu script .py
INPUT_DIR="./memorias/2023"
OUTPUT_DIR="./resultados_kmedoids/2023"

K=25                # número de clusters por archivo
MAX_K=35           # máximo K para el elbow
MODEL="dccuchile/bert-base-spanish-wwm-cased"
METRIC="cosine"    # cosine | euclidean
SEED=42

# --- Preparación ---
mkdir -p "$OUTPUT_DIR"

echo "==========================================================="
echo "  Procesando .md desde:  $INPUT_DIR"
echo "  Resultados en:         $OUTPUT_DIR"
echo "  Script:                $SCRIPT"
echo "==========================================================="

# Recolectar archivos de forma segura (soporta espacios)
mapfile -d '' FILES < <(find "$INPUT_DIR" -maxdepth 1 -type f -name '*.md' -print0 | sort -z)

if [[ ${#FILES[@]} -eq 0 ]]; then
  echo "No se encontraron .md en $INPUT_DIR"
  exit 0
fi

# --- Bucle principal ---
for FILE in "${FILES[@]}"; do
  BASENAME="$(basename "$FILE")"
  STEM="${BASENAME%.*}"
  echo -e "\n>>> Procesando: $BASENAME ..."

  python "$SCRIPT" \
    --file "$FILE" \
    --k "$K" \
    --max_k "$MAX_K" \
    --model "$MODEL" \
    --seed "$SEED" \
    --save_dir "$OUTPUT_DIR" \
    --distance_metric "$METRIC"

  echo ">>> ✅ Completado: $STEM"
done

echo -e "\n🚀 Listo. Archivos madre Excel generados en: $OUTPUT_DIR/"
echo " (por archivo: _archivo_madre_etiquetado.xlsx)"

# --- (Opcional) Ejecución en paralelo con 4 procesos ---
# find "$INPUT_DIR" -maxdepth 1 -type f -name '*.txt' -print0 \
# | xargs -0 -n 1 -P 4 -I {} \
#   python "$SCRIPT" --file "{}" --k "$K" --max_k "$MAX_K" --model "$MODEL" \
#                    --seed "$SEED" --save_dir "$OUTPUT_DIR" --distance_metric "$METRIC"
