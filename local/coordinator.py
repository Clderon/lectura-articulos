#!/usr/bin/env python3
"""
Coordinador - Determina qué artículos procesar
"""
import sys
import os
from pathlib import Path
from dotenv import load_dotenv
import json

# UTF-8 en Windows
if sys.platform == "win32":
    os.environ["PYTHONIOENCODING"] = "utf-8"
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

load_dotenv()

PDF_FOLDER = os.getenv("PDF_FOLDER")
EXCEL_PATH = os.getenv("EXCEL_PATH")
OUTPUT_DIR = Path(__file__).parent / "output"

from shared_utils import load_analyzed_articles


def get_pdf_files() -> list:
    """Obtiene lista de PDFs en la carpeta."""
    pdf_folder = Path(PDF_FOLDER)
    if not pdf_folder.exists():
        print(f"[ERROR] Carpeta no existe: {PDF_FOLDER}")
        sys.exit(1)

    pdfs = sorted(pdf_folder.glob("*.pdf"))
    return pdfs


def get_processed_folders() -> set:
    """Obtiene carpetas ya procesadas en output/."""
    if not OUTPUT_DIR.exists():
        OUTPUT_DIR.mkdir(exist_ok=True)
        return set()

    return {d.name for d in OUTPUT_DIR.iterdir() if d.is_dir()}


def normalize_title(title: str) -> str:
    """Normaliza título para comparación."""
    return str(title).lower().strip()[:50]


def coordinator_main():
    """Coordina verificación y genera lista de trabajo (respeta fórmulas del Excel)."""
    print("=" * 70)
    print("COORDINADOR - Verificación de artículos")
    print("=" * 70)

    # 1. Cargar artículos analizados en Excel
    print("\n[1/3] Leyendo artículos analizados en Excel...")
    analyzed_in_excel = load_analyzed_articles(EXCEL_PATH)
    print(f"  ✓ {len(analyzed_in_excel)} artículos encontrados en Excel\n")

    if analyzed_in_excel:
        print("  " + "─" * 66)
        print("  ARTÍCULOS EN EXCEL:")
        print("  " + "─" * 66)
        for num, title in sorted(analyzed_in_excel.values(), key=lambda x: x[0]):
            print(f"  {num:2d}. {title[:60]}")
        print("  " + "─" * 66)

    # 2. Obtener PDFs disponibles
    print("\n[2/3] Escaneando PDFs disponibles...")
    pdf_files = get_pdf_files()
    print(f"  ✓ {len(pdf_files)} PDFs encontrados\n")

    print("  " + "─" * 66)
    print("  PDFs EN CARPETA:")
    print("  " + "─" * 66)
    for i, pdf_path in enumerate(pdf_files, 1):
        print(f"  {i:2d}. {pdf_path.name}")
    print("  " + "─" * 66)

    # 3. Obtener carpetas procesadas
    print("\n[3/3] Verificando carpetas procesadas...")
    processed_folders = get_processed_folders()
    print(f"  ✓ {len(processed_folders)} carpetas con .md")

    if processed_folders:
        print("\n  " + "─" * 66)
        print("  CARPETAS PROCESADAS:")
        print("  " + "─" * 66)
        for folder in sorted(processed_folders):
            print(f"  ✓ {folder}/")
        print("  " + "─" * 66)

    # 4. Determinar qué procesar
    print("\n" + "=" * 70)
    print("ANÁLISIS DETALLADO")
    print("=" * 70)

    work_list = []
    already_processed = []
    in_excel_only = []
    pending = []

    for pdf_path in pdf_files:
        pdf_name = pdf_path.stem
        folder_name = pdf_name

        # ¿Ya está procesado (carpeta existe)?
        if folder_name in processed_folders:
            already_processed.append(pdf_name)
            continue

        # ¿Ya está en Excel?
        title_normalized = normalize_title(pdf_name)
        in_excel = any(normalize_title(t[1]) == title_normalized for t in analyzed_in_excel.values())

        if in_excel:
            in_excel_only.append(pdf_name)
            continue

        # Necesita procesamiento
        pending.append(pdf_name)
        work_list.append(str(pdf_path))

    # Mostrar resumen
    print(f"\n✓ YA PROCESADOS (con carpeta .md): {len(already_processed)}")
    if already_processed:
        print("  " + "─" * 66)
        for name in already_processed:
            print(f"    ✓ {name}")
        print("  " + "─" * 66)

    print(f"\n✓ EN EXCEL (sin carpeta .md): {len(in_excel_only)}")
    if in_excel_only:
        print("  " + "─" * 66)
        for name in in_excel_only:
            print(f"    • {name}")
        print("  " + "─" * 66)

    print(f"\n⚠ PENDIENTES (necesitan procesar): {len(pending)}")
    if pending:
        print("  " + "─" * 66)
        for name in pending:
            print(f"    ○ {name}")
        print("  " + "─" * 66)

    # 5. Preguntar al usuario cuántos procesar
    print("\n" + "=" * 70)
    print("SELECCIÓN INTERACTIVA")
    print("=" * 70)

    if not pending:
        print("\n✓ Todos los artículos ya están procesados o en Excel.")
        work_file = Path(__file__).parent / "work_list.json"
        with open(work_file, "w", encoding="utf-8") as f:
            json.dump({"todo": [], "count": 0}, f, indent=2, ensure_ascii=False)
        print("=" * 70)
        return []

    print(f"\nHay {len(pending)} artículos pendientes de procesar.")
    print(f"(Máximo a procesar: {len(pending)})\n")

    while True:
        try:
            cantidad = int(input("¿Cuántos artículos deseas procesar? "))
            if cantidad <= 0:
                print("❌ Debe ser un número mayor a 0")
                continue
            if cantidad > len(pending):
                print(f"❌ Solo hay {len(pending)} pendientes")
                continue
            break
        except ValueError:
            print("❌ Ingresa un número válido")

    # Seleccionar los primeros N artículos pendientes
    selected_work = work_list[:cantidad]
    selected_names = pending[:cantidad]

    print(f"\n✓ Seleccionados {cantidad} artículos:")
    print("  " + "─" * 66)
    for name in selected_names:
        print(f"    ○ {name}")
    print("  " + "─" * 66)

    # 6. Guardar lista de trabajo
    print("\n" + "=" * 70)
    print("LISTA DE TRABAJO GENERADA")
    print("=" * 70)

    work_file = Path(__file__).parent / "work_list.json"
    with open(work_file, "w", encoding="utf-8") as f:
        json.dump({"todo": selected_work, "count": cantidad}, f, indent=2, ensure_ascii=False)

    print(f"\n✓ Lista guardada: {work_file}")
    print(f"✓ Próximo paso: python mineru_extractor.py")
    print("=" * 70)

    return selected_work


if __name__ == "__main__":
    work_list = coordinator_main()
    sys.exit(0 if work_list else 1)
