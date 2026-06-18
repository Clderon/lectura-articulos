#!/usr/bin/env python3
"""
MinerU Extractor - Procesa PDF → carpeta con .md (una por página)
"""
import sys
import os
import json
import time
from pathlib import Path
from dotenv import load_dotenv
import fitz
import io

# UTF-8 en Windows
if sys.platform == "win32":
    os.environ["PYTHONIOENCODING"] = "utf-8"
    import io as iomod
    sys.stdout = iomod.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

load_dotenv()

OUTPUT_DIR = Path(__file__).parent / "output"
WORK_FILE = Path(__file__).parent / "work_list.json"


def extract_pdf_to_md(pdf_path: str, output_folder: Path) -> bool:
    """Convierte PDF a .md usando MinerU VL."""
    try:
        from transformers import AutoProcessor, Qwen2VLForConditionalGeneration
        from mineru_vl_utils import MinerUClient
        from PIL import Image

        pdf_path = Path(pdf_path)
        print(f"\n  Procesando: {pdf_path.name}")

        # Crear carpeta de salida
        output_folder.mkdir(parents=True, exist_ok=True)

        # Abrir PDF
        pdf_doc = fitz.open(pdf_path)
        num_pages = len(pdf_doc)
        print(f"  Total páginas: {num_pages}")

        # Cargar modelo (una vez por PDF, pero libera memoria después)
        print(f"  Cargando modelo MinerU VL...")
        model = Qwen2VLForConditionalGeneration.from_pretrained(
            "opendatalab/MinerU2.5-2509-1.2B",
            dtype="auto",
            device_map="auto"
        )
        processor = AutoProcessor.from_pretrained(
            "opendatalab/MinerU2.5-2509-1.2B",
            use_fast=True
        )
        client = MinerUClient(backend="transformers", model=model, processor=processor)

        # Procesar cada página
        start_time = time.time()
        for page_num in range(num_pages):
            # Convertir página a imagen
            pix = pdf_doc[page_num].get_pixmap(matrix=fitz.Matrix(2, 2))
            img_data = pix.tobytes("png")
            image = Image.open(io.BytesIO(img_data))

            # Procesar con MinerU
            print(f"    Página {page_num + 1}/{num_pages}...", end="", flush=True)
            blocks = client.two_step_extract(image)
            print(f" OK ({len(blocks)} bloques)")

            # Guardar a .md
            output_file = output_folder / f"page_{page_num + 1:02d}.md"
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(f"# Página {page_num + 1}\n\n")
                for block in blocks:
                    if block.content:
                        if block.type != "text":
                            f.write(f"**[{block.type}]** {block.content}\n\n")
                        else:
                            f.write(f"{block.content}\n\n")

        pdf_doc.close()
        elapsed = time.time() - start_time

        print(f"  OK Completado en {elapsed:.2f}s")
        print(f"  OK Archivos guardados en: {output_folder}")

        return True

    except Exception as e:
        print(f"  ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False


def extractor_main():
    """Procesa artículos de la lista de trabajo."""
    print("=" * 70)
    print("MINERU EXTRACTOR - PDF → .md")
    print("=" * 70)

    # Cargar lista de trabajo
    if not WORK_FILE.exists():
        print(f"[ERROR] No hay lista de trabajo. Ejecuta coordinator.py primero")
        sys.exit(1)

    with open(WORK_FILE, "r", encoding="utf-8") as f:
        work_data = json.load(f)

    work_list = work_data.get("todo", [])
    if not work_list:
        print("[INFO] No hay artículos para procesar")
        sys.exit(0)

    print(f"\nProcesando {len(work_list)} artículos...\n")

    # Procesar cada PDF
    for i, pdf_path in enumerate(work_list, 1):
        pdf_path = Path(pdf_path)
        pdf_name = pdf_path.stem

        print(f"\n[{i}/{len(work_list)}] {pdf_name}")

        # Crear carpeta de salida
        output_folder = OUTPUT_DIR / pdf_name

        # Procesar
        success = extract_pdf_to_md(pdf_path, output_folder)

        if not success:
            print(f"  WARNING Saltado debido a error")
            continue

        # Crear archivo de confirmación
        done_file = output_folder / ".DONE"
        done_file.touch()

        print(f"  Ready Listo para fragmentar")

    print("\n" + "=" * 70)
    print("OK Extraccion completada")
    print("OK Proximo paso: ejecutar fragmenter.py")
    print("=" * 70)


if __name__ == "__main__":
    extractor_main()
