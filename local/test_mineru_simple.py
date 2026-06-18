#!/usr/bin/env python3
"""
Test simple de MinerU VL - sin Excel, solo extracción
"""
import sys
import os
from pathlib import Path

# UTF-8 en Windows
if sys.platform == "win32":
    os.environ["PYTHONIOENCODING"] = "utf-8"
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

print("=" * 70)
print("TEST MINERU VL - Extracción simple")
print("=" * 70)

# 1. Verificar GPU
print("\n[1/5] Verificando GPU...")
try:
    import torch
    gpu_available = torch.cuda.is_available()
    if gpu_available:
        gpu_name = torch.cuda.get_device_name(0)
        print(f"  ✓ GPU encontrada: {gpu_name}")
    else:
        print(f"  ✗ GPU NO disponible (usando CPU - MUY LENTO)")
except Exception as e:
    print(f"  ✗ Error: {e}")
    sys.exit(1)

# 2. Cargar modelo
print("\n[2/5] Cargando modelo MinerU VL...")
try:
    from transformers import AutoProcessor, Qwen2VLForConditionalGeneration
    from mineru_vl_utils import MinerUClient
    from PIL import Image
    import fitz
    import io

    print("  Descargando modelo (1.2B parámetros, ~2.5GB)...")
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
    print("  ✓ Modelo cargado")
except Exception as e:
    print(f"  ✗ Error: {e}")
    sys.exit(1)

# 3. Seleccionar PDF
print("\n[3/5] Seleccionando PDF...")
pdf_folder = Path("C:/Analisis-Prueba")
pdf_files = list(pdf_folder.glob("*.pdf"))

if not pdf_files:
    print(f"  ✗ No hay PDFs en {pdf_folder}")
    sys.exit(1)

pdf_path = pdf_files[0]
print(f"  ✓ Usando: {pdf_path.name}")

# 4. Convertir PDF a imágenes y procesar
print("\n[4/5] Procesando PDF...")
try:
    pdf_doc = fitz.open(pdf_path)
    num_pages = len(pdf_doc)
    print(f"  Total de páginas: {num_pages}")

    output_dir = Path("mineru_test_output")
    output_dir.mkdir(exist_ok=True)

    import time
    start_time = time.time()

    for page_num in range(1):  # Solo primera página
        print(f"\n  Procesando página {page_num + 1}/1...")

        # Convertir página a imagen
        pix = pdf_doc[page_num].get_pixmap(matrix=fitz.Matrix(2, 2))
        img_data = pix.tobytes("png")
        image = Image.open(io.BytesIO(img_data))

        # Procesar con MinerU
        print(f"    Analizando con MinerU VL...")
        blocks = client.two_step_extract(image)

        print(f"    ✓ {len(blocks)} bloques detectados")

        # Guardar resultados
        output_file = output_dir / f"page_{page_num + 1}.md"
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(f"# Página {page_num + 1}\n\n")

            for i, block in enumerate(blocks):
                f.write(f"## Bloque {i + 1} ({block.type})\n")
                if block.content:
                    f.write(f"{block.content}\n\n")
                else:
                    f.write(f"[Sin contenido de texto]\n\n")

        print(f"    → Guardado: {output_file}")

    pdf_doc.close()
    elapsed = time.time() - start_time
    print(f"\n  ✓ Archivos guardados en: {output_dir.absolute()}")
    print(f"  ⏱ Tiempo total: {elapsed:.2f}s")

except Exception as e:
    print(f"  ✗ Error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# 5. Resumen
print("\n[5/5] Resumen")
print("=" * 70)
print("✓ MinerU VL funcionando correctamente")
print(f"✓ Archivos guardados en: mineru_test_output/")
print("✓ Próximo paso: integrar con DeepSeek + Excel")
print("=" * 70)
