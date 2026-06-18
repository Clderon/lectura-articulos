#!/usr/bin/env python3
"""Debug script para MinerU"""
import sys
import traceback

try:
    print("[1] Iniciando...", flush=True)

    print("[2] Importando transformers...", flush=True)
    from transformers import AutoProcessor, Qwen2VLForConditionalGeneration
    print("[2] OK Transformers OK", flush=True)

    print("[3] Importando MinerU...", flush=True)
    from mineru_vl_utils import MinerUClient
    print("[3] OK MinerU OK", flush=True)

    print("[4] Cargando modelo (intentando CPU first)...", flush=True)
    try:
        model = Qwen2VLForConditionalGeneration.from_pretrained(
            "opendatalab/MinerU2.5-2509-1.2B",
            dtype="float16",
            device_map="cpu"
        )
        print("[4] OK Modelo cargado en CPU", flush=True)
    except Exception as e:
        print(f"[4] CPU fallo: {e}, intentando auto...", flush=True)
        model = Qwen2VLForConditionalGeneration.from_pretrained(
            "opendatalab/MinerU2.5-2509-1.2B",
            dtype="auto",
            device_map="auto"
        )
        print("[4] OK Modelo cargado auto", flush=True)

    print("[5] Cargando processor...", flush=True)
    processor = AutoProcessor.from_pretrained(
        "opendatalab/MinerU2.5-2509-1.2B",
        use_fast=True
    )
    print("[5] OK Processor OK", flush=True)

    print("[6] Creando cliente...", flush=True)
    client = MinerUClient(backend="transformers", model=model, processor=processor)
    print("[6] OK Cliente OK", flush=True)

    print("\nOK TODOS LOS PASOS EXITOSOS")

except Exception as e:
    print(f"\nERROR: {e}", flush=True)
    traceback.print_exc()
    sys.exit(1)
