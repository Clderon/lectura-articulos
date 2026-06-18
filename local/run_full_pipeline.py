#!/usr/bin/env python3
"""
Pipeline maestro - Ejecuta los 3 pasos en secuencia automáticamente
"""
import subprocess
import sys

def run_step(script_name: str) -> bool:
    """Ejecuta un paso y espera a que termine."""
    print(f"\n{'='*70}")
    print(f"INICIANDO: {script_name}")
    print(f"{'='*70}\n")

    try:
        result = subprocess.run([sys.executable, script_name], check=True)
        print(f"\n✓ {script_name} completado exitosamente\n")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n✗ {script_name} falló con código {e.returncode}\n")
        return False
    except Exception as e:
        print(f"\n✗ Error ejecutando {script_name}: {e}\n")
        return False

def main():
    print("\n" + "="*70)
    print("PIPELINE COMPLETO - Procesa artículos de trabajo_list.json")
    print("="*70)

    steps = [
        "mineru_extractor.py",
        "fragmenter.py",
        "deepseek_analyzer.py",
    ]

    for i, step in enumerate(steps, 1):
        print(f"\n[{i}/{len(steps)}] Ejecutando {step}...")

        if not run_step(step):
            print(f"⚠ Pipeline detenido en {step}")
            sys.exit(1)

    print("\n" + "="*70)
    print("✓ PIPELINE COMPLETADO EXITOSAMENTE")
    print("="*70)

if __name__ == "__main__":
    main()
