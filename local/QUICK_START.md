# Quick Start - Pipeline Local

Guía rápida para procesar artículos. Para detalles completos, ve a `README_PRINCIPAL.md`.

---

## 🎯 Setup Inicial (Primera vez)

### 1. Configurar `.env`
```powershell
cd local
cp .env.example .env  # Si no existe
```

Edita `local/.env`:
```env
DEEPSEEK_API_KEY=sk-tu_key_aqui
EXCEL_PATH=C:\ruta\completa\Excel.xlsx
PDF_FOLDER=C:\ruta\completa\PDFs
```

### 2. Verificar Instalación
```powershell
python debug_mineru.py
```

Debería mostrar "OK" para todos los pasos. Si falla, revisa README_PRINCIPAL.md → Problemas Comunes.

---

## 🚀 Procesamiento Normal (Cada vez que quieras analizar artículos)

### Opción A: Automático (recomendado)
```powershell
python run_full_pipeline.py
```
Ejecuta los 4 pasos automáticamente.

### Opción B: Manual (si quieres control)

**Paso 1: Selecciona artículos**
```powershell
python coordinator.py
```
→ Crea `work_list.json` con PDFs sin procesar

**Paso 2: Extrae texto**
```powershell
python mineru_extractor.py
```
→ Crea `output/{nombre}/*.md` (páginas extraídas)
⏱ ~60s/página (si tienes GPU, más rápido)

**Paso 3: Agrupa secciones**
```powershell
python fragmenter.py
```
→ Crea `output/{nombre}/consolidated.md` (secciones unidas)

**Paso 4: Analiza con IA**
```powershell
python deepseek_analyzer.py
```
→ Llena Excel con análisis

---

## 🔍 Verificar Progreso

### ¿Qué PDFs quedan por procesar?
```powershell
python coordinator.py
# Muestra qué hay en work_list.json
```

### ¿Dónde están los archivos?
```
output/
├── {nombre_articulo_1}/
│   ├── page_01.md, page_02.md, ... (extracción)
│   ├── consolidated.md             (secciones agrupadas)
│   ├── .DONE                        (marcador: PDF procesado)
│   └── .FRAGMENTS_DONE              (marcador: secciones agrupadas)
├── {nombre_articulo_2}/
└── ...
```

### ¿El Excel se actualizó?
Abre Excel → "Revisión de Artículos" → verifica filas 4+

---

## ⚠️ Problemas Rápidos

| Problema | Solución |
|----------|----------|
| `ModuleNotFoundError: mineru_vl_utils` | `pip install mineru` |
| `CUDA out of memory` | Usa CPU (código ya lo hace automáticamente) |
| Datos en fila equivocada | Verifica Excel encabezado en fila 3 |
| `DeepSeek 429 error` | Espera 60s, ejecuta de nuevo |
| El Excel se "repara" | Haz clic "Sí", sigue usando |

Para más detalles → README_PRINCIPAL.md

---

## 📊 Qué Esperar

| Paso | Tiempo | Archivo |
|------|--------|---------|
| Coordinator | <1s | `work_list.json` |
| MinerU (4 PDFs) | ~8 min | `output/*/page_XX.md` |
| Fragmenter (4 PDFs) | ~2s | `output/*/consolidated.md` |
| DeepSeek (3 artículos) | ~2-3 min | Excel actualizado |

---

## 💡 Consejos

1. **Primeros artículos:** Revisa el Excel manualmente para verificar que los análisis sean correctos
2. **Muchos artículos:** Ejecuta `deepseek_analyzer.py` múltiples veces (procesa máximo 3/ejecución)
3. **Sin GPU:** Los pasos 1 y 3 serán más lentos pero funcionan
4. **Cambió el Excel:** Actualiza `EXCEL_PATH` en `.env` y reinicia

---

**Última actualización:** 2026-06-17
