# Local - MinerU VL + DeepSeek + Excel

Procesamiento de PDFs usando MinerU Vision-Language Model en tu máquina con GPU RTX 4050.

## 📋 Flujo de Procesamiento

```
1️⃣ coordinator.py
   └─ Verifica qué artículos faltan
   └─ Genera work_list.json

2️⃣ mineru_extractor.py
   └─ Lee work_list.json
   └─ Procesa PDFs con MinerU VL
   └─ Genera carpetas con page_*.md
   └─ Crea .DONE flag

3️⃣ fragmenter.py
   └─ Lee carpetas con .DONE
   └─ Divide .md en chunks por capítulos
   └─ Guarda en carpeta "chunks"
   └─ Crea .FRAGMENTS_DONE flag

4️⃣ deepseek_analyzer.py
   └─ Lee carpetas con .FRAGMENTS_DONE
   └─ Envía chunks a DeepSeek
   └─ Consolida análisis
   └─ Guarda en Excel
```

## 🚀 Ejecución

### Paso 1: Verificar qué procesar
```bash
python coordinator.py
```
**Output:** `work_list.json` con PDFs a procesar

### Paso 2: Extraer PDFs → .md
```bash
python mineru_extractor.py
```
⏱ **Tiempo:** ~60s por página (ejemplo: PDF de 10 páginas = 10 minutos)

**Output:** Carpetas en `output/`
```
output/
├── articulo_1/
│   ├── page_01.md
│   ├── page_02.md
│   └── .DONE
├── articulo_2/
│   └── ...
```

### Paso 3: Fragmentar en chunks
```bash
python fragmenter.py
```
⏱ **Tiempo:** Rápido (segundos)

**Output:** Carpetas `chunks/` dentro de cada artículo
```
output/articulo_1/
├── chunks/
│   ├── Inicio.md
│   ├── 1._Introduction.md
│   ├── 2._Methodology.md
│   ├── 2._Methodology_Pt1.md
│   ├── 2._Methodology_Pt2.md
│   └── ...
└── .FRAGMENTS_DONE
```

### Paso 4: Analizar con DeepSeek
```bash
python deepseek_analyzer.py
```
⏱ **Tiempo:** ~30s por chunk (varia con tamaño)

**Output:** Excel actualizado con resultados

---

## 📁 Estructura de Archivos

```
local/
├── .env                      # Configuración (PDF_FOLDER, EXCEL_PATH, etc)
├── procesar_articulos.py     # Script original (no usar, solo referencia)
├── shared_utils.py           # Funciones compartidas
├── coordinator.py            # Verifica artículos
├── mineru_extractor.py       # Extrae PDF → .md
├── fragmenter.py             # Divide .md en chunks
├── deepseek_analyzer.py      # Analiza con DeepSeek
├── output/                   # Resultados (carpetas por artículo)
└── work_list.json            # Lista de trabajo (generado)
```

---

## ⚙️ Configuración (.env)

Edita `local/.env`:
```env
DEEPSEEK_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
PDF_FOLDER=C:/Analisis-Prueba
EXCEL_PATH=C:/Cursos/Diseño de Investigacion I/Revición de Artículos.xlsx
```

---

## 📊 Tiempos Estimados

Para **37 PDFs × 10 páginas = 370 páginas**:

| Fase | Tiempo | Paralelizable |
|---|---|---|
| Coordinador | ~10s | N/A |
| MinerU (370 págs) | **~370 min (6+ h)** | No |
| Fragmenter | ~30s | N/A |
| DeepSeek (chunks) | ~20-30 min | No (secuencial por artículo) |
| **Total** | **~6-7 horas** | - |

---

## 🔄 Reintentos y Recuperación

Si algo falla:

1. **coordinator.py falla** → No hay cambios, reintentar
2. **mineru_extractor.py falla** → Busca `.DONE` flag, reinicia solo sin procesado
3. **fragmenter.py falla** → Busca `.FRAGMENTS_DONE`, reinicia
4. **deepseek_analyzer.py falla** → Reintentos automáticos, busca `.ANALYSIS_DONE`

---

## ⚡ Optimizaciones

- **GPU:** Usa RTX 4050 automáticamente (device_map="auto")
- **Modelo:** MinerU 1.2B (mejor calidad, más lento)
- **Chunks:** Max 4000 palabras cada uno (ajustable en fragmenter.py)
- **Reintentos:** DeepSeek reintentos 3 veces con backoff

---

## 🐛 Troubleshooting

**"GPU not found"**
→ Revisa CUDA: `python -c "import torch; print(torch.cuda.is_available())"`

**"DeepSeek timeout"**
→ Aumenta retries en deepseek_analyzer.py

**"Excel locked"**
→ Cierra el Excel antes de ejecutar analyzer

**"No PDFs found"**
→ Verifica PDF_FOLDER en .env

---

## 📝 Notas

- Los scripts son **secuenciales** - ejecuta en orden
- Cada paso genera **flags** (.DONE, .FRAGMENTS_DONE, .ANALYSIS_DONE)
- Los resultados se almacenan en **output/** localmente
- El Excel se actualiza al final del análisis
- Los chunks .md son **reutilizables** (no necesita reprocesar PDF)

---

**Actualizado:** 2026-06-02
