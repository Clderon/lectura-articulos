# Lectura Selectiva - Procesamiento de Artículos PDF → Excel

Sistema automatizado para extraer, fragmentar y analizar artículos académicos sobre Infrastructure as Code usando MinerU VL + DeepSeek.

**Status:** 🟢 Producción (MinerU local + DeepSeek activo)

---

## 📋 Tabla de Contenidos

1. [Instalación Rápida](#instalación-rápida)
2. [Configuración](#configuración)
3. [Pipeline Completo](#pipeline-completo)
4. [Plantilla Excel](#plantilla-excel)
5. [Problemas Comunes](#problemas-comunes)
6. [Estructura del Proyecto](#estructura-del-proyecto)

---

## 🚀 Instalación Rápida

### Requisitos Previos
- **Python 3.11+**
- **Git** (para clonar el repo)
- **CUDA 11.8+** (solo si usas GPU local)
- **DeepSeek API Key** (gratis en https://platform.deepseek.com/)

### Paso 1: Clonar y Setup
```powershell
git clone <repo_url>
cd Lectura-Selectiva

# Crear entorno virtual
python -m venv venv
.\venv\Scripts\Activate.ps1

# Instalar dependencias
pip install -r requirements.txt
```

### Paso 2: Obtener Credenciales
1. **DeepSeek API Key**
   - Ve a https://platform.deepseek.com/
   - Crea cuenta → API Keys
   - Copia la key
   - Agrega a `local/.env` como `DEEPSEEK_API_KEY=sk-...`

### Paso 3: Crear Archivo .env en `local/`
```bash
cp local/.env.example local/.env  # (si existe)
```

Edita `local/.env` con:
```env
DEEPSEEK_API_KEY=sk-xxxxxxxxxxxxx
EXCEL_PATH=C:\ruta\a\tu\Revisión de Artículos - V3.xlsx
PDF_FOLDER=C:\ruta\a\tus\PDFs
```

### Paso 4: Verificar Instalación
```powershell
cd local
python debug_mineru.py
```

Debería mostrar "OK" para todos los pasos.

---

## ⚙️ Configuración Detallada

### Variables de Entorno (`.env`)

| Variable | Ejemplo | Descripción |
|----------|---------|-------------|
| `DEEPSEEK_API_KEY` | `sk-...` | Token de DeepSeek API |
| `EXCEL_PATH` | `C:\...\Revisión de Artículos - V3.xlsx` | Ruta absoluta al Excel |
| `PDF_FOLDER` | `C:\...\PDFs` | Carpeta con PDFs a procesar |

### Estructura del Excel

El sistema espera un Excel con **2 hojas**:

#### 📊 Hoja 1: "Revisión de Artículos"

```
┌────┬───────────────┬─────────┬──────┬────────────┬───────────┬────────┬──────────┬────────┬────────┬────────┬────────────┬──────────┐
│ N° │ Artículo      │ Autor   │ Año  │ Pertinencia│ Actualidad│ Rigor  │ Calidad  │ Impacto│ Ética  │Puntaje │ Veredicto  │Justif.   │
│ A  │ B             │ C       │ D    │ E          │ F         │ G      │ H        │ I      │ J      │ K      │ L          │ M        │
├────┼───────────────┼─────────┼──────┼────────────┼───────────┼────────┼──────────┼────────┼────────┼────────┼────────────┼──────────┤
│ 1  │               │         │      │            │           │        │          │        │        │        │            │          │
│ 2  │               │         │      │            │           │        │          │        │        │        │            │          │
│ 3  │               │         │      │            │           │        │          │        │        │        │            │          │
└────┴───────────────┴─────────┴──────┴────────────┴───────────┴────────┴──────────┴────────┴────────┴────────┴────────────┴──────────┘

Fórmulas:
- Columna A (N°): =SI.ERROR(INDICE(TablaArticulos[N°];FILA()-3);"")
- Columna K (Puntaje): =SUMA(E:J)
```

**⚠️ IMPORTANTE:** Las columnas A y K tienen fórmulas. No las sobrescribas.

#### 📖 Hoja 2: "Lectura Selectiva"

```
┌────┬────┬──────────┬────────┬────────┬─────────────┬──────────────┬─────────────┬────────────┬────────────┬─────────┬───────────┬──────────┬───────────┬─────────────┐
│ N° │    │ Objetivo │ Método │Hallazgo│ Limitación  │ Ref. textual │ Paradigma   │ Herramienta│ Tipo estud.│Variables│Dimensiones│Indicadores│Trab.futuro │¿Generaliz.? │
│ A  │ B  │ C        │ D      │ E      │ F           │ G            │ H           │ I          │ J          │ K       │ L         │ M         │ N          │ O           │
├────┼────┼──────────┼────────┼────────┼─────────────┼──────────────┼─────────────┼────────────┼────────────┼─────────┼───────────┼──────────┼───────────┼─────────────┤
│ 1  │    │          │        │        │             │              │             │            │            │         │           │          │           │             │
│ 2  │    │          │        │        │             │              │             │            │            │         │           │          │           │             │
│ 3  │    │          │        │        │             │              │             │            │            │         │           │          │           │             │
└────┴────┴──────────┴────────┴────────┴─────────────┴──────────────┴─────────────┴────────────┴────────────┴─────────┴───────────┴──────────┴───────────┴─────────────┘
```

**Formato recomendado:**
- Columnas C-P: Ajuste de texto (Word Wrap) + altura de fila 60
- Todas las celdas: Alineación superior izquierda

---

## 🔄 Pipeline Completo

El sistema está compuesto por 4 scripts que se ejecutan **secuencialmente**:

### 1️⃣ Coordinator (Selecciona artículos)
```powershell
cd local
python coordinator.py
```
- Lee la carpeta `PDF_FOLDER`
- Identifica PDFs sin procesar
- Crea `work_list.json` con lista de tareas

**Output:** `work_list.json`

---

### 2️⃣ MinerU Extractor (PDF → .md)
```powershell
python mineru_extractor.py
```
- Lee PDFs de `work_list.json`
- Extrae texto + estructura con MinerU VL
- Crea carpeta `output/{nombre_pdf}/` con archivos `page_XX.md`
- Marca con `.DONE` cuando termina

**Input:** `work_list.json`  
**Output:** `output/{nombre}/*.md`

**Tiempo:** ~60s por página (si tienes GPU disponible)

---

### 3️⃣ Fragmenter (Agrupa secciones)
```powershell
python fragmenter.py
```
- Lee archivos `page_XX.md` de cada artículo
- Detecta secciones (Abstract, Introduction, Methodology, Results, etc.)
- Crea `consolidated.md` con formato:
  ```
  ### [SECCIÓN 1/5 - PRINCIPAL]
  **Título:** Introduction
  **Contenido:** ...
  ---
  (contenido después de separador)
  ```
- Marca con `.FRAGMENTS_DONE`

**Input:** `output/{nombre}/*.md`  
**Output:** `output/{nombre}/consolidated.md`

---

### 4️⃣ DeepSeek Analyzer (Análisis IA)
```powershell
python deepseek_analyzer.py
```
- Lee Excel → busca artículos sin analizar
- Envía secciones de `consolidated.md` a DeepSeek **una por una**
- DeepSeek responde "Sección recibida" hasta recibir "CERRAR ARTÍCULO"
- Parsea 2 tablas Markdown de la respuesta
- Escribe en Excel (Revisión de Artículos + Lectura Selectiva)

**Límite:** Procesa máximo 3 artículos por ejecución (configurable en código)

**Input:** Excel + `consolidated.md`  
**Output:** Excel actualizado

---

## 🎯 Flujo de Uso Completo

### Primera vez:
```powershell
cd local

# 1. Selecciona artículos
python coordinator.py

# 2. Extrae PDFs
python mineru_extractor.py

# 3. Fragmenta secciones
python fragmenter.py

# 4. Analiza con IA
python deepseek_analyzer.py
```

### Después (nuevos artículos):
```powershell
# Solo repite los 4 pasos
# El sistema salta automáticamente lo que ya está hecho
python coordinator.py
python mineru_extractor.py
python fragmenter.py
python deepseek_analyzer.py
```

---

## 📊 Plantilla Excel (Descargar)

**Estructura mínima necesaria:**

### "Revisión de Artículos" (Encabezado en fila 3)
```
Fila 3: N° | Artículo (Título) | Autor(es) | Año | Pertinencia | Actualidad | Rigor | Fuente (Calidad) | Impacto | Ética | Puntaje | Veredicto | Justificación
Fila 4+: (filas para datos)
```

### "Lectura Selectiva" (Encabezado en fila 3)
```
Fila 3: N° | (vacío) | Objetivo | Método | Hallazgo | Limitación | Ref. textual | Paradigma IaC | Herramienta IaC | Tipo de estudio | Variables | Dimensiones | Indicadores | Trabajo futuro | ¿Generalizable?
Fila 4+: (filas para datos)
```

---

## 🔍 Problemas Comunes y Soluciones

### ❌ "ModuleNotFoundError: No module named 'mineru_vl_utils'"

**Causa:** MinerU no está instalado correctamente.

**Solución:**
```powershell
pip install mineru
python debug_mineru.py  # Verifica instalación
```

Si sigue fallando:
```powershell
pip uninstall mineru torch transformers -y
pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu118
pip install mineru
```

---

### ❌ "CUDA out of memory" / "Exit code -1073741819"

**Causa:** GPU sin suficiente VRAM.

**Solución:** El código ya está configurado para usar CPU con `device_map="cpu"` y `dtype="float16"`:

```python
# En mineru_extractor.py (línea ~46)
model = Qwen2VLForConditionalGeneration.from_pretrained(
    "opendatalab/MinerU2.5-2509-1.2B",
    dtype="float16",          # ← Precisión media (consume menos)
    device_map="cpu"          # ← Fuerza CPU (pero es lento)
)
```

Si aún tienes problemas:
```powershell
# Reducir tamaño de batch en código (no hay flag, edita mineru_extractor.py línea 67)
```

---

### ❌ "Excel row calculation wrong" / Datos en fila equivocada

**Causa:** El script no detectó el encabezado correctamente.

**Solución:** Verifica que el Excel tenga estructura:
```
Fila 1: (vacía o encabezado genérico)
Fila 2: (vacía)
Fila 3: N° | Artículo | Autor | ...  ← AQUÍ DEBE ESTAR EL ENCABEZADO
Fila 4+: (datos)
```

Si tu Excel está en fila diferente, edita `deepseek_analyzer.py` línea ~454:
```python
# Cambiar este rango si tu encabezado está en otra fila
for row_idx in range(1, 10):
```

---

### ❌ "DeepSeek API error: 429 / Rate limit"

**Causa:** Demasiadas solicitudes en poco tiempo.

**Solución:** 
- Espera 60 segundos
- Ejecuta de nuevo (el script retoma desde donde paró)
- Reduce `MAX_ARTICLES` en `deepseek_analyzer.py` línea 504 (actualmente es 3)

---

### ❌ "El Excel se "repara" al abrir"

**Causa:** openpyxl escribió en formato no completamente compatible.

**Solución:** Haz clic en "Sí" cuando Excel pregunta reparar. Luego:
```powershell
# El archivo se corrige automáticamente
# Guarda normalmente con Ctrl+S
```

---

### ⚠️ "Todos los artículos tienen puntaje 11 / veredicto Aprobado"

**Causa:** El prompt de DeepSeek es muy genérico → respuestas uniformes.

**Solución:** El código ya tiene criterios detallados (líneas 64-120 en `deepseek_analyzer.py`). Si sigue pasando:

1. Verifica que DeepSeek esté recibiendo contenido diferente:
   ```powershell
   # El script muestra [DEBUG] RESPUESTA FINAL en consola
   # Copia eso y verifica que sea análisis real, no copypaste
   ```

2. Si ves las mismas respuestas palabra por palabra → usa otro modelo o ajusta `TEMPERATURE` en `shared_utils.py` línea 19:
   ```python
   TEMPERATURE = 0.1  # Prueba 0.3 o 0.5 para más variabilidad
   ```

---

## 📁 Estructura del Proyecto

```
Lectura-Selectiva/
├── local/                          🔬 PIPELINE ACTIVO
│   ├── .env                        (Credenciales: DEEPSEEK_API_KEY, EXCEL_PATH, PDF_FOLDER)
│   ├── coordinator.py              (1. Selecciona PDFs)
│   ├── mineru_extractor.py         (2. Extrae con MinerU VL)
│   ├── fragmenter.py               (3. Agrupa secciones)
│   ├── deepseek_analyzer.py        (4. Analiza con DeepSeek)
│   ├── shared_utils.py             (Funciones compartidas)
│   ├── debug_mineru.py             (Verificar instalación)
│   ├── output/                     (Archivos generados)
│   │   ├── {nombre_articulo}/
│   │   │   ├── page_01.md
│   │   │   ├── page_02.md
│   │   │   ├── consolidated.md
│   │   │   ├── .DONE
│   │   │   └── .FRAGMENTS_DONE
│   │   └── ...
│   ├── work_list.json              (Artículos pendientes)
│   └── README.md                   (Notas internas)
│
├── pymupdf/                        ⚡ ALTERNATIVA RÁPIDA (no usado actualmente)
│   ├── .env
│   ├── procesar_articulos.py
│   └── output/
│
├── .env.example                    (Template para .env)
├── .gitignore                      (Ignora PDFs, output/, Excel, etc.)
├── requirements.txt                (Dependencias Python)
├── README_PRINCIPAL.md             (Este archivo)
└── .git/                           (Control de versiones)
```

---

## 🔧 Requisitos Técnicos Detallados

### Para usar "local" (MinerU):

**Requisitos:**
- Python 3.11+
- CUDA 11.8+ (opcional, usa CPU si no disponible pero es lento)
- 16GB RAM mínimo (8GB tight, 32GB recomendado)

**Instalación CUDA en Windows (si tienes GPU NVIDIA):**
1. Descarga CUDA Toolkit 11.8: https://developer.nvidia.com/cuda-11-8-0-download-archive
2. Ejecuta instalador
3. Verifica:
   ```powershell
   python -c "import torch; print(torch.cuda.is_available())"
   # Debe mostrar: True
   ```

**Si no tienes GPU:**
- El código usa CPU automáticamente (lento pero funciona)
- Cada página tardará ~60s en lugar de 1-2s

---

## 📈 Optimizaciones

### Reducir tiempo de procesamiento:

1. **Si tienes GPU RTX 40X0+:**
   - Cambiar en `mineru_extractor.py` línea 49:
   ```python
   device_map="auto"  # En lugar de "cpu"
   ```

2. **Aumentar artículos analizados por ejecución:**
   - Cambiar en `deepseek_analyzer.py` línea 504:
   ```python
   MAX_ARTICLES = 10  # (en lugar de 3)
   ```

3. **Paralelizar (avanzado):**
   - El código actual es secuencial (más seguro)
   - Se puede paralelizar con `concurrent.futures` si se necesita

---

## 📞 Soporte

**Si algo no funciona:**

1. Ejecuta `python debug_mineru.py` y copia el error completo
2. Revisa la sección [Problemas Comunes](#problemas-comunes)
3. Verifica que `.env` tenga todas las variables
4. Comprueba que los PDFs existen en `PDF_FOLDER`
5. Verifica que el Excel tiene las 2 hojas correctas

---

**Última actualización:** 2026-06-17  
**Versión:** 4.0 (Local MinerU + DeepSeek, Sin API, Sin Docker)
