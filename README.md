# Lectura Selectiva - Automatización de Revisión Sistemática

Sistema de automatización para procesar artículos académicos en PDF y llenar Excel usando DeepSeek como modelo de análisis.

## Descripción

Script Python que:
- Extrae texto de PDFs (maneja 2 columnas)
- Detecta secciones automáticamente (I., A., 1., etc.)
- Envía secciones a DeepSeek API para análisis
- Procesa 2 artículos por conversación (128K tokens)
- Detecta automáticamente duplicados
- Asigna números secuenciales a artículos nuevos
- Guarda resultados en Excel (preserva formato)

## Instalación

```bash
pip install -r requirements.txt
```

## Configuración

### 1. Crear archivo `.env`

```bash
# Copia el template
cp .env.example .env
```

### 2. Editar `.env` con tus valores

```env
DEEPSEEK_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
PDF_FOLDER=/ruta/a/los/pdfs
EXCEL_PATH=/ruta/al/excel.xlsx
MODO_VERIFICACION=0
MAX_WORDS_PER_CHUNK=6000
```

### Variables disponibles

| Variable | Obligatoria | Default | Descripción |
|----------|:-----------:|---------|-------------|
| `DEEPSEEK_API_KEY` | ✅ | - | API key de DeepSeek |
| `PDF_FOLDER` | ✅ | - | Carpeta con PDFs |
| `EXCEL_PATH` | ✅ | - | Ruta del Excel |
| `MODO_VERIFICACION` | ❌ | 0 | 0=directo, 1=preview |
| `MAX_WORDS_PER_CHUNK` | ❌ | 6000 | Palabras por chunk |
| `SHEET1_NAME` | ❌ | "Revisión de Artículos" | Nombre hoja 1 |
| `SHEET2_NAME` | ❌ | "Lectura Selectiva" | Nombre hoja 2 |

⚠️ **IMPORTANTE:** El archivo `.env` NO se sube a git (está en `.gitignore`)

## Uso

```bash
python procesar_articulos.py
```

El script te pedirá:
1. Cuántos artículos nuevos procesar (o Enter para todos)

Luego:
- Busca PDFs nuevos (omite duplicados automáticamente)
- Asigna números secuenciales
- Procesa en pares con DeepSeek
- Guarda en Excel inmediatamente

## Flujo de Procesamiento

```
Carpeta PDFs
    ↓
Extrae títulos (metadata, contenido, nombre)
    ↓
Compara contra Excel (detecta duplicados)
    ↓
PDFs nuevos → asigna N°17, 18, 19...
    ↓
Procesa en pares [17,18], [19,20]...
    ↓
Cada par = 1 conversación DeepSeek
    ↓
Extrae secciones → envía por chunks
    ↓
DeepSeek analiza → genera 2 tablas
    ↓
Escribe en Excel inmediatamente
```

## Archivos

- `procesar_articulos.py` - Script principal
- `requirements.txt` - Dependencias
- `reporte_procesamiento.txt` - Generado al finalizar

## Excel

### Hoja 1: "Revisión de Artículos"
Script llena: N°, Título, Autores, Pertinencia, Actualidad, Rigor, Fuente Calidad, Impacto, Ética, Puntaje, Decisión, Motivo Rechazo

### Hoja 2: "Lectura Selectiva"
Script llena: N°, Autor/Año, Objetivo, Método, Hallazgo, Limitación, Uso en tu tesis

Si decisión = RECHAZADO → N° en Hoja 2 aparece con fondo rojo

## Notas

- Los números asignados dependen de cuántos artículos ya hay en Excel
- Detecta duplicados por similitud de título
- Procesa máximo 2 por chat para mantener contexto
- Si pide 3 artículos: genera 2 chats (pares [17,18] y [19])
