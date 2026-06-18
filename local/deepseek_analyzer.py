#!/usr/bin/env python3
"""
DeepSeek Analyzer - Lee chunks → analiza → guarda en Excel
"""
import sys
import os
import json
from pathlib import Path
from dotenv import load_dotenv
import re

# UTF-8 en Windows
if sys.platform == "win32":
    os.environ["PYTHONIOENCODING"] = "utf-8"
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

load_dotenv()

OUTPUT_DIR = Path(__file__).parent / "output"
WORK_FILE = Path(__file__).parent / "work_list.json"
CHUNKS_SUBDIR = "chunks"

from shared_utils import (
    send_to_deepseek,
    write_article_to_excel,
    create_deepseek_client,
    split_into_sections,
    load_analyzed_articles,
)


MASTER_PROMPT_TEMPLATE = """Eres un asistente de investigación especializado en revisiones sistemáticas.

CONTEXTO DEL ESTUDIO:
Título: "Estudio de mapeo sistemático de los enfoques declarativos e imperativos en infraestructura como código: tendencias de investigación, paradigmas y soporte de herramientas"

En este chat analizarás artículos completos. Se te enviarán sección por sección (introducción, metodología, resultados, conclusión, etc.).

INSTRUCCIONES DE RECEPCIÓN:
- Cuando recibas una sección responde SOLO: "Sección recibida."
- NO analices ni generes tablas hasta que recibas el mensaje: "CERRAR ARTÍCULO"
- Cuando recibas ese mensaje genera las DOS tablas siguientes.

TABLA 1 — RÚBRICA DE VALIDACIÓN
Completa esta tabla DESPUÉS de recibir TODO el contenido del artículo (cuando recibas "CERRAR ARTÍCULO").
Extrae Autor(es) y Año de la información que hayas recibido en las secciones.

| Campo | Valor |
|---|---|
| Artículo (Título) | |
| Autor(es) | |
| Año | |
| Pertinencia (0-2) | |
| Actualidad (0-2) | |
| Rigor (0-2) | |
| Fuente Calidad (0-2) | |
| Impacto (0-2) | |
| Ética (0-2) | |
| Puntaje Total (/12) | |
| Decisión | |
| Justificación | |

CRITERIOS DE PUNTUACIÓN DETALLADOS:

- Pertinencia (0-2): ¿IaC es el OBJETO DE ESTUDIO o solo una HERRAMIENTA?
  0 = IaC es medio/soporte. Artículo investiga otra cosa (ML, seguridad, DevOps general) y menciona IaC de paso
    Ejemplo: "Usamos Terraform para desplegar nuestro modelo ML"
  1 = IaC comparte protagonismo. Artículo trata IaC + otras cosas con igual énfasis
    Ejemplo: "Comparamos DevOps tools (Jenkins, Docker, Terraform, Kubernetes)"
  2 = IaC es el tema central. Artículo investiga paradigmas, herramientas, prácticas o desafíos DE IaC
    Ejemplo: "Análisis comparativo de Terraform vs Ansible en infraestructura híbrida"
    Pregunta guía: ¿El investigador hace experimentos CON IaC o SOBRE IaC?

- Actualidad (0-2): ¿Qué tan reciente es el contenido respecto a 2025?
  0 = Pre-2018. Desactualizado, cambios mayores en IaC desde entonces
  1 = 2018-2022. Aceptable pero las herramientas/prácticas evolucionaron
  2 = 2023-2025. Reciente, incluye evoluciones recientes de IaC
    Pregunta guía: Si el artículo habla de Terraform, ¿menciona features de 2023+? ¿O es con versiones antiguas?

- Rigor (0-2): ¿Qué tan bien está documentado el método?
  0 = Sin método claro. No describe cómo hizo el estudio, o es principalmente opinión
  1 = Método parcial. Describe qué hizo pero falta: población de estudio, variables, métricas concretas, o no reconoce limitaciones
  2 = Método completo. Describe: población/muestra, variables independiente y dependiente, métricas concretas, limitaciones del estudio
    Pregunta guía: ¿Puedo replicar el estudio con la información dada? ¿Reconoce limitaciones?

- Fuente Calidad (0-2): ¿Dónde se publicó? ¿Es confiable?
  0 = No indexada. Publicación desconocida, blog personal, white-paper sin revisión
  1 = Indexada pero bajo perfil. Conferencia menor, revista indexada pero no top-tier, springer, researchgate
  2 = Verificable y confiable. Top-tier conference (ICSE, FSE, ASE), revista indexada reputada (ACM TOSE, IEEE TSE)
    Pregunta guía: ¿Está en ACM/IEEE/ScienceDirect? ¿Tuvo revisión por pares?

- Impacto (0-2): ¿Tiene influencia en el área? ¿Es citado?
  0 = Sin impacto observable. <5 citaciones o indetectable
  1 = Impacto moderado. 5-30 citaciones o es referencia menor en algunos trabajos
  2 = Alto impacto. >30 citaciones O es referencia frecuente en papers recientes sobre IaC
    Pregunta guía: ¿Este paper se cita frecuentemente en trabajos sobre IaC?

- Ética (0-2): ¿Hay transparencia y completitud en afiliaciones y datos?
  0 = Ausente. Sin DOI, sin afiliaciones claras, sin datos abiertos (si aplica), editorial dudosa
  1 = Parcial. Tiene DOI pero sin afiliaciones, o editorial aceptable pero sin transparencia completa
  2 = Completa. DOI presente, afiliaciones de autores claras, datos abiertos o disponibles, sin conflicto de intereses aparente
    Pregunta guía: ¿Puedo rastrear a los autores? ¿Tienen conflicto de intereses? ¿Los datos son replicables?

CRITERIO DE DECISIÓN:
- Aprobado: IaC es el OBJETO DE ESTUDIO del artículo.
  El artículo investiga, analiza, evalúa o compara IaC, sus paradigmas (declarativo, imperativo, híbrido), sus herramientas, prácticas, desafíos o adopción.

- Excluido: IaC aparece como herramienta de soporte o medio para desplegar otra cosa (una plataforma, un sistema ML, una app, etc.) pero NO es lo que el artículo investiga.

EXCEPCIONES - EXCLUIR SIEMPRE aunque mencionen IaC:
- Dockerfiles y estudios centrados en Docker como herramienta principal
- Herramientas IaC válidas: Terraform, Ansible, Puppet, Chef, Pulumi, CloudFormation, OpenTofu, AWS CDK y similares

TABLA 2 — FICHA DEL ARTÍCULO
Completa esta tabla DESPUÉS de haber recibido TODO el contenido (al recibir "CERRAR ARTÍCULO").

| Campo | Contenido |
|---|---|
| Autor / Año | |
| Objetivo | |
| Método | |
| Hallazgo | |
| Limitación | |
| Referencia textual / Sección | |
| Paradigma IaC cubierto | |
| Herramienta IaC analizada | |
| Tipo de estudio | |
| Variables definidas | |
| Dimensiones | |
| Indicadores medibles | |
| Trabajo futuro declarado | |
| ¿Generalizable a híbrido? | |
| Uso en tu tesis | |

CRITERIOS TABLA 2:
- Objetivo: qué problema o pregunta aborda
- Método: cómo lo investiga
- Hallazgo: resultado o conclusión principal
- Limitación: restricciones que el artículo reconoce
- Referencia textual / Sección: frase exacta del artículo + en qué sección aparece (ej: "Results / Section 4.1")
- Paradigma IaC cubierto: Declarativo / Imperativo / Híbrido / No especificado
- Herramienta IaC analizada: Terraform / Ansible / Puppet / Pulumi / CloudFormation / Varias / No especificado
- Tipo de estudio: Mining study / Experimento controlado / Revisión sistemática / Propuesta de herramienta / Caso de estudio / Encuesta
- Variables definidas: variable independiente y dependiente (ej: VI: tipo módulo, VD: idempotencia)
- Dimensiones: categorías en que agrupa hallazgos. Formato: lista con guiones y saltos de línea reales, NO HTML (ej: "- Security" ENTER "- Maintainability")
- Indicadores medibles: métricas concretas usadas. Formato: lista con guiones (ej: "- % roles con módulos imperativos\n- Cyclomatic complexity")
- Trabajo futuro declarado: qué propone el autor como investigación pendiente. Si hay múltiples, usar lista con guiones.
- ¿Generalizable a híbrido?: Sí / No / Parcial (+ cita textual si existe)
- Uso en tu tesis: cómo aporta al SMS, qué enfoque cubre, qué pregunta responde. Mantener conciso.

IMPORTANTE:
- No inventes datos que no estén en el texto
- Si un dato no está disponible escribe: No especificado

Responde ahora: "Listo, esperando secciones"
"""


def read_consolidated_sections(article_folder: Path) -> list:
    """Lee secciones del consolidated.md EN EL ORDEN ORIGINAL (respetando el consolidado)."""
    consolidated_file = article_folder / "consolidated.md"

    if not consolidated_file.exists():
        return []

    with open(consolidated_file, "r", encoding="utf-8") as f:
        content = f.read()

    sections = []
    # Buscar patrones de secciones: ### [SECCIÓN X/Y - TIPO]
    # Captura: metadatos ANTES de ---, y contenido DESPUÉS de ---
    section_pattern = r"### \[SECCIÓN (\d+)/(\d+) - (PRINCIPAL|SUBSECCIÓN)\]\n\n(.*?)---\n\n(.*?)(?=### \[SECCIÓN|$)"

    matches = re.finditer(section_pattern, content, re.DOTALL)

    for match in matches:
        section_num = int(match.group(1))
        total_sections = int(match.group(2))
        section_type = match.group(3)
        metadata_text = match.group(4)
        section_content_raw = match.group(5)

        # Extraer metadatos
        metadata = {}
        lines = metadata_text.split("\n")
        for line in lines:
            if line.startswith("**"):
                key_match = re.match(r"\*\*(.+?):\*\*\s+(.*)", line)
                if key_match:
                    metadata[key_match.group(1)] = key_match.group(2)

        # El contenido real está DESPUÉS del ---
        section_content = section_content_raw.strip()

        sections.append({
            "number": section_num,
            "total": total_sections,
            "type": section_type,
            "title": metadata.get("Título", ""),
            "content": section_content,
            "metadata": metadata
        })

    # IMPORTANTE: No reordenar - mantener orden original del consolidado
    return sections


def extract_author_year_from_content(sections: list) -> tuple:
    """Extrae Autor(es) y Año del contenido del artículo."""
    authors = ""
    year = ""

    # Buscar en la PRIMERA SECCIÓN (contenido principal)
    if sections:
        first_section_content = sections[0].get("content", "")
        lines = first_section_content.split('\n')

        # Buscar año (2020-2026)
        for line in lines:
            year_match = re.search(r'\b(20\d{2})\b', line)
            if year_match:
                year = year_match.group(1)
                break

        # Extraer autores: primeras líneas que contienen nombres propios
        # Ignorar metadatos (líneas que empiezan con **)
        author_lines = []
        for i, line in enumerate(lines):
            stripped = line.strip()
            # Saltear metadatos y líneas vacías
            if not stripped or stripped.startswith('**'):
                continue
            # Si encuentra email, parar (indica fin de autores)
            if '@' in stripped:
                break
            # Si encuentra una línea que parece ser contenido principal (mayúscula), parar
            if i > 0 and stripped and len(stripped) > 50 and stripped[0].isupper():
                # Probablemente es el inicio del contenido real
                break
            # Agregar líneas con nombres/institutos (no son números puros ni muy largas)
            if stripped and not stripped.isdigit() and len(stripped) < 100:
                author_lines.append(stripped)

        # Tomar primeros 1-2 autores encontrados
        if author_lines:
            authors = ", ".join(author_lines[:2])

    return authors or "No especificado", year or "No especificado"


def analyze_article(client, article_name: str, article_folder: Path) -> tuple:
    """Analiza un artículo enviando secciones EN ORDEN a DeepSeek."""
    print(f"\n  Analizando: {article_name}")

    # Inicializar conversación
    messages = []

    # Enviar prompt maestro
    print(f"    Enviando prompt maestro...", end="", flush=True)
    response = send_to_deepseek(client, messages, MASTER_PROMPT_TEMPLATE)
    print(f" ✓")

    # Leer secciones del consolidated.md
    sections = read_consolidated_sections(article_folder)

    if not sections:
        print(f"    ⚠ Sin secciones para analizar")
        return None, None

    print(f"    Secciones encontradas: {len(sections)}")


    # Enviar cada sección EN ORDEN a DeepSeek
    for section in sections:
        section_info = f"[SECCIÓN {section['number']}/{section['total']} - {section['type']}]\n"
        section_info += f"Título: {section['title']}\n\n"
        section_info += section['content']

        # DEBUG: Mostrar primeras 2 secciones (donde están autores/año)
        if section['number'] <= 2:
            print(f"\n[DEBUG] SECCIÓN {section['number']} A ENVIAR (primeros 800 chars):")
            print(f"{section_info[:800]}")
            print(f"[DEBUG] Total caracteres: {len(section_info)}\n")

        print(f"    [{section['number']}/{section['total']}] {section['title'][:40]}...", end="", flush=True)
        try:
            response = send_to_deepseek(client, messages, section_info)
            print(f" ✓")
        except Exception as e:
            print(f" ✗ Error: {e}")
            continue

    # Cerrar artículo y recibir tablas
    print(f"    Compilando análisis...", end="", flush=True)
    try:
        final_response = send_to_deepseek(
            client,
            messages,
            "CERRAR ARTÍCULO"
        )
        print(f" ✓")
    except Exception as e:
        print(f"    ✗ Error en análisis final: {e}")
        return None, None

    # DEBUG: Mostrar respuesta INMEDIATAMENTE
    print(f"\n\n{'='*70}")
    print(f"[DEBUG] RESPUESTA FINAL DE DEEPSEEK (primeros 2000 caracteres):")
    print(f"{'='*70}")
    print(final_response[:2000])
    print(f"{'='*70}\n")

    # Parsear respuesta (extraer tablas con Autor(es) y Año incluidos)
    tabla1, tabla2 = parse_deepseek_response(final_response, article_name)

    return tabla1, tabla2


def parse_markdown_table(block: str) -> dict:
    """Extrae datos de tabla Markdown."""
    lines = block.strip().split("\n")
    if len(lines) < 2:
        return {}

    # Extraer encabezados
    header_line = lines[0]
    headers = [h.strip() for h in header_line.split("|")[1:-1]]

    result = {}
    for i in range(2, len(lines)):
        cells = [c.strip() for c in lines[i].split("|")[1:-1]]
        if len(cells) >= 2:
            key = cells[0]
            value = cells[1] if len(cells) > 1 else ""
            result[key] = value

    return result


def parse_deepseek_response(response: str, article_name: str) -> tuple:
    """Parsea tablas Markdown de la respuesta de DeepSeek."""
    # DEBUG: Mostrar respuesta completa
    print(f"\n[DEBUG] ═════════════════════════════════════════")
    print(f"[DEBUG] RESPUESTA COMPLETA DE DEEPSEEK:")
    print(f"[DEBUG] ═════════════════════════════════════════")
    print(response)
    print(f"[DEBUG] ═════════════════════════════════════════\n")

    # Buscar bloques de tabla (| ... | )
    table_blocks = re.findall(r"(\|[^\n]+(?:\n\|[^\n]+)+)", response)
    print(f"[DEBUG] Tablas encontradas: {len(table_blocks)}")
    for i, block in enumerate(table_blocks, 1):
        print(f"[DEBUG] TABLA {i}:\n{block}\n")

    tabla1 = {}
    tabla2 = {}

    for block in table_blocks:
        parsed = parse_markdown_table(block)
        keys_lower = {k.lower() for k in parsed}

        # Identificar tabla1 por claves características
        if any(k in keys_lower for k in ("artículo (título)", "artículo", "pertinencia", "puntaje")):
            tabla1 = parsed
        # Identificar tabla2 por claves características
        elif any(k in keys_lower for k in ("autor / año", "objetivo", "hallazgo")):
            tabla2 = parsed

    # Fallback por posición
    if not tabla1 and len(table_blocks) >= 1:
        tabla1 = parse_markdown_table(table_blocks[0])
    if not tabla2 and len(table_blocks) >= 2:
        tabla2 = parse_markdown_table(table_blocks[1])

    # Normalizar claves y convertir valores (DeepSeek devuelve Autor(es) y Año en TABLA 1)
    tabla1_final = {
        "Artículo (Título)": tabla1.get("Artículo (Título)", tabla1.get("Artículo", article_name)),
        "Autor(es)": tabla1.get("Autor(es)", tabla1.get("Autores", "No especificado")),
        "Año": tabla1.get("Año", "No especificado"),
        "Pertinencia": int(tabla1.get("Pertinencia (0-2)", tabla1.get("Pertinencia", 1)) or 1),
        "Actualidad": int(tabla1.get("Actualidad (0-2)", tabla1.get("Actualidad", 1)) or 1),
        "Rigor": int(tabla1.get("Rigor (0-2)", tabla1.get("Rigor", 1)) or 1),
        "Fuente (Calidad)": int(tabla1.get("Fuente Calidad (0-2)", tabla1.get("Fuente (Calidad)", 1)) or 1),
        "Impacto": int(tabla1.get("Impacto (0-2)", tabla1.get("Impacto", 1)) or 1),
        "Ética": int(tabla1.get("Ética (0-2)", tabla1.get("Ética", 1)) or 1),
        "Puntaje": int(tabla1.get("Puntaje Total (/12)", tabla1.get("Puntaje", 6)) or 6),
        "Veredicto": tabla1.get("Decisión", "Excluido"),
        "Justificación": tabla1.get("Justificación", tabla1.get("Motivo Rechazo F2", ""))
    }

    # Función para limpiar HTML y convertir a saltos de línea reales
    def clean_html_breaks(text):
        """Convierte <br> y <br/> a saltos de línea reales."""
        if isinstance(text, str):
            text = text.replace("<br>", "\n")
            text = text.replace("<br/>", "\n")
            text = text.replace("<br />", "\n")
        return text

    tabla2_final = {
        "Objetivo": clean_html_breaks(tabla2.get("Objetivo", "No especificado")),
        "Método": clean_html_breaks(tabla2.get("Método", tabla2.get("Metodo", "No especificado"))),
        "Hallazgo": clean_html_breaks(tabla2.get("Hallazgo", tabla2.get("Hallazgos", "No especificado"))),
        "Limitación": clean_html_breaks(tabla2.get("Limitación", tabla2.get("Limitaciones", "No especificado"))),
        "Referencia textual / Sección": clean_html_breaks(tabla2.get("Referencia textual / Sección", "No especificado")),
        "Paradigma IaC cubierto": clean_html_breaks(tabla2.get("Paradigma IaC cubierto", "No especificado")),
        "Herramienta IaC analizada": clean_html_breaks(tabla2.get("Herramienta IaC analizada", "No especificado")),
        "Tipo de estudio": clean_html_breaks(tabla2.get("Tipo de estudio", "No especificado")),
        "Variables definidas": clean_html_breaks(tabla2.get("Variables definidas", "No especificado")),
        "Dimensiones": clean_html_breaks(tabla2.get("Dimensiones", "No especificado")),
        "Indicadores medibles": clean_html_breaks(tabla2.get("Indicadores medibles", "No especificado")),
        "Trabajo futuro declarado": clean_html_breaks(tabla2.get("Trabajo futuro declarado", "No especificado")),
        "¿Generalizable a híbrido?": clean_html_breaks(tabla2.get("¿Generalizable a híbrido?", "No especificado")),
        "Uso en tu tesis": clean_html_breaks(tabla2.get("Uso en tu tesis", tabla2.get("Uso en la tesis", "No especificado")))
    }

    return tabla1_final, tabla2_final


def check_if_analyzed_in_lectura_selectiva(ws_lectura, article_num: int) -> bool:
    """Verifica si un artículo ya tiene datos en Lectura Selectiva (columnas C-P)."""
    row = article_num  # N° + 3 para saltar encabezados

    # Verificar si hay algo en las columnas C-P (Objetivo hasta Generalizable)
    for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P']:
        cell_value = ws_lectura[f"{col}{row}"].value
        if cell_value:  # Si encuentra algo, ya está analizado
            return True

    return False


def analyzer_main():
    """Analiza artículos con DeepSeek - Lee Excel y analiza máximo 3 nuevos."""
    print("=" * 70)
    print("DEEPSEEK ANALYZER - Analiza artículos pendientes (máximo 3)")
    print("=" * 70)

    if not OUTPUT_DIR.exists():
        print("[INFO] No hay artículos en output/")
        sys.exit(0)

    # Leer Excel para ver qué artículos ya están
    excel_path = os.getenv("EXCEL_PATH")
    try:
        import openpyxl
        wb_read = openpyxl.load_workbook(excel_path)
        ws_revision = wb_read["Revisión de Artículos"]
        ws_lectura = wb_read["Lectura Selectiva"]
    except Exception as e:
        print(f"[ERROR] No se pudo leer Excel: {e}")
        sys.exit(1)

    # Obtener artículos ya en Excel (para no repetir)
    articulos_en_excel = set()
    for row_idx in range(4, 44):
        title_cell = ws_revision[f"B{row_idx}"].value
        if title_cell:
            articulos_en_excel.add(str(title_cell).strip().lower())

    # 1. DETECTAR DINÁMICAMENTE dónde está el encabezado (fila con "N°" o equivalente)
    header_row = None
    for row_idx in range(1, 10):
        cell_a = ws_revision[f"A{row_idx}"].value
        cell_b = ws_revision[f"B{row_idx}"].value

        # Buscar indicios de encabezado: "N°", "Autor", "Título", etc.
        if cell_a and isinstance(cell_a, str) and any(x in str(cell_a).lower() for x in ["n°", "numero"]):
            header_row = row_idx
            break
        elif cell_b and isinstance(cell_b, str) and any(x in str(cell_b).lower() for x in ["autor", "artículo", "título"]):
            header_row = row_idx
            break

    # 2. El contenido REAL empieza DESPUÉS del encabezado
    content_start = (header_row + 1) if header_row else 2  # Default: fila 2

    # 3. Buscar primera fila VACÍA: solo verificar columna B (Artículo/Título)
    # Si B está vacío = disponible para nuevo artículo
    next_row = content_start
    for row_idx in range(content_start, 44):
        num_en_a = ws_revision[f"A{row_idx}"].value
        titulo_en_b = ws_revision[f"B{row_idx}"].value

        # Si hay número en A pero B está vacío → FILA VACÍA, escribir aquí
        if num_en_a is not None and not titulo_en_b:
            next_row = row_idx
            break

    # Buscar artículos en output/ que NO estén en Excel
    article_folders = []

    for folder in OUTPUT_DIR.iterdir():
        if folder.is_dir() and (folder / "consolidated.md").exists():
            folder_name = folder.name.lower()

            # Verificar si NO está en Excel
            if not any(folder_name in art for art in articulos_en_excel):
                article_folders.append((next_row, folder))
                next_row += 1

    wb_read.close()

    if not article_folders:
        if articulos_encontrados > 0:
            print("[INFO] Todos los artículos en Excel están analizados")
        else:
            print("[INFO] No hay artículos en Excel ni en output/ para procesar")
        sys.exit(0)

    print(f"\nEncontrados {len(article_folders)} artículos pendientes")
    print(f"Analizando máximo 3...\n")

    # Crear cliente DeepSeek
    client = create_deepseek_client()

    # Procesar cada artículo (máximo 3)
    MAX_ARTICLES = 3
    for i, (art_num, article_folder) in enumerate(article_folders):
        if i >= MAX_ARTICLES:
            print(f"\n[INFO] Límite de {MAX_ARTICLES} artículos alcanzado (para testing)")
            break

        article_name = article_folder.name

        print(f"\n[{art_num}] {article_name}")

        # Analizar
        tabla1, tabla2 = analyze_article(client, article_name, article_folder)

        if tabla1 is None:
            print(f"  WARNING No se pudo analizar")
            continue

        # Guardar en Excel
        try:
            write_article_to_excel(excel_path, art_num, tabla1, tabla2)
            print(f"  OK Guardado en Excel (N°{art_num})")
            print(f"  OK Veredicto: {tabla1['Veredicto']} | Puntaje: {tabla1['Puntaje']}")

        except Exception as e:
            print(f"  ERROR escribiendo Excel: {e}")

    print("\n" + "=" * 70)
    print("✓ Análisis completado")
    print("=" * 70)


if __name__ == "__main__":
    analyzer_main()
