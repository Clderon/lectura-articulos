#!/usr/bin/env python3
"""
Fragmenter - Divide .md en chunks por capítulos/secciones
"""
import sys
import os
from pathlib import Path
from dotenv import load_dotenv
import re
import json

from shared_utils import cut_at_references, SECTION_HEADER_PATTERNS

# UTF-8 en Windows
if sys.platform == "win32":
    os.environ["PYTHONIOENCODING"] = "utf-8"
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

load_dotenv()

OUTPUT_DIR = Path(__file__).parent / "output"
CHUNKS_SUBDIR = "chunks"
MAX_CHUNK_SIZE = 4000  # palabras máximo por chunk


def sanitize_filename(filename: str) -> str:
    """Elimina caracteres inválidos en nombres de archivo."""
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '', filename)
    sanitized = re.sub(r'\s+', ' ', sanitized).strip()
    return sanitized[:100]


def roman_to_int(s: str) -> int:
    """Convierte números romanos a enteros para ordenamiento."""
    roman_map = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}
    result = 0
    prev_value = 0
    for char in reversed(s.upper()):
        if char not in roman_map:
            return float('inf')
        value = roman_map[char]
        if value < prev_value:
            result -= value
        else:
            result += value
        prev_value = value
    return result


def sort_sections(sections: list) -> list:
    """Ordena secciones según su numeración (I, II, A, B, etc.)."""
    def get_sort_key(section_title: str):
        # Extraer el número/letra al inicio
        # Buscar números romanos: I, II, III, IV, V, VI, etc.
        roman_match = re.match(r"^([IVX]+)\.", section_title)
        if roman_match:
            roman = roman_match.group(1)
            return (roman_to_int(roman), 0, "")  # (número romano, 0 para principal, "")

        # Buscar letras: A, B, C, D, etc.
        letter_match = re.match(r"^([A-Z])\.", section_title)
        if letter_match:
            letter = letter_match.group(1)
            return (1000, ord(letter) - ord('A'), "")  # (gran número, posición letra, "")

        # Si no tiene numeración clara, enviar al final
        return (10000, 0, section_title)

    return sorted(sections, key=lambda s: get_sort_key(s["title"]))


def extract_sections_from_pages(article_folder: Path) -> list:
    """Extrae secciones de las páginas EN ORDEN - todas, aunque estén vacías."""
    sections = []
    md_files = sorted(article_folder.glob("page_*.md"))

    current_section = None  # Sin sección inicial

    for md_file in md_files:
        with open(md_file, "r", encoding="utf-8") as f:
            content = f.read()

        # Cortar antes de referencias
        content = cut_at_references(content)

        lines = content.split("\n")

        for line in lines:
            # SOLO crear nueva sección si encuentra **[title]**
            if "**[title]**" in line:
                # Guardar sección anterior AUNQUE ESTÉ VACÍA (importante para estructura)
                if current_section is not None:
                    section_content = "\n".join(current_section["content"]).strip()
                    current_section["content"] = section_content
                    sections.append(current_section)

                # Extraer título
                title_match = re.search(r"\*\*\[title\]\*\*\s+(.+)", line)
                section_title = title_match.group(1).strip() if title_match else ""

                if section_title:
                    current_section = {"title": section_title, "content": []}
                else:
                    current_section = None
            else:
                # Agregar línea a la sección actual (si existe)
                if current_section is not None and line.strip():
                    current_section["content"].append(line)

    # Guardar última sección (AUNQUE ESTÉ VACÍA)
    if current_section is not None:
        section_content = "\n".join(current_section["content"]).strip()
        current_section["content"] = section_content
        sections.append(current_section)

    return sections


def split_section_into_chunks(section_title: str, section_content: str) -> list:
    """Divide una sección en chunks si es muy grande. Si se divide: nombre_2, nombre_3, etc."""
    chunks = []
    words = section_content.split()
    chunk_words = []
    chunk_num = 1

    for word in words:
        chunk_words.append(word)

        if len(chunk_words) >= MAX_CHUNK_SIZE:
            # Guardar chunk
            chunk_content = " ".join(chunk_words)
            chunks.append({"title": section_title, "content": chunk_content, "num": chunk_num})

            chunk_words = []
            chunk_num += 1

    # Guardar chunk final
    if chunk_words:
        chunk_content = " ".join(chunk_words)
        chunks.append({"title": section_title, "content": chunk_content, "num": chunk_num})

    # Si hay múltiples chunks, agregar número al nombre
    if len(chunks) > 1:
        for i, chunk in enumerate(chunks, 1):
            chunk["title"] = f"{chunk['title']}_{i}"

    return chunks


def fragmenter_main():
    """Fragmenta artículos por secciones, dividiendo si son muy grandes."""
    print("=" * 70)
    print("FRAGMENTER - .md → chunks por secciones")
    print("=" * 70)

    # Encontrar carpetas procesadas
    if not OUTPUT_DIR.exists():
        print("[INFO] No hay artículos procesados")
        sys.exit(0)

    article_folders = [d for d in OUTPUT_DIR.iterdir() if d.is_dir() and (d / ".DONE").exists()]

    if not article_folders:
        print("[INFO] No hay artículos con estado .DONE")
        sys.exit(0)

    print(f"\nFragmentando {len(article_folders)} artículos...\n")

    # Procesar cada artículo
    for article_folder in article_folders:
        article_name = article_folder.name
        print(f"\n{article_name}")

        # Extraer secciones de las páginas EN ORDEN (simple, directo)
        sections = extract_sections_from_pages(article_folder)

        if not sections:
            print(f"  ⚠ Sin secciones")
            continue

        print(f"  Secciones extraídas: {len(sections)}")
        for i, sec in enumerate(sections, 1):
            print(f"    {i}. {sec['title'][:60]}")

        # Crear archivo consolidado con metadatos para DeepSeek
        consolidated_file = article_folder / "consolidated.md"
        total_sections = len(sections)

        with open(consolidated_file, "w", encoding="utf-8") as f:
            f.write(f"# {article_name}\n\n")
            f.write(f"**Total de secciones:** {total_sections}\n\n")
            f.write("---\n\n")

            # Escribir cada sección con metadatos
            for idx, section in enumerate(sections, 1):
                # Determinar tipo (principal o subsección)
                is_subsection = re.match(r"^[A-Z]\.", section['title'])
                section_type = "Subsección" if is_subsection else "Principal"

                # Escribir metadatos de la sección
                f.write(f"### [SECCIÓN {idx}/{total_sections} - {section_type.upper()}]\n\n")
                f.write(f"**Tipo:** {section_type}\n")
                f.write(f"**Título:** {section['title']}\n")
                f.write(f"**Orden en Secuencia:** {idx}/{total_sections}\n")
                f.write(f"**Contenido Original:** Sin modificar\n\n")
                f.write("---\n\n")

                # Escribir contenido original exactamente como mineru lo sacó
                f.write(f"{section['content']}\n\n")
                f.write("---\n\n")

        print(f"  Secciones con metadatos: {total_sections}")
        print(f"  → Guardado en: {consolidated_file}")

        # Crear archivo de confirmación
        done_file = article_folder / ".FRAGMENTS_DONE"
        done_file.touch()

    print("\n" + "=" * 70)
    print("✓ Fragmentación completada")
    print("✓ Próximo paso: ejecutar deepseek_analyzer.py")
    print("=" * 70)


if __name__ == "__main__":
    fragmenter_main()
