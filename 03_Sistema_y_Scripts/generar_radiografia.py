#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import sys
from pathlib import Path
from collections import Counter, defaultdict

import fitz  # PyMuPDF
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
import openpyxl

BASE_DIR = Path(__file__).resolve().parent
CERT_DIR = BASE_DIR.parent / "Certificados"


def generar_arbol(directorio: Path) -> str:
    """Genera el árbol visual de directorios y archivos dentro de Certificados/."""
    lineas = [f"{directorio.name}/"]
    for root, dirs, files in os.walk(directorio):
        dirs.sort()
        files.sort()
        rel_path = Path(root).relative_to(directorio)
        nivel = 0 if str(rel_path) == "." else len(rel_path.parts)
        indent = "│   " * nivel
        if str(rel_path) != ".":
            lineas.append(f"{'│   ' * (nivel - 1)}├── 📁 {Path(root).name}/ ({len(files)} archivos)")
        for i, f in enumerate(files):
            conector = "└──" if i == len(files) - 1 and not dirs else "├──"
            lineas.append(f"{indent}{conector} 📄 {f}")
    return "\n".join(lineas)


def extraer_pagina_pdf_estructurada(page) -> str:
    """
    Extrae el texto de una página PDF agrupando elementos por coordenada Y (renglón visual)
    y separando columnas/campos horizontales con ' | ' a velocidad C.
    """
    try:
        data = page.get_text("dict")
    except Exception:
        return page.get_text("text", sort=True)

    items = []
    for block in data.get("blocks", []):
        if block.get("type") != 0:  # Solo bloques de texto
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            # Unir spans muy pegados dentro de la misma línea y separar si hay salto horizontal grande
            curr_x0, curr_y0, curr_x1, curr_y1 = spans[0]["bbox"]
            curr_text = spans[0]["text"]
            for sp in spans[1:]:
                sx0, sy0, sx1, sy1 = sp["bbox"]
                stext = sp["text"]
                if not stext.strip():
                    curr_text += " "
                    curr_x1 = max(curr_x1, sx1)
                    continue
                if sx0 - curr_x1 > 8.0:
                    if curr_text.strip():
                        items.append((curr_x0, (curr_y0 + curr_y1) / 2.0, curr_x1, curr_text.strip()))
                    curr_x0, curr_y0, curr_x1, curr_y1 = sx0, sy0, sx1, sy1
                    curr_text = stext
                else:
                    curr_text += stext
                    curr_x1 = max(curr_x1, sx1)
            if curr_text.strip():
                items.append((curr_x0, (curr_y0 + curr_y1) / 2.0, curr_x1, curr_text.strip()))

    if not items:
        return ""

    # Ordenar por posición vertical (Y) y luego horizontal (X)
    items.sort(key=lambda x: (round(x[1], 1), x[0]))

    # Agrupar en renglones visuales (tolerancia de 3.5 puntos en Y)
    filas = []
    fila_actual = [items[0]]
    y_ref = items[0][1]

    for item in items[1:]:
        if abs(item[1] - y_ref) <= 3.5:
            fila_actual.append(item)
        else:
            filas.append(fila_actual)
            fila_actual = [item]
            y_ref = item[1]
    filas.append(fila_actual)

    lineas_salida = []
    for fila in filas:
        fila.sort(key=lambda x: x[0])
        partes = [fila[0][3]]
        prev_x1 = fila[0][2]
        for x0, _, x1, txt in fila[1:]:
            if x0 - prev_x1 > 6.0:
                partes.append(" | ")
            else:
                partes.append(" ")
            partes.append(txt)
            prev_x1 = max(prev_x1, x1)
        linea = "".join(partes).strip()
        if linea:
            lineas_salida.append(linea)

    return "\n".join(lineas_salida)


def procesar_pdf(filepath: Path):
    """Devuelve (estado, num_paginas, paginas_texto, paginas_imagen, contenido_texto)."""
    try:
        doc = fitz.open(filepath)
    except Exception as e:
        return ("ERROR_PDF", 0, 0, 0, f"[ERROR AL ABRIR PDF: {e}]")

    num_pages = len(doc)
    pag_texto = 0
    pag_img = 0
    bloques_pagina = []

    for idx, page in enumerate(doc, start=1):
        texto_pag = extraer_pagina_pdf_estructurada(page)
        solo_alfanum = re.sub(r"\s+", "", texto_pag)
        if len(solo_alfanum) >= 15:
            pag_texto += 1
            if num_pages > 1:
                bloques_pagina.append(f"--- [Página {idx}/{num_pages}] ---\n{texto_pag}")
            else:
                bloques_pagina.append(texto_pag)
        else:
            pag_img += 1
            bloques_pagina.append(f"--- [Página {idx}/{num_pages}: IMAGEN / SIN TEXTO SELECCIONABLE] ---")

    doc.close()

    if pag_texto == 0:
        estado = "IMAGEN_ESCANEADO"
    elif pag_img > 0:
        estado = "MIXTO_TEXTO_E_IMAGEN"
    else:
        estado = "TEXTO_NATIVO"

    return (estado, num_pages, pag_texto, pag_img, "\n".join(bloques_pagina))


def procesar_docx(filepath: Path):
    """Extrae párrafos y tablas de archivos .docx en orden."""
    try:
        doc = docx.Document(filepath)
    except Exception as e:
        return ("ERROR_DOCX", f"[ERROR AL ABRIR DOCX: {e}]")

    lineas = []
    # Intentar mantener el orden exacto de párrafos y tablas
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("p"):
            p = Paragraph(child, doc)
            txt = p.text.strip()
            if txt:
                lineas.append(txt)
        elif child.tag.endswith("tbl"):
            t = Table(child, doc)
            lineas.append("[INICIO TABLA DOCX]")
            for row in t.rows:
                celdas = [re.sub(r"\s+", " ", c.text.strip()) for c in row.cells]
                # Evitar celdas combinadas repetidas consecutivas idénticas si toda la fila es igual
                if any(celdas):
                    lineas.append(" | ".join(celdas))
            lineas.append("[FIN TABLA DOCX]")

    contenido = "\n".join(lineas)
    if len(re.sub(r"\s+", "", contenido)) < 15:
        return ("DOCX_IMAGEN_O_VACIO", "[DOCX SIN TEXTO EXTRAÍBLE - POSIBLE IMAGEN PEGADA ADENTRO]")
    return ("DOCX_TEXTO", contenido)


def procesar_xlsx(filepath: Path):
    """Extrae el contenido de hojas Excel .xlsx."""
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    except Exception as e:
        return ("ERROR_XLSX", f"[ERROR AL ABRIR XLSX: {e}]")

    lineas = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        lineas.append(f"--- [HOJA EXCEL: {sheet}] ---")
        filas_leidas = 0
        for row in ws.iter_rows(values_only=True):
            valores = [str(v).strip() if v is not None else "" for v in row]
            if any(valores):
                lineas.append(" | ".join(valores))
                filas_leidas += 1
                if filas_leidas >= 300:
                    lineas.append("... [TRUNCADO A 300 FILAS] ...")
                    break
    wb.close()
    return ("XLSX_DATOS", "\n".join(lineas))


def procesar_doc_antiguo(filepath: Path):
    """Intenta extraer texto legible de archivos .doc (Word 97-2003) sin dependencias externas."""
    try:
        raw = filepath.read_bytes()
    except Exception as e:
        return ("ERROR_DOC", f"[ERROR AL LEER DOC: {e}]")

    # Extraer secuencias ASCII/Latin1 y UTF-16LE legibles
    chunks_utf16 = [m.group(0).decode("utf-16le", errors="ignore")
                    for m in re.finditer(rb"(?:[\x20-\x7e\xa0-\xff]\x00){6,}", raw)]
    chunks_latin = [m.group(0).decode("latin-1", errors="ignore")
                    for m in re.finditer(rb"[\x20-\x7e\xa0-\xff]{8,}", raw)]

    texto_combinado = "\n".join(chunks_utf16 if len(chunks_utf16) > 5 else chunks_latin)
    # Limpiar basura de metadatos comunes de OLE2
    lineas_limpias = []
    for l in texto_combinado.splitlines():
        l_strip = l.strip()
        if not l_strip:
            continue
        if any(bad in l_strip for bad in ["Microsoft Word", "Normal.dot", "MSWordDoc", "Word.Document", "Root Entry", "CompObj"]):
            continue
        lineas_limpias.append(l_strip)

    resultado = "\n".join(lineas_limpias)
    if len(re.sub(r"\s+", "", resultado)) < 20:
        return ("DOC_ANTIGUO_SIN_TEXTO", "[ARCHIVO .DOC ANTIGUO SIN TEXTO CLARO O IMAGEN]")
    return ("DOC_ANTIGUO_TEXTO", resultado)


def es_pdf_encubierto(filepath: Path) -> bool:
    """Detecta si un archivo con extensión rara empieza con %PDF-."""
    try:
        with open(filepath, "rb") as f:
            return f.read(5) == b"%PDF-"
    except Exception:
        return False


def main():
    if not CERT_DIR.exists():
        print(f"❌ No se encontró el directorio: {CERT_DIR}")
        sys.exit(1)

    print(f"🔍 Escaneando árbol de directorios en: {CERT_DIR} ...")
    arbol_str = generar_arbol(CERT_DIR)

    # Recolectar todos los archivos
    todos_archivos = []
    for root, dirs, files in os.walk(CERT_DIR):
        dirs.sort()
        files.sort()
        for f in files:
            todos_archivos.append(Path(root) / f)

    total = len(todos_archivos)
    print(f"📦 Total de archivos encontrados: {total}\n")

    conteo_estados = Counter()
    pendientes_imagen = []
    archivos_especiales = []
    textos_por_categoria = defaultdict(list)

    for idx, fpath in enumerate(todos_archivos, start=1):
        if idx % 100 == 0 or idx == total:
            print(f"⏳ Procesando [{idx}/{total}] archivos...")

        rel_path = fpath.relative_to(CERT_DIR)
        categoria = rel_path.parts[0] if len(rel_path.parts) > 1 else "Raiz"
        ext = fpath.suffix.lower()

        # Verificar si es PDF (por extensión o por cabecera mágica %PDF-)
        if ext == ".pdf" or es_pdf_encubierto(fpath):
            estado, n_pag, p_txt, p_img, contenido = procesar_pdf(fpath)
            conteo_estados[estado] += 1

            if estado in ("IMAGEN_ESCANEADO", "MIXTO_TEXTO_E_IMAGEN"):
                pendientes_imagen.append(f"{estado} | Págs(Txt:{p_txt}/Img:{p_img}) | {rel_path}")

            if estado in ("TEXTO_NATIVO", "MIXTO_TEXTO_E_IMAGEN"):
                bloque = (
                    f"{'='*80}\n"
                    f"ARCHIVO: {rel_path}\n"
                    f"TIPO: PDF | ESTADO: {estado} | PAGINAS: {n_pag} (Texto: {p_txt}, Imagen: {p_img})\n"
                    f"{'='*80}\n"
                    f"{contenido}\n\n"
                )
                textos_por_categoria[categoria].append(bloque)

        elif ext == ".docx":
            estado, contenido = procesar_docx(fpath)
            conteo_estados[estado] += 1
            if estado == "DOCX_IMAGEN_O_VACIO":
                pendientes_imagen.append(f"{estado} | {rel_path}")
            else:
                bloque = (
                    f"{'='*80}\n"
                    f"ARCHIVO: {rel_path}\n"
                    f"TIPO: DOCX | ESTADO: {estado}\n"
                    f"{'='*80}\n"
                    f"{contenido}\n\n"
                )
                textos_por_categoria[categoria].append(bloque)

        elif ext == ".doc":
            estado, contenido = procesar_doc_antiguo(fpath)
            conteo_estados[estado] += 1
            if estado == "DOC_ANTIGUO_SIN_TEXTO":
                pendientes_imagen.append(f"{estado} | {rel_path}")
            else:
                bloque = (
                    f"{'='*80}\n"
                    f"ARCHIVO: {rel_path}\n"
                    f"TIPO: DOC_97_2003 | ESTADO: {estado}\n"
                    f"{'='*80}\n"
                    f"{contenido}\n\n"
                )
                textos_por_categoria[categoria].append(bloque)

        elif ext == ".xlsx":
            estado, contenido = procesar_xlsx(fpath)
            conteo_estados[estado] += 1
            bloque = (
                f"{'='*80}\n"
                f"ARCHIVO: {rel_path}\n"
                f"TIPO: XLSX | ESTADO: {estado}\n"
                f"{'='*80}\n"
                f"{contenido}\n\n"
            )
            textos_por_categoria[categoria].append(bloque)

        elif ext in (".txt", ".csv"):
            conteo_estados["TXT_CSV"] += 1
            try:
                contenido = fpath.read_text(encoding="utf-8", errors="replace")
            except Exception as e:
                contenido = f"[ERROR LECTURA TXT: {e}]"
            bloque = (
                f"{'='*80}\n"
                f"ARCHIVO: {rel_path}\n"
                f"TIPO: {ext.upper()} | ESTADO: TEXTO\n"
                f"{'='*80}\n"
                f"{contenido}\n\n"
            )
            textos_por_categoria[categoria].append(bloque)

        else:
            conteo_estados[f"OTRO_{ext or 'SIN_EXT'}"] += 1
            size_kb = fpath.stat().st_size / 1024.0
            archivos_especiales.append(f"{ext or 'SIN_EXT'} ({size_kb:.1f} KB) | {rel_path}")
            if ext in (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"):
                pendientes_imagen.append(f"IMAGEN_SUELTA ({ext}) | {rel_path}")

    # 1. Guardar lista de pendientes de imagen (OCR)
    ruta_pendientes = BASE_DIR / "pendientes_imagen.txt"
    with open(ruta_pendientes, "w", encoding="utf-8") as f:
        f.write(f"TOTAL ARCHIVOS CON IMAGEN / ESCANEADOS: {len(pendientes_imagen)}\n")
        f.write("=" * 80 + "\n")
        for linea in pendientes_imagen:
            f.write(linea + "\n")

    # 2. Construir cabecera de resumen
    resumen_lines = [
        "=" * 80,
        "RESUMEN GENERAL DE RADIOGRAFÍA - CHIRCA / CERTIFICADOS",
        "=" * 80,
        f"Total de archivos escaneados: {total}",
        "",
        "DESGLOSE POR ESTADO:",
    ]
    for est, cant in conteo_estados.most_common():
        resumen_lines.append(f"  - {est}: {cant}")

    resumen_lines.append("\nARCHIVOS ESPECIALES / EXTENSIONES NO ESTÁNDAR:")
    for esp in archivos_especiales:
        resumen_lines.append(f"  * {esp}")

    resumen_lines.append("\n" + "=" * 80)
    resumen_lines.append("ÁRBOL COMPLETO DE DIRECTORIOS Y ARCHIVOS")
    resumen_lines.append("=" * 80)
    resumen_lines.append(arbol_str)
    resumen_lines.append("\n")

    cabecera_resumen = "\n".join(resumen_lines)

    # Guardar 00_RESUMEN_Y_ARBOL.txt
    ruta_resumen = BASE_DIR / "00_RESUMEN_Y_ARBOL.txt"
    ruta_resumen.write_text(cabecera_resumen, encoding="utf-8")

    # 3. Guardar archivos por categoría y el archivo completo
    ruta_completa = BASE_DIR / "radiografia_COMPLETA.txt"
    with open(ruta_completa, "w", encoding="utf-8") as f_comp:
        f_comp.write(cabecera_resumen)
        for cat, bloques in sorted(textos_por_categoria.items()):
            nombre_limpio = re.sub(r"[^a-zA-Z0-9_]", "_", cat)
            ruta_cat = BASE_DIR / f"radiografia_{nombre_limpio}.txt"
            contenido_cat = (
                f"{'='*80}\n"
                f"CATEGORÍA: {cat} ({len(bloques)} documentos con texto)\n"
                f"{'='*80}\n\n"
                + "".join(bloques)
            )
            ruta_cat.write_text(contenido_cat, encoding="utf-8")
            f_comp.write(contenido_cat)

    print("\n✅ ¡PROCESO TERMINADO CON ÉXITO!")
    print("-" * 60)
    for est, cant in conteo_estados.most_common():
        print(f"  {est:28s}: {cant}")
    print("-" * 60)
    print("📂 Archivos generados en ~/Documents/Chirca/:")
    for p in sorted(BASE_DIR.glob("*.txt")):
        mb = p.stat().st_size / (1024 * 1024)
        tokens_est = int(p.stat().st_size / 3.8)
        print(f"  - {p.name:32s} -> {mb:6.2f} MB (~{tokens_est:,} tokens est.)")


if __name__ == "__main__":
    main()
