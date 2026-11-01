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
LOCAL_CERT_DIR = BASE_DIR.parent / "Certificados"

# Rutas posibles de SharePoint sincronizado en Windows
CANDIDATOS_SHAREPOINT = [
    Path("C:/Users/Pasantias/OVC SA/GAMAS - GAMAS"),
    Path.home() / "OVC SA" / "GAMAS - GAMAS",
    Path.home() / "Chirca OVC S.A" / "GAMAS - GAMAS",
]

# Mapeo exacto: (Categoría, Prefijo Virtual compatible con los parsers, Ruta en GAMAS - GAMAS)
MAPA_SHAREPOINT = [
    (
        "Arcillas",
        Path("Arcillas/PUREMIN/CERTIFICADO DE CALIDAD"),
        Path("5. ARCILLAS/PUREMIN/CERTIFICADO DE CALIDAD"),
    ),
    (
        "Arcillas",
        Path("Arcillas/SEPIOL S.A/CERTIFICADOS DE CALIDAD"),
        Path("5. ARCILLAS/SEPIOL S.A/CERTIFICADOS DE CALIDAD"),
    ),
    (
        "Biotecnología",
        Path("Biotecnología/FERMENTIS/CERTIFICADOS DE CALIDAD Y DE EXTENSIÓN"),
        Path("2. BIOTECNOLOGÍA/FERMENTIS/CERTIFICADOS DE CALIDAD Y DE EXTENSIÓN"),
    ),
    (
        "Biotecnología",
        Path("Biotecnología/VINEXPERT/CERTIFICADO DE CALIDAD"),
        Path("2. BIOTECNOLOGÍA/VINEXPERT/CERTIFICADO DE CALIDAD"),
    ),
    (
        "Carbon",
        Path("Carbon/ASPIRE CHEMICAL-OVC Y VINEXPERT CARBONES/CERTIFICADOS DE CALIDAD"),
        Path("4. CARBÓN ACTIVADO/ASPIRE CHEMICAL-OVC Y VINEXPERT CARBONES/CERTIFICADOS DE CALIDAD"),
    ),
    (
        "Carbon",
        Path("Carbon/NORIT/CERTIFICADOS DE CALIDAD"),
        Path("4. CARBÓN ACTIVADO/NORIT/CERTIFICADOS DE CALIDAD"),
    ),
    (
        "FILTRACIÓN",
        Path("FILTRACIÓN"),
        Path("3. FILTRACIÓN/PRECOAT/CERTIFICADOS DE CALIDAD"),
    ),
]

CARPETAS_EXCLUIDAS = (
    "TENDENCIA",
    "TENDENDIA",
    "EDITABLE",
    "NO USAR",
    "VENCIDO",
    "OBSOLETO",
    "COAS DE BOBINAS",
    "CERTIFICADOS DE EXTENSION",
)


def detectar_origen_datos():
    """Detecta si estamos en SharePoint (Windows) o en carpeta local Certificados (Linux)."""
    for sp_dir in CANDIDATOS_SHAREPOINT:
        if sp_dir.exists():
            return "SHAREPOINT", sp_dir
    if LOCAL_CERT_DIR.exists():
        return "LOCAL", LOCAL_CERT_DIR
    return None, None


def recolectar_archivos_objetivo(modo: str, raiz_dir: Path):
    """
    Devuelve una lista de tuplas: (fpath_real, rel_str_virtual, categoria)
    usando siempre barras '/' para mantener compatibilidad total entre Windows y Linux.
    """
    archivos = []

    if modo == "LOCAL":
        for root, dirs, files in os.walk(raiz_dir):
            dirs[:] = sorted([
                d for d in dirs
                if not any(ex in d.upper() for ex in CARPETAS_EXCLUIDAS)
            ])
            files.sort()
            for f in files:
                if f.startswith("~$") or f in (".DS_Store", "Thumbs.db"):
                    continue
                fpath = Path(root) / f
                rel_path = fpath.relative_to(raiz_dir)
                rel_str = rel_path.as_posix()
                categoria = rel_path.parts[0] if len(rel_path.parts) > 1 else "Raiz"
                archivos.append((fpath, rel_str, categoria))
    else:
        for categoria, prefijo_virtual, sub_sp in MAPA_SHAREPOINT:
            dir_objetivo = raiz_dir / sub_sp
            if not dir_objetivo.exists():
                print(f"⚠️ Subcarpeta no encontrada en SharePoint: {dir_objetivo}")
                continue
            for root, dirs, files in os.walk(dir_objetivo):
                dirs[:] = sorted([
                    d for d in dirs
                    if not any(ex in d.upper() for ex in CARPETAS_EXCLUIDAS)
                ])
                files.sort()
                for f in files:
                    if f.startswith("~$") or f in (".DS_Store", "Thumbs.db"):
                        continue
                    fpath = Path(root) / f
                    rel_interno = fpath.relative_to(dir_objetivo)
                    rel_virtual = prefijo_virtual / rel_interno
                    archivos.append((fpath, rel_virtual.as_posix(), categoria))

    return archivos


def cargar_cache_radiografias_previas():
    """
    Carga los bloques ya extraídos en radiografia_*.txt y pendientes_imagen.txt
    para NO volver a abrir ni descargar de OneDrive los ~2.000 PDFs históricos.
    """
    bloques_por_cat = defaultdict(list)
    conocidos_rel = set()
    conocidos_cat_nombre = set()
    conteo_estados = Counter()
    pendientes_previos = []

    mapa_cat_archivos = [
        ("Arcillas", "radiografia_Arcillas.txt"),
        ("Biotecnología", "radiografia_Biotecnolog_a.txt"),
        ("Carbon", "radiografia_Carbon.txt"),
        ("FILTRACIÓN", "radiografia_FILTRACI_N.txt"),
    ]

    for cat, nom_txt in mapa_cat_archivos:
        p = BASE_DIR / nom_txt
        if not p.exists():
            p = BASE_DIR / "datos_radiografia" / nom_txt
        if not p.exists():
            continue

        txt = p.read_text(encoding="utf-8", errors="replace")
        partes = txt.split("=" * 80 + "\nARCHIVO: ")[1:]
        for b in partes:
            lineas = b.splitlines()
            if not lineas:
                continue
            rel_str = lineas[0].strip().replace("\\", "/")
            meta = lineas[1].strip() if len(lineas) > 1 else ""
            m_est = re.search(r"ESTADO:\s*([^|\n]+)", meta)
            estado = m_est.group(1).strip() if m_est else "TEXTO_NATIVO"

            bloque_completo = "=" * 80 + "\nARCHIVO: " + b
            if not bloque_completo.endswith("\n\n"):
                bloque_completo = bloque_completo.rstrip() + "\n\n"

            bloques_por_cat[cat].append(bloque_completo)
            conocidos_rel.add(rel_str)
            conocidos_cat_nombre.add((cat, Path(rel_str).name.lower()))
            conteo_estados[estado] += 1

    p_pend = BASE_DIR / "pendientes_imagen.txt"
    if p_pend.exists():
        for lin in p_pend.read_text(encoding="utf-8", errors="replace").splitlines():
            if "|" not in lin or lin.startswith("TOTAL") or lin.startswith("="):
                continue
            partes_p = [x.strip() for x in lin.split("|")]
            est_p = partes_p[0]
            rel_p = partes_p[-1].replace("\\", "/")
            cat_p = rel_p.split("/")[0] if "/" in rel_p else "Raiz"
            pendientes_previos.append(lin)
            if rel_p not in conocidos_rel:
                conocidos_rel.add(rel_p)
                conocidos_cat_nombre.add((cat_p, Path(rel_p).name.lower()))
                conteo_estados[est_p] += 1

    return bloques_por_cat, conocidos_rel, conocidos_cat_nombre, conteo_estados, pendientes_previos


def extraer_pagina_pdf_estructurada(page) -> str:
    try:
        data = page.get_text("dict")
    except Exception:
        return page.get_text("text", sort=True)

    items = []
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
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

    items.sort(key=lambda x: (round(x[1], 1), x[0]))

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
    try:
        doc = docx.Document(filepath)
    except Exception as e:
        return ("ERROR_DOCX", f"[ERROR AL ABRIR DOCX: {e}]")

    lineas = []
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
                if any(celdas):
                    lineas.append(" | ".join(celdas))
            lineas.append("[FIN TABLA DOCX]")

    contenido = "\n".join(lineas)
    if len(re.sub(r"\s+", "", contenido)) < 15:
        return ("DOCX_IMAGEN_O_VACIO", "[DOCX SIN TEXTO EXTRAÍBLE - POSIBLE IMAGEN PEGADA ADENTRO]")
    return ("DOCX_TEXTO", contenido)


def procesar_xlsx(filepath: Path):
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
    try:
        raw = filepath.read_bytes()
    except Exception as e:
        return ("ERROR_DOC", f"[ERROR AL LEER DOC: {e}]")

    chunks_utf16 = [m.group(0).decode("utf-16le", errors="ignore")
                    for m in re.finditer(rb"(?:[\x20-\x7e\xa0-\xff]\x00){6,}", raw)]
    chunks_latin = [m.group(0).decode("latin-1", errors="ignore")
                    for m in re.finditer(rb"[\x20-\x7e\xa0-\xff]{8,}", raw)]

    texto_combinado = "\n".join(chunks_utf16 if len(chunks_utf16) > 5 else chunks_latin)
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
    try:
        with open(filepath, "rb") as f:
            return f.read(5) == b"%PDF-"
    except Exception:
        return False


def main():
    modo, raiz_dir = detectar_origen_datos()
    if modo is None:
        print("❌ No se encontró ni la carpeta de SharePoint sincronizada ni la carpeta local 'Certificados'.")
        sys.exit(1)

    print(f"🔍 Modo detectado: {modo} -> Escaneando {raiz_dir} ...")

    (
        textos_por_categoria,
        conocidos_rel,
        conocidos_cat_nombre,
        conteo_estados,
        pendientes_imagen,
    ) = cargar_cache_radiografias_previas()

    todos_archivos = recolectar_archivos_objetivo(modo, raiz_dir)
    total = len(todos_archivos)
    print(f"📦 Archivos en carpetas objetivo: {total} (En caché histórico: {len(conocidos_rel)})")

    nuevos_procesados = 0
    archivos_especiales = []

    for idx, (fpath, rel_str, categoria) in enumerate(todos_archivos, start=1):
        # Si ya está en la radiografía previa, NO lo abrimos (evita que OneDrive descargue el PDF)
        if rel_str in conocidos_rel or (categoria, fpath.name.lower()) in conocidos_cat_nombre:
            continue

        nuevos_procesados += 1
        print(f"🆕 Procesando archivo nuevo [{nuevos_procesados}]: {rel_str}")
        ext = fpath.suffix.lower()

        if ext == ".pdf" or (ext not in (".docx", ".doc", ".xlsx", ".txt", ".csv", ".jpg", ".jpeg", ".png") and es_pdf_encubierto(fpath)):
            estado, n_pag, p_txt, p_img, contenido = procesar_pdf(fpath)
            conteo_estados[estado] += 1

            if estado in ("IMAGEN_ESCANEADO", "MIXTO_TEXTO_E_IMAGEN"):
                pendientes_imagen.append(f"{estado} | Págs(Txt:{p_txt}/Img:{p_img}) | {rel_str}")

            if estado in ("TEXTO_NATIVO", "MIXTO_TEXTO_E_IMAGEN"):
                bloque = (
                    f"{'='*80}\n"
                    f"ARCHIVO: {rel_str}\n"
                    f"TIPO: PDF | ESTADO: {estado} | PAGINAS: {n_pag} (Texto: {p_txt}, Imagen: {p_img})\n"
                    f"{'='*80}\n"
                    f"{contenido}\n\n"
                )
                textos_por_categoria[categoria].append(bloque)

        elif ext == ".docx":
            estado, contenido = procesar_docx(fpath)
            conteo_estados[estado] += 1
            if estado == "DOCX_IMAGEN_O_VACIO":
                pendientes_imagen.append(f"{estado} | {rel_str}")
            else:
                bloque = (
                    f"{'='*80}\n"
                    f"ARCHIVO: {rel_str}\n"
                    f"TIPO: DOCX | ESTADO: {estado}\n"
                    f"{'='*80}\n"
                    f"{contenido}\n\n"
                )
                textos_por_categoria[categoria].append(bloque)

        elif ext == ".doc":
            estado, contenido = procesar_doc_antiguo(fpath)
            conteo_estados[estado] += 1
            if estado == "DOC_ANTIGUO_SIN_TEXTO":
                pendientes_imagen.append(f"{estado} | {rel_str}")
            else:
                bloque = (
                    f"{'='*80}\n"
                    f"ARCHIVO: {rel_str}\n"
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
                f"ARCHIVO: {rel_str}\n"
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
                f"ARCHIVO: {rel_str}\n"
                f"TIPO: {ext.upper()} | ESTADO: TEXTO\n"
                f"{'='*80}\n"
                f"{contenido}\n\n"
            )
            textos_por_categoria[categoria].append(bloque)

        else:
            conteo_estados[f"OTRO_{ext or 'SIN_EXT'}"] += 1
            if ext in (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"):
                pendientes_imagen.append(f"IMAGEN_SUELTA ({ext}) | {rel_str}")

    if nuevos_procesados == 0:
        print("⚡ No hay certificados nuevos sin procesar (0 descargas de OneDrive requeridas).")
    else:
        print(f"✅ Se extrajeron {nuevos_procesados} certificados nuevos.")

    ruta_pendientes = BASE_DIR / "pendientes_imagen.txt"
    with open(ruta_pendientes, "w", encoding="utf-8") as f:
        f.write(f"TOTAL ARCHIVOS CON IMAGEN / ESCANEADOS: {len(pendientes_imagen)}\n")
        f.write("=" * 80 + "\n")
        for linea in pendientes_imagen:
            f.write(linea + "\n")

    ruta_completa = BASE_DIR / "radiografia_COMPLETA.txt"
    with open(ruta_completa, "w", encoding="utf-8") as f_comp:
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


if __name__ == "__main__":
    main()
