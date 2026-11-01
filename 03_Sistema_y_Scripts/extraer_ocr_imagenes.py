#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import time
import re
from pathlib import Path
from collections import defaultdict

import fitz  # PyMuPDF
from rapidocr_onnxruntime import RapidOCR

BASE_DIR = Path(__file__).resolve().parent
CERT_DIR = BASE_DIR.parent / "Certificados"
PENDIENTES_FILE = BASE_DIR / "pendientes_imagen.txt"
CACHE_FILE = BASE_DIR / "cache_ocr_imagenes.json"


def reconstruir_tabla_desde_cajas_ocr(result_ocr) -> str:
    """
    Toma las cajas [box, texto, score] de RapidOCR y las agrupa por coordenada Y
    (renglón visual), ordenando de izquierda a derecha con ' | '.
    """
    if not result_ocr:
        return ""

    items = []
    for box, txt, score in result_ocr:
        txt_clean = str(txt).strip()
        if not txt_clean or score < 0.35:
            continue
        # box = [[x0, y0], [x1, y0], [x1, y1], [x0, y1]]
        xs = [pt[0] for pt in box]
        ys = [pt[1] for pt in box]
        x0, x1 = min(xs), max(xs)
        y_center = (min(ys) + max(ys)) / 2.0
        alto = max(ys) - min(ys)
        items.append((x0, y_center, x1, alto, txt_clean))

    if not items:
        return ""

    items.sort(key=lambda x: (x[1], x[0]))

    # Tolerancia vertical dinámica (~mitad de la altura de letra, aprox 10 px a 150 DPI)
    filas = []
    fila_actual = [items[0]]
    y_ref = items[0][1]
    tol_y = max(8.0, items[0][3] * 0.55)

    for it in items[1:]:
        if abs(it[1] - y_ref) <= tol_y:
            fila_actual.append(it)
            y_ref = sum(x[1] for x in fila_actual) / len(fila_actual)
        else:
            filas.append(fila_actual)
            fila_actual = [it]
            y_ref = it[1]
            tol_y = max(8.0, it[3] * 0.55)
    filas.append(fila_actual)

    lineas_salida = []
    for fila in filas:
        fila.sort(key=lambda x: x[0])
        partes = [fila[0][4]]
        prev_x1 = fila[0][2]
        for x0, _, x1, _, txt in fila[1:]:
            if x0 - prev_x1 > 14.0:
                partes.append(" | ")
            else:
                partes.append(" ")
            partes.append(txt)
            prev_x1 = max(prev_x1, x1)
        lin = "".join(partes).strip()
        if lin:
            lineas_salida.append(lin)

    return "\n".join(lineas_salida)


def ocr_archivo(fpath: Path, engine: RapidOCR) -> str:
    ext = fpath.suffix.lower()
    if ext in (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"):
        res, _ = engine(str(fpath))
        return reconstruir_tabla_desde_cajas_ocr(res)

    # Es un PDF (o .2 que empieza con %PDF)
    try:
        doc = fitz.open(fpath)
    except Exception as e:
        return f"[ERROR AL ABRIR PDF PARA OCR: {e}]"

    num_pages = len(doc)
    bloques_pag = []

    for idx, page in enumerate(doc, start=1):
        # Si alguna página ya tiene texto nativo, usar el texto nativo directo
        txt_nativo = page.get_text("text").strip()
        if len(re.sub(r"\s+", "", txt_nativo)) >= 25:
            bloques_pag.append(f"--- [Página {idx}/{num_pages} (TEXTO NATIVO)] ---\n{txt_nativo}")
            continue

        # Renderizar página a 150 DPI en escala de grises (rápido y nítido para certificados)
        pix = page.get_pixmap(dpi=150, colorspace=fitz.csGRAY)
        img_bytes = pix.tobytes("png")
        res, _ = engine(img_bytes)
        texto_ocr = reconstruir_tabla_desde_cajas_ocr(res)
        if num_pages > 1:
            bloques_pag.append(f"--- [Página {idx}/{num_pages} (OCR)] ---\n{texto_ocr}")
        else:
            bloques_pag.append(texto_ocr)

    doc.close()
    return "\n".join(bloques_pag)


def main():
    if not PENDIENTES_FILE.exists():
        print(f"❌ No se encontró {PENDIENTES_FILE}")
        return

    # Cargar caché incremental si ya existe de una corrida previa
    cache = {}
    if CACHE_FILE.exists():
        try:
            cache = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    lineas_p = PENDIENTES_FILE.read_text(encoding="utf-8").splitlines()
    archivos_procesar = []

    for l in lineas_p:
        if "|" not in l or l.startswith("TOTAL") or l.startswith("="):
            continue
        partes = [p.strip() for p in l.split("|")]
        estado = partes[0]
        rel_path = partes[-1]

        # Ignorar el DOCX vacío de tendencias y el compilado repetido de 76 páginas
        if estado == "DOCX_IMAGEN_O_VACIO":
            continue
        if "reseumen coas 2025.pdf" in rel_path.lower():
            continue
        if "certificados de analisis aspire chemicals.pdf" in rel_path.lower():
            continue

        archivos_procesar.append(rel_path)

    total = len(archivos_procesar)
    ya_en_cache = sum(1 for r in archivos_procesar if r in cache)
    print("=" * 75)
    print(f"🔎 INICIANDO OCR INTELIGENTE SOBRE CERTIFICADOS ESCANEADOS")
    print(f"   Total de archivos objetivo : {total}")
    print(f"   Ya procesados en caché     : {ya_en_cache}")
    print(f"   Pendientes por escanear hoy: {total - ya_en_cache}")
    print("=" * 75)

    engine = RapidOCR()
    t_inicio = time.time()
    procesados_hoy = 0

    for idx, rel_str in enumerate(archivos_procesar, start=1):
        if rel_str in cache and len(cache[rel_str].strip()) > 10:
            continue

        fpath = CERT_DIR / rel_str
        if not fpath.exists():
            continue

        t0 = time.time()
        texto_extraido = ocr_archivo(fpath, engine)
        dt = time.time() - t0
        procesados_hoy += 1

        cache[rel_str] = texto_extraido
        # Guardar caché en disco cada 5 archivos para no perder nada si se cancela
        if procesados_hoy % 5 == 0 or idx == total:
            CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")

        restantes = (total - ya_en_cache) - procesados_hoy
        vel_media = (time.time() - t_inicio) / max(1, procesados_hoy)
        eta_min = (restantes * vel_media) / 60.0
        nombre_corto = Path(rel_str).name[:38]
        print(f"⏳ [{idx:3d}/{total}] ({dt:4.1f}s | ETA: {eta_min:4.1f} min) -> {nombre_corto}")

    # Guardado final del caché
    CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")

    # Exportar archivos radiografia_OCR_<Categoria>.txt con el mismo formato que los de texto
    por_categoria = defaultdict(list)
    for rel_str, texto_ocr in sorted(cache.items()):
        cat = rel_str.split("/")[0] if "/" in rel_str else "Otros"
        bloque = (
            f"{'='*80}\n"
            f"ARCHIVO: {rel_str}\n"
            f"TIPO: PDF_OCR | ESTADO: IMAGEN_PROCESADA_CON_OCR\n"
            f"{'='*80}\n"
            f"{texto_ocr}\n\n"
        )
        por_categoria[cat].append(bloque)

    print("\n📂 Generando archivos de radiografía OCR por categoría:")
    for cat, bloques in sorted(por_categoria.items()):
        nom_limpio = re.sub(r"[^a-zA-Z0-9_]", "_", cat)
        out_txt = BASE_DIR / f"radiografia_OCR_{nom_limpio}.txt"
        contenido = (
            f"{'='*80}\n"
            f"CATEGORÍA OCR: {cat} ({len(bloques)} documentos escaneados leídos)\n"
            f"{'='*80}\n\n"
            + "".join(bloques)
        )
        out_txt.write_text(contenido, encoding="utf-8")
        print(f"   ✔ {out_txt.name:34s} ({len(bloques)} certificados recuperados)")

    print("\n✅ ¡OCR COMPLETADO CON ÉXITO! Todos los textos quedaron guardados en caché.")


if __name__ == "__main__":
    main()
