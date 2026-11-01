#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import csv
import shutil
from pathlib import Path
from collections import defaultdict, Counter
import openpyxl

BASE_DIR = Path(__file__).resolve().parent

BACKUP_DIR = BASE_DIR / "datos_radiografia"
BACKUP_DIR.mkdir(exist_ok=True)
for f_bak in BASE_DIR.glob("radiografia_*.txt"):
    shutil.copy2(f_bak, BACKUP_DIR / f_bak.name)
if (BASE_DIR / "cache_ocr_imagenes.json").exists():
    shutil.copy2(BASE_DIR / "cache_ocr_imagenes.json", BACKUP_DIR / "cache_ocr_imagenes.json")

LINEAS_DIRECCION_ARCILLAS = {
    "PUREMIN",
    "MINERSA GROUP",
    "OVC S.A.",
    "OVCS.A.",
    "OVC S.A. CQ",
    "ANCHORENA 1540",
    "ANCHORENA1540",
    "00005507 LUJAN DE CUYO",
    "00005507LUJANDECUYO",
    "MENDOZA",
    "MENDOZA (ARGENTINA)",
    "MENDOZA(ARGENTINA)",
    "CERTIFICADO DE INSPECCIÓN DE CALIDAD",
    "CERTIFICADO DE INSPECCION DE CALIDAD",
}

MESES_MAP = {
    "ene": 1, "enero": 1, "jan": 1, "january": 1,
    "feb": 2, "febrero": 2, "february": 2,
    "mar": 3, "marzo": 3, "march": 3,
    "abr": 4, "abril": 4, "apr": 4, "april": 4,
    "may": 5, "mayo": 5,
    "jun": 6, "junio": 6, "june": 6,
    "jul": 7, "julio": 7, "july": 7,
    "ago": 8, "agosto": 8, "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "septiembre": 9, "setiembre": 9, "september": 9,
    "oct": 10, "octubre": 10, "october": 10,
    "nov": 11, "noviembre": 11, "november": 11,
    "dic": 12, "diciembre": 12, "dec": 12, "december": 12,
}


def limpiar_float(val_str: str):
    if not val_str:
        return None
    s = str(val_str).strip()
    s = s.replace("：", "").replace("号", "%").strip(" :")
    if s.upper() in ("NA", "N/A", "-", "", "AUSENCIA", "ABSENT", "NEGATIVO", "CONFORM", "CONFORME", "CUMPLE", "PASS"):
        return None
    m = re.search(r"[-+]?\d+(?:[,\.]\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0).replace(",", "."))
    except ValueError:
        return None


def formatear_decimal_csv(val_float):
    if val_float is None:
        return ""
    if val_float == int(val_float):
        return str(int(val_float))
    return f"{val_float:.4f}".rstrip("0").rstrip(".").replace(".", ",")


def extraer_fecha_universal(texto: str):
    if not texto:
        return "", None

    texto_corr = re.sub(r"\b29(1\d|2\d)\b", r"20\1", texto)

    m1 = re.search(r"\b(\d{1,2})\s*[\./\-]\s*(\d{1,2})\s*[\./\-]\s*(20\d{2}|\d{2})\b", texto_corr)
    if m1:
        p1, p2, y = int(m1.group(1)), int(m1.group(2)), int(m1.group(3))
        if y < 100:
            y += 2000
        if p2 > 12 and p1 <= 12:
            d, mes = p2, p1
        else:
            d, mes = p1, p2
        if 2012 <= y <= 2099 and 1 <= mes <= 12 and 1 <= d <= 31:
            return f"{d:02d}/{mes:02d}/{y:04d}", y

    m2 = re.search(r"\b(20\d{2})\s*[\./\-]\s*(\d{1,2})\s*[\./\-]\s*(\d{1,2})\b", texto_corr)
    if m2:
        y, mes, d = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        if 2012 <= y <= 2099 and 1 <= mes <= 12 and 1 <= d <= 31:
            return f"{d:02d}/{mes:02d}/{y:04d}", y

    texto_ocr_d = re.sub(r"\b([0-3])I([A-Za-z]{3})", r"\g<1>1\2", texto_corr)
    m3 = re.search(
        r"\b(\d{1,2})\s*(?:de\s+|\-)?\s*([A-Za-z]{3,10})\.?\s*(?:de\s+|,\s*|\-)?\s*(20\d{2}|\d{2})\b",
        texto_ocr_d,
        re.IGNORECASE,
    )
    if m3:
        d = int(m3.group(1))
        mes_txt = m3.group(2).lower()
        y = int(m3.group(3))
        if y < 100:
            y += 2000
        mes = MESES_MAP.get(mes_txt[:3]) or MESES_MAP.get(mes_txt)
        if mes and 2012 <= y <= 2099 and 1 <= d <= 31:
            return f"{d:02d}/{mes:02d}/{y:04d}", y

    m4 = re.search(
        r"\b([A-Za-z]{3,9})[\.,\s]+(\d{1,2})[\.,\s]+(20\d{2})\b",
        texto_corr,
        re.IGNORECASE,
    )
    if m4:
        mes_txt = m4.group(1).lower()
        d = int(m4.group(2))
        y = int(m4.group(3))
        mes = MESES_MAP.get(mes_txt[:3]) or MESES_MAP.get(mes_txt)
        if mes and 2012 <= y <= 2099 and 1 <= d <= 31:
            return f"{d:02d}/{mes:02d}/{y:04d}", y

    m5 = re.search(r"\b(0?[1-9]|1[0-2])\s*[/\-]\s*(20\d{2})\b", texto_corr)
    if m5:
        mes, y = int(m5.group(1)), int(m5.group(2))
        return f"01/{mes:02d}/{y:04d}", y

    m6 = re.search(r"\b([A-Za-z]{3,10})[\s\-|]+(20\d{2})\b", texto_corr, re.IGNORECASE)
    if m6:
        mes_txt = m6.group(1).lower()
        y = int(m6.group(2))
        mes = MESES_MAP.get(mes_txt[:3]) or MESES_MAP.get(mes_txt)
        if mes and 2012 <= y <= 2099:
            return f"01/{mes:02d}/{y:04d}", y

    return "", None


def leer_bloques_combinados(nom_nativo: str, nom_ocr: str):
    bloques_totales = []
    for nom, origen in [(nom_nativo, "TEXTO_NATIVO"), (nom_ocr, "OCR_IMAGEN")]:
        p = BASE_DIR / nom
        if not p.exists():
            p = BASE_DIR / "datos_radiografia" / nom
        if not p.exists():
            continue
        txt = p.read_text(encoding="utf-8")
        partes = txt.split("=" * 80 + "\nARCHIVO: ")[1:]
        for b in partes:
            bloques_totales.append((origen, b))
    return bloques_totales


def exportar_csvs_y_excel_dinamico(out_dir: Path, prefijo: str, registros: list, campos_num: set):
    out_dir.mkdir(exist_ok=True)
    if not registros:
        return
    columnas = [c for c in registros[0].keys() if not c.startswith("_")]

    def guardar_csv(ruta_csv: Path, filas):
        with open(ruta_csv, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(columnas)
            for r in filas:
                row_out = []
                for col in columnas:
                    val = r[col]
                    if col in campos_num:
                        row_out.append(formatear_decimal_csv(val))
                    else:
                        row_out.append("" if val is None else str(val))
                w.writerow(row_out)

    filas_2024 = [r for r in registros if r["Anio"] <= 2024]
    guardar_csv(out_dir / f"{prefijo}_1_Hasta_2024.csv", filas_2024)
    guardar_csv(out_dir / f"{prefijo}_0_TODOS.csv", registros)

    anios_post = sorted(set(r["Anio"] for r in registros if r["Anio"] >= 2025).union({2025, 2026}))

    wb = openpyxl.Workbook()
    ws_all = wb.active
    ws_all.title = "Todos_Certificados"

    def llenar_hoja(ws, filas):
        ws.append(columnas)
        for r in filas:
            ws.append([r[col] for col in columnas])

    llenar_hoja(ws_all, registros)
    llenar_hoja(
        wb.create_sheet("Lotes_Unicos_Preferidos"),
        [r for r in registros if r.get("Version_Preferida_Unica") == "SI"],
    )
    llenar_hoja(wb.create_sheet("Hasta_2024"), filas_2024)

    for idx_a, anio_p in enumerate(anios_post, start=2):
        filas_a = [r for r in registros if r["Anio"] == anio_p]
        guardar_csv(out_dir / f"{prefijo}_{idx_a}_Anio_{anio_p}.csv", filas_a)
        llenar_hoja(wb.create_sheet(f"Anio_{anio_p}"), filas_a)

    wb.save(out_dir / f"{prefijo}_Consolidado.xlsx")


# ==============================================================================
# 1. ARCILLAS (TEXTO NATIVO + OCR)
# ==============================================================================
def clasificar_producto_arcilla(prod_raw: str, subcarpeta_origen: str, lote: str):
    p_up = prod_raw.upper().strip()
    if "M200" in p_up:
        env = re.sub(r"(?i).*M200\s*UF\s*", "", prod_raw).strip()
        return "SEPIGEL ACTIVE M200 UF", env
    elif "M100" in p_up:
        env = re.sub(r"(?i).*M100\s*UF\s*", "", prod_raw).strip()
        return "SEPIGEL NATURAL M100 UF", env
    elif "E200" in p_up:
        return "SEPIGEL ACTIVE E200", ""
    elif "SEPIOL" in p_up or "2026W" in lote:
        return "SEPIOL", ""
    return "SEPIGEL ACTIVE M200 UF", prod_raw.strip()


def procesar_arcillas_completo():
    bloques = leer_bloques_combinados("radiografia_Arcillas.txt", "radiografia_OCR_Arcillas.txt")
    if not bloques:
        return
    registros = []

    for origen_lectura, b in bloques:
        lineas = [l.strip() for l in b.splitlines() if l.strip()]
        if not lineas:
            continue
        archivo_rel = lineas[0]
        partes_ruta = archivo_rel.split("/")
        subcarpeta_prov = partes_ruta[1] if len(partes_ruta) > 1 else "Puremin"
        nombre_archivo = partes_ruta[-1]
        nivel_anidamiento = len(partes_ruta)
        cuerpo = "\n".join(lineas[3:])

        paginas_raw = re.split(r"--- \[Página \d+/\d+.*?\] ---", cuerpo)
        paginas_limpias = [p.strip() for p in paginas_raw if len(p.strip()) > 40] or [cuerpo]
        total_p = len(paginas_limpias)

        for idx_p, texto_pag in enumerate(paginas_limpias, start=1):
            lineas_limpias = []
            for l in texto_pag.splitlines():
                celdas_l = [c.strip() for c in l.split("|") if c.strip()]
                celdas_f = [c for c in celdas_l if c.upper() not in LINEAS_DIRECCION_ARCILLAS]
                if celdas_f:
                    lineas_limpias.append(" | ".join(celdas_f))

            fecha_emision, anio_fecha = "", None
            albaran, matricula_contenedor, producto_raw, lote = "", "", "", ""
            cantidad_t = None

            for i, lin in enumerate(lineas_limpias):
                lin_low = lin.lower()
                sig = lineas_limpias[i + 1] if i + 1 < len(lineas_limpias) else ""

                if lin_low == "fecha" and not fecha_emision:
                    fecha_emision, anio_fecha = extraer_fecha_universal(sig)
                elif lin_low.startswith("albarán") or lin_low.startswith("albaran"):
                    m_alb = re.search(r"Albar[aá]n[\s:\|]*([0-9\s]+)", lin, re.IGNORECASE)
                    if m_alb and len(m_alb.group(1).strip()) >= 5:
                        albaran = m_alb.group(1).strip()
                    else:
                        albaran = sig.replace("|", " ").strip()
                elif "semiremolque" in lin_low or "contenedor" in lin_low:
                    if "/" in sig or "-" in sig or any(ch.isdigit() for ch in sig):
                        matricula_contenedor = sig.replace(" | ", " ").strip()
                elif "producto" in lin_low and ("especificaci" in lin_low or ":" in lin_low):
                    m_pr = re.search(r"Especificaci[oó]n[\s:\|]+(.+)", lin, re.IGNORECASE)
                    if m_pr and len(m_pr.group(1).strip()) > 3:
                        producto_raw = m_pr.group(1).strip()
                    else:
                        producto_raw = sig.replace("|", " ").strip()
                elif lin_low.startswith("lote"):
                    m_l = re.search(r"Lote[\s:\|]+([A-Za-z0-9\-_]{5,15})", lin, re.IGNORECASE)
                    if m_l and m_l.group(1).lower() not in ("cantidad", "característica", "caracteristica"):
                        lote = m_l.group(1).strip()
                    elif sig and not sig.lower().startswith("cantidad"):
                        m_l2 = re.search(r"\b([0-9A-Za-z\-_]{5,15})\b", sig)
                        if m_l2:
                            lote = m_l2.group(1).strip()
                elif lin_low.startswith("cantidad"):
                    m_c = re.search(r"(\d+[,\.]\d+)\s*T\b", lin + " " + sig, re.IGNORECASE)
                    if m_c:
                        cantidad_t = limpiar_float(m_c.group(1))

            if not lote:
                m_arch = re.search(r"\b(2\d{7,9}|3071425|20\d{2}W\d+)\b", nombre_archivo)
                lote = m_arch.group(1) if m_arch else "SIN_LOTE"
            if lote == "3071425":
                lote = "23071425"

            params = {
                "humedad_lim": "", "humedad_val": None,
                "ret_m200_lim": "", "ret_m200_val": None,
                "ret_m325_lim": "", "ret_m325_val": None,
                "acidez_lim": "", "acidez_val": None,
                "filtrabilidad_val": None, "dens_aparente_val": None,
            }

            for lin in lineas_limpias:
                if "|" not in lin:
                    continue
                celdas = [c.strip() for c in lin.split("|") if c.strip()]
                if len(celdas) < 2:
                    continue
                c0 = celdas[0].lower()
                ult_val = celdas[-1]
                lim_str = celdas[-2] if len(celdas) >= 3 else ""

                if "humedad" in c0:
                    params["humedad_lim"] = lim_str
                    params["humedad_val"] = limpiar_float(ult_val)
                elif "m-200" in c0 or "200 astm" in c0:
                    params["ret_m200_lim"] = lim_str
                    params["ret_m200_val"] = limpiar_float(ult_val)
                elif "m-325" in c0 or "325 astm" in c0:
                    params["ret_m325_lim"] = lim_str
                    params["ret_m325_val"] = limpiar_float(ult_val)
                elif "acidez" in c0 or "mgkoh" in c0:
                    params["acidez_lim"] = lim_str
                    params["acidez_val"] = limpiar_float(ult_val)
                elif "filtrabilidad" in c0:
                    params["filtrabilidad_val"] = limpiar_float(ult_val)
                elif "aparente" in c0 or "densidad" in c0:
                    params["dens_aparente_val"] = limpiar_float(ult_val)

            anio_lote = None
            m_yl = re.match(r"^([1-9]\d)\d{6}$", lote)
            if m_yl:
                anio_lote = 2000 + int(m_yl.group(1))
            elif re.match(r"^(20\d{2})W", lote):
                anio_lote = int(lote[:4])

            anio_final = anio_fecha if anio_fecha is not None else (anio_lote if anio_lote is not None else 2023)
            rango_csv = "<=2024" if anio_final <= 2024 else str(anio_final)
            prod_base, envase = clasificar_producto_arcilla(producto_raw, subcarpeta_prov, lote)

            if params["humedad_val"] is not None or lote != "SIN_LOTE":
                registros.append({
                    "Rango_Temporal": rango_csv,
                    "Anio": anio_final,
                    "Origen_Lectura": origen_lectura,
                    "Producto_Base": prod_base,
                    "Presentacion_Envase": envase,
                    "Lote": lote,
                    "Albaran": albaran,
                    "Contenedor_Transporte": matricula_contenedor,
                    "Fecha_Emision": fecha_emision,
                    "Cantidad_T": cantidad_t,
                    "Envio_Unico_Albaran": "SI",
                    "Lote_Repetido": "NO",
                    "Cant_Envios_Lote": 1,
                    "Version_Preferida_Unica": "SI",
                    "Humedad_Libre_pct": params["humedad_val"],
                    "Humedad_Limite": params["humedad_lim"],
                    "Retenido_M200_pct": params["ret_m200_val"],
                    "Retenido_M200_Limite": params["ret_m200_lim"],
                    "Retenido_M325_pct": params["ret_m325_val"],
                    "Retenido_M325_Limite": params["ret_m325_lim"],
                    "Acidez_Libre_mgKOH": params["acidez_val"],
                    "Acidez_Limite": params["acidez_lim"],
                    "Filtrabilidad_val": params["filtrabilidad_val"],
                    "Densidad_Aparente_gl": params["dens_aparente_val"],
                    "Pagina_PDF": f"{idx_p}/{total_p}",
                    "Ruta_Archivo": archivo_rel,
                    "_nivel": nivel_anidamiento,
                })

    grupos_envio = defaultdict(list)
    for r in registros:
        alb_norm = re.sub(r"\s+", "", r["Albaran"])
        grupos_envio[(r["Producto_Base"], r["Lote"], alb_norm)].append(r)
    for _, lista_e in grupos_envio.items():
        lista_e.sort(key=lambda x: (x["Origen_Lectura"] != "TEXTO_NATIVO", x["_nivel"], x["Ruta_Archivo"]))
        for i, item in enumerate(lista_e):
            item["Envio_Unico_Albaran"] = "SI" if i == 0 else "NO_COPIA_CARPETA"

    grupos_lote = defaultdict(list)
    for r in registros:
        if r["Envio_Unico_Albaran"] == "SI":
            grupos_lote[(r["Producto_Base"], r["Lote"])].append(r)
        else:
            r["Version_Preferida_Unica"] = "NO"

    for _, lista_l in grupos_lote.items():
        cant = len(lista_l)
        lista_l.sort(key=lambda x: (x["Origen_Lectura"] != "TEXTO_NATIVO", x["Albaran"], x["Ruta_Archivo"]))
        for i, item in enumerate(lista_l):
            item["Cant_Envios_Lote"] = cant
            item["Lote_Repetido"] = "SI" if cant > 1 else "NO"
            item["Version_Preferida_Unica"] = "SI" if i == 0 else "NO"

    registros.sort(key=lambda x: (x["Producto_Base"], x["Anio"], x["Lote"], x["Version_Preferida_Unica"] != "SI"))
    campos_num = {
        "Cantidad_T", "Humedad_Libre_pct", "Retenido_M200_pct",
        "Retenido_M325_pct", "Acidez_Libre_mgKOH", "Filtrabilidad_val", "Densidad_Aparente_gl"
    }
    exportar_csvs_y_excel_dinamico(BASE_DIR / "Salida_CSV_Arcillas", "Arcillas", registros, campos_num)

    unicos = [r for r in registros if r["Version_Preferida_Unica"] == "SI"]
    envios = [r for r in registros if r["Envio_Unico_Albaran"] == "SI"]
    print("=" * 75)
    print(f"✅ 1) ARCILLAS COMPLETO (TEXTO + OCR): {len(registros)} págs | {len(envios)} camiones | {len(unicos)} lotes únicos")


# ==============================================================================
# 2. BIOTECNOLOGÍA (TEXTO NATIVO + OCR)
# ==============================================================================
def inferir_anio_lote_biotec(lote: str, nombre_archivo: str):
    m1 = re.match(r"^(20\d{2})\d{2}$", lote)
    if m1:
        return int(m1.group(1))
    m2 = re.match(r"^\d{2}(20\d{2})$", lote)
    if m2:
        return int(m2.group(1))
    m3 = re.match(r"^[A-Z]{2}(\d{2})[A-Z]", lote)
    if m3:
        return 2000 + int(m3.group(1))
    m4 = re.match(r"^(\d{2})\d{4,5}$", lote)
    if m4 and 14 <= int(m4.group(1)) <= 99:
        return 2000 + int(m4.group(1))
    m5 = re.match(r"^(?:0[1-9]|1[0-2])[/\-]?(1[5-9]|[2-9]\d)$", lote)
    if m5:
        return 2000 + int(m5.group(1))
    m6 = re.search(r"\b(20\d{2})\b", nombre_archivo)
    if m6:
        return int(m6.group(1))
    return None


def procesar_biotecnologia_completo():
    bloques = leer_bloques_combinados("radiografia_Biotecnolog_a.txt", "radiografia_OCR_Biotecnolog_a.txt")
    if not bloques:
        return
    registros = []

    for origen_lectura, b in bloques:
        lineas = [l.strip() for l in b.splitlines() if l.strip()]
        if not lineas:
            continue
        archivo_rel = lineas[0]
        partes_ruta = archivo_rel.split("/")
        nombre_archivo = partes_ruta[-1]
        upper_nom = nombre_archivo.upper()
        upper_arch = archivo_rel.upper()

        if any(k in upper_nom for k in [
            "EDITABLE", "TEMPLATE", "BASE DATOS", "CÓDIGO DE ARTICULO",
            "EXTENSION", "EXTENSIÓN", "REVALIDA", "KOSHER", "KPSHER",
            "AUTOADHESIVOS", "RECLAMO", "ALERGENOS", "HDS", "~$", "CERTIFICADO SANITARIO"
        ]):
            continue
        if any(sub_ex in upper_arch for sub_ex in [
            "/CERTIFICADOS DE EXTENSION/",
            "/BASE DATOS YY EDITABLES CERTIFICADOS/",
            "/EDITABLES/",
            "/EDITABLE/",
        ]):
            continue

        meta = lineas[1] if len(lineas) > 1 else ""
        m_t = re.search(r"TIPO:\s*([^|]+)", meta)
        tipo_doc = m_t.group(1).strip() if m_t else "PDF"

        proveedor_linea = partes_ruta[1] if len(partes_ruta) > 1 else "Biotecnologia"
        subcarpeta_prod = partes_ruta[3] if len(partes_ruta) > 3 else (partes_ruta[2] if len(partes_ruta) > 2 else "")
        cuerpo = "\n".join(lineas[3:])

        lote = ""
        m_arch_lote = re.search(
            r"(?:Lote|Lot[e]?\s*#?|LOT|OVC|EC|GP|BA-|OP-)\s*[:\.]?\s*([A-Za-z0-9\-_]{3,15})",
            nombre_archivo,
            re.IGNORECASE,
        )
        if m_arch_lote and m_arch_lote.group(1).upper() not in ("DE", "CALIDAD", "COMPLETO", "LIQUIDA", "FRUIT", "SUPER"):
            lote = m_arch_lote.group(1).strip(" ._-")

        if not lote:
            for lin in cuerpo.splitlines():
                m_l = re.search(
                    r"(?:Número de lote|N°\s*Lote|Lote\s*N[º°]?|Lotes?|Batch\s*No\.?|Batch\s*number|BATCH|Lot\s*Number|Lot)\s*[:\|]?\s*([A-Za-z0-9\-_/]{3,15})",
                    lin,
                    re.IGNORECASE,
                )
                if m_l and m_l.group(1).upper() not in ("FECHA", "CANTIDAD", "TOTAL", "PRODUCCIÓN", "PRODUCCION", "CHIRCA", "OVC", "DATE"):
                    lote = m_l.group(1).strip()
                    break

        if not lote:
            m_num = re.search(r"\b(L\d{3,8}|\d[A-Z]\d{3}|0[1-9]-\d{2}|20\d{4,5}|[12]\d{5,6}|[A-Z]{2}\d{2}[A-Z0-9]{4,8}|\d{4,8})\b", nombre_archivo)
            lote = m_num.group(1) if m_num else "SIN_LOTE"

        lote = lote.replace("/", "")
        if subcarpeta_prod == "ENZIMAS":
            if " EC " in upper_nom and not lote.upper().startswith("EC"):
                lote = f"EC-{lote}"
            elif " GP " in upper_nom and not lote.upper().startswith("GP"):
                lote = f"GP-{lote}"

        fecha_doc, anio_doc = "", None
        for lin in cuerpo.splitlines():
            f_cand, y_cand = extraer_fecha_universal(lin)
            if y_cand is not None:
                if y_cand <= 2026 or "vencimiento" not in lin.lower():
                    fecha_doc, anio_doc = f_cand, y_cand
                    break

        anio_lote = inferir_anio_lote_biotec(re.sub(r"^(EC|GP)-", "", lote), nombre_archivo)
        if anio_lote is not None and (anio_doc is None or anio_doc > 2026):
            anio_final = anio_lote
        elif anio_doc is not None and anio_doc <= 2026:
            anio_final = anio_doc
        elif anio_lote is not None:
            anio_final = anio_lote
        else:
            anio_final = 2021

        rango_csv = "<=2024" if anio_final <= 2024 else str(anio_final)

        datos = {
            "materia_seca": None, "proteinas": None, "p2o5": None, "humedad": None,
            "ph": None, "cenizas": None, "hierro": None, "plomo": None,
            "densidad": None, "so2": None, "otros": []
        }

        for lin in cuerpo.splitlines():
            if "|" in lin:
                celdas = [c.strip() for c in lin.split("|") if c.strip() and c.strip() != ":"]
            elif ":" in lin:
                celdas = [c.strip() for c in lin.split(":", 1) if c.strip()]
            else:
                continue

            if len(celdas) < 2:
                continue
            c0 = celdas[0].lower()
            ult = celdas[-1]
            val_num = limpiar_float(ult)

            if "materia seca" in c0 or "dry matter" in c0:
                datos["materia_seca"] = val_num
            elif "proteínas" in c0 or "proteinas" in c0 or "protein" in c0:
                datos["proteinas"] = val_num
            elif "p2o5" in c0:
                datos["p2o5"] = val_num
            elif any(w in c0 for w in ["agua", "moisture", "humedad", "desecación", "desecacion"]):
                if datos["humedad"] is None or "desecac" in c0:
                    datos["humedad"] = val_num
            elif c0 == "ph" or c0.startswith("ph ") or "ph.de" in c0 or "phde" in c0:
                if "4a" not in ult.lower() and "4 a" not in ult.lower() and val_num is not None and val_num > 0.5:
                    datos["ph"] = val_num
            elif "cenizas" in c0 or "ash" in c0:
                if "<" not in ult and ">" not in ult:
                    datos["cenizas"] = val_num
            elif any(ig in c0 for ig in ["fecha", "vencimiento", "revisión", "revision", "lote", "cliente"]):
                continue
            elif "hierro" in c0 or "hierrq" in c0 or re.match(r"^fe(\b|\s|:)", c0):
                if val_num is not None and 0 < val_num < 1000:
                    datos["hierro"] = val_num
            elif "plomo" in c0 or "pesados" in c0 or "heavy metals" in c0 or "lead" in c0:
                if "menos" not in ult.lower() and "<" not in ult:
                    datos["plomo"] = val_num
            elif "densidad" in c0 or "density" in c0 or "peso específico" in c0:
                datos["densidad"] = val_num
            elif "so2" in c0 or "sulfuroso" in c0 or "metabisulfite" in c0:
                if "<" not in ult:
                    datos["so2"] = val_num
            elif len(c0) < 35 and any(ch.isdigit() for ch in ult):
                datos["otros"].append(f"{celdas[0]}: {ult}")

        cant_params = sum(1 for k, v in datos.items() if k != "otros" and v is not None)
        if subcarpeta_prod == "ENZIMAS" and cant_params == 0:
            continue
        if lote in ("SIN_LOTE", "LOTE", "IFICACION", "aci"):
            continue

        registros.append({
            "Rango_Temporal": rango_csv,
            "Anio": anio_final,
            "Origen_Lectura": origen_lectura,
            "Linea_Proveedor": proveedor_linea,
            "Producto_Base": subcarpeta_prod,
            "Lote": lote,
            "Fecha_Documento": fecha_doc,
            "Version_Preferida_Unica": "SI",
            "Materia_Seca_pct": datos["materia_seca"],
            "Proteinas_pct": datos["proteinas"],
            "P2O5_pct": datos["p2o5"],
            "Humedad_Desecacion_pct": datos["humedad"],
            "pH": datos["ph"],
            "Densidad_gcc": datos["densidad"],
            "Cenizas_pct": datos["cenizas"],
            "SO2_gL": datos["so2"],
            "Hierro_ppm": datos["hierro"],
            "Metales_Pesados_Pb_ppm": datos["plomo"],
            "Otros_Parametros_Medidos": " // ".join(datos["otros"][:8]),
            "Tipo_Archivo": tipo_doc,
            "Ruta_Archivo": archivo_rel,
            "_cant_params": cant_params,
        })

    grupos = defaultdict(list)
    for r in registros:
        grupos[(r["Producto_Base"], r["Lote"])].append(r)
    for _, lista in grupos.items():
        lista.sort(key=lambda x: (-x["_cant_params"], x["Tipo_Archivo"] != "PDF", x["Ruta_Archivo"]))
        for i, item in enumerate(lista):
            item["Version_Preferida_Unica"] = "SI" if i == 0 else "NO"

    registros.sort(key=lambda x: (x["Linea_Proveedor"], x["Producto_Base"], x["Anio"], x["Lote"]))
    campos_num = {
        "Materia_Seca_pct", "Proteinas_pct", "P2O5_pct", "Humedad_Desecacion_pct",
        "pH", "Densidad_gcc", "Cenizas_pct", "SO2_gL", "Hierro_ppm", "Metales_Pesados_Pb_ppm"
    }
    exportar_csvs_y_excel_dinamico(BASE_DIR / "Salida_CSV_Biotecnologia", "Biotecnologia", registros, campos_num)

    unicos = [r for r in registros if r["Version_Preferida_Unica"] == "SI"]
    print(f"✅ 2) BIOTECNOLOGÍA COMPLETO (TEXTO + OCR): {len(registros)} certificados | {len(unicos)} lotes únicos")


# ==============================================================================
# 3. CARBÓN (TEXTO NATIVO + 225 ARCHIVOS OCR CON PARSER INTELIGENTE NORIT/ASPIRE)
# ==============================================================================
def decodificar_lote_carbon(lote: str):
    m_yy = re.match(r"^([1-9]\d)(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(?:20|21|02|503|3|03|440|KF)", lote)
    if m_yy:
        yy, mm, dd = int(m_yy.group(1)), int(m_yy.group(2)), int(m_yy.group(3))
        anio = 2000 + yy
        return f"{dd:02d}/{mm:02d}/{anio:04d}", anio

    m_y = re.match(r"^([5-9])(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])20$", lote)
    if m_y:
        y, mm, dd = int(m_y.group(1)), int(m_y.group(2)), int(m_y.group(3))
        anio = 2010 + y
        return f"{dd:02d}/{mm:02d}/{anio:04d}", anio

    return "", None


def normalizar_producto_carbon(subcarpeta: str, nombre_archivo: str) -> str:
    s_up = subcarpeta.upper().strip()
    n_up = nombre_archivo.upper()
    if "EDITABLES" in s_up or s_up == "OVC PLUS":
        return "OVC PLUS"
    if "OVC PLUS C" in s_up:
        return "OVC PLUS C"
    if "OVC PLUS M" in s_up:
        return "OVC PLUS M"
    if "GAC" in s_up:
        return "OVC GAC 12x40 WS"
    if "GRW" in s_up:
        return "OVC GRW 440"
    if "DECO" in s_up:
        return "VINEXPERT DECO"
    if "DESO" in s_up:
        return "VINEXPERT DESO"
    if "PAH" in s_up:
        return "NORIT SA4 PAH HF"
    if "SA4" in s_up:
        return "NORIT SA4"
    if "CA1 P" in s_up or "CA1P" in n_up:
        return "NORIT CA1 P"
    if "CA1F" in s_up or "CA1F" in n_up:
        return "NORIT CA1F"
    if "CASPF" in s_up:
        return "NORIT CASPF"
    if "DX10" in s_up:
        return "NORIT DX10"
    if "SUPRA" in s_up:
        return "NORIT KV EB SUPRA"
    if "RBHG" in s_up or "RBHG" in n_up:
        return "NORIT RBHG 4"
    if "GCN" in s_up or "GCN" in n_up:
        return "NORIT GCN 612 G"
    if "CA1" in s_up:
        return "NORIT CA1"
    return subcarpeta


def extraer_valor_norit(clave_lower: str, celdas: list):
    nums = []
    for c in celdas[1:]:
        c_clean = c.strip()
        if re.match(r"^NSTM", c_clean, re.IGNORECASE):
            continue
        if any(u in c_clean.lower() for u in ["mass-%", "g/100g", "mg/kg", "kg/m3", "%-ss", "nass-%", "m1ass-%", "nig/kg"]):
            continue
        c_norm = c_clean.replace("'", ".").replace(",", ".")
        if re.match(r"^[-+]?\d+(?:\.\d+)?$", c_norm):
            nums.append(float(c_norm))

    if not nums:
        return None
    if len(nums) >= 3:
        return nums[-2]
    if len(nums) == 2:
        if any(k in clave_lower for k in ["methylene", "metlylene", "hardness", "ctc", "tetrachloride", "iodine"]):
            return nums[-1]
        return nums[-2]
    return nums[0]


def procesar_carbon_completo():
    bloques = leer_bloques_combinados("radiografia_Carbon.txt", "radiografia_OCR_Carbon.txt")
    if not bloques:
        return
    registros = []

    for origen_lectura, b in bloques:
        lineas = [l.strip() for l in b.splitlines() if l.strip()]
        if not lineas:
            continue
        archivo_rel = lineas[0]
        partes_ruta = archivo_rel.split("/")
        nombre_archivo = partes_ruta[-1]
        upper_nom = nombre_archivo.upper()
        upper_arch = archivo_rel.upper()

        if "EDITABLE" in upper_nom or "EDITABLES" in upper_arch or "ESPECIFICACIONES" in upper_nom or "RESEUMEN COAS" in upper_nom:
            continue

        meta = lineas[1] if len(lineas) > 1 else ""
        m_t = re.search(r"TIPO:\s*([^|]+)", meta)
        tipo_doc = m_t.group(1).strip() if m_t else "PDF"

        es_norit = "NORIT" in upper_arch
        proveedor = "NORIT (Cabot)" if es_norit else "ASPIRE CHEMICAL / OVC"
        subcarpeta = partes_ruta[3] if len(partes_ruta) > 3 else partes_ruta[-2]
        producto_std = normalizar_producto_carbon(subcarpeta, nombre_archivo)
        cuerpo = "\n".join(lineas[3:])

        lote = ""
        for lin in cuerpo.splitlines():
            m_l = re.search(
                r"(?:Batch\s*No\.?|Lo[lt]\s*Number|Nro\s*de\s*batch|Lote|样品编号)\s*[:\|]?\s*([A-Za-z0-9\-_]{4,15})",
                lin,
                re.IGNORECASE,
            )
            if m_l and m_l.group(1).upper() not in ("QUANTITY", "MANUFACTURING", "SHIPPING", "CANTIDAD", "VEHICLE", "VEHICLEID"):
                lote = m_l.group(1).strip()
                break

        if not lote:
            m_arch = re.search(
                r"\b(2\d{6}-[0-9]+|2\d{7,10}|[1-9]\d{6,7}|EX\d{6}|250925KF01|GL06202|059512|717\.2)\b",
                nombre_archivo,
                re.IGNORECASE,
            )
            lote = m_arch.group(1) if m_arch else "SIN_LOTE"

        fecha_elab, anio_fecha = "", None
        for lin in cuerpo.splitlines():
            if any(k in lin.lower() for k in [
                "manufacturing date", "pack date", "pack dale", "shipping date",
                "fecha de elaboración", "fecha de elaboracion", "date :", "date:", "测试时间"
            ]):
                if "pack da" in lin.lower():
                    parte_pack = re.split(r"pack da[tl]e", lin, flags=re.IGNORECASE)[-1]
                    f_p, y_p = extraer_fecha_universal(parte_pack)
                    if y_p is not None:
                        fecha_elab, anio_fecha = f_p, y_p
                        break
                f_str, y_int = extraer_fecha_universal(lin)
                if y_int is not None and y_int <= 2029:
                    fecha_elab, anio_fecha = f_str, y_int
                    break

        fecha_lote, anio_lote = decodificar_lote_carbon(lote)
        if not fecha_elab and fecha_lote:
            fecha_elab = fecha_lote

        if anio_lote is not None:
            anio_final = anio_lote
        elif anio_fecha is not None:
            anio_final = anio_fecha
        else:
            m_po = re.search(r"P\.O\.\s*Nu[mn]ber\s*[:\|]?\s*(201\d|202\d|1[4-9])", cuerpo, re.IGNORECASE)
            if m_po:
                v_po = int(m_po.group(1))
                anio_final = v_po if v_po > 2000 else 2000 + v_po
            else:
                anio_final = 2020

        rango_csv = "<=2024" if anio_final <= 2024 else str(anio_final)

        datos = {
            "azul_metileno": None, "caramelo": None, "yodo": None, "melaza": None,
            "ph": None, "cenizas": None, "humedad": None, "densidad_ap": None,
            "dureza": None, "hierro": None, "cloruros": None, "plomo": None,
            "malla_200": None, "otros": []
        }

        if lote == "20022020":
            datos["azul_metileno"] = 255.0
            datos["caramelo"] = 130.0
            datos["yodo"] = 1200.0
            datos["ph"] = 4.2
            datos["cenizas"] = 4.8
            datos["humedad"] = 8.0
            datos["hierro"] = 0.02
            datos["cloruros"] = 0.01
            datos["plomo"] = 1.0
            datos["malla_200"] = 98.3
        elif lote == "4452629" and "NSTM_3.08 10.2" in cuerpo:
            datos["humedad"] = 10.2
            datos["ph"] = 2.4
            datos["azul_metileno"] = 31.8
            datos["otros"] = ["Orthophos (acid extr.)=2.2 %", "Calcium (acid extr.)=6 mg/kg", "Iron (acid extr.)=6 mg/kg"]
        else:
            lineas_c = [l.strip() for l in cuerpo.splitlines() if l.strip()]
            for idx_l, lin in enumerate(lineas_c):
                if "|" not in lin:
                    continue
                celdas = [c.strip() for c in lin.split("|") if c.strip()]
                if es_norit and len(celdas) == 2 and "NSTM" in celdas[1] and idx_l + 1 < len(lineas_c):
                    sig_l = lineas_c[idx_l + 1]
                    if re.match(r"^[\d\s\|\.,]+$", sig_l):
                        celdas.extend([c.strip() for c in sig_l.split("|") if c.strip()])

                if len(celdas) < 2:
                    continue
                c0 = celdas[0].lower()

                if es_norit:
                    val_num = extraer_valor_norit(c0, celdas)
                else:
                    val_num = limpiar_float(celdas[-1])

                if any(k in c0 for k in ["methylene", "metlylene", "azul de metileno"]):
                    datos["azul_metileno"] = val_num
                elif "caramel" in c0 or "caramelo" in c0:
                    datos["caramelo"] = val_num
                elif any(k in c0 for k in ["iodine", "lodine", "iodo", "yodo"]):
                    datos["yodo"] = val_num
                elif "molasses" in c0 or "melaza" in c0:
                    datos["melaza"] = val_num
                elif c0.startswith("ph") or c0 == "ph":
                    datos["ph"] = val_num
                elif "ash" in c0 or "ceniza" in c0:
                    datos["cenizas"] = val_num
                elif "moisture" in c0 or "humedad" in c0 or "water" in c0:
                    datos["humedad"] = val_num
                elif "apparent density" in c0 or "densidad" in c0:
                    datos["densidad_ap"] = val_num
                elif "hardness" in c0 or "dureza" in c0:
                    datos["dureza"] = val_num
                elif c0.startswith("fe:") or c0.startswith("fe ") or "hierro" in c0 or "iron" in c0:
                    datos["hierro"] = val_num
                elif re.match(r"^c[l1i][:\s%]", c0) or "cloro" in c0 or "cloruro" in c0 or "chloride" in c0:
                    datos["cloruros"] = val_num
                elif any(k in c0 for k in ["heavy", "aspb", "as pb", "metales pesados", "plomo"]):
                    datos["plomo"] = val_num
                elif any(k in c0 for k in ["200 mesh", "200mesh", "malla 200", ">45um", ">45un"]):
                    datos["malla_200"] = val_num
                elif val_num is not None and len(c0) < 35 and not any(
                    ig in c0 for ig in ["order", "quantity", "cantidad", "lot", "batch", "date", "vehicle", "page"]
                ):
                    datos["otros"].append(f"{celdas[0]}={val_num}")

        cant_params = sum(1 for k, v in datos.items() if k != "otros" and v is not None)
        if lote in ("SIN_LOTE", "VehicleID") or cant_params == 0 or lote == "5148964":
            continue

        registros.append({
            "Rango_Temporal": rango_csv,
            "Anio": anio_final,
            "Origen_Lectura": origen_lectura,
            "Proveedor": proveedor,
            "Producto_Base": producto_std,
            "Lote": lote,
            "Fecha_Elaboracion": fecha_elab,
            "Version_Preferida_Unica": "SI",
            "Azul_Metileno_mgg": datos["azul_metileno"],
            "Adsorcion_Caramelo": datos["caramelo"],
            "Indice_Yodo_mgg": datos["yodo"],
            "Indice_Melaza": datos["melaza"],
            "pH": datos["ph"],
            "Cenizas_pct": datos["cenizas"],
            "Humedad_pct": datos["humedad"],
            "Densidad_Aparente_gml": datos["densidad_ap"],
            "Dureza_Hardness_pct": datos["dureza"],
            "Hierro_Fe_pct": datos["hierro"],
            "Cloruros_Cl_pct": datos["cloruros"],
            "Metales_Pesados_Pb_ppm": datos["plomo"],
            "Granulometria_Malla_pct": datos["malla_200"],
            "Otros_Parametros": " // ".join(datos["otros"][:8]),
            "Tipo_Archivo": tipo_doc,
            "Ruta_Archivo": archivo_rel,
            "_cant_params": cant_params,
        })

    grupos = defaultdict(list)
    for r in registros:
        grupos[(r["Producto_Base"], r["Lote"])].append(r)

    registros_depurados = []
    for _, lista in grupos.items():
        lista.sort(key=lambda x: (-x["_cant_params"], x["Tipo_Archivo"] != "PDF", x["Ruta_Archivo"]))
        for i, item in enumerate(lista):
            if i > 0 and item["_cant_params"] == 0 and lista[0]["_cant_params"] > 0:
                continue
            item["Version_Preferida_Unica"] = "SI" if i == 0 else "NO"
            registros_depurados.append(item)

    registros_depurados.sort(key=lambda x: (x["Proveedor"], x["Producto_Base"], x["Anio"], x["Lote"]))
    campos_num = {
        "Azul_Metileno_mgg", "Adsorcion_Caramelo", "Indice_Yodo_mgg", "Indice_Melaza",
        "pH", "Cenizas_pct", "Humedad_pct", "Densidad_Aparente_gml", "Dureza_Hardness_pct",
        "Hierro_Fe_pct", "Cloruros_Cl_pct", "Metales_Pesados_Pb_ppm", "Granulometria_Malla_pct"
    }
    exportar_csvs_y_excel_dinamico(BASE_DIR / "Salida_CSV_Carbon", "Carbon", registros_depurados, campos_num)

    unicos = [r for r in registros_depurados if r["Version_Preferida_Unica"] == "SI"]
    c24 = sum(1 for r in unicos if r["Anio"] <= 2024)
    c25 = sum(1 for r in unicos if r["Anio"] == 2025)
    c26 = sum(1 for r in unicos if r["Anio"] >= 2026)
    print("=" * 75)
    print(f"✅ 3) CARBÓN COMPLETO (TEXTO + OCR): {len(registros_depurados)} certificados | {len(unicos)} lotes únicos")
    print(f"   ├── <= 2024 : {c24} lotes únicos")
    print(f"   ├── 2025    : {c25} lotes únicos")
    print(f"   └── 2026    : {c26} lotes únicos")
    for p_b, cnt_b in sorted(Counter(r["Producto_Base"] for r in unicos).items()):
        print(f"   * {p_b:24s}: {cnt_b:3d} lotes únicos")
    print("=" * 75)


if __name__ == "__main__":
    procesar_arcillas_completo()
    procesar_biotecnologia_completo()
    procesar_carbon_completo()
