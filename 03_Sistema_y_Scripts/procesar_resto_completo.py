#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import csv
from pathlib import Path
from collections import defaultdict, Counter
import openpyxl

BASE_DIR = Path(__file__).resolve().parent

LINEAS_DIRECCION_ARCILLAS = {
    "PUREMIN",
    "MINERSA GROUP",
    "OVC S.A.",
    "OVC S.A. CQ",
    "ANCHORENA 1540",
    "00005507 LUJAN DE CUYO",
    "MENDOZA",
    "MENDOZA (ARGENTINA)",
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
    if s.upper() in ("NA", "N/A", "-", "", "AUSENCIA", "NEGATIVO", "CONFORM", "CONFORME", "CUMPLE", "PASS"):
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

    # 1. DD/MM/YYYY o DD.MM.YYYY
    m1 = re.search(r"\b(\d{1,2})\s*[\./\-]\s*(\d{1,2})\s*[\./\-]\s*(20[12]\d|\d{2})\b", texto)
    if m1:
        p1, p2, y = int(m1.group(1)), int(m1.group(2)), int(m1.group(3))
        if y < 100:
            y += 2000
        if p2 > 12 and p1 <= 12:
            d, mes = p2, p1
        else:
            d, mes = p1, p2
        if 2012 <= y <= 2035 and 1 <= mes <= 12 and 1 <= d <= 31:
            return f"{d:02d}/{mes:02d}/{y:04d}", y

    # 2. YYYY-MM-DD
    m2 = re.search(r"\b(20[12]\d)\s*[\./\-]\s*(\d{1,2})\s*[\./\-]\s*(\d{1,2})\b", texto)
    if m2:
        y, mes, d = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        if 2012 <= y <= 2035 and 1 <= mes <= 12 and 1 <= d <= 31:
            return f"{d:02d}/{mes:02d}/{y:04d}", y

    # 3. '08 de Febrero de 2019' o '27 Sep 2019' o '09Sep2019'
    m3 = re.search(
        r"\b(\d{1,2})\s*(?:de\s+)?([A-Za-z]{3,10})\.?\s*(?:de\s+|,\s*)?(20[12]\d)\b",
        texto,
        re.IGNORECASE,
    )
    if m3:
        d = int(m3.group(1))
        mes_txt = m3.group(2).lower()
        y = int(m3.group(3))
        mes = MESES_MAP.get(mes_txt[:3]) or MESES_MAP.get(mes_txt)
        if mes and 2012 <= y <= 2035 and 1 <= d <= 31:
            return f"{d:02d}/{mes:02d}/{y:04d}", y

    # 4. 'Jan. 18, 2021' o 'Jun. 27, 2025'
    m4 = re.search(
        r"\b([A-Za-z]{3,9})\.?\s+(\d{1,2})\s*,?\s*(20[12]\d)\b",
        texto,
        re.IGNORECASE,
    )
    if m4:
        mes_txt = m4.group(1).lower()
        d = int(m4.group(2))
        y = int(m4.group(3))
        mes = MESES_MAP.get(mes_txt[:3]) or MESES_MAP.get(mes_txt)
        if mes and 2012 <= y <= 2035 and 1 <= d <= 31:
            return f"{d:02d}/{mes:02d}/{y:04d}", y

    # 5. 'MM-YYYY' o 'MM/YYYY'
    m5 = re.search(r"\b(0?[1-9]|1[0-2])\s*[/\-]\s*(20[12]\d)\b", texto)
    if m5:
        mes, y = int(m5.group(1)), int(m5.group(2))
        return f"01/{mes:02d}/{y:04d}", y

    # 6. 'Octubre 2018' o 'Nov 2018'
    m6 = re.search(r"\b([A-Za-z]{3,10})\s+(20[12]\d)\b", texto, re.IGNORECASE)
    if m6:
        mes_txt = m6.group(1).lower()
        y = int(m6.group(2))
        mes = MESES_MAP.get(mes_txt[:3]) or MESES_MAP.get(mes_txt)
        if mes and 2012 <= y <= 2035:
            return f"01/{mes:02d}/{y:04d}", y

    return "", None


def exportar_csvs_y_excel(out_dir: Path, prefijo: str, registros: list, campos_num: set):
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

    filas_2024 = [r for r in registros if r["Rango_Temporal"] == "<=2024"]
    filas_2025 = [r for r in registros if r["Rango_Temporal"] == "2025"]
    filas_2026 = [r for r in registros if r["Rango_Temporal"] == "2026"]

    guardar_csv(out_dir / f"{prefijo}_1_Hasta_2024.csv", filas_2024)
    guardar_csv(out_dir / f"{prefijo}_2_Anio_2025.csv", filas_2025)
    guardar_csv(out_dir / f"{prefijo}_3_Anio_2026.csv", filas_2026)
    guardar_csv(out_dir / f"{prefijo}_0_TODOS.csv", registros)

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
    llenar_hoja(wb.create_sheet("Anio_2025"), filas_2025)
    llenar_hoja(wb.create_sheet("Anio_2026"), filas_2026)
    wb.save(out_dir / f"{prefijo}_Consolidado.xlsx")


# ==============================================================================
# 1. PROCESADOR DE ARCILLAS (CORREGIDO LOTE Y FAMILIAS DE PRODUCTO)
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


def procesar_arcillas():
    txt_path = BASE_DIR / "radiografia_Arcillas.txt"
    if not txt_path.exists():
        return
    txt = txt_path.read_text(encoding="utf-8")
    bloques = txt.split("=" * 80 + "\nARCHIVO: ")[1:]
    registros = []

    for b in bloques:
        lineas = [l.strip() for l in b.splitlines() if l.strip()]
        if not lineas:
            continue
        archivo_rel = lineas[0]
        partes_ruta = archivo_rel.split("/")
        subcarpeta_prov = partes_ruta[1] if len(partes_ruta) > 1 else "Puremin"
        nombre_archivo = partes_ruta[-1]
        nivel_anidamiento = len(partes_ruta)
        cuerpo = "\n".join(lineas[3:])

        paginas_raw = re.split(r"--- \[Página \d+/\d+\] ---", cuerpo)
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
                m_arch = re.search(r"\b(2[0-6]\d{6,8}|3071425|202\dW\d+)\b", nombre_archivo)
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
            m_yl = re.match(r"^(2[0-6])\d{6}$", lote)
            if m_yl:
                anio_lote = 2000 + int(m_yl.group(1))
            elif re.match(r"^(202\d)W", lote):
                anio_lote = int(lote[:4])

            anio_final = anio_fecha if anio_fecha is not None else (anio_lote if anio_lote is not None else 2023)
            rango_csv = "<=2024" if anio_final <= 2024 else ("2025" if anio_final == 2025 else "2026")
            prod_base, envase = clasificar_producto_arcilla(producto_raw, subcarpeta_prov, lote)

            if params["humedad_val"] is not None or lote != "SIN_LOTE":
                registros.append({
                    "Rango_Temporal": rango_csv,
                    "Anio": anio_final,
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
        grupos_envio[(r["Producto_Base"], r["Lote"], r["Albaran"], r["Contenedor_Transporte"])].append(r)
    for _, lista_e in grupos_envio.items():
        lista_e.sort(key=lambda x: (x["_nivel"], x["Ruta_Archivo"]))
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
        lista_l.sort(key=lambda x: (x["Albaran"], x["Ruta_Archivo"]))
        for i, item in enumerate(lista_l):
            item["Cant_Envios_Lote"] = cant
            item["Lote_Repetido"] = "SI" if cant > 1 else "NO"
            item["Version_Preferida_Unica"] = "SI" if i == 0 else "NO"

    registros.sort(key=lambda x: (x["Producto_Base"], x["Anio"], x["Lote"], x["Version_Preferida_Unica"] != "SI"))
    campos_num = {
        "Cantidad_T", "Humedad_Libre_pct", "Retenido_M200_pct",
        "Retenido_M325_pct", "Acidez_Libre_mgKOH", "Filtrabilidad_val", "Densidad_Aparente_gl"
    }
    exportar_csvs_y_excel(BASE_DIR / "Salida_CSV_Arcillas", "Arcillas", registros, campos_num)

    unicos = [r for r in registros if r["Version_Preferida_Unica"] == "SI"]
    envios = [r for r in registros if r["Envio_Unico_Albaran"] == "SI"]
    print("=" * 75)
    print("✅ 1) ARCILLAS RE-PROCESADO CON LOTES CORREGIDOS")
    print(f"   Páginas leídas: {len(registros)} | Camiones/Albaranes únicos: {len(envios)} | Lotes únicos reales: {len(unicos)}")
    for p_b, c_b in sorted(Counter(r["Producto_Base"] for r in unicos).items()):
        c24 = sum(1 for r in unicos if r["Producto_Base"] == p_b and r["Rango_Temporal"] == "<=2024")
        c25 = sum(1 for r in unicos if r["Producto_Base"] == p_b and r["Rango_Temporal"] == "2025")
        c26 = sum(1 for r in unicos if r["Producto_Base"] == p_b and r["Rango_Temporal"] == "2026")
        print(f"   - {p_b:28s}: {c_b:3d} lotes únicos (<=2024: {c24}, 2025: {c25}, 2026: {c26})")


# ==============================================================================
# 2. PROCESADOR DE BIOTECNOLOGÍA (FECHAS/LOTES MEJORADOS + EXCEL)
# ==============================================================================
def inferir_anio_lote_biotec(lote: str, nombre_archivo: str):
    m1 = re.match(r"^(20[12]\d)\d{2}$", lote)
    if m1:
        return int(m1.group(1))
    m2 = re.match(r"^\d{2}(201[6-9]|202[0-6])$", lote)
    if m2:
        return int(m2.group(1))
    m3 = re.match(r"^[A-Z]{2}(1[7-9]|2[0-6])[A-Z]", lote)
    if m3:
        return 2000 + int(m3.group(1))
    m4 = re.match(r"^(1[4-9]|2[0-6])\d{4,5}$", lote)
    if m4:
        return 2000 + int(m4.group(1))
    m5 = re.match(r"^(?:0[1-9]|1[0-2])-?(19|20|21|22|23|24|25|26)$", lote)
    if m5:
        return 2000 + int(m5.group(1))
    m6 = re.search(r"\b(201[5-9]|202[0-6])\b", nombre_archivo)
    if m6:
        return int(m6.group(1))
    return None


def procesar_biotecnologia():
    txt_path = BASE_DIR / "radiografia_Biotecnolog_a.txt"
    if not txt_path.exists():
        return
    txt = txt_path.read_text(encoding="utf-8")
    bloques = txt.split("=" * 80 + "\nARCHIVO: ")[1:]
    registros = []

    for b in bloques:
        lineas = [l.strip() for l in b.splitlines() if l.strip()]
        if not lineas:
            continue
        archivo_rel = lineas[0]
        upper_arch = archivo_rel.upper()

        if any(k in upper_arch for k in [
            "EDITABLE", "TEMPLATE", "BASE DATOS", "CÓDIGO DE ARTICULO",
            "EXTENSION", "EXTENSIÓN", "REVALIDA", "KOSHER", "KPSHER",
            "AUTOADHESIVOS", "RECLAMO", "ALERGENOS", "HDS", "~$"
        ]):
            continue

        meta = lineas[1] if len(lineas) > 1 else ""
        m_t = re.search(r"TIPO:\s*([^|]+)", meta)
        tipo_doc = m_t.group(1).strip() if m_t else "PDF"

        partes_ruta = archivo_rel.split("/")
        proveedor_linea = partes_ruta[1] if len(partes_ruta) > 1 else "Biotecnologia"
        subcarpeta_prod = partes_ruta[3] if len(partes_ruta) > 3 else (partes_ruta[2] if len(partes_ruta) > 2 else "")
        nombre_archivo = partes_ruta[-1]
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
                    r"(?:Número de lote|Lote\s*N[º°]?|Lote|Batch\s*No\.?|Lot\s*Number|Lot)\s*[:\|]?\s*([A-Za-z0-9\-_]{3,15})",
                    lin,
                    re.IGNORECASE,
                )
                if m_l and m_l.group(1).upper() not in ("FECHA", "CANTIDAD", "TOTAL", "PRODUCCIÓN", "PRODUCCION", "CHIRCA", "OVC"):
                    lote = m_l.group(1).strip()
                    break

        if not lote:
            m_num = re.search(r"\b(20[12]\d{3,4}|[12]\d{5,6}|[A-Z]{2}\d{2}[A-Z0-9]{4,8}|\d{4,8})\b", nombre_archivo)
            lote = m_num.group(1) if m_num else "SIN_LOTE"

        fecha_doc, anio_doc = "", None
        for lin in cuerpo.splitlines():
            f_cand, y_cand = extraer_fecha_universal(lin)
            if y_cand is not None:
                if y_cand <= 2026 or "vencimiento" not in lin.lower():
                    fecha_doc, anio_doc = f_cand, y_cand
                    break

        anio_lote = inferir_anio_lote_biotec(lote, nombre_archivo)
        if anio_lote is not None and (anio_doc is None or anio_doc > 2026):
            anio_final = anio_lote
        elif anio_doc is not None and anio_doc <= 2026:
            anio_final = anio_doc
        elif anio_lote is not None:
            anio_final = anio_lote
        else:
            anio_final = 2021

        rango_csv = "<=2024" if anio_final <= 2024 else ("2025" if anio_final == 2025 else "2026")

        datos = {
            "materia_seca": None, "proteinas": None, "p2o5": None, "humedad": None,
            "ph": None, "cenizas": None, "hierro": None, "plomo": None,
            "densidad": None, "so2": None, "otros": []
        }

        for lin in cuerpo.splitlines():
            if "|" not in lin:
                continue
            celdas = [c.strip() for c in lin.split("|") if c.strip() and c.strip() != ":"]
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
                datos["humedad"] = val_num
            elif c0 == "ph" or c0.startswith("ph "):
                datos["ph"] = val_num
            elif "cenizas" in c0 or "ash" in c0:
                datos["cenizas"] = val_num
            elif "hierro" in c0 or c0.startswith("fe"):
                datos["hierro"] = val_num
            elif "plomo" in c0 or "pesados" in c0 or "heavy metals" in c0 or "lead" in c0:
                datos["plomo"] = val_num
            elif "densidad" in c0 or "density" in c0 or "peso específico" in c0:
                datos["densidad"] = val_num
            elif "so2" in c0 or "anhídrido sulfuroso" in c0 or "sulfuroso" in c0:
                datos["so2"] = val_num
            elif len(c0) < 35 and any(ch.isdigit() for ch in ult):
                datos["otros"].append(f"{celdas[0]}: {ult}")

        registros.append({
            "Rango_Temporal": rango_csv,
            "Anio": anio_final,
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
        })

    grupos = defaultdict(list)
    for r in registros:
        grupos[(r["Producto_Base"], r["Lote"])].append(r)
    for _, lista in grupos.items():
        lista.sort(key=lambda x: (x["Tipo_Archivo"] != "PDF", x["Ruta_Archivo"]))
        for i, item in enumerate(lista):
            item["Version_Preferida_Unica"] = "SI" if i == 0 else "NO"

    registros.sort(key=lambda x: (x["Linea_Proveedor"], x["Producto_Base"], x["Anio"], x["Lote"]))
    campos_num = {
        "Materia_Seca_pct", "Proteinas_pct", "P2O5_pct", "Humedad_Desecacion_pct",
        "pH", "Densidad_gcc", "Cenizas_pct", "SO2_gL", "Hierro_ppm", "Metales_Pesados_Pb_ppm"
    }
    exportar_csvs_y_excel(BASE_DIR / "Salida_CSV_Biotecnologia", "Biotecnologia", registros, campos_num)

    unicos = [r for r in registros if r["Version_Preferida_Unica"] == "SI"]
    c24 = sum(1 for r in registros if r["Rango_Temporal"] == "<=2024")
    c25 = sum(1 for r in registros if r["Rango_Temporal"] == "2025")
    c26 = sum(1 for r in registros if r["Rango_Temporal"] == "2026")
    print("=" * 75)
    print("✅ 2) BIOTECNOLOGÍA RE-PROCESADO CON FECHAS Y LOTES MEJORADOS")
    print(f"   Certificados extraídos: {len(registros)} (Lotes únicos: {len(unicos)})")
    print(f"   ├── <= 2024 : {c24} certificados (Lotes únicos: {sum(1 for r in unicos if r['Rango_Temporal']=='<=2024')})")
    print(f"   ├── 2025    : {c25} certificados (Lotes únicos: {sum(1 for r in unicos if r['Rango_Temporal']=='2025')})")
    print(f"   └── 2026    : {c26} certificados (Lotes únicos: {sum(1 for r in unicos if r['Rango_Temporal']=='2026')})")


# ==============================================================================
# 3. PROCESADOR DE CARBÓN (ASPIRE CHEMICAL + CABOT NORIT)
# ==============================================================================
def inferir_anio_carbon(lote: str, nombre_archivo: str):
    m1 = re.match(r"^(19|20|21|22|23|24|25|26)\d{4,6}", lote)
    if m1:
        return 2000 + int(m1.group(1))
    m2 = re.match(r"^([5-9])\d{6}$", lote)
    if m2:
        return 2010 + int(m2.group(1))
    m3 = re.search(r"\b(201[4-9]|202[0-6])\b", nombre_archivo)
    if m3:
        return int(m3.group(1))
    return None


def procesar_carbon():
    txt_path = BASE_DIR / "radiografia_Carbon.txt"
    if not txt_path.exists():
        return
    txt = txt_path.read_text(encoding="utf-8")
    bloques = txt.split("=" * 80 + "\nARCHIVO: ")[1:]
    registros = []

    for b in bloques:
        lineas = [l.strip() for l in b.splitlines() if l.strip()]
        if not lineas:
            continue
        archivo_rel = lineas[0]
        upper_arch = archivo_rel.upper()

        if "EDITABLE" in upper_arch or "ESPECIFICACIONES" in upper_arch:
            continue

        meta = lineas[1] if len(lineas) > 1 else ""
        m_t = re.search(r"TIPO:\s*([^|]+)", meta)
        tipo_doc = m_t.group(1).strip() if m_t else "PDF"

        partes_ruta = archivo_rel.split("/")
        proveedor = "NORIT (Cabot)" if "NORIT" in upper_arch else "ASPIRE CHEMICAL / OVC"
        subcarpeta = partes_ruta[3] if len(partes_ruta) > 3 else partes_ruta[-2]
        nombre_archivo = partes_ruta[-1]
        cuerpo = "\n".join(lineas[3:])

        lote = ""
        for lin in cuerpo.splitlines():
            m_l = re.search(
                r"(?:Batch\s*No\.?|Lot\s*Number|Nro\s*de\s*batch|Lote|样品编号)\s*[:\|]?\s*([A-Za-z0-9\-_]{4,15})",
                lin,
                re.IGNORECASE,
            )
            if m_l and m_l.group(1).upper() not in ("QUANTITY", "MANUFACTURING", "SHIPPING", "CANTIDAD"):
                lote = m_l.group(1).strip()
                break

        if not lote:
            m_arch = re.search(r"\b(2[0-6]\d{6}|[456789]\d{6}|EX\d{6}|250925KF01)\b", nombre_archivo)
            lote = m_arch.group(1) if m_arch else "SIN_LOTE"

        fecha_elab, anio_fecha = "", None
        for lin in cuerpo.splitlines():
            if any(k in lin.lower() for k in [
                "manufacturing date", "pack date", "shipping date",
                "fecha de elaboración", "fecha de elaboracion", "date :", "测试时间"
            ]):
                f_str, y_int = extraer_fecha_universal(lin)
                if y_int is not None and y_int <= 2026:
                    fecha_elab, anio_fecha = f_str, y_int
                    break

        anio_lote = inferir_anio_carbon(lote, nombre_archivo)
        anio_final = anio_fecha if anio_fecha is not None else (anio_lote if anio_lote is not None else 2021)
        rango_csv = "<=2024" if anio_final <= 2024 else ("2025" if anio_final == 2025 else "2026")

        datos = {
            "azul_metileno": None, "caramelo": None, "yodo": None, "melaza": None,
            "ph": None, "cenizas": None, "humedad": None, "densidad_ap": None,
            "hierro": None, "cloruros": None, "plomo": None, "malla_200": None,
            "otros": []
        }

        for lin in cuerpo.splitlines():
            if "|" not in lin:
                continue
            celdas = [c.strip() for c in lin.split("|") if c.strip()]
            if len(celdas) < 2:
                continue
            c0 = celdas[0].lower()
            ult = celdas[-1]
            val_num = limpiar_float(ult)

            if "methylene blue" in c0 or "azul de metileno" in c0:
                datos["azul_metileno"] = val_num
            elif "caramel" in c0 or "caramelo" in c0:
                datos["caramelo"] = val_num
            elif "iodine" in c0 or "yodo" in c0:
                datos["yodo"] = val_num
            elif "molasses" in c0 or "melaza" in c0:
                datos["melaza"] = val_num
            elif "ph value" in c0 or c0 == "ph":
                datos["ph"] = val_num
            elif "ash" in c0 or "ceniza" in c0:
                datos["cenizas"] = val_num
            elif "moisture" in c0 or "humedad" in c0 or "water" in c0:
                datos["humedad"] = val_num
            elif "apparent density" in c0 or "densidad" in c0:
                datos["densidad_ap"] = val_num
            elif c0.startswith("fe:") or c0.startswith("fe ") or "hierro" in c0 or "iron" in c0:
                datos["hierro"] = val_num
            elif c0.startswith("cl:") or c0.startswith("cl ") or "cloruro" in c0 or "chloride" in c0:
                datos["cloruros"] = val_num
            elif "heavy metals" in c0 or "metales pesados" in c0 or "plomo" in c0:
                datos["plomo"] = val_num
            elif "200 mesh" in c0 or "malla 200" in c0 or "325 mesh" in c0:
                datos["malla_200"] = val_num
            elif len(c0) < 35 and val_num is not None and not any(k in c0 for k in ["order", "quantity", "lot", "batch"]):
                datos["otros"].append(f"{celdas[0]}={ult}")

        registros.append({
            "Rango_Temporal": rango_csv,
            "Anio": anio_final,
            "Proveedor": proveedor,
            "Producto_Base": subcarpeta,
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
            "Hierro_Fe_pct": datos["hierro"],
            "Cloruros_Cl_pct": datos["cloruros"],
            "Metales_Pesados_Pb_ppm": datos["plomo"],
            "Granulometria_Malla_pct": datos["malla_200"],
            "Otros_Parametros": " // ".join(datos["otros"][:8]),
            "Tipo_Archivo": tipo_doc,
            "Ruta_Archivo": archivo_rel,
        })

    grupos = defaultdict(list)
    for r in registros:
        grupos[(r["Producto_Base"], r["Lote"])].append(r)
    for _, lista in grupos.items():
        lista.sort(key=lambda x: (x["Tipo_Archivo"] != "PDF", x["Ruta_Archivo"]))
        for i, item in enumerate(lista):
            item["Version_Preferida_Unica"] = "SI" if i == 0 else "NO"

    registros.sort(key=lambda x: (x["Proveedor"], x["Producto_Base"], x["Anio"], x["Lote"]))
    campos_num = {
        "Azul_Metileno_mgg", "Adsorcion_Caramelo", "Indice_Yodo_mgg", "Indice_Melaza",
        "pH", "Cenizas_pct", "Humedad_pct", "Densidad_Aparente_gml",
        "Hierro_Fe_pct", "Cloruros_Cl_pct", "Metales_Pesados_Pb_ppm", "Granulometria_Malla_pct"
    }
    exportar_csvs_y_excel(BASE_DIR / "Salida_CSV_Carbon", "Carbon", registros, campos_num)

    unicos = [r for r in registros if r["Version_Preferida_Unica"] == "SI"]
    c24 = sum(1 for r in registros if r["Rango_Temporal"] == "<=2024")
    c25 = sum(1 for r in registros if r["Rango_Temporal"] == "2025")
    c26 = sum(1 for r in registros if r["Rango_Temporal"] == "2026")
    print("=" * 75)
    print("✅ 3) CARBÓN (DOCUMENTOS CON TEXTO NATIVO) PROCESADO")
    print(f"   Certificados extraídos: {len(registros)} (Lotes únicos: {len(unicos)})")
    print(f"   ├── <= 2024 : {c24} certificados (Lotes únicos: {sum(1 for r in unicos if r['Rango_Temporal']=='<=2024')})")
    print(f"   ├── 2025    : {c25} certificados (Lotes únicos: {sum(1 for r in unicos if r['Rango_Temporal']=='2025')})")
    print(f"   └── 2026    : {c26} certificados (Lotes únicos: {sum(1 for r in unicos if r['Rango_Temporal']=='2026')})")
    print("=" * 75)


if __name__ == "__main__":
    procesar_arcillas()
    procesar_biotecnologia()
    procesar_carbon()
