#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import csv
from pathlib import Path
from collections import defaultdict, Counter
import openpyxl

BASE_DIR = Path(__file__).resolve().parent
TXT_PATH = BASE_DIR / "radiografia_FILTRACI_N.txt"
OUT_DIR = BASE_DIR / "Salida_CSV_Filtracion"
OUT_DIR.mkdir(exist_ok=True)


def limpiar_numero_float(val_str: str):
    """Convierte '2,44', '2.399', '19,01 %' a float real (o None si no es numérico)."""
    if not val_str:
        return None
    s = val_str.strip()
    if s.upper() in ("NA", "N/A", "-", "", "AUSENCIA", "BLANCO", "POLVO"):
        return None
    # Quitar números de nota al pie pegados con espacio (ej. '1.253 1')
    s = re.sub(r"^(\d+[,\.]\d+)\s+\d+$", r"\1", s)
    m = re.search(r"[-+]?\d+(?:[,\.]\d+)?", s)
    if not m:
        return None
    num_txt = m.group(0).replace(",", ".")
    try:
        return float(num_txt)
    except ValueError:
        return None


def formatear_decimal_csv(val_float):
    """Formatea float con coma decimal para Excel en español (o vacío si es None)."""
    if val_float is None:
        return ""
    if val_float == int(val_float):
        return str(int(val_float))
    return f"{val_float:.4f}".rstrip("0").rstrip(".").replace(".", ",")


def normalizar_producto(subcarpeta: str, tipo_pdf: str) -> str:
    """Estandariza el nombre del producto base a partir de la subcarpeta."""
    s = subcarpeta.upper().strip()
    mapa = {
        "PRECOAT 11 F": "PRECOAT 11F",
        "PRECOAT 12 H": "PRECOAT 12H",
        "PRECOAT 12F": "PRECOAT 12F",
        "PRECOAT 14 H": "PRECOAT 14H",
        "PRECOAT 14C": "PRECOAT 14C",
        "PRECOAT 15 H": "PRECOAT 15H",
        "PRECOAT 15F": "PRECOAT 15F",
        "PRECOAT 4F": "PRECOAT 4F",
        "PRECOAT DWG": "PRECOAT DWG",
        "PRECOAT DWG - BIG BAG": "PRECOAT DWG (BIG BAG)",
        "PRECOAT DWG-B": "PRECOAT DWG-B",
        "PRECOAT DWP3": "PRECOAT DWP3",
        "PRECOAT QSB": "PRECOAT QSB",
    }
    return mapa.get(s, s)


def extraer_lote(cuerpo: str, nombre_archivo: str):
    """Extrae el número de lote priorizando el texto interno y validando con el nombre de archivo."""
    lote_texto = ""
    for linea in cuerpo.splitlines():
        l_clean = linea.strip()
        if re.match(r"^LOTES\s*:", l_clean, re.IGNORECASE):
            continue
        m = re.search(
            r"Lote\s*(?:de\s*elaboraci[oó]n)?\s*[:\.\|]?\s*([0-9]{3,8}[A-Za-z]?)\b",
            l_clean,
            re.IGNORECASE,
        )
        if m:
            lote_texto = m.group(1).strip()
            break

    lote_archivo = ""
    m_arch_20 = re.search(r"(20[12]\d{3,5})", nombre_archivo)
    if m_arch_20:
        lote_archivo = m_arch_20.group(1)
    else:
        m_arch_gen = re.search(r"(?:Lote|LOT[E]?|OVC)\s*([0-9]{3,7})", nombre_archivo, re.IGNORECASE)
        if m_arch_gen:
            lote_archivo = m_arch_gen.group(1)

    if lote_texto:
        return lote_texto, "TEXTO_PDF"
    elif lote_archivo:
        return lote_archivo, "NOMBRE_ARCHIVO"
    return "SIN_LOTE", "NO_DETECTADO"


def extraer_fecha(cuerpo: str):
    """Extrae y normaliza la fecha de elaboración/emisión a DD/MM/YYYY y devuelve (fecha_str, anio_int)."""
    for linea in cuerpo.splitlines():
        if any(w in linea.lower() for w in ["fecha", "elaborac", "producc"]):
            m = re.search(r"(\d{1,2})\s*[/\-\.]\s*(\d{1,2})\s*[/\-\.]?\s*(\d{2,4})", linea)
            if m:
                d, mes, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if y < 100:
                    y += 2000
                if 2015 <= y <= 2030 and 1 <= mes <= 12 and 1 <= d <= 31:
                    return f"{d:02d}/{mes:02d}/{y:04d}", y
    return "", None


def detectar_cliente_y_version(cuerpo: str, nombre_archivo: str):
    """Detecta el cliente destinatario y clasifica el tipo de versión del archivo."""
    cliente_pdf = ""
    for linea in cuerpo.splitlines():
        if "cliente" in linea.lower():
            m = re.search(r"Cliente\s*:\s*([^|\n]+)", linea, re.IGNORECASE)
            if m:
                c_raw = m.group(1).strip()
                c_raw = re.sub(r"Pa[ií]s\s*:.*", "", c_raw, flags=re.IGNORECASE).strip()
                if c_raw and c_raw.upper() not in ("ARGENTINA", "PAÍS:"):
                    cliente_pdf = c_raw
                    break

    clientes_conocidos = [
        ("FECOVITA", ["FECOVITA", "FE.CO.VIT"]),
        ("ARCOR / INGRECOR", ["INGRECOR", "ARCOR"]),
        ("GRUPO PEÑAFLOR", ["PEÑAFLOR", "PENAFLOR"]),
        ("GLUCOVIL", ["GLUCOVIL"]),
        ("ENAV", ["ENAV"]),
        ("CEPAS ARGENTINAS", ["CEPAS"]),
        ("JVA / J Y V", ["JVA", "J Y V"]),
    ]

    cliente_final = cliente_pdf
    upper_arch = nombre_archivo.upper()
    for nombre_std, variantes in clientes_conocidos:
        if any(v in upper_arch or v in cliente_pdf.upper() for v in variantes):
            cliente_final = nombre_std
            break

    tags = []
    if re.search(r"OK\s*OK|OKK", upper_arch):
        tags.append("OK_DOBLE")
    elif re.search(r"\bOK\b|OK\.|\d+OK", upper_arch):
        tags.append("OK")
    if "CORREC" in upper_arch:
        tags.append("CORRECCION")
    if re.search(r"\bR\.PDF|\(R\)|\(REAL\)", upper_arch):
        tags.append("REANALISIS_R")
    if cliente_final:
        tags.append(f"CLIENTE_{cliente_final.split()[0]}")
    if nombre_archivo.lower().endswith((".docx", ".doc")):
        tags.append("WORD_DOC")

    version_str = " + ".join(tags) if tags else "ESTANDAR"

    prioridad = 10
    if "OK_DOBLE" in tags:
        prioridad = 60
    elif "CORRECCION" in tags:
        prioridad = 55
    elif "OK" in tags:
        prioridad = 50
    elif "REANALISIS_R" in tags:
        prioridad = 40
    elif not cliente_final:
        prioridad = 30
    else:
        prioridad = 20

    if nombre_archivo.lower().endswith(".pdf"):
        prioridad += 2

    return cliente_final, version_str, prioridad


def parsear_certificado(archivo_rel: str, tipo_doc: str, estado_doc: str, cuerpo: str):
    partes_ruta = archivo_rel.split("/")
    subcarpeta = partes_ruta[1] if len(partes_ruta) > 1 else "Raiz"
    nombre_archivo = partes_ruta[-1]

    m_tipo = re.search(r"Tipo\s+de\s+Filtrante\s*:\s*([^\n|]+)", cuerpo, re.IGNORECASE)
    tipo_pdf = m_tipo.group(1).strip() if m_tipo else ""
    producto_std = normalizar_producto(subcarpeta, tipo_pdf)

    planta = ""
    if "Anchorena" in cuerpo:
        planta = "Luján de Cuyo (Anchorena 1540)"
    elif "Lateral Sur" in cuerpo or "Gllen" in cuerpo:
        planta = "Guaymallén (Acc. Este Lat. Sur 2497)"

    m_val = re.search(r"Validez\s+del\s+producto\s*:\s*([^\n]+)", cuerpo, re.IGNORECASE)
    validez = m_val.group(1).strip() if m_val else ""

    lote, origen_lote = extraer_lote(cuerpo, nombre_archivo)
    fecha_str, anio_fecha = extraer_fecha(cuerpo)

    anio_lote = None
    m_ylote = re.match(r"^(20[12]\d)\d{2,3}$", lote)
    if m_ylote:
        y_cand = int(m_ylote.group(1))
        if 2018 <= y_cand <= 2026:
            anio_lote = y_cand

    if anio_lote is not None:
        anio_final = anio_lote
        origen_anio = "LOTE_YYYYWW" if anio_fecha is None else (
            "COINCIDE_LOTE_Y_FECHA" if anio_lote == anio_fecha else f"LOTE_{anio_lote}_(FECHA_DICE_{anio_fecha})"
        )
    elif anio_fecha is not None:
        anio_final = anio_fecha
        origen_anio = "FECHA_PDF"
    else:
        anio_final = 2019
        origen_anio = "HISTORICO_LOTE_CORTO"

    if anio_final <= 2024:
        rango_csv = "<=2024"
    elif anio_final == 2025:
        rango_csv = "2025"
    else:
        rango_csv = "2026"

    cliente, version_cert, prioridad = detectar_cliente_y_version(cuerpo, nombre_archivo)

    datos = {
        "color": "",
        "apariencia": "",
        "part_extranias": "",
        "part_metalicas": "",
        "permeabilidad": None,
        "densidad_humeda": None,
        "ret_malla_150": None,
        "ret_malla_80": None,
        "humedad": None,
        "impurezas": None,
        "oiv_calc_espec": "",
        "oiv_calc_res": "",
        "oiv_as_espec": "",
        "oiv_as_res": None,
        "oiv_fe_espec": "",
        "oiv_fe_res": None,
        "oiv_pb_espec": "",
        "oiv_pb_res": None,
        "notas_extra": [],
    }

    for linea in cuerpo.splitlines():
        l_strip = linea.strip()
        if not l_strip:
            continue

        if any(k in l_strip.upper() for k in ["DIACTIV", "FIBRA CELULOSA", "FABRICADO CON"]):
            datos["notas_extra"].append(l_strip)
            continue

        if "|" in l_strip:
            celdas = [c.strip() for c in l_strip.split("|") if c.strip() != ""]
            if len(celdas) >= 2:
                clave = celdas[0].lower()
                ult_val = celdas[-1]

                if clave == "color":
                    datos["color"] = ult_val
                elif clave == "apariencia":
                    datos["apariencia"] = ult_val
                elif "extrañas" in clave or "extranas" in clave:
                    datos["part_extranias"] = ult_val
                elif "metálicas" in clave or "metalicas" in clave:
                    datos["part_metalicas"] = ult_val
                elif "permeabilidad" in clave:
                    datos["permeabilidad"] = limpiar_numero_float(ult_val)
                elif "densidad" in clave:
                    datos["densidad_humeda"] = limpiar_numero_float(ult_val)
                elif "malla 150" in clave:
                    datos["ret_malla_150"] = limpiar_numero_float(ult_val)
                elif "malla 80" in clave:
                    datos["ret_malla_80"] = limpiar_numero_float(ult_val)
                elif clave.startswith("humedad"):
                    datos["humedad"] = limpiar_numero_float(ult_val)
                elif "impurezas" in clave:
                    datos["impurezas"] = limpiar_numero_float(ult_val)
                elif "calcinaci" in clave:
                    if len(celdas) >= 4:
                        datos["oiv_calc_espec"] = celdas[2]
                        datos["oiv_calc_res"] = celdas[3]
                    elif len(celdas) == 3:
                        datos["oiv_calc_espec"] = celdas[2]
                    elif len(celdas) == 2:
                        datos["oiv_calc_espec"] = celdas[1]
                elif "arsénico" in clave or "arsenico" in clave:
                    if len(celdas) >= 4:
                        datos["oiv_as_espec"] = celdas[2]
                        datos["oiv_as_res"] = limpiar_numero_float(celdas[3])
                    elif len(celdas) == 3:
                        datos["oiv_as_espec"] = celdas[2]
                elif clave.startswith("hierro"):
                    if len(celdas) >= 4:
                        datos["oiv_fe_espec"] = celdas[2]
                        datos["oiv_fe_res"] = limpiar_numero_float(celdas[3])
                    elif len(celdas) == 3:
                        datos["oiv_fe_espec"] = celdas[2]
                elif "metales pesados" in clave:
                    if len(celdas) >= 4:
                        datos["oiv_pb_espec"] = celdas[2]
                        datos["oiv_pb_res"] = limpiar_numero_float(celdas[3])
                    elif len(celdas) == 3:
                        datos["oiv_pb_espec"] = celdas[2]

        if "Perdida por Calcinación" in l_strip and "|" not in l_strip:
            m_as = re.search(r"Ars[eé]nico\s+ppm\s+([<\d]+)\s+([\d,\.]+)", l_strip)
            if m_as:
                datos["oiv_as_espec"] = m_as.group(1)
                datos["oiv_as_res"] = limpiar_numero_float(m_as.group(2))
            m_fe = re.search(r"Hierro\s+ppm\s+([<\d]+)\s+([\d,\.]+)", l_strip)
            if m_fe:
                datos["oiv_fe_espec"] = m_fe.group(1)
                datos["oiv_fe_res"] = limpiar_numero_float(m_fe.group(2))
            m_pb = re.search(r"Metales pesados \(Pb\)\s+ppm\s+([<\d]+)\s+([\d,\.]+)", l_strip)
            if m_pb:
                datos["oiv_pb_espec"] = m_pb.group(1)
                datos["oiv_pb_res"] = limpiar_numero_float(m_pb.group(2))

    return {
        "Rango_Temporal": rango_csv,
        "Anio": anio_final,
        "Origen_Anio": origen_anio,
        "Producto_Base": producto_std,
        "Subcarpeta": subcarpeta,
        "Tipo_Filtrante_PDF": tipo_pdf,
        "Lote": lote,
        "Origen_Lote": origen_lote,
        "Fecha_Emision": fecha_str,
        "Cliente": cliente,
        "Version_Certificado": version_cert,
        "_prioridad": prioridad,
        "Lote_Repetido": "NO",
        "Cant_Archivos_Lote": 1,
        "Version_Preferida_Unica": "SI",
        "Permeabilidad_Darcy": datos["permeabilidad"],
        "Densidad_Humeda_gcc": datos["densidad_humeda"],
        "Retencion_Malla_150_pct": datos["ret_malla_150"],
        "Retencion_Malla_80_pct": datos["ret_malla_80"],
        "Humedad_pct": datos["humedad"],
        "Impurezas_pct": datos["impurezas"],
        "Color": datos["color"],
        "Apariencia": datos["apariencia"],
        "Particulas_Extranias": datos["part_extranias"],
        "Particulas_Metalicas": datos["part_metalicas"],
        "OIV_Calcinacion_Espec": datos["oiv_calc_espec"],
        "OIV_Calcinacion_Res": datos["oiv_calc_res"],
        "OIV_Arsenico_ppm_Espec": datos["oiv_as_espec"],
        "OIV_Arsenico_ppm_Res": datos["oiv_as_res"],
        "OIV_Hierro_ppm_Espec": datos["oiv_fe_espec"],
        "OIV_Hierro_ppm_Res": datos["oiv_fe_res"],
        "OIV_Plomo_ppm_Espec": datos["oiv_pb_espec"],
        "OIV_Plomo_ppm_Res": datos["oiv_pb_res"],
        "Validez": validez,
        "Planta": planta,
        "Notas_Materia_Prima": " // ".join(datos["notas_extra"]),
        "Tipo_Archivo": tipo_doc,
        "Ruta_Archivo": archivo_rel,
    }


def main():
    txt = TXT_PATH.read_text(encoding="utf-8")
    bloques = txt.split("=" * 80 + "\nARCHIVO: ")[1:]

    registros = []
    excluidos = []

    for b in bloques:
        lineas = [l.strip() for l in b.splitlines() if l.strip()]
        archivo_rel = lineas[0]
        meta_linea = lineas[1] if len(lineas) > 1 else ""
        m_tipo = re.search(r"TIPO:\s*([^|]+)\|\s*ESTADO:\s*([^|]+)", meta_linea)
        tipo_doc = m_tipo.group(1).strip() if m_tipo else "DESCONOCIDO"
        estado_doc = m_tipo.group(2).strip() if m_tipo else "DESCONOCIDO"

        upper_arch = archivo_rel.upper()
        if any(k in upper_arch for k in ["TENDENCIA", "EDITABLE", "TABLA ", "DOCUMENTO.DOCX", "SOLO CAMBIAR"]):
            excluidos.append((archivo_rel, "PLANTILLA_O_INFORME_TENDENCIA"))
            continue
        if "ANALISIS OVC 4919 QSB LABORATORIO.DOC" in upper_arch:
            excluidos.append((archivo_rel, "DOC_ANTIGUO_OBJETO_INCRUSTADO"))
            continue

        cuerpo = "\n".join(lineas[3:])
        reg = parsear_certificado(archivo_rel, tipo_doc, estado_doc, cuerpo)
        registros.append(reg)

    grupos_lote = defaultdict(list)
    for r in registros:
        clave = (r["Producto_Base"], r["Lote"])
        grupos_lote[clave].append(r)

    for clave, lista in grupos_lote.items():
        cant = len(lista)
        if cant > 1:
            lista.sort(key=lambda x: (x["_prioridad"], x["Fecha_Emision"], x["Ruta_Archivo"]), reverse=True)
            for i, item in enumerate(lista):
                item["Lote_Repetido"] = "SI"
                item["Cant_Archivos_Lote"] = cant
                item["Version_Preferida_Unica"] = "SI" if i == 0 else "NO"

    registros.sort(key=lambda x: (x["Producto_Base"], x["Anio"], x["Lote"], x["Version_Preferida_Unica"] == "NO"))

    columnas = [c for c in registros[0].keys() if not c.startswith("_")]
    campos_numericos = {
        "Permeabilidad_Darcy",
        "Densidad_Humeda_gcc",
        "Retencion_Malla_150_pct",
        "Retencion_Malla_80_pct",
        "Humedad_pct",
        "Impurezas_pct",
        "OIV_Arsenico_ppm_Res",
        "OIV_Hierro_ppm_Res",
        "OIV_Plomo_ppm_Res",
    }

    def guardar_csv(ruta_csv: Path, filas):
        with open(ruta_csv, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(columnas)
            for r in filas:
                row_out = []
                for col in columnas:
                    val = r[col]
                    if col in campos_numericos:
                        row_out.append(formatear_decimal_csv(val))
                    else:
                        row_out.append("" if val is None else str(val))
                writer.writerow(row_out)

    filas_2024 = [r for r in registros if r["Rango_Temporal"] == "<=2024"]
    filas_2025 = [r for r in registros if r["Rango_Temporal"] == "2025"]
    filas_2026 = [r for r in registros if r["Rango_Temporal"] == "2026"]

    guardar_csv(OUT_DIR / "Filtracion_1_Hasta_2024.csv", filas_2024)
    guardar_csv(OUT_DIR / "Filtracion_2_Anio_2025.csv", filas_2025)
    guardar_csv(OUT_DIR / "Filtracion_3_Anio_2026.csv", filas_2026)
    guardar_csv(OUT_DIR / "Filtracion_0_TODOS.csv", registros)

    wb = openpyxl.Workbook()
    ws_todos = wb.active
    ws_todos.title = "Todos_Certificados"

    def llenar_hoja(ws, filas):
        ws.append(columnas)
        for r in filas:
            ws.append([r[col] for col in columnas])

    llenar_hoja(ws_todos, registros)
    llenar_hoja(wb.create_sheet("Hasta_2024"), filas_2024)
    llenar_hoja(wb.create_sheet("Anio_2025"), filas_2025)
    llenar_hoja(wb.create_sheet("Anio_2026"), filas_2026)
    llenar_hoja(
        wb.create_sheet("Sin_Duplicados_Preferidos"),
        [r for r in registros if r["Version_Preferida_Unica"] == "SI"],
    )

    ws_excl = wb.create_sheet("Archivos_Excluidos")
    ws_excl.append(["Ruta_Archivo", "Motivo_Exclusion"])
    for ruta_e, mot_e in excluidos:
        ws_excl.append([ruta_e, mot_e])

    wb.save(OUT_DIR / "Filtracion_Consolidado.xlsx")

    print("=" * 75)
    print("✅ EXTRACCIÓN Y AUDITORÍA DE FILTRACIÓN COMPLETADA")
    print("=" * 75)
    print(f"Total bloques leídos               : {len(bloques)}")
    print(f"Archivos excluidos (plantillas/inf): {len(excluidos)}")
    print(f"Certificados reales extraídos      : {len(registros)}")
    print(f"  ├── CSV <= 2024                  : {len(filas_2024)} certificados (Lotes únicos: {sum(1 for r in filas_2024 if r['Version_Preferida_Unica']=='SI')})")
    print(f"  ├── CSV 2025                     : {len(filas_2025)} certificados (Lotes únicos: {sum(1 for r in filas_2025 if r['Version_Preferida_Unica']=='SI')})")
    print(f"  └── CSV 2026                     : {len(filas_2026)} certificados (Lotes únicos: {sum(1 for r in filas_2026 if r['Version_Preferida_Unica']=='SI')})")
    print("-" * 75)
    print(f"{'PRODUCTO BASE':24s} | {'<=2024':>7s} | {'2025':>6s} | {'2026':>6s} | {'TOTAL':>6s} | {'SIN DARCY':>9s}")
    print("-" * 75)

    por_prod = defaultdict(lambda: Counter())
    sin_darcy = []
    discrepancias_anio = []

    for r in registros:
        p = r["Producto_Base"]
        por_prod[p][r["Rango_Temporal"]] += 1
        por_prod[p]["TOTAL"] += 1
        if r["Permeabilidad_Darcy"] is None:
            por_prod[p]["SIN_DARCY"] += 1
            sin_darcy.append(r["Ruta_Archivo"])
        if "FECHA_DICE" in r["Origen_Anio"]:
            discrepancias_anio.append(f"{r['Ruta_Archivo']} -> Lote:{r['Lote']} ({r['Origen_Anio']})")

    for p in sorted(por_prod.keys()):
        c = por_prod[p]
        print(f"{p:24s} | {c['<=2024']:7d} | {c['2025']:6d} | {c['2026']:6d} | {c['TOTAL']:6d} | {c['SIN_DARCY']:9d}")

    print("-" * 75)
    if sin_darcy:
        print(f"⚠️ Certificados sin valor de Permeabilidad ({len(sin_darcy)}):")
        for sd in sin_darcy[:15]:
            print(f"   - {sd}")
    else:
        print("🎯 ¡100% de los certificados extraídos tienen su valor de Permeabilidad (Darcy)!")

    if discrepancias_anio:
        print(f"\n📌 Casos donde el Año del Lote difiere del Año de la Fecha escrita ({len(discrepancias_anio)}):")
        for da in discrepancias_anio[:15]:
            print(f"   - {da}")

    print(f"\n📂 Archivos guardados en: {OUT_DIR}/")


if __name__ == "__main__":
    main()
