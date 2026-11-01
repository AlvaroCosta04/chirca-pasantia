#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
from pathlib import Path
import numpy as np
import pandas as pd
import docx
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import generar_informes_filtracion as gif

BASE_DIR = Path(__file__).resolve().parent

CATEGORIAS_CONFIG = [
    {
        "nombre": "ARCILLAS",
        "csv_dir": BASE_DIR / "Salida_CSV_Arcillas",
        "prefijo": "Arcillas",
        "out_dir": BASE_DIR / "Informes_Tendencia_Arcillas",
        "col_fecha": "Fecha_Emision",
        "parametros": [
            ("Humedad_Libre_pct", "Humedad Libre", "%", False),
            ("Retenido_M200_pct", "Retenido Malla 200", "%", False),
            ("Retenido_M325_pct", "Retenido Malla 325", "%", False),
            ("Acidez_Libre_mgKOH", "Acidez Libre", "mg KOH/g", False),
            ("Filtrabilidad_val", "Filtrabilidad", "-", False),
            ("Densidad_Aparente_gl", "Densidad Aparente", "g/l", False),
            ("Cantidad_T", "Cantidad Despachada", "T", True),
        ],
    },
    {
        "nombre": "BIOTECNOLOGÍA",
        "csv_dir": BASE_DIR / "Salida_CSV_Biotecnologia",
        "prefijo": "Biotecnologia",
        "out_dir": BASE_DIR / "Informes_Tendencia_Biotecnologia",
        "col_fecha": "Fecha_Documento",
        "parametros": [
            ("Materia_Seca_pct", "Materia Seca (M.S.)", "%", False),
            ("Proteinas_pct", "Proteínas (/M.S.)", "%", False),
            ("P2O5_pct", "P2O5 (/M.S.)", "%", False),
            ("Humedad_Desecacion_pct", "Humedad / Pérdida por Desecación", "%", False),
            ("pH", "pH", "u. pH", False),
            ("Densidad_gcc", "Densidad / Peso Específico", "g/cc", False),
            ("Cenizas_pct", "Cenizas", "%", False),
            ("SO2_gL", "Anhídrido Sulfuroso (SO2)", "g/L", False),
            ("Hierro_ppm", "Hierro (Fe)", "ppm", True),
            ("Metales_Pesados_Pb_ppm", "Metales Pesados (Pb)", "ppm", True),
        ],
    },
    {
        "nombre": "CARBÓN",
        "csv_dir": BASE_DIR / "Salida_CSV_Carbon",
        "prefijo": "Carbon",
        "out_dir": BASE_DIR / "Informes_Tendencia_Carbon",
        "col_fecha": "Fecha_Elaboracion",
        "parametros": [
            ("Azul_Metileno_mgg", "Adsorción Azul de Metileno", "mg/g", False),
            ("Adsorcion_Caramelo", "Adsorción de Caramelo", "-", False),
            ("Indice_Yodo_mgg", "Índice de Yodo", "mg/g", False),
            ("Indice_Melaza", "Índice de Melaza", "-", False),
            ("pH", "pH", "u. pH", False),
            ("Cenizas_pct", "Contenido de Cenizas", "%", False),
            ("Humedad_pct", "Humedad", "%", False),
            ("Densidad_Aparente_gml", "Densidad Aparente", "g/ml", False),
            ("Dureza_Hardness_pct", "Dureza (Hardness)", "%", False),
            ("Granulometria_Malla_pct", "Granulometría Malla 200", "%", False),
            ("Hierro_Fe_pct", "Hierro (Fe)", "%", True),
            ("Cloruros_Cl_pct", "Cloruros (Cl)", "%", True),
            ("Metales_Pesados_Pb_ppm", "Metales Pesados (Pb)", "ppm", True),
        ],
    },
]


def cargar_csv_categoria(ruta_csv: Path, col_fecha: str, parametros: list) -> pd.DataFrame:
    if not ruta_csv.exists():
        return pd.DataFrame()
    df = pd.read_csv(ruta_csv, sep=";", encoding="utf-8-sig")
    if len(df) == 0:
        return df
    if "Version_Preferida_Unica" in df.columns:
        df = df[df["Version_Preferida_Unica"] == "SI"].copy()

    if col_fecha in df.columns and "Fecha_Emision" not in df.columns:
        df["Fecha_Emision"] = df[col_fecha]
    elif "Fecha_Emision" not in df.columns:
        df["Fecha_Emision"] = ""

    if "Cliente" not in df.columns:
        if "Linea_Proveedor" in df.columns:
            df["Cliente"] = df["Linea_Proveedor"]
        elif "Proveedor" in df.columns:
            df["Cliente"] = df["Proveedor"]
        elif "Presentacion_Envase" in df.columns:
            df["Cliente"] = df["Presentacion_Envase"]
        else:
            df["Cliente"] = "-"

    for col, _, _, _ in parametros:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", ".", regex=False)
                .replace({"nan": np.nan, "None": np.nan, "": np.nan})
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df["Anio"] = pd.to_numeric(df["Anio"], errors="coerce").fillna(2021).astype(int)
    return gif.ordenar_df_cronologico(df)


def personalizar_banner_word(ruta_docx: Path, nombre_cat: str):
    doc = docx.Document(ruta_docx)
    if doc.tables:
        p0 = doc.tables[0].cell(0, 0).paragraphs[0]
        if p0.runs:
            p0.runs[0].text = f"DIVISIÓN {nombre_cat} — CONTROL DE CALIDAD (COAs)\n"
    doc.save(ruta_docx)


def actualizar_hoja_bd(ruta_xlsx: Path, df_hist: pd.DataFrame, df_2025: pd.DataFrame, parametros_cfg: list):
    wb = openpyxl.load_workbook(ruta_xlsx)
    ws_bd = wb["Base_Datos_Unica"]
    ws_bd.delete_rows(1, ws_bd.max_row)

    df_comb = gif.ordenar_df_cronologico(pd.concat([df_hist, df_2025], ignore_index=True))
    cols_base = ["Rango_Temporal", "Anio", "Origen_Lectura", "Producto_Base", "Lote", "Fecha_Emision", "Cliente"]
    cols_param = [c for c, _, _, _ in parametros_cfg if c in df_comb.columns and df_comb[c].notna().sum() > 0]
    cols_extra = [c for c in ["Otros_Parametros_Medidos", "Otros_Parametros", "Ruta_Archivo"] if c in df_comb.columns]
    cols_export = [c for c in cols_base if c in df_comb.columns] + cols_param + cols_extra

    ws_bd.append(cols_export)
    for _, row in df_comb.iterrows():
        ws_bd.append([None if pd.isna(row[c]) else row[c] for c in cols_export])

    fill_navy = PatternFill(start_color=gif.HEX_NAVY, end_color=gif.HEX_NAVY, fill_type="solid")
    fill_zebra = PatternFill(start_color=gif.HEX_LIGHT, end_color=gif.HEX_LIGHT, fill_type="solid")
    font_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style="thin", color="D0D7DE"),
        right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"),
        bottom=Side(style="thin", color="D0D7DE"),
    )

    for cell in ws_bd[1]:
        cell.fill = fill_navy
        cell.font = font_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx in range(2, ws_bd.max_row + 1):
        for c_idx in range(1, ws_bd.max_column + 1):
            c = ws_bd.cell(row=r_idx, column=c_idx)
            c.font = font_body
            c.border = thin_border
            if r_idx % 2 == 0:
                c.fill = fill_zebra

    for col_cells in ws_bd.columns:
        col_letter = col_cells[0].column_letter
        max_len = max(len(str(c.value or "")) for c in col_cells[:25])
        ws_bd.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)

    wb.save(ruta_xlsx)


def procesar_todas_categorias():
    total_word = 0
    total_excel = 0

    for cfg in CATEGORIAS_CONFIG:
        nombre_cat = cfg["nombre"]
        csv_dir = cfg["csv_dir"]
        prefijo = cfg["prefijo"]
        out_dir = cfg["out_dir"]
        charts_dir = out_dir / "_graficos_temp"
        out_dir.mkdir(exist_ok=True)
        charts_dir.mkdir(exist_ok=True)

        gif.PARAMETROS_CONFIG = cfg["parametros"]
        gif.CHARTS_DIR = charts_dir

        df_hist_all = cargar_csv_categoria(csv_dir / f"{prefijo}_1_Hasta_2024.csv", cfg["col_fecha"], cfg["parametros"])
        df_2025_all = cargar_csv_categoria(csv_dir / f"{prefijo}_2_Anio_2025.csv", cfg["col_fecha"], cfg["parametros"])

        productos = sorted(
            set(df_hist_all["Producto_Base"].dropna().unique() if len(df_hist_all) else [])
            .union(set(df_2025_all["Producto_Base"].dropna().unique() if len(df_2025_all) else []))
        )

        print("=" * 75)
        print(f"📊 GENERANDO INFORMES DE TENDENCIA: {nombre_cat} ({len(productos)} familias)")
        print("=" * 75)

        for prod in productos:
            if nombre_cat == "CARBÓN":
                if "NORIT" in str(prod).upper():
                    gif.PARAMETROS_CONFIG = [
                        ("Azul_Metileno_mgg", "Adsorción Azul de Metileno", "g/100g", False),
                        ("Adsorcion_Caramelo", "Adsorción de Caramelo", "-", False),
                        ("Indice_Yodo_mgg", "Índice de Yodo", "mg/g", False),
                        ("Indice_Melaza", "Índice de Melaza (EUR)", "-", False),
                        ("pH", "pH", "u. pH", False),
                        ("Cenizas_pct", "Contenido de Cenizas", "%", False),
                        ("Humedad_pct", "Humedad (as packed)", "%", False),
                        ("Densidad_Aparente_gml", "Densidad Aparente", "kg/m3", False),
                        ("Dureza_Hardness_pct", "Dureza (Ball-pan Hardness)", "%", False),
                        ("Granulometria_Malla_pct", "Granulometría >45 µm", "%", False),
                        ("Hierro_Fe_pct", "Hierro soluble en ácido (Fe)", "mg/kg (ppm)", True),
                        ("Cloruros_Cl_pct", "Cloruros (Cl)", "%", True),
                        ("Metales_Pesados_Pb_ppm", "Metales Pesados (Pb)", "ppm", True),
                    ]
                else:
                    gif.PARAMETROS_CONFIG = cfg["parametros"]
            prod_slug = re.sub(r"[^A-Za-z0-9]+", "_", str(prod)).strip("_")
            prod_dir = out_dir / prod_slug

            sub_hist = df_hist_all[df_hist_all["Producto_Base"] == prod].copy() if len(df_hist_all) else pd.DataFrame()
            sub_2025 = df_2025_all[df_2025_all["Producto_Base"] == prod].copy() if len(df_2025_all) else pd.DataFrame()

            cols_num = [c for c, _, _, _ in cfg["parametros"]]
            tiene_num_hist = len(sub_hist) > 0 and any(c in sub_hist.columns and sub_hist[c].notna().sum() > 0 for c in cols_num)
            tiene_num_2025 = len(sub_2025) > 0 and any(c in sub_2025.columns and sub_2025[c].notna().sum() > 0 for c in cols_num)

            if not (tiene_num_hist or tiene_num_2025):
                continue

            prod_dir.mkdir(exist_ok=True)

            if tiene_num_hist:
                ruta_doc_h = prod_dir / f"Informe_Tendencia_{prod_slug}_Historico_Hasta_2024.docx"
                gif.crear_informe_word(sub_hist, str(prod), "HISTORICO", ruta_doc_h)
                personalizar_banner_word(ruta_doc_h, nombre_cat)
                total_word += 1
                print(f"  ✔ [WORD HISTÓRICO <=2024] {str(prod):32s} -> {len(sub_hist):3d} lotes únicos")

            if tiene_num_2025:
                ruta_doc_25 = prod_dir / f"Informe_Tendencia_{prod_slug}_Anual_2025.docx"
                gif.crear_informe_word(sub_2025, str(prod), "ANUAL_2025", ruta_doc_25)
                personalizar_banner_word(ruta_doc_25, nombre_cat)
                total_word += 1
                print(f"  ✔ [WORD ANUAL 2025]       {str(prod):32s} -> {len(sub_2025):3d} lotes únicos")

            ruta_xl = prod_dir / f"Base_y_Estadistica_{prod_slug}.xlsx"
            gif.crear_excel_producto(sub_hist, sub_2025, str(prod), ruta_xl)
            actualizar_hoja_bd(ruta_xl, sub_hist, sub_2025, cfg["parametros"])
            total_excel += 1

    print("\n" + "=" * 75)
    print(f"✅ ¡PROCESO COMPLETADO! Se generaron {total_word} informes Word (.docx)")
    print(f"   y {total_excel} libros Excel (.xlsx) para Arcillas, Biotecnología y Carbón.")
    print("=" * 75)


if __name__ == "__main__":
    procesar_todas_categorias()
