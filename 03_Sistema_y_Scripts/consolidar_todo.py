#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import shutil
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ROOT_DIR = BASE_DIR.parent if BASE_DIR.name == "03_Sistema_y_Scripts" else BASE_DIR
CSV_MASTER_DIR = ROOT_DIR / "01_Bases_de_Datos_CSV"
CSV_MASTER_DIR.mkdir(exist_ok=True)

categorias = [
    ("FILTRACIÓN", "Filtracion", BASE_DIR / "Salida_CSV_Filtracion"),
    ("ARCILLAS", "Arcillas", BASE_DIR / "Salida_CSV_Arcillas"),
    ("BIOTECNOLOGÍA", "Biotecnologia", BASE_DIR / "Salida_CSV_Biotecnologia"),
    ("CARBÓN", "Carbon", BASE_DIR / "Salida_CSV_Carbon"),
]

dfs_globales = []

for nombre_cat, prefijo, carpeta_origen in categorias:
    archivo_todos = carpeta_origen / f"{prefijo}_0_TODOS.csv"
    if not archivo_todos.exists():
        print(f"⚠️ No se encontró {archivo_todos}")
        continue

    sub_destino = CSV_MASTER_DIR / f"Por_Categoria_{prefijo}"
    sub_destino.mkdir(exist_ok=True)
    for f_csv in sorted(carpeta_origen.glob("*")):
        if f_csv.suffix.lower() in (".csv", ".xlsx"):
            shutil.copy2(f_csv, sub_destino / f_csv.name)

    df_cat = pd.read_csv(archivo_todos, sep=";", encoding="utf-8-sig", dtype=str)
    df_cat.insert(0, "Categoria_General", nombre_cat)

    if "Fecha_Emision" not in df_cat.columns:
        if "Fecha_Documento" in df_cat.columns:
            df_cat["Fecha_Emision"] = df_cat["Fecha_Documento"]
        elif "Fecha_Elaboracion" in df_cat.columns:
            df_cat["Fecha_Emision"] = df_cat["Fecha_Elaboracion"]
        else:
            df_cat["Fecha_Emision"] = ""

    if "Origen_Lectura" not in df_cat.columns:
        df_cat["Origen_Lectura"] = "TEXTO_NATIVO"

    dfs_globales.append(df_cat)

df_global = pd.concat(dfs_globales, ignore_index=True)

cols_primarias = [
    "Categoria_General",
    "Rango_Temporal",
    "Anio",
    "Origen_Lectura",
    "Producto_Base",
    "Lote",
    "Fecha_Emision",
    "Version_Preferida_Unica",
]
cols_restantes = [c for c in df_global.columns if c not in cols_primarias and c != "Ruta_Archivo"]
orden_final_cols = [c for c in cols_primarias if c in df_global.columns] + cols_restantes + ["Ruta_Archivo"]
df_global = df_global[orden_final_cols]

df_global.to_csv(CSV_MASTER_DIR / "GLOBAL_0_TODOS_LOS_CERTIFICADOS.csv", sep=";", index=False, encoding="utf-8-sig")

anios_num = pd.to_numeric(df_global["Anio"], errors="coerce").fillna(2020).astype(int)
df_global[anios_num <= 2024].to_csv(CSV_MASTER_DIR / "GLOBAL_1_Hasta_2024.csv", sep=";", index=False, encoding="utf-8-sig")
df_global[anios_num == 2025].to_csv(CSV_MASTER_DIR / "GLOBAL_2_Anio_2025.csv", sep=";", index=False, encoding="utf-8-sig")
df_global[anios_num >= 2026].to_csv(CSV_MASTER_DIR / "GLOBAL_3_Anio_2026.csv", sep=";", index=False, encoding="utf-8-sig")

df_unicos = df_global[df_global["Version_Preferida_Unica"] == "SI"]
df_unicos.to_csv(CSV_MASTER_DIR / "GLOBAL_LOTES_UNICOS_SIN_DUPLICADOS.csv", sep=";", index=False, encoding="utf-8-sig")

for carpeta_inf in [
    BASE_DIR / "Informes_Tendencia_Arcillas",
    BASE_DIR / "Informes_Tendencia_Biotecnologia",
    BASE_DIR / "Informes_Tendencia_Carbon",
]:
    if not carpeta_inf.exists():
        continue
    for sub_prod in sorted(carpeta_inf.iterdir()):
        if not sub_prod.is_dir() or sub_prod.name.startswith("_"):
            continue
        xl_files = list(sub_prod.glob("Base_y_Estadistica_*.xlsx"))
        if xl_files:
            try:
                df_bd = pd.read_excel(xl_files[0], "Base_Datos_Unica")
                if len(df_bd) < 5 or "VARIOS PRODUCTOS" in sub_prod.name or "OBSOLETOS" in sub_prod.name:
                    shutil.rmtree(sub_prod)
            except Exception:
                pass

print("=" * 75)
print(f"✅ ¡BASE DE DATOS MAESTRA GLOBAL GENERADA EN {CSV_MASTER_DIR}!")
print("=" * 75)
