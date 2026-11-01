#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ORQUESTADOR MAESTRO DEL SISTEMA DE CALIDAD (CHIRCA PASANTÍA)
Activa en secuencia los scripts modulares guardados en 03_Sistema_y_Scripts/
y genera todos los resultados dentro de este directorio:
  - 01_Bases_de_Datos_CSV/
  - 02_Informes_de_Tendencia/

Uso:
  python3 EJECUTAR_SISTEMA.py             -> Regenera CSVs e Informes desde textos/OCR (~25 seg)
  python3 EJECUTAR_SISTEMA.py --completo  -> Escanea PDFs nuevos en Certificados/ y corre OCR incremental
"""

import re
import sys
import shutil
import subprocess
from datetime import datetime
from pathlib import Path
import pandas as pd

ROOT_DIR = Path(__file__).resolve().parent
SYS_DIR = ROOT_DIR / "03_Sistema_y_Scripts"
CSV_DIR = ROOT_DIR / "01_Bases_de_Datos_CSV"
INF_DIR = ROOT_DIR / "02_Informes_de_Tendencia"


def preparar_directorios():
    SYS_DIR.mkdir(exist_ok=True)
    CSV_DIR.mkdir(exist_ok=True)
    INF_DIR.mkdir(exist_ok=True)
    wa_dir = ROOT_DIR / "PARA_ENVIAR_WHATSAPP"
    if wa_dir.exists():
        shutil.rmtree(wa_dir, ignore_errors=True)


def aplicar_correcciones_en_scripts():
    # 1. Asegurar que generar_radiografia.py y extraer_ocr_imagenes.py apunten a ROOT_DIR / "Certificados"
    for nom_s in ["generar_radiografia.py", "extraer_ocr_imagenes.py"]:
        p_s = SYS_DIR / nom_s
        if p_s.exists():
            txt = p_s.read_text(encoding="utf-8")
            txt = txt.replace('BASE_DIR / "Certificados"', 'BASE_DIR.parent / "Certificados"')
            p_s.write_text(txt, encoding="utf-8")

    # 2. Asegurar que integrar_ocr_y_csvs.py aplique el filtro de ENZIMAS solo en Biotecnología (y no en Carbón)
    p_int = SYS_DIR / "integrar_ocr_y_csvs.py"
    if p_int.exists():
        txt_i = p_int.read_text(encoding="utf-8")

        # Limpiar el bloque si se había insertado también en Carbón en la corrida previa
        bloque_limpiar = (
            '\n        if subcarpeta_prod == "ENZIMAS" and cant_params == 0:\n'
            '            continue'
        )
        txt_i = txt_i.replace(bloque_limpiar, "")

        if 'if subcarpeta_prod == "ENZIMAS":' not in txt_i:
            reemplazo_lote = (
                'lote = lote.replace("/", "")\n'
                '        if subcarpeta_prod == "ENZIMAS":\n'
                '            if " EC " in upper_nom and not lote.upper().startswith("EC"):\n'
                '                lote = f"EC-{lote}"\n'
                '            elif " GP " in upper_nom and not lote.upper().startswith("GP"):\n'
                '                lote = f"GP-{lote}"'
            )
            txt_i = txt_i.replace('lote = lote.replace("/", "")', reemplazo_lote, 1)

        txt_i = txt_i.replace(
            "anio_lote = inferir_anio_lote_biotec(lote, nombre_archivo)",
            'anio_lote = inferir_anio_lote_biotec(re.sub(r"^(EC|GP)-", "", lote), nombre_archivo)'
        )

        viejo_fe = 'elif "hierro" in c0 or "hierrq" in c0 or c0.startswith("fe"):'
        nuevo_fe = (
            'elif any(ig in c0 for ig in ["fecha", "vencimiento", "revisión", "revision", "lote", "cliente"]):\n'
            '                continue\n'
            '            elif "hierro" in c0 or "hierrq" in c0 or re.match(r"^fe(\\b|\\s|:)", c0):'
        )
        txt_i = txt_i.replace(viejo_fe, nuevo_fe)

        txt_i = txt_i.replace(
            'if "menos" not in ult.lower() and "<" not in ult:\n                    datos["hierro"] = val_num',
            'if val_num is not None and 0 < val_num < 1000:\n                    datos["hierro"] = val_num'
        )
        txt_i = txt_i.replace(
            'if "4a" not in ult.lower() and "4 a" not in ult.lower():\n                    datos["ph"] = val_num',
            'if "4a" not in ult.lower() and "4 a" not in ult.lower() and val_num is not None and val_num > 0.5:\n                    datos["ph"] = val_num'
        )

        # Insertar únicamente en la 1ra aparición (función de Biotecnología)
        viejo_cant = 'cant_params = sum(1 for k, v in datos.items() if k != "otros" and v is not None)'
        nuevo_cant = (
            'cant_params = sum(1 for k, v in datos.items() if k != "otros" and v is not None)\n'
            '        if subcarpeta_prod == "ENZIMAS" and cant_params == 0:\n'
            '            continue'
        )
        txt_i = txt_i.replace(viejo_cant, nuevo_cant, 1)

        p_int.write_text(txt_i, encoding="utf-8")


def organizar_resultados():
    # 1. Mover Informes de Tendencia generados hacia 02_Informes_de_Tendencia/
    for carpeta_inf_local in SYS_DIR.glob("Informes_Tendencia_*"):
        if carpeta_inf_local.is_dir():
            dest_inf = INF_DIR / carpeta_inf_local.name
            if dest_inf.exists():
                shutil.rmtree(dest_inf)
            shutil.copytree(carpeta_inf_local, dest_inf)
            shutil.rmtree(carpeta_inf_local)

    # Limpiar subcarpetas con < 5 lotes en Arcillas, Biotecnología y Carbón
    for nombre_c in ["Informes_Tendencia_Arcillas", "Informes_Tendencia_Biotecnologia", "Informes_Tendencia_Carbon"]:
        carpeta_inf = INF_DIR / nombre_c
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

    # 2. Consolidar Bases de Datos CSV en 01_Bases_de_Datos_CSV/
    categorias = [
        ("FILTRACIÓN", "Filtracion", SYS_DIR / "Salida_CSV_Filtracion"),
        ("ARCILLAS", "Arcillas", SYS_DIR / "Salida_CSV_Arcillas"),
        ("BIOTECNOLOGÍA", "Biotecnologia", SYS_DIR / "Salida_CSV_Biotecnologia"),
        ("CARBÓN", "Carbon", SYS_DIR / "Salida_CSV_Carbon"),
    ]
    dfs_globales = []
    for nombre_cat, prefijo, carpeta_origen in categorias:
        archivo_todos = carpeta_origen / f"{prefijo}_0_TODOS.csv"
        if not archivo_todos.exists():
            continue
        sub_destino = CSV_DIR / f"Por_Categoria_{prefijo}"
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

    if dfs_globales:
        df_global = pd.concat(dfs_globales, ignore_index=True)
        cols_primarias = [
            "Categoria_General", "Rango_Temporal", "Anio", "Origen_Lectura",
            "Producto_Base", "Lote", "Fecha_Emision", "Version_Preferida_Unica",
        ]
        cols_restantes = [c for c in df_global.columns if c not in cols_primarias and c != "Ruta_Archivo"]
        orden_final_cols = [c for c in cols_primarias if c in df_global.columns] + cols_restantes + ["Ruta_Archivo"]
        df_global = df_global[orden_final_cols]

        df_global.to_csv(CSV_DIR / "GLOBAL_0_TODOS_LOS_CERTIFICADOS.csv", sep=";", index=False, encoding="utf-8-sig")
        anios_num = pd.to_numeric(df_global["Anio"], errors="coerce").fillna(2020).astype(int)
        df_global[anios_num <= 2024].to_csv(CSV_DIR / "GLOBAL_1_Hasta_2024.csv", sep=";", index=False, encoding="utf-8-sig")
        for idx_a, anio_p in enumerate(sorted(set(anios_num[anios_num >= 2025]).union({2025, 2026})), start=2):
            df_global[anios_num == anio_p].to_csv(CSV_DIR / f"GLOBAL_{idx_a}_Anio_{anio_p}.csv", sep=";", index=False, encoding="utf-8-sig")

        df_unicos = df_global[df_global["Version_Preferida_Unica"] == "SI"]
        df_unicos.to_csv(CSV_DIR / "GLOBAL_LOTES_UNICOS_SIN_DUPLICADOS.csv", sep=";", index=False, encoding="utf-8-sig")


def main():
    modo_completo = "--completo" in sys.argv
    print("=" * 75)
    print(f"🚀 EJECUTANDO SISTEMA EN: {ROOT_DIR}")
    print("=" * 75)

    preparar_directorios()
    aplicar_correcciones_en_scripts()

    pasos = []
    if modo_completo:
        pasos.extend([
            ("Escaneando Certificados/ (Texto nativo)", "generar_radiografia.py"),
            ("Procesando PDFs escaneados nuevos con OCR (Caché incremental)", "extraer_ocr_imagenes.py"),
        ])
    pasos.extend([
        ("Extrayendo CSVs de Filtración", "procesar_filtracion.py"),
        ("Extrayendo CSVs de Arcillas, Biotecnología y Carbón (Texto + OCR)", "integrar_ocr_y_csvs.py"),
        ("Generando Informes Word/Excel de Filtración", "generar_informes_filtracion.py"),
        ("Generando Informes Word/Excel de Arcillas, Biotecnología y Carbón", "generar_informes_resto.py"),
    ])

    for idx, (desc, script_name) in enumerate(pasos, start=1):
        script_path = SYS_DIR / script_name
        if script_path.exists():
            print(f"\n▶ Paso {idx}/{len(pasos)}: {desc} ({script_name})...")
            subprocess.run([sys.executable, str(script_path)], cwd=str(SYS_DIR), check=False)
        else:
            print(f"⚠️ No se encontró {script_path}")

    print("\n▶ Paso Final: Consolidando 01_Bases_de_Datos_CSV/ y 02_Informes_de_Tendencia/...")
    organizar_resultados()

    print("\n" + "=" * 75)
    print("✅ ¡PROCESO COMPLETADO! Archivos principales en 01_Bases_de_Datos_CSV/:")
    print("=" * 75)
    for f_csv in sorted(CSV_DIR.glob("GLOBAL_*.csv")):
        hora_mod = datetime.fromtimestamp(f_csv.stat().st_mtime).strftime("%d/%m/%Y %H:%M:%S")
        tam_kb = f_csv.stat().st_size / 1024.0
        print(f"  🕒 [{hora_mod}] ({tam_kb:6.1f} KB) -> {f_csv.name}")
    print("=" * 75)


if __name__ == "__main__":
    main()
