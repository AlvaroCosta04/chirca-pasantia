#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ORQUESTADOR MAESTRO MULTIPLATAFORMA Y EJECUTABLE (.EXE) - CHIRCA COAs
"""
import os
import runpy
import shutil
import subprocess
import sys
from pathlib import Path

os.environ["PYTHONIOENCODING"] = "utf-8"

# Forzar inclusión de todas las librerías en PyInstaller
if getattr(sys, "frozen", False):
    import fitz
    import docx
    import openpyxl
    import pandas
    import numpy
    import matplotlib
    from rapidocr_onnxruntime import RapidOCR

# Modo hijo cuando corre como .exe compilado
if len(sys.argv) >= 3 and sys.argv[1] == "--ejecutar-modulo":
    ruta_mod = Path(sys.argv[2]).resolve()
    os.chdir(str(ruta_mod.parent))
    sys.path.insert(0, str(ruta_mod.parent))
    runpy.run_path(str(ruta_mod), run_name="__main__")
    sys.exit(0)


def main():
    print("=" * 75)
    print("  SISTEMA AUTOMÁTICO DE CERTIFICADOS E INFORMES (CHIRCA OVC S.A.)")
    print("=" * 75)

    # Detectar carpeta raíz (tanto en .py normal como en .exe compilado)
    if getattr(sys, "frozen", False):
        ACTUAL_DIR = Path(sys.executable).resolve().parent
    else:
        ACTUAL_DIR = Path(__file__).resolve().parent

    SCRIPTS_DIR = ACTUAL_DIR / "03_Sistema_y_Scripts" if (ACTUAL_DIR / "03_Sistema_y_Scripts").exists() else ACTUAL_DIR
    ROOT_DIR = SCRIPTS_DIR.parent if SCRIPTS_DIR.name == "03_Sistema_y_Scripts" else ACTUAL_DIR

    CANDIDATOS_SHAREPOINT = [
        Path("C:/Users/Pasantias/OVC SA/GAMAS - GAMAS"),
        Path.home() / "OVC SA" / "GAMAS - GAMAS",
        Path.home() / "Chirca OVC S.A" / "GAMAS - GAMAS",
    ]

    destino_sharepoint = None
    for sp in CANDIDATOS_SHAREPOINT:
        if sp.exists():
            destino_sharepoint = sp / "00. INFORMES DE TENDENCIAS (CONTROL DE CALIDAD)"
            break

    if destino_sharepoint is not None:
        destino_informes = destino_sharepoint
        print(f"☁️  Conectado a SharePoint: {destino_informes}")
    else:
        destino_informes = ROOT_DIR / "02_Informes_de_Tendencia"
        print(f"💻 Modo local: {destino_informes}")

    destino_informes.mkdir(parents=True, exist_ok=True)

    pasos = [
        ("1/5 - Escaneando Certificados nuevos (SharePoint / Local)", "generar_radiografia.py"),
        ("2/5 - Procesando PDFs escaneados nuevos con OCR", "extraer_ocr_imagenes.py"),
        ("3/5 - Actualizando CSVs de Filtración", "procesar_filtracion.py"),
        ("4/5 - Actualizando CSVs de Arcillas, Biotecnología y Carbón", "integrar_ocr_y_csvs.py"),
        ("5/5a - Generando Informes Word/Excel de Filtración", "generar_informes_filtracion.py"),
        ("5/5b - Actualizando Informes de Arcillas, Biotecnología y Carbón", "generar_informes_resto.py"),
        ("5/5c - Consolidando Base de Datos Maestra Global", "consolidar_todo.py"),
    ]

    for titulo, script in pasos:
        ruta_script = SCRIPTS_DIR / script
        if not ruta_script.exists():
            print(f"⚠️ No se encontró {ruta_script}, omitiendo...")
            continue

        print(f"\n🚀 {titulo}...")
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--ejecutar-modulo", str(ruta_script)]
        else:
            cmd = [sys.executable, str(ruta_script)]

        res = subprocess.run(cmd, cwd=str(SCRIPTS_DIR))
        if res.returncode != 0:
            print(f"\n❌ Error al ejecutar {script}.")
            if sys.platform.startswith("win"):
                input("\nPresione ENTER para salir...")
            sys.exit(res.returncode)

    print("\n📦 Guardando informes actualizados en destino...")
    for carpeta_inf in SCRIPTS_DIR.glob("Informes_Tendencia_*"):
        if carpeta_inf.is_dir():
            nombre_limpio = carpeta_inf.name.replace("Informes_Tendencia_", "Tendencias_")
            carpeta_final = destino_informes / nombre_limpio
            shutil.copytree(carpeta_inf, carpeta_final, dirs_exist_ok=True)
            shutil.rmtree(carpeta_inf, ignore_errors=True)
            print(f"  ✔ Actualizado: {carpeta_final.name}")

    print("\n" + "=" * 75)
    print("✅ ¡PROCESO COMPLETADO CON ÉXITO!")
    print(f"📂 Informes disponibles en:\n   {destino_informes}")
    print("=" * 75)

    # Dejar la ventana abierta en Windows si le hicieron doble clic al .exe
    if getattr(sys, "frozen", False) or sys.platform.startswith("win"):
        input("\nPresione ENTER para cerrar esta ventana...")


if __name__ == "__main__":
    main()
