#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ORQUESTADOR MAESTRO MULTIPLATAFORMA (WINDOWS / LINUX) - CHIRCA COAs
Ejecutar con: python EJECUTAR_SISTEMA.py
"""
import shutil
import subprocess
import sys
from pathlib import Path

ACTUAL_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = ACTUAL_DIR / "03_Sistema_y_Scripts" if (ACTUAL_DIR / "03_Sistema_y_Scripts").exists() else ACTUAL_DIR
ROOT_DIR = SCRIPTS_DIR.parent if SCRIPTS_DIR.name == "03_Sistema_y_Scripts" else SCRIPTS_DIR

pasos = [
    ("1/5 - Escaneando Certificados nuevos (SharePoint / Local)", "generar_radiografia.py"),
    ("2/5 - Procesando PDFs escaneados nuevos con OCR (Caché incremental)", "extraer_ocr_imagenes.py"),
    ("3/5 - Actualizando CSVs de Filtración", "procesar_filtracion.py"),
    ("4/5 - Actualizando CSVs de Arcillas, Biotecnología y Carbón", "integrar_ocr_y_csvs.py"),
    ("5/5a - Generando Informes Word/Excel de Filtración", "generar_informes_filtracion.py"),
    ("5/5b - Actualizando Informes de Arcillas, Biotecnología y Carbón", "generar_informes_resto.py"),
    ("5/5c - Consolidando Base de Datos Maestra Global", "consolidar_todo.py"),
]

for titulo, script in pasos:
    ruta_script = SCRIPTS_DIR / script
    if ruta_script.exists():
        print(f"\n🚀 {titulo} ({script})...")
        subprocess.run([sys.executable, str(ruta_script)], cwd=str(SCRIPTS_DIR), check=False)

# Mover carpetas de informes generadas hacia 02_Informes_de_Tendencia de forma multiplataforma
destino_informes = ROOT_DIR / "02_Informes_de_Tendencia"
destino_informes.mkdir(exist_ok=True)

for carpeta_inf in SCRIPTS_DIR.glob("Informes_Tendencia_*"):
    if carpeta_inf.is_dir():
        shutil.copytree(carpeta_inf, destino_informes / carpeta_inf.name, dirs_exist_ok=True)
        shutil.rmtree(carpeta_inf, ignore_errors=True)

print("\n✅ ¡ACTUALIZACIÓN COMPLETA FINALIZADA EN WINDOWS/LINUX!")
