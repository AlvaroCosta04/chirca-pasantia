#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ORQUESTADOR MAESTRO MULTIPLATAFORMA (WINDOWS / LINUX) - CHIRCA COAs
Ejecutar con: python EJECUTAR_SISTEMA.py
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path

# Configurar UTF-8 en la terminal para soportar caracteres en Windows
os.environ["PYTHONIOENCODING"] = "utf-8"

ACTUAL_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = ACTUAL_DIR / "03_Sistema_y_Scripts" if (ACTUAL_DIR / "03_Sistema_y_Scripts").exists() else ACTUAL_DIR
ROOT_DIR = SCRIPTS_DIR.parent if SCRIPTS_DIR.name == "03_Sistema_y_Scripts" else ACTUAL_DIR

# Posibles rutas de SharePoint en Windows
CANDIDATOS_SHAREPOINT = [
    Path("C:/Users/Pasantias/OVC SA/GAMAS - GAMAS"),
    Path.home() / "OVC SA" / "GAMAS - GAMAS",
    Path.home() / "Chirca OVC S.A" / "GAMAS - GAMAS",
]

# Detectar destino de informes: SharePoint (Windows) o Local (Linux)
destino_sharepoint = None
for sp in CANDIDATOS_SHAREPOINT:
    if sp.exists():
        destino_sharepoint = sp / "00. INFORMES DE TENDENCIAS (CONTROL DE CALIDAD)"
        break

if destino_sharepoint is not None:
    destino_informes = destino_sharepoint
    print(f"☁️ Destino de informes configurado en SharePoint: {destino_informes}")
else:
    destino_informes = ROOT_DIR / "02_Informes_de_Tendencia"
    print(f"💻 Destino de informes local: {destino_informes}")

destino_informes.mkdir(parents=True, exist_ok=True)

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
    if not ruta_script.exists():
        print(f"⚠️ Archivo no encontrado, omitiendo: {ruta_script}")
        continue

    print(f"\n🚀 {titulo} ({script})...")
    resultado = subprocess.run([sys.executable, str(ruta_script)], cwd=str(SCRIPTS_DIR))
    if resultado.returncode != 0:
        print(f"\n❌ ERROR crítico al ejecutar {script}. Proceso detenido.")
        sys.exit(resultado.returncode)

# Mover las carpetas generadas hacia el destino (SharePoint o Local)
print("\n📦 Sincronizando carpetas de informes hacia el destino...")
for carpeta_inf in SCRIPTS_DIR.glob("Informes_Tendencia_*"):
    if carpeta_inf.is_dir():
        nombre_limpio = carpeta_inf.name.replace("Informes_Tendencia_", "Tendencias_")
        carpeta_final = destino_informes / nombre_limpio
        shutil.copytree(carpeta_inf, carpeta_final, dirs_exist_ok=True)
        shutil.rmtree(carpeta_inf, ignore_errors=True)
        print(f"  ✔ Carpeta copiada: {carpeta_final.name}")

print(f"\n✅ ¡ACTUALIZACIÓN COMPLETA FINALIZADA!")
print(f"📂 Los informes se encuentran listos en: {destino_informes}")
