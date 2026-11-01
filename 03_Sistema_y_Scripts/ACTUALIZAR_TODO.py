#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SCRIPT MAESTRO DE ACTUALIZACIÓN AUTOMÁTICA (CHIRCA - COAs)
Ejecutar con: python3 ACTUALIZAR_TODO.py
1. Escanea la carpeta Certificados/ buscando PDFs/Word/Excel nuevos.
2. Si hay PDFs escaneados nuevos (imagen), les pasa OCR usando el caché incremental.
3. Regenera todos los CSVs por categoría y los CSVs Maestros Globales (<=2024, 2025, 2026, 2027...).
4. Actualiza todos los Informes de Tendencia Word (.docx) y Excel (.xlsx) con gráficos.
"""
import subprocess
import sys
from pathlib import Path

pasos = [
    ("1/5 - Escaneando carpeta Certificados/ (Texto nativo)", "generar_radiografia.py"),
    ("2/5 - Procesando PDFs escaneados nuevos con OCR (Caché incremental)", "extraer_ocr_imagenes.py"),
    ("3/5 - Actualizando CSVs de Filtración", "procesar_filtracion.py"),
    ("4/5 - Actualizando CSVs de Arcillas, Biotecnología y Carbón (Texto + OCR)", "integrar_ocr_y_csvs.py"),
    ("5/5 - Generando Informes Word/Excel y CSV Maestro Global", "generar_informes_filtracion.py"),
    ("5/5b - Actualizando Informes del resto de categorías", "generar_informes_resto.py"),
    ("5/5c - Consolidando Base de Datos Maestra Global", "consolidar_todo.py"),
]

for titulo, script in pasos:
    if Path(script).exists():
        print(f"\n🚀 {titulo} ({script})...")
        subprocess.run([sys.executable, script], check=False)

# Sincronizar carpetas de informes
subprocess.run("mkdir -p 02_Informes_de_Tendencia && cp -rf Informes_Tendencia_* 02_Informes_de_Tendencia/ 2>/dev/null && rm -rf Informes_Tendencia_*", shell=True)
print("\n✅ ¡ACTUALIZACIÓN COMPLETA FINALIZADA!")
