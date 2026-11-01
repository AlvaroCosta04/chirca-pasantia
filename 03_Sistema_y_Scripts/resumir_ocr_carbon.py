#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
from collections import defaultdict, Counter

txt = Path("radiografia_OCR_Carbon.txt").read_text(encoding="utf-8")
bloques = txt.split("=" * 80 + "\nARCHIVO: ")[1:]

por_sub = defaultdict(list)
terminos_frecuentes = defaultdict(Counter)

for b in bloques:
    lineas = [l.strip() for l in b.splitlines() if l.strip()]
    if not lineas:
        continue
    ruta = lineas[0]
    partes = ruta.split("/")
    sub = partes[3] if len(partes) > 3 else partes[-2]
    cuerpo = "\n".join(lineas[3:])
    por_sub[sub].append((ruta, cuerpo))

    for l in lineas[3:]:
        c0 = l.split("|")[0].split(":")[0].strip()
        if 1 < len(c0) < 38:
            terminos_frecuentes[sub][c0] += 1

out = [f"TOTAL CERTIFICADOS OCR EN CARBÓN: {len(bloques)}\n" + "=" * 75]
for sub, lista in sorted(por_sub.items()):
    top_t = [f"{k}({v})" for k, v in terminos_frecuentes[sub].most_common(12)]
    out.append(f"\n📂 SUBCARPETA: {sub} ({len(lista)} archivos)")
    out.append(f"   Etiquetas más frecuentes: {', '.join(top_t)}")
    for ruta_m, cuerpo_m in lista[:2]:
        out.append(f"\n   --- Muestra: {ruta_m} ---")
        out.append("\n".join("   " + l for l in cuerpo_m.splitlines()[:25]))

Path("resumen_ocr_carbon.txt").write_text("\n".join(out), encoding="utf-8")
print("✅ Archivo liviano generado: resumen_ocr_carbon.txt")
