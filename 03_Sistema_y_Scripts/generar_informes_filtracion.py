#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage

# ==============================================================================
# CONFIGURACIÓN DE RUTAS Y PARÁMETROS
# ==============================================================================
BASE_DIR = Path(__file__).resolve().parent
CSV_DIR = BASE_DIR / "Salida_CSV_Filtracion"
OUT_DIR = BASE_DIR / "Informes_Tendencia_Filtracion"
CHARTS_DIR = OUT_DIR / "_graficos_temp"

OUT_DIR.mkdir(exist_ok=True)
CHARTS_DIR.mkdir(exist_ok=True)

PARAMETROS_CONFIG = [
    ("Permeabilidad_Darcy", "Permeabilidad", "Darcy", False),
    ("Densidad_Humeda_gcc", "Densidad Húmeda", "g/cc", False),
    ("Retencion_Malla_150_pct", "Retención Malla 150 Tyler", "%", False),
    ("Retencion_Malla_80_pct", "Retención Malla 80 Tyler", "%", False),
    ("Humedad_pct", "Humedad", "%", False),
    ("Impurezas_pct", "Impurezas", "%", False),
    ("OIV_Arsenico_ppm_Res", "Arsénico (OIV)", "ppm", True),
    ("OIV_Hierro_ppm_Res", "Hierro (OIV)", "ppm", True),
    ("OIV_Plomo_ppm_Res", "Metales Pesados - Pb (OIV)", "ppm", True),
]

HEX_NAVY = "1B365D"
HEX_STEEL = "2C5E8A"
HEX_LIGHT = "F0F4F8"
COLOR_NAVY = RGBColor(27, 54, 93)
COLOR_STEEL = RGBColor(44, 94, 138)
COLOR_GRAY = RGBColor(90, 90, 90)


# ==============================================================================
# ORDENAMIENTO CRONOLÓGICO INTELIGENTE DE LOTES
# ==============================================================================
def clave_orden_lote(row):
    """
    Ordena cronológicamente los lotes reconociendo que 2019171 es el sub-lote 1
    de la semana 17 de 2019 (2019-17-1) para que quede justo después de 201917.
    """
    anio = int(row.get("Anio", 2020))
    lote_str = str(row.get("Lote", "")).strip()

    # Caso estándar YYYYWW (6 dígitos, ej. 201917)
    m6 = re.match(r"^(20[12]\d)(\d{2})$", lote_str)
    if m6:
        return (int(m6.group(1)), int(m6.group(2)), 0, lote_str)

    # Caso sub-lote YYYYWWX (7 dígitos, ej. 2019171 -> 2019, semana 17, sub 1)
    m7 = re.match(r"^(20[12]\d)(\d{2})(\d)$", lote_str)
    if m7:
        return (int(m7.group(1)), int(m7.group(2)), int(m7.group(3)), lote_str)

    # Otros lotes cortos (ej. 646, 713) se ubican al inicio de su año asignado
    nums = re.findall(r"\d+", lote_str)
    num_val = int(nums[0]) if nums else 0
    return (anio, 0, num_val, lote_str)


def ordenar_df_cronologico(df: pd.DataFrame) -> pd.DataFrame:
    if len(df) == 0:
        return df
    df_c = df.copy()
    df_c["_sort_key"] = [clave_orden_lote(r) for _, r in df_c.iterrows()]
    df_c = df_c.sort_values("_sort_key").drop(columns=["_sort_key"]).reset_index(drop=True)
    return df_c


# ==============================================================================
# UTILIDADES DE ESTILO WORD (PYTHON-DOCX)
# ==============================================================================
def sombrear_celda(cell, hex_color: str):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def estilizar_tabla_word(table, col_widths=None):
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_cells = table.rows[0].cells
    for cell in hdr_cells:
        sombrear_celda(cell, HEX_NAVY)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(9)
                run.font.name = "Calibri"

    for r_idx, row in enumerate(table.rows[1:], start=1):
        bg_color = HEX_LIGHT if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, cell in enumerate(row.cells):
            sombrear_celda(cell, bg_color)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.size = Pt(8.5)
                    run.font.name = "Calibri"
                    if c_idx == 0:
                        run.font.bold = True

    if col_widths:
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                if idx < len(row.cells):
                    row.cells[idx].width = Inches(width)


def agregar_titulo_seccion(doc, texto: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(texto)
    run.font.name = "Calibri"
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = COLOR_NAVY
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="8" w:space="4" w:color="{HEX_STEEL}"/>'
        f'</w:pBdr>'
    )
    p._p.get_or_add_pPr().append(pBdr)


def fmt_num(val, decimales=3) -> str:
    if val is None or pd.isna(val):
        return "-"
    if abs(val - round(val)) < 1e-6 and abs(val) >= 10:
        return f"{int(round(val))}"
    return f"{val:.{decimales}f}".rstrip("0").rstrip(".").replace(".", ",")


# ==============================================================================
# GENERACIÓN DE GRÁFICOS SIN SUPERPOSICIÓN DE ETIQUETAS
# ==============================================================================
def anotar_outliers_sin_solapar(ax, out_indices, out_vals, out_lotes, media_global, total_puntos):
    """
    Coloca etiquetas a los valores atípicos alternando posiciones verticales y
    horizontales cuando hay dos puntos cercanos en el eje X.
    """
    if len(out_indices) == 0:
        return

    # Si hay más de 6 outliers, priorizar los 6 más extremos pero ordenados por X
    if len(out_indices) > 6:
        top_idx = np.argsort(np.abs(out_vals - media_global))[::-1][:6]
        out_indices = out_indices[top_idx]
        out_vals = out_vals[top_idx]
        out_lotes = out_lotes[top_idx]

    orden_x = np.argsort(out_indices)
    out_indices = out_indices[orden_x]
    out_vals = out_vals[orden_x]
    out_lotes = out_lotes[orden_x]

    dist_min_x = max(3, int(total_puntos * 0.09))
    offsets_arriba = [(0, 10), (-16, 22), (16, 22), (0, 32)]
    offsets_abajo = [(0, -18), (-16, -28), (16, -28), (0, -36)]

    prev_x = -999
    nivel_cluster = 0

    for xi, yi, lote_lbl in zip(out_indices, out_vals, out_lotes):
        if abs(xi - prev_x) <= dist_min_x:
            nivel_cluster += 1
        else:
            nivel_cluster = 0
        prev_x = xi

        if yi >= media_global:
            dx, dy = offsets_arriba[nivel_cluster % len(offsets_arriba)]
        else:
            dx, dy = offsets_abajo[nivel_cluster % len(offsets_abajo)]

        ax.annotate(
            f"L.{lote_lbl} ({fmt_num(yi)})",
            (xi, yi),
            textcoords="offset points",
            xytext=(dx, dy),
            ha="center",
            fontsize=7,
            fontweight="bold",
            color="#B02A2A",
            bbox=dict(boxstyle="round,pad=0.18", fc="#FFF8F8", ec="#E6B0AA", lw=0.6, alpha=0.92),
            arrowprops=dict(arrowstyle="-", color="#C0392B", lw=0.6, alpha=0.7),
            zorder=6,
        )


def generar_grafico_parametro(df_prod: pd.DataFrame, col: str, nombre_param: str, unidad: str,
                              producto: str, modo: str, ruta_img: Path) -> bool:
    sub = df_prod[["Anio", "Lote", "Fecha_Emision", col]].dropna(subset=[col]).copy()
    if len(sub) < 2:
        return False

    sub = ordenar_df_cronologico(sub)
    valores = sub[col].values
    media_global = np.mean(valores)
    std_global = np.std(valores, ddof=1) if len(valores) > 1 else 0.0

    plt.rcParams["font.sans-serif"] = "DejaVu Sans"
    plt.rcParams["axes.edgecolor"] = "#CCCCCC"
    plt.rcParams["axes.linewidth"] = 0.8

    anios_unicos = sorted(sub["Anio"].unique())

    if modo == "HISTORICO" and len(anios_unicos) >= 2:
        fig, (ax1, ax2) = plt.subplots(
            1, 2, figsize=(11.5, 4.3), dpi=220, gridspec_kw={"width_ratios": [1.05, 1.35]}
        )
        fig.patch.set_facecolor("white")

        # --- PANEL 1: Evolución Anual ---
        stats_anual = sub.groupby("Anio")[col].agg(["count", "mean", "std"]).reset_index()
        stats_anual["std"] = stats_anual["std"].fillna(0)

        x_anios = stats_anual["Anio"].astype(str).tolist()
        y_means = stats_anual["mean"].values
        y_stds = stats_anual["std"].values

        ax1.plot(x_anios, y_means, marker="o", color="#1B365D", linewidth=2.2,
                 markersize=6, label="Promedio Anual", zorder=4)
        if np.any(y_stds > 0):
            ax1.fill_between(x_anios, y_means - y_stds, y_means + y_stds,
                             color="#2C5E8A", alpha=0.18, label="Banda ±1 Desv. Est.")

        ax1.axhline(media_global, color="#7F8C8D", linestyle="--", linewidth=1.1,
                    label=f"Media Histórica ({fmt_num(media_global)})")

        for xa, ym, cnt in zip(x_anios, y_means, stats_anual["count"]):
            ax1.annotate(
                f"{fmt_num(ym)}\n(n={cnt})", (xa, ym),
                textcoords="offset points", xytext=(0, 8),
                ha="center", fontsize=7.5, fontweight="bold", color="#1B365D"
            )

        ax1.set_title(f"Evolución Anual del Promedio ({unidad})", fontsize=10,
                      fontweight="bold", color="#1B365D", pad=10)
        ax1.set_xlabel("Año de Elaboración", fontsize=8.5, fontweight="bold", color="#333333")
        ax1.set_ylabel(f"{nombre_param} ({unidad})", fontsize=8.5, fontweight="bold", color="#333333")
        ax1.margins(y=0.26)
        ax1.grid(True, linestyle=":", alpha=0.6, color="#BBBBBB")
        ax1.legend(loc="best", fontsize=7.5, frameon=True, facecolor="white", framealpha=0.9)
        ax1.tick_params(axis="both", labelsize=8)

        # --- PANEL 2: Serie Lote a Lote ---
        x_idx = np.arange(len(sub))
        ax2.plot(x_idx, valores, marker=".", color="#2C5E8A", linewidth=1.2,
                 markersize=4.5, alpha=0.85, label="Valor por Lote", zorder=3)
        ax2.axhline(media_global, color="#1B365D", linestyle="--", linewidth=1.2,
                    label=f"Promedio ({fmt_num(media_global)})")

        if std_global > 1e-9:
            lim_sup = media_global + 2 * std_global
            lim_inf = media_global - 2 * std_global
            ax2.axhline(lim_sup, color="#D9534F", linestyle=":", linewidth=1.0, label="Umbral ±2σ (Atípicos)")
            ax2.axhline(lim_inf, color="#D9534F", linestyle=":", linewidth=1.0)

            mask_out = np.abs(valores - media_global) > 2 * std_global
            if np.any(mask_out):
                out_indices = x_idx[mask_out]
                out_vals = valores[mask_out]
                out_lotes = sub.loc[mask_out, "Lote"].astype(str).values
                ax2.scatter(out_indices, out_vals, color="#D9534F", s=38, zorder=5, label="Valor Atípico")
                anotar_outliers_sin_solapar(ax2, out_indices, out_vals, out_lotes, media_global, len(sub))

        step = max(1, len(sub) // 10)
        ax2.set_xticks(x_idx[::step])
        ax2.set_xticklabels(sub["Lote"].astype(str).iloc[::step].tolist(), rotation=45, ha="right", fontsize=7.5)

        ax2.set_title(f"Trazabilidad Lote a Lote (n = {len(sub)} lotes)", fontsize=10,
                      fontweight="bold", color="#1B365D", pad=10)
        ax2.set_xlabel("Secuencia Cronológica de Lotes", fontsize=8.5, fontweight="bold", color="#333333")
        ax2.margins(y=0.30)
        ax2.grid(True, linestyle=":", alpha=0.6, color="#BBBBBB")
        ax2.legend(loc="best", fontsize=7.5, frameon=True, facecolor="white", framealpha=0.9)
        ax2.tick_params(axis="y", labelsize=8)

        fig.suptitle(f"{producto} — Análisis Histórico de {nombre_param} (Hasta Dic. 2024)",
                     fontsize=11.5, fontweight="bold", color="#1B365D", y=0.99)
        plt.tight_layout()
        fig.savefig(ruta_img, dpi=220, bbox_inches="tight")
        plt.close(fig)
        return True

    else:
        fig, ax = plt.subplots(figsize=(9.5, 4.1), dpi=220)
        fig.patch.set_facecolor("white")

        x_lotes = sub["Lote"].astype(str).tolist()
        x_idx = np.arange(len(sub))

        ax.plot(x_idx, valores, marker="o", color="#1B365D", linewidth=2.0,
                markersize=5.5, label=f"{nombre_param} ({unidad})", zorder=4)
        ax.axhline(media_global, color="#2C5E8A", linestyle="--", linewidth=1.3,
                   label=f"Promedio ({fmt_num(media_global)} {unidad})")

        if std_global > 1e-9:
            ax.fill_between(x_idx, media_global - std_global, media_global + std_global,
                            color="#2C5E8A", alpha=0.15, label="Banda Normal (±1σ)")
            lim_sup = media_global + 2 * std_global
            lim_inf = media_global - 2 * std_global
            ax.axhline(lim_sup, color="#D9534F", linestyle=":", linewidth=1.0, label="Umbral Alerta (±2σ)")
            ax.axhline(lim_inf, color="#D9534F", linestyle=":", linewidth=1.0)

            mask_out = np.abs(valores - media_global) > 2 * std_global
            if np.any(mask_out):
                ax.scatter(x_idx[mask_out], valores[mask_out], color="#D9534F", s=45, zorder=5, label="Valor Atípico")

        if len(sub) <= 25:
            for i_pt, (xi, val) in enumerate(zip(x_idx, valores)):
                dy_pt = 8 if i_pt % 2 == 0 else -14 if len(sub) > 16 else 8
                ax.annotate(fmt_num(val), (xi, val), textcoords="offset points",
                            xytext=(0, dy_pt), ha="center", fontsize=7.5, fontweight="bold", color="#1B365D")

        step = 1 if len(sub) <= 25 else max(1, len(sub) // 15)
        ax.set_xticks(x_idx[::step])
        ax.set_xticklabels(x_lotes[::step], rotation=45, ha="right", fontsize=8)

        titulo_periodo = "Desempeño Anual 2025" if modo == "ANUAL_2025" else f"Serie Histórica ({anios_unicos[0]})"
        ax.set_title(f"{producto} — {nombre_param} ({unidad}) | {titulo_periodo} (n = {len(sub)} lotes)",
                     fontsize=10.5, fontweight="bold", color="#1B365D", pad=12)
        ax.set_xlabel("Lote de Elaboración", fontsize=8.5, fontweight="bold", color="#333333")
        ax.set_ylabel(f"{nombre_param} ({unidad})", fontsize=8.5, fontweight="bold", color="#333333")
        ax.margins(y=0.28)
        ax.grid(True, linestyle=":", alpha=0.6, color="#BBBBBB")
        ax.legend(loc="best", fontsize=8, frameon=True, facecolor="white", framealpha=0.9)

        plt.tight_layout()
        fig.savefig(ruta_img, dpi=220, bbox_inches="tight")
        plt.close(fig)
        return True


# ==============================================================================
# REDACCIÓN AUTOMÁTICA DE CONCLUSIONES, ETAPAS Y DETECCIÓN DE ANOMALÍAS
# ==============================================================================
def redactar_conclusion_parametro(df_prod: pd.DataFrame, col: str, nombre_param: str,
                                  unidad: str, modo: str) -> dict:
    sub = df_prod[["Anio", "Lote", "Fecha_Emision", col]].dropna(subset=[col]).copy()
    sub = ordenar_df_cronologico(sub)
    n = len(sub)
    if n == 0:
        return None

    valores = sub[col].values
    v_min = float(np.min(valores))
    v_max = float(np.max(valores))
    v_mean = float(np.mean(valores))
    v_std = float(np.std(valores, ddof=1)) if n > 1 else 0.0
    cv_pct = (v_std / v_mean * 100.0) if v_mean != 0 else 0.0

    lote_min = str(sub.loc[sub[col].idxmin(), "Lote"])
    lote_max = str(sub.loc[sub[col].idxmax(), "Lote"])

    anios = sorted(sub["Anio"].unique())
    if modo == "HISTORICO" and len(anios) >= 2:
        por_anio = sub.groupby("Anio")[col].agg(["count", "mean", "std"]).reset_index()
        primer_anio, ult_anio = int(por_anio.iloc[0]["Anio"]), int(por_anio.iloc[-1]["Anio"])
        media_ini, media_fin = float(por_anio.iloc[0]["mean"]), float(por_anio.iloc[-1]["mean"])
        delta_pct = ((media_fin - media_ini) / media_ini * 100.0) if media_ini != 0 else 0.0

        if abs(delta_pct) < 5.0:
            sentido = (
                f"un comportamiento general estable entre puntas ({primer_anio}: {fmt_num(media_ini)} {unidad} "
                f"vs. {ult_anio}: {fmt_num(media_fin)} {unidad}; variación neta de {fmt_num(delta_pct, 1)}%)"
            )
        elif delta_pct > 0:
            sentido = (
                f"una tendencia creciente del {fmt_num(delta_pct, 1)}% entre el inicio de la serie en {primer_anio} "
                f"({fmt_num(media_ini)} {unidad}) y el cierre en {ult_anio} ({fmt_num(media_fin)} {unidad})"
            )
        else:
            sentido = (
                f"una reducción progresiva de nivel del {fmt_num(abs(delta_pct), 1)}% desde {primer_anio} "
                f"({fmt_num(media_ini)} {unidad}) hacia {ult_anio} ({fmt_num(media_fin)} {unidad})"
            )

        # Análisis de variación entre etapas (Primera mitad de años vs Segunda mitad)
        texto_etapas = ""
        if len(por_anio) >= 3:
            mitad = len(por_anio) // 2
            etapa1 = por_anio.iloc[:mitad]
            etapa2 = por_anio.iloc[mitad:]
            m_e1 = float(etapa1["mean"].mean())
            m_e2 = float(etapa2["mean"].mean())
            rango_e1 = f"{int(etapa1.iloc[0]['Anio'])}–{int(etapa1.iloc[-1]['Anio'])}" if len(etapa1) > 1 else f"{int(etapa1.iloc[0]['Anio'])}"
            rango_e2 = f"{int(etapa2.iloc[0]['Anio'])}–{int(etapa2.iloc[-1]['Anio'])}" if len(etapa2) > 1 else f"{int(etapa2.iloc[0]['Anio'])}"
            cambio_etapas_pct = ((m_e2 - m_e1) / m_e1 * 100.0) if m_e1 != 0 else 0.0
            if abs(cambio_etapas_pct) >= 8.0:
                dir_etapa = "superior" if cambio_etapas_pct > 0 else "inferior"
                texto_etapas = (
                    f" Al comparar por etapas operativas, se distingue un cambio de régimen entre el período "
                    f"{rango_e1} (media de etapa: {fmt_num(m_e1)} {unidad}) y el período {rango_e2} "
                    f"(media de etapa: {fmt_num(m_e2)} {unidad}, un {fmt_num(abs(cambio_etapas_pct), 1)}% {dir_etapa})."
                )
            else:
                texto_etapas = (
                    f" Comparando la etapa inicial ({rango_e1}: media {fmt_num(m_e1)} {unidad}) frente a la etapa "
                    f"reciente ({rango_e2}: media {fmt_num(m_e2)} {unidad}), el promedio se mantiene consistente sin saltos abruptos de nivel."
                )

        resumen_anios = ", ".join(
            [f"{int(r['Anio'])}: {fmt_num(r['mean'])} {unidad} (n={int(r['count'])})" for _, r in por_anio.iterrows()]
        )
        texto_tendencia = (
            f"Sobre un total de {n} lotes únicos evaluados ({primer_anio}–{ult_anio}), el parámetro {nombre_param} "
            f"registró una media histórica de {fmt_num(v_mean)} {unidad} (±{fmt_num(v_std)} {unidad}; CV = {fmt_num(cv_pct, 1)}%), "
            f"con mínimo de {fmt_num(v_min)} {unidad} (Lote {lote_min}) y máximo de {fmt_num(v_max)} {unidad} (Lote {lote_max}). "
            f"La trayectoria interanual muestra {sentido}.{texto_etapas} Promedios anuales: {resumen_anios}."
        )
    else:
        estabilidad = (
            "muy alta consistencia operativa" if cv_pct < 5.0
            else ("dispersión moderada dentro de parámetros normales" if cv_pct < 15.0 else "alta variabilidad entre lotes")
        )
        texto_tendencia = (
            f"En el período evaluado ({n} lotes únicos), {nombre_param} presentó un valor promedio de "
            f"{fmt_num(v_mean)} {unidad} con una desviación estándar de {fmt_num(v_std)} {unidad} (CV = {fmt_num(cv_pct, 1)}%, "
            f"indicando {estabilidad}). El mínimo fue de {fmt_num(v_min)} {unidad} (Lote {lote_min}) "
            f"y el máximo alcanzó {fmt_num(v_max)} {unidad} (Lote {lote_max})."
        )

    outliers_lista = []
    if n >= 4 and v_std > 1e-9:
        umbral_sup = v_mean + 2 * v_std
        umbral_inf = v_mean - 2 * v_std
        df_out = sub[(sub[col] > umbral_sup) | (sub[col] < umbral_inf)]
        for _, row_o in df_out.iterrows():
            val_o = float(row_o[col])
            z_score = (val_o - v_mean) / v_std
            tipo_desv = "por encima" if val_o > v_mean else "por debajo"
            nota_extra = ""
            if v_mean > 0 and val_o > v_mean * 6:
                nota_extra = " [Posible error de tipeo/coma decimal en el certificado original]"
            outliers_lista.append(
                f"Lote {row_o['Lote']} (Año {int(row_o['Anio'])}): {fmt_num(val_o)} {unidad} "
                f"({fmt_num(abs(z_score), 1)}σ {tipo_desv} de la media de {fmt_num(v_mean)} {unidad}){nota_extra}"
            )

    if outliers_lista:
        texto_anomalias = (
            f"Se identificaron {len(outliers_lista)} lotes con valores estadísticamente atípicos "
            f"(fuera de la banda de normalidad μ ± 2σ = [{fmt_num(v_mean - 2*v_std)} ; {fmt_num(v_mean + 2*v_std)}] {unidad}): "
            + "; ".join(outliers_lista) + "."
        )
    else:
        texto_anomalias = (
            "No se registraron anomalías ni valores atípicos fuera del umbral estadístico (±2σ); "
            "el 100% de los lotes se mantuvo dentro de la banda normal de proceso."
        )

    return {
        "col": col,
        "nombre": nombre_param,
        "unidad": unidad,
        "n": n,
        "min": v_min,
        "lote_min": lote_min,
        "max": v_max,
        "lote_max": lote_max,
        "mean": v_mean,
        "std": v_std,
        "cv": cv_pct,
        "texto_tendencia": texto_tendencia,
        "texto_anomalias": texto_anomalias,
        "cant_outliers": len(outliers_lista),
    }


# ==============================================================================
# CONSTRUCCIÓN DEL DOCUMENTO WORD (.DOCX)
# ==============================================================================
def crear_informe_word(df_prod: pd.DataFrame, producto: str, modo: str, ruta_docx: Path):
    doc = Document()
    df_prod = ordenar_df_cronologico(df_prod)

    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)

    anios = sorted(df_prod["Anio"].unique())
    rango_anios_str = f"{anios[0]} – {anios[-1]}" if len(anios) > 1 else f"{anios[0]}"
    titulo_alcance = (
        f"INFORME HISTÓRICO DE TENDENCIAS DE CALIDAD ({rango_anios_str})"
        if modo == "HISTORICO"
        else "INFORME ANUAL DE TENDENCIAS DE CALIDAD (AÑO 2025)"
    )

    tbl_banner = doc.add_table(rows=1, cols=1)
    tbl_banner.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_b = tbl_banner.cell(0, 0)
    sombrear_celda(cell_b, HEX_NAVY)

    p_b1 = cell_b.paragraphs[0]
    p_b1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_b1.paragraph_format.space_before = Pt(6)
    p_b1.paragraph_format.space_after = Pt(2)
    r_b1 = p_b1.add_run("DIVISIÓN FILTRACIÓN — CONTROL DE CALIDAD (COAs)\n")
    r_b1.font.name = "Calibri"
    r_b1.font.size = Pt(9.5)
    r_b1.font.bold = True
    r_b1.font.color.rgb = RGBColor(190, 215, 240)

    r_b2 = p_b1.add_run(f"{producto}\n")
    r_b2.font.name = "Calibri"
    r_b2.font.size = Pt(18)
    r_b2.font.bold = True
    r_b2.font.color.rgb = RGBColor(255, 255, 255)

    r_b3 = p_b1.add_run(titulo_alcance)
    r_b3.font.name = "Calibri"
    r_b3.font.size = Pt(11.5)
    r_b3.font.bold = True
    r_b3.font.color.rgb = RGBColor(255, 255, 255)

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(6)
    p_meta.paragraph_format.space_after = Pt(10)
    r_meta = p_meta.add_run(
        f"Lotes únicos auditados: {len(df_prod)} certificados validados (sin duplicados)   |   "
        f"Período cubierto: {rango_anios_str}   |   "
        f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y')}"
    )
    r_meta.font.name = "Calibri"
    r_meta.font.size = Pt(9)
    r_meta.font.italic = True
    r_meta.font.color.rgb = COLOR_GRAY

    resultados_params = []
    for col, nom, uni, _ in PARAMETROS_CONFIG:
        if col in df_prod.columns and df_prod[col].notna().sum() > 0:
            info = redactar_conclusion_parametro(df_prod, col, nom, uni, modo)
            if info:
                resultados_params.append(info)

    agregar_titulo_seccion(doc, "1. Resumen Estadístico de Parámetros de Calidad")

    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.space_after = Pt(6)
    r_intro = p_intro.add_run(
        f"La siguiente tabla consolida los indicadores estadísticos obligatorios (cantidad de lotes, "
        f"promedio, desviación estándar, valor mínimo y valor máximo registrado) para cada parámetro "
        f"analizado en los certificados de calidad de {producto}."
    )
    r_intro.font.name = "Calibri"
    r_intro.font.size = Pt(10)

    tbl_stat = doc.add_table(rows=1, cols=8)
    hdr_s = [
        "Parámetro Analizado", "Unidad", "Lotes (n)", "Promedio",
        "Desv. Est. (σ)", "CV (%)", "Mínimo (Lote)", "Máximo (Lote)"
    ]
    for i, h in enumerate(hdr_s):
        tbl_stat.rows[0].cells[i].text = h

    for info in resultados_params:
        row_cells = tbl_stat.add_row().cells
        row_cells[0].text = info["nombre"]
        row_cells[1].text = info["unidad"]
        row_cells[2].text = str(info["n"])
        row_cells[3].text = fmt_num(info["mean"])
        row_cells[4].text = fmt_num(info["std"])
        row_cells[5].text = f"{fmt_num(info['cv'], 1)}%"
        row_cells[6].text = f"{fmt_num(info['min'])} (L.{info['lote_min']})"
        row_cells[7].text = f"{fmt_num(info['max'])} (L.{info['lote_max']})"

    estilizar_tabla_word(tbl_stat, col_widths=[1.8, 0.55, 0.55, 0.75, 0.75, 0.6, 0.95, 0.95])

    if modo == "HISTORICO" and len(anios) >= 2:
        agregar_titulo_seccion(doc, "2. Evolución de Promedios y Desvíos por Año")

        params_principales = [p for p in resultados_params if not p["nombre"].endswith("(OIV)")]
        if not params_principales:
            params_principales = resultados_params

        cols_tabla_anual = ["Año", "Lotes (n)"]
        for p in params_principales:
            cols_tabla_anual.append(f"{p['nombre']} ({p['unidad']})\nMedia ± σ [Mín–Máx]")

        tbl_anual = doc.add_table(rows=1, cols=len(cols_tabla_anual))
        for i, h in enumerate(cols_tabla_anual):
            tbl_anual.rows[0].cells[i].text = h

        for anio in anios:
            df_a = df_prod[df_prod["Anio"] == anio]
            rc = tbl_anual.add_row().cells
            rc[0].text = str(int(anio))
            rc[1].text = str(len(df_a))
            for idx_p, p in enumerate(params_principales, start=2):
                s_a = df_a[p["col"]].dropna()
                if len(s_a) > 0:
                    m_a = s_a.mean()
                    sd_a = s_a.std(ddof=1) if len(s_a) > 1 else 0.0
                    rc[idx_p].text = f"{fmt_num(m_a)} ± {fmt_num(sd_a)} [{fmt_num(s_a.min())}–{fmt_num(s_a.max())}]"
                else:
                    rc[idx_p].text = "-"

        estilizar_tabla_word(tbl_anual)

    num_sec_graf = "3" if (modo == "HISTORICO" and len(anios) >= 2) else "2"
    agregar_titulo_seccion(doc, f"{num_sec_graf}. Gráficos de Tendencia y Conclusiones Analíticas por Parámetro")

    for info in resultados_params:
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(10)
        p_sub.paragraph_format.space_after = Pt(4)
        p_sub.paragraph_format.keep_with_next = True
        r_sub = p_sub.add_run(f"▪ {info['nombre']} ({info['unidad']})")
        r_sub.font.name = "Calibri"
        r_sub.font.size = Pt(11.5)
        r_sub.font.bold = True
        r_sub.font.color.rgb = COLOR_STEEL

        nombre_img = re.sub(r"[^A-Za-z0-9_]", "_", f"{producto}_{modo}_{info['col']}") + ".png"
        ruta_img = CHARTS_DIR / nombre_img
        ok_chart = generar_grafico_parametro(
            df_prod, info["col"], info["nombre"], info["unidad"], producto, modo, ruta_img
        )
        if ok_chart and ruta_img.exists():
            info["ruta_img"] = ruta_img
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(4)
            ancho_pulg = 6.4 if (modo == "HISTORICO" and len(anios) >= 2) else 5.8
            p_img.add_run().add_picture(str(ruta_img), width=Inches(ancho_pulg))

        p_tend = doc.add_paragraph()
        p_tend.paragraph_format.space_before = Pt(2)
        p_tend.paragraph_format.space_after = Pt(3)
        r_t_lbl = p_tend.add_run("Evolución y tendencia: ")
        r_t_lbl.font.name = "Calibri"
        r_t_lbl.font.size = Pt(9.5)
        r_t_lbl.font.bold = True
        r_t_txt = p_tend.add_run(info["texto_tendencia"])
        r_t_txt.font.name = "Calibri"
        r_t_txt.font.size = Pt(9.5)

        p_anom = doc.add_paragraph()
        p_anom.paragraph_format.space_before = Pt(1)
        p_anom.paragraph_format.space_after = Pt(8)
        r_a_lbl = p_anom.add_run("Detección de anomalías: ")
        r_a_lbl.font.name = "Calibri"
        r_a_lbl.font.size = Pt(9.5)
        r_a_lbl.font.bold = True
        if info["cant_outliers"] > 0:
            r_a_lbl.font.color.rgb = RGBColor(180, 40, 40)
        r_a_txt = p_anom.add_run(info["texto_anomalias"])
        r_a_txt.font.name = "Calibri"
        r_a_txt.font.size = Pt(9.5)

    num_sec_anexo = "4" if (modo == "HISTORICO" and len(anios) >= 2) else "3"
    agregar_titulo_seccion(doc, f"{num_sec_anexo}. Anexo: Base de Datos Consolidada de Certificados ({len(df_prod)} Lotes)")

    cols_anexo = [("Lote", "Lote"), ("Anio", "Año"), ("Fecha_Emision", "Fecha")]
    for info in resultados_params:
        cols_anexo.append((info["col"], f"{info['nombre']}\n({info['unidad']})"))
    cols_anexo.append(("Cliente", "Cliente"))

    tbl_lotes = doc.add_table(rows=1, cols=len(cols_anexo))
    for i, (_, titulo_col) in enumerate(cols_anexo):
        tbl_lotes.rows[0].cells[i].text = titulo_col

    for _, row in df_prod.iterrows():
        rc = tbl_lotes.add_row().cells
        for i, (col_key, _) in enumerate(cols_anexo):
            val = row.get(col_key, "")
            if col_key in ("Lote", "Anio"):
                rc[i].text = "" if pd.isna(val) else str(val)
            elif col_key in ("Fecha_Emision", "Cliente"):
                rc[i].text = "-" if (pd.isna(val) or str(val).strip() == "") else str(val)
            else:
                rc[i].text = fmt_num(val)

    estilizar_tabla_word(tbl_lotes)
    doc.save(ruta_docx)


# ==============================================================================
# CONSTRUCCIÓN DEL LIBRO EXCEL CON LAS 4 HOJAS DE KARINA (.XLSX)
# ==============================================================================
def crear_excel_producto(df_hist: pd.DataFrame, df_2025: pd.DataFrame, producto: str, ruta_xlsx: Path):
    """
    Crea el libro Excel con 4 pestañas:
    1) Base_Datos_Unica
    2) Analisis_Estadistico (Global + Promedios por Año)
    3) Graficos_Tendencia (Imágenes incrustadas)
    4) Conclusion_Tendencias (Evolución + Anomalías)
    """
    wb = openpyxl.Workbook()

    fill_navy = PatternFill(start_color=HEX_NAVY, end_color=HEX_NAVY, fill_type="solid")
    fill_steel = PatternFill(start_color=HEX_STEEL, end_color=HEX_STEEL, fill_type="solid")
    fill_zebra = PatternFill(start_color=HEX_LIGHT, end_color=HEX_LIGHT, fill_type="solid")
    font_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Calibri", size=10)
    font_bold = Font(name="Calibri", size=10, bold=True)
    thin_border = Border(
        left=Side(style="thin", color="D0D7DE"),
        right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"),
        bottom=Side(style="thin", color="D0D7DE"),
    )

    # --- HOJA 1: Base_Datos_Unica ---
    ws_bd = wb.active
    ws_bd.title = "Base_Datos_Unica"

    df_comb = ordenar_df_cronologico(pd.concat([df_hist, df_2025], ignore_index=True))
    cols_export = [
        "Rango_Temporal", "Anio", "Lote", "Fecha_Emision", "Cliente", "Version_Certificado",
        "Permeabilidad_Darcy", "Densidad_Humeda_gcc", "Retencion_Malla_150_pct",
        "Retencion_Malla_80_pct", "Humedad_pct", "Impurezas_pct",
        "OIV_Arsenico_ppm_Res", "OIV_Hierro_ppm_Res", "OIV_Plomo_ppm_Res", "Ruta_Archivo"
    ]
    cols_export = [c for c in cols_export if c in df_comb.columns]

    ws_bd.append(cols_export)
    for _, row in df_comb.iterrows():
        ws_bd.append([None if pd.isna(row[c]) else row[c] for c in cols_export])

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

    # --- HOJA 2: Analisis_Estadistico ---
    ws_stat = wb.create_sheet("Analisis_Estadistico")
    headers_stat = [
        "Período", "Parámetro", "Unidad", "Lotes (n)", "Promedio",
        "Desv. Estándar (σ)", "CV (%)", "Mínimo", "Lote Mín", "Máximo", "Lote Máx"
    ]
    ws_stat.append(headers_stat)
    for cell in ws_stat[1]:
        cell.fill = fill_navy
        cell.font = font_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center")

    infos_para_conclusiones = []
    imagenes_para_excel = []

    for etiqueta_p, df_sub, modo_p in [
        ("Histórico (<=2024)", df_hist, "HISTORICO"),
        ("Anual 2025", df_2025, "ANUAL_2025"),
    ]:
        if len(df_sub) == 0:
            continue
        for col, nom, uni, _ in PARAMETROS_CONFIG:
            if col in df_sub.columns and df_sub[col].notna().sum() > 0:
                info = redactar_conclusion_parametro(df_sub, col, nom, uni, modo_p)
                if info:
                    ws_stat.append([
                        etiqueta_p, info["nombre"], info["unidad"], info["n"],
                        round(info["mean"], 4), round(info["std"], 4), round(info["cv"], 2),
                        round(info["min"], 4), info["lote_min"], round(info["max"], 4), info["lote_max"]
                    ])
                    infos_para_conclusiones.append((etiqueta_p, info))
                    nombre_img = re.sub(r"[^A-Za-z0-9_]", "_", f"{producto}_{modo_p}_{col}") + ".png"
                    ruta_img = CHARTS_DIR / nombre_img
                    if ruta_img.exists():
                        imagenes_para_excel.append((f"{etiqueta_p} — {nom} ({uni})", ruta_img))

    # Agregar tabla de Promedios por Año debajo en Analisis_Estadistico
    ws_stat.append([])
    ws_stat.append(["PROMEDIOS Y DESVÍOS POR AÑO CALENDARIO"])
    fila_titulo_anual = ws_stat.max_row
    ws_stat.cell(row=fila_titulo_anual, column=1).font = Font(name="Calibri", size=11, bold=True, color=HEX_NAVY)

    params_activos = [
        (col, nom, uni) for col, nom, uni, _ in PARAMETROS_CONFIG
        if col in df_comb.columns and df_comb[col].notna().sum() > 0
    ]
    hdr_anual = ["Año", "Lotes Únicos (n)"]
    for _, nom, uni in params_activos:
        hdr_anual.extend([f"Promedio {nom} ({uni})", f"Desv.Est {nom} ({uni})"])
    ws_stat.append(hdr_anual)
    fila_hdr_anual = ws_stat.max_row
    for cell in ws_stat[fila_hdr_anual]:
        cell.fill = fill_steel
        cell.font = font_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for anio in sorted(df_comb["Anio"].unique()):
        df_a = df_comb[df_comb["Anio"] == anio]
        row_a = [int(anio), len(df_a)]
        for col, _, _ in params_activos:
            s_a = df_a[col].dropna()
            if len(s_a) > 0:
                row_a.append(round(float(s_a.mean()), 4))
                row_a.append(round(float(s_a.std(ddof=1)) if len(s_a) > 1 else 0.0, 4))
            else:
                row_a.extend([None, None])
        ws_stat.append(row_a)

    for r_idx in range(2, ws_stat.max_row + 1):
        if r_idx in (fila_titulo_anual, fila_hdr_anual, fila_titulo_anual - 1):
            continue
        for c_idx in range(1, ws_stat.max_column + 1):
            c = ws_stat.cell(row=r_idx, column=c_idx)
            c.font = font_bold if c_idx in (1, 2) else font_body
            c.border = thin_border
            if r_idx % 2 == 0:
                c.fill = fill_zebra

    # --- HOJA 3: Graficos_Tendencia ---
    ws_graf = wb.create_sheet("Graficos_Tendencia")
    ws_graf["A1"] = f"GRÁFICOS DE EVOLUCIÓN Y CONTROL DE CALIDAD — {producto}"
    ws_graf["A1"].font = Font(name="Calibri", size=13, bold=True, color=HEX_NAVY)

    fila_actual_img = 3
    for titulo_g, ruta_img in imagenes_para_excel:
        ws_graf.cell(row=fila_actual_img, column=1, value=f"▪ {titulo_g}").font = Font(
            name="Calibri", size=11, bold=True, color=HEX_STEEL
        )
        try:
            img_xl = XLImage(str(ruta_img))
            img_xl.width = 760
            img_xl.height = 280
            ws_graf.add_image(img_xl, f"A{fila_actual_img + 1}")
            fila_actual_img += 16
        except Exception:
            fila_actual_img += 2

    # --- HOJA 4: Conclusion_Tendencias ---
    ws_conc = wb.create_sheet("Conclusion_Tendencias")
    headers_conc = [
        "Período", "Parámetro", "Unidad", "Lotes (n)",
        "Conclusión de Evolución y Etapas", "Detección de Anomalías (Criterio ±2σ)"
    ]
    ws_conc.append(headers_conc)
    for cell in ws_conc[1]:
        cell.fill = fill_navy
        cell.font = font_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for etiqueta_p, info in infos_para_conclusiones:
        ws_conc.append([
            etiqueta_p, info["nombre"], info["unidad"], info["n"],
            info["texto_tendencia"], info["texto_anomalias"]
        ])

    for r_idx in range(2, ws_conc.max_row + 1):
        ws_conc.row_dimensions[r_idx].height = 58
        for c_idx in range(1, ws_conc.max_column + 1):
            c = ws_conc.cell(row=r_idx, column=c_idx)
            c.font = font_bold if c_idx in (1, 2) else font_body
            c.border = thin_border
            c.alignment = Alignment(vertical="center", wrap_text=(c_idx >= 5))
            if r_idx % 2 == 0:
                c.fill = fill_zebra

    # Ajustar anchos de columna
    for ws in [ws_bd, ws_stat, ws_conc]:
        for col_cells in ws.columns:
            col_letter = col_cells[0].column_letter
            max_len = max(len(str(c.value or "")) for c in col_cells[:25])
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 55)

    wb.save(ruta_xlsx)


# ==============================================================================
# FLUJO PRINCIPAL
# ==============================================================================
def cargar_csv_limpio(ruta_csv: Path) -> pd.DataFrame:
    df = pd.read_csv(ruta_csv, sep=";", encoding="utf-8-sig")
    df = df[df["Version_Preferida_Unica"] == "SI"].copy()
    for col, _, _, _ in PARAMETROS_CONFIG:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", ".", regex=False)
                .replace({"nan": np.nan, "None": np.nan, "": np.nan})
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")
    df["Anio"] = pd.to_numeric(df["Anio"], errors="coerce").fillna(2020).astype(int)
    return ordenar_df_cronologico(df)


def main():
    csv_2024 = CSV_DIR / "Filtracion_1_Hasta_2024.csv"
    csv_2025 = CSV_DIR / "Filtracion_2_Anio_2025.csv"

    if not csv_2024.exists() or not csv_2025.exists():
        print(f"❌ No se encontraron los archivos CSV en {CSV_DIR}. Ejecutá primero procesar_filtracion.py")
        return

    print("📂 Cargando bases de datos de Filtración (lotes únicos ordenados cronológicamente)...")
    df_hist_all = cargar_csv_limpio(csv_2024)
    df_2025_all = cargar_csv_limpio(csv_2025)

    df_hist_all = df_hist_all[df_hist_all["Producto_Base"] != "PRECOAT DWG-B"].copy()

    productos = sorted(
        set(df_hist_all["Producto_Base"].unique()).union(set(df_2025_all["Producto_Base"].unique()))
    )

    print(f"📊 Generando informes v2 (.docx + .xlsx con gráficos) para {len(productos)} productos...\n")
    informes_creados = 0

    for prod in productos:
        prod_slug = re.sub(r"[^A-Za-z0-9]+", "_", prod).strip("_")
        prod_dir = OUT_DIR / prod_slug
        prod_dir.mkdir(exist_ok=True)

        sub_hist = df_hist_all[df_hist_all["Producto_Base"] == prod].copy()
        sub_2025 = df_2025_all[df_2025_all["Producto_Base"] == prod].copy()

        if len(sub_hist) > 0:
            ruta_doc_h = prod_dir / f"Informe_Tendencia_{prod_slug}_Historico_Hasta_2024.docx"
            crear_informe_word(sub_hist, prod, "HISTORICO", ruta_doc_h)
            informes_creados += 1
            print(f"  ✔ [WORD HISTÓRICO <=2024] {prod:22s} -> {len(sub_hist):3d} lotes únicos ({ruta_doc_h.name})")

        if len(sub_2025) > 0:
            ruta_doc_25 = prod_dir / f"Informe_Tendencia_{prod_slug}_Anual_2025.docx"
            crear_informe_word(sub_2025, prod, "ANUAL_2025", ruta_doc_25)
            informes_creados += 1
            print(f"  ✔ [WORD ANUAL 2025]       {prod:22s} -> {len(sub_2025):3d} lotes únicos ({ruta_doc_25.name})")

        ruta_xl = prod_dir / f"Base_y_Estadistica_{prod_slug}.xlsx"
        crear_excel_producto(sub_hist, sub_2025, prod, ruta_xl)

    print("\n" + "=" * 75)
    print(f"✅ ¡VERSIÓN 2 COMPLETADA! Se actualizaron los {informes_creados} informes Word (.docx)")
    print(f"   y los {len(productos)} libros Excel (.xlsx con 4 hojas y gráficos) en: {OUT_DIR}/")
    print("=" * 75)


if __name__ == "__main__":
    main()
