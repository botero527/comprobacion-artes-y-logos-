#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
TEST — Auditoría Layers K / K2  (AutoCAD COM)
Selecciona archivos DWG manualmente para validar la lógica
antes de correr el escaneo completo sobre toda la red.
"""

import os
import sys
import time
from datetime import datetime

# ──────────────────────────────────────────────────────────
# ARCHIVOS A PROBAR — agrega rutas aquí o déjalo vacío
# para que el script las pida en consola
# ──────────────────────────────────────────────────────────
ARCHIVOS_PRUEBA = [
    # r"\\192.168.2.37\ingenieria\...\archivo1.dwg",
    # r"C:\Users\abotero\Desktop\prueba.dwg",
]
# ──────────────────────────────────────────────────────────

COLOR_AZUL_ACI  = 5
COLOR_VERDE_ACI = 3

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("Falta pywin32. Ejecuta: pip install pywin32")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    print("Falta openpyxl. Ejecuta: pip install openpyxl")
    sys.exit(1)


# ──────────────────────────────────────────────────────────
# HELPERS CONSOLA
# ──────────────────────────────────────────────────────────
def sep(c="─", n=65): print(c * n)
def ts(): return time.strftime("%H:%M:%S")
def ok(m):   print(f"{ts()}    ✓  {m}")
def warn(m): print(f"{ts()}    !  {m}")
def err(m):  print(f"{ts()}    ✗  {m}")
def info(m): print(f"{ts()}       {m}")


# ──────────────────────────────────────────────────────────
# NOMBRES DE COLORES ACI
# ──────────────────────────────────────────────────────────
_COLORES = {
    1: "Rojo", 2: "Amarillo", 3: "Verde", 4: "Cyan",
    5: "Azul", 6: "Magenta",  7: "Blanco/Negro",
    8: "Gris oscuro", 9: "Gris claro",
    0: "ByBlock", 256: "ByLayer",
}
def nombre_color(aci): return _COLORES.get(abs(aci), f"ACI {abs(aci)}")


# ──────────────────────────────────────────────────────────
# AUTOCAD COM
# ──────────────────────────────────────────────────────────
def conectar_autocad():
    pythoncom.CoInitialize()
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        # Silenciar diálogos y errores de AutoCAD
        acad.Application.Preferences.OpenSave.DemandLoadARXApp = 2
        ok(f"AutoCAD conectado: {acad.Version}")
        return acad
    except Exception as e:
        err(f"No hay AutoCAD abierto: {e}")
        err("Abre AutoCAD primero (no hace falta abrir ningún archivo) y vuelve a ejecutar.")
        sys.exit(1)


def abrir_dwg(acad, ruta):
    """Abre el DWG en modo solo-lectura para que sea más rápido y no bloquee el archivo."""
    ruta_abs = os.path.abspath(ruta)
    for intento in range(1, 4):
        try:
            # True = read-only: más rápido y no modifica el archivo
            doc = acad.Documents.Open(ruta_abs, True)
            time.sleep(0.5)   # mínimo para que AutoCAD termine de cargar
            return doc
        except Exception as e:
            if intento < 3:
                warn(f"  Reintento {intento}/3: {str(e)[:60]}")
                time.sleep(intento * 1.5)
    return None


def cerrar_dwg(doc):
    try:
        doc.Close(False)
        time.sleep(0.1)
    except Exception:
        pass

def leer_layers(doc):
    layers = []
    try:
        col = doc.Layers
        for i in range(col.Count):
            try:
                l   = col.Item(i)
                aci = abs(l.Color)
                layers.append({
                    "nombre":      l.Name,
                    "color_aci":   aci,
                    "color_texto": nombre_color(aci),
                })
            except Exception:
                continue
    except Exception as e:
        warn(f"Error leyendo layers: {e}")
    return layers


# ──────────────────────────────────────────────────────────
# LÓGICA DE VALIDACIÓN
# ──────────────────────────────────────────────────────────
def evaluar(layers):
    layer_k  = next((l for l in layers if l["nombre"].upper().strip() == "K"),  None)
    layer_k2 = next((l for l in layers if l["nombre"].upper().strip() == "K2"), None)

    tiene_k     = layer_k  is not None
    tiene_k2    = layer_k2 is not None
    k_es_azul   = tiene_k  and layer_k["color_aci"]  == COLOR_AZUL_ACI
    k2_es_verde = tiene_k2 and layer_k2["color_aci"] == COLOR_VERDE_ACI

    if k_es_azul and k2_es_verde:
        estado = "ACTUALIZADA"
    elif not tiene_k and not tiene_k2:
        estado = "VIEJA"
    else:
        estado = "INCOMPLETA"

    return {
        "estado":      estado,
        "tiene_k":     tiene_k,   "layer_k":  layer_k,  "k_es_azul":   k_es_azul,
        "tiene_k2":    tiene_k2,  "layer_k2": layer_k2, "k2_es_verde": k2_es_verde,
    }


# ──────────────────────────────────────────────────────────
# ANÁLISIS INDIVIDUAL CON SALIDA DETALLADA EN CONSOLA
# ──────────────────────────────────────────────────────────
def analizar(acad, ruta, numero, total):
    nombre = os.path.basename(ruta)
    sep()
    info(f"[{numero}/{total}] {nombre}")
    info(f"Ruta: {ruta}")
    sep("·")

    if not os.path.exists(ruta):
        err("Archivo NO encontrado en disco")
        return {"archivo": nombre, "ruta": ruta, "estado": "ERROR",
                "detalle": {}, "layers": [], "detalle_error": "Archivo no existe"}

    info("Abriendo en AutoCAD (solo lectura)...")
    t0  = time.time()
    doc = abrir_dwg(acad, ruta)
    ms  = (time.time() - t0) * 1000

    if doc is None:
        err("No se pudo abrir el archivo")
        return {"archivo": nombre, "ruta": ruta, "estado": "ERROR",
                "detalle": {}, "layers": [], "detalle_error": "No se pudo abrir"}

    layers = leer_layers(doc)
    cerrar_dwg(doc)

    info(f"Abierto en {ms:.0f} ms  —  {len(layers)} layers encontrados")
    print()

    # Tabla de layers en consola
    print(f"  {'NOMBRE LAYER':<32} {'COLOR ACI':>9}  COLOR")
    print("  " + "─" * 58)
    for l in sorted(layers, key=lambda x: x["nombre"]):
        u = l["nombre"].upper().strip()
        marcador = ""
        if u == "K":
            estado_color = "AZUL ✓" if l["color_aci"] == COLOR_AZUL_ACI \
                           else f"NO ES AZUL ✗  (es {l['color_texto']})"
            marcador = f"  ◄── LAYER K  →  {estado_color}"
        elif u == "K2":
            estado_color = "VERDE ✓" if l["color_aci"] == COLOR_VERDE_ACI \
                           else f"NO ES VERDE ✗  (es {l['color_texto']})"
            marcador = f"  ◄── LAYER K2 →  {estado_color}"
        print(f"  {l['nombre']:<32} {l['color_aci']:>9}  {l['color_texto']}{marcador}")

    print()
    diag = evaluar(layers)
    lk   = diag["layer_k"]  or {}
    lk2  = diag["layer_k2"] or {}

    sep("═")
    ICONOS = {"ACTUALIZADA": "✓", "VIEJA": "✗", "INCOMPLETA": "!"}
    print(f"\n  {ICONOS.get(diag['estado'],'?')}  RESULTADO: {diag['estado']}\n")

    print(f"  Layer K    : {'SÍ' if diag['tiene_k'] else 'NO existe'}", end="")
    if lk:
        print(f"  →  nombre='{lk.get('nombre','')}' / color={lk.get('color_texto','')}", end="")
    print()
    if diag["tiene_k"]:
        print(f"  K es azul  : {'SÍ ✓' if diag['k_es_azul'] else 'NO ✗  ← necesita ser AZUL (ACI 5)'}")

    print(f"  Layer K2   : {'SÍ' if diag['tiene_k2'] else 'NO existe'}", end="")
    if lk2:
        print(f"  →  nombre='{lk2.get('nombre','')}' / color={lk2.get('color_texto','')}", end="")
    print()
    if diag["tiene_k2"]:
        print(f"  K2 es verde: {'SÍ ✓' if diag['k2_es_verde'] else 'NO ✗  ← necesita ser VERDE (ACI 3)'}")

    EXPLICACIONES = {
        "ACTUALIZADA": "Layer K (azul) + Layer K2 (verde) presentes → arte ACTUALIZADA",
        "VIEJA":       "Sin layer K ni K2 → arte VIEJA / OBSOLETA",
        "INCOMPLETA":  "Solo cumple una condición o color incorrecto → revisar",
    }
    print(f"\n  → {EXPLICACIONES[diag['estado']]}")
    print()

    return {
        "archivo":       nombre,
        "ruta":          ruta,
        "estado":        diag["estado"],
        "detalle":       diag,
        "layers":        layers,
        "detalle_error": "",
    }


# ──────────────────────────────────────────────────────────
# EXCEL DE PRUEBA
# ──────────────────────────────────────────────────────────
_thin   = Side(style="thin", color="BBBBBB")
_border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

BG_ESTADO = {
    "ACTUALIZADA": "C6EFCE",
    "VIEJA":       "FFCCCC",
    "INCOMPLETA":  "FFEB9C",
    "ERROR":       "FFC7CE",
}

def _c(ws, r, col, val, bold=False, bg=None, center=False,
       color_txt="000000", mono=False, size=None):
    cell = ws.cell(r, col, val)
    s    = size or (8 if mono else 9)
    cell.font      = Font(name="Courier New" if mono else "Arial",
                          bold=bold, size=s, color=color_txt)
    cell.fill      = PatternFill("solid", start_color=bg) if bg else PatternFill()
    cell.alignment = Alignment(horizontal="center" if center else "left",
                               vertical="center")
    cell.border    = _border
    return cell


def generar_excel(resultados, ruta_salida):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RESULTADOS PRUEBA"
    ws.sheet_view.showGridLines = False

    fecha = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    HEADERS = [
        "#", "ARCHIVO", "ESTADO",
        "TIENE K", "NOMBRE K", "COLOR K", "K ES AZUL",
        "TIENE K2", "NOMBRE K2", "COLOR K2", "K2 ES VERDE",
        "TOTAL LAYERS", "TODOS LOS LAYERS",
        "RUTA COMPLETA", "DETALLE ERROR",
    ]
    ANCHOS = [5, 38, 16, 10, 14, 16, 11, 10, 14, 16, 12, 13, 70, 80, 40]

    ws.merge_cells(f"A1:{openpyxl.utils.get_column_letter(len(HEADERS))}1")
    t = ws.cell(1, 1, f"TEST AUDITORÍA LAYERS K/K2  —  {fecha}")
    t.font      = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    t.fill      = PatternFill("solid", start_color="1F3864")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for col, h in enumerate(HEADERS, 1):
        c = ws.cell(2, col, h)
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill      = PatternFill("solid", start_color="2E75B6")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = _border
    ws.row_dimensions[2].height = 20

    for idx, r in enumerate(resultados, 1):
        row  = idx + 2
        alt  = row % 2 == 0
        bg_a = "EEF4FF" if alt else None
        d    = r.get("detalle", {})
        lk   = d.get("layer_k")  or {}
        lk2  = d.get("layer_k2") or {}

        lista = " | ".join(
            f"{l['nombre']}({l['color_texto']})"
            for l in sorted(r.get("layers", []), key=lambda x: x["nombre"])
        )

        def yn(val):    return ("SÍ",   "C6EFCE") if val else ("NO",   "FFCCCC")
        def ytick(val): return ("SÍ ✓", "C6EFCE") if val else ("NO ✗", "FFCCCC")

        _c(ws, row,  1, idx,           center=True, bg=bg_a)
        _c(ws, row,  2, r["archivo"],  bg=bg_a)
        _c(ws, row,  3, r["estado"],   center=True, bold=True,
           bg=BG_ESTADO.get(r["estado"], "FFFFFF"))

        v, b = yn(d.get("tiene_k"));    _c(ws, row, 4, v, center=True, bg=b)
        _c(ws, row,  5, lk.get("nombre", ""),      bg=bg_a)
        _c(ws, row,  6, lk.get("color_texto", ""), bg=bg_a)
        v, b = ytick(d.get("k_es_azul"));   _c(ws, row, 7, v, center=True, bg=b)

        v, b = yn(d.get("tiene_k2"));   _c(ws, row, 8, v, center=True, bg=b)
        _c(ws, row,  9, lk2.get("nombre", ""),      bg=bg_a)
        _c(ws, row, 10, lk2.get("color_texto", ""), bg=bg_a)
        v, b = ytick(d.get("k2_es_verde")); _c(ws, row, 11, v, center=True, bg=b)

        _c(ws, row, 12, len(r.get("layers", [])), center=True, bg=bg_a)
        _c(ws, row, 13, lista,     mono=True, bg=bg_a)
        _c(ws, row, 14, r["ruta"], mono=True, color_txt="0070C0", bg=bg_a)
        es_err = r["estado"] == "ERROR"
        _c(ws, row, 15, r.get("detalle_error", ""),
           bg="FFC7CE" if es_err else bg_a)

    for i, w in enumerate(ANCHOS, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A3"
    wb.save(ruta_salida)
    ok(f"Excel guardado: {ruta_salida}")


# ──────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────
def main():
    sep("═")
    print("  TEST — AUDITORÍA LAYERS K / K2  (AutoCAD COM)")
    print("  IMPORTANTE: AutoCAD debe estar abierto antes de ejecutar esto.")
    sep("═")

    archivos = [a for a in ARCHIVOS_PRUEBA if a.strip()]
    if not archivos:
        print("\n  Ingresa rutas DWG una por una. Enter vacío para terminar.\n")
        while True:
            ruta = input(f"  Archivo {len(archivos)+1}: ").strip().strip('"').strip("'")
            if not ruta:
                break
            if not ruta.lower().endswith(".dwg"):
                warn("No es .dwg, ignorado")
                continue
            archivos.append(ruta)
            ok(f"Agregado: {os.path.basename(ruta)}")

    if not archivos:
        warn("Sin archivos para probar.")
        return

    print(f"\n  {len(archivos)} archivo(s) a analizar\n")

    acad = conectar_autocad()

    resultados = []
    t0 = time.time()
    for i, ruta in enumerate(archivos, 1):
        r = analizar(acad, ruta, i, len(archivos))
        resultados.append(r)

    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass

    total_s = time.time() - t0

    sep("═")
    print("\n  RESUMEN\n")
    ICONOS = {"ACTUALIZADA": "✓", "VIEJA": "✗", "INCOMPLETA": "!", "ERROR": "E"}
    for r in resultados:
        print(f"  {ICONOS.get(r['estado'],'?')}  {r['estado']:<14}  {r['archivo']}")

    print(f"\n  Total: {len(resultados)} archivos  —  {total_s:.1f} s")
    for estado in ["ACTUALIZADA", "VIEJA", "INCOMPLETA", "ERROR"]:
        n = sum(1 for r in resultados if r["estado"] == estado)
        if n:
            print(f"  {estado:<14}: {n}")

    nombre_excel = f"TEST_layers_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    print()
    generar_excel(resultados, nombre_excel)

    sep("═")
    print(f"\n  Listo. Excel: {nombre_excel}\n")
    sep("═")


if __name__ == "__main__":
    main()
