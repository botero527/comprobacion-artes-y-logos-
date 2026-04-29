#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AUDITORÍA LAYERS K / K2 — AGP PLANOS TECNICOS
Estructura: RUTA_BASE > Vehículo > Modelo > Versión > ARTES
Recolecta: DWGs sueltos en ARTES + DWGs en subcarpetas con "BN" en el nombre
Excluye: subcarpetas "OBSOLETOS" (nombre contiene "obsoleto" sin importar mayúsculas)
Valida por cada DWG:
  - Layer K existe y es AZUL (ACI 5)?
  - Layer K2 existe y es VERDE (ACI 3)?
Estado final:
  ACTUALIZADA  → Layer K existe y es azul  Y  Layer K2 existe y es verde
  VIEJA        → Ninguna de las dos condiciones se cumple
  INCOMPLETA   → Solo una condición se cumple (o colores incorrectos)
"""

import os
import sys
import time
from collections import defaultdict
from datetime import datetime

# ──────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────────────────
RUTA_BASE     = r"\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS\CHEVROLET"
ARCHIVO_EXCEL = f"Auditoria_Layers_K_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
COLOR_VERDE_ACI = 3   # Índice de color verde en AutoCAD (ACI)
COLOR_AZUL_ACI  = 5   # Índice de color azul  en AutoCAD (ACI)
# ──────────────────────────────────────────────────────────

class Logger:
    def _ts(self):
        return time.strftime("%H:%M:%S")
    def info(self, msg):   print(f"{self._ts()}  {msg}")
    def warn(self, msg):   print(f"{self._ts()}  [WARN]  {msg}")
    def error(self, msg):  print(f"{self._ts()}  [ERROR] {msg}")

log = Logger()

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
# PALETA Y ESTILOS
# ──────────────────────────────────────────────────────────
C = {
    "hdr":         "1F3864",
    "titulo":      "2E75B6",
    "white":       "FFFFFF",
    "alt":         "EEF4FF",
    "actualizada": "C6EFCE",  # verde claro
    "vieja":       "FFCCCC",  # rojo claro
    "incompleta":  "FFEB9C",  # amarillo claro
    "error":       "FFC7CE",  # rojo intenso
    "gris":        "D9D9D9",
}

_thin   = Side(style="thin", color="BBBBBB")
_border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

def _cel_header(cell, bg=None):
    bg = bg or C["hdr"]
    cell.font      = Font(name="Arial", bold=True, color=C["white"], size=10)
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = _border

def _cel_dato(cell, alt=False, center=False, bold=False, color_txt=None, bg=None):
    cell.font      = Font(name="Arial", size=9, bold=bold,
                          color=color_txt or "000000")
    cell.alignment = Alignment(horizontal="center" if center else "left",
                               vertical="center", wrap_text=False)
    cell.border    = _border
    if bg:
        cell.fill = PatternFill("solid", start_color=bg)
    elif alt:
        cell.fill = PatternFill("solid", start_color=C["alt"])

# ──────────────────────────────────────────────────────────
# MOTOR AUTOCAD
# ──────────────────────────────────────────────────────────
class AutoCADMotor:
    def __init__(self):
        self.acad = None
        pythoncom.CoInitialize()
        try:
            self.acad = win32com.client.GetActiveObject("AutoCAD.Application")
            log.info("[AutoCAD] Conectado a instancia existente")
        except Exception:
            log.error("No hay AutoCAD abierto. Abre AutoCAD primero y vuelve a ejecutar.")
            sys.exit(1)

    def abrir(self, ruta):
        ruta_abs = os.path.abspath(ruta)
        for intento, espera in enumerate([1.0, 2.0, 3.0, 5.0, 8.0], 1):
            try:
                doc = self.acad.Documents.Open(ruta_abs)
                time.sleep(0.8)
                return doc
            except Exception as e:
                if intento < 5:
                    log.warn(f"  Reintento {intento}/5 abriendo {os.path.basename(ruta)}: {str(e)[:40]}")
                    time.sleep(espera)
        return None

    def cerrar(self, doc):
        try:
            doc.Close(False)
        except Exception:
            pass

    def inspeccionar_layers(self, doc):
        """
        Retorna dict con toda la info de layers del documento.
        {nombre: {color_aci, color_nombre, es_k, es_k2, k2_verde}}
        """
        resultado = {}
        try:
            layers = doc.Layers
            for i in range(layers.Count):
                try:
                    layer = layers.Item(i)
                    nombre = layer.Name
                    color_aci = layer.Color
                    resultado[nombre] = {
                        "color_aci":    color_aci,
                        "color_nombre": _nombre_color(color_aci),
                    }
                except Exception:
                    continue
        except Exception as e:
            log.warn(f"  Error leyendo layers: {str(e)[:60]}")
        return resultado

    def quit(self):
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _nombre_color(aci):
    """Convierte índice ACI a nombre descriptivo para los más comunes."""
    mapa = {
        1: "Rojo",    2: "Amarillo", 3: "Verde",   4: "Cyan",
        5: "Azul",    6: "Magenta",  7: "Blanco",  8: "Gris oscuro",
        9: "Gris",    0: "ByBlock", 256: "ByLayer",
    }
    return mapa.get(aci, f"Color {aci}")


# ──────────────────────────────────────────────────────────
# LÓGICA DE ESTADO
# ──────────────────────────────────────────────────────────
def evaluar_estado(layers_dict):
    """
    Reglas:
      ACTUALIZADA  → Layer K existe y es AZUL (ACI 5)
                     Y  Layer K2 existe y es VERDE (ACI 3)
      VIEJA        → Ninguna condición se cumple
      INCOMPLETA   → Solo una condición se cumple, o existe pero con color incorrecto
    """
    nombre_k  = ""
    nombre_k2 = ""
    color_k   = ""
    color_k2  = ""
    k_azul    = False
    k2_verde  = False
    tiene_k   = False
    tiene_k2  = False

    for nombre, info in layers_dict.items():
        upper = nombre.upper().strip()
        if upper == "K":
            tiene_k  = True
            nombre_k = nombre
            color_k  = info["color_nombre"]
            if info["color_aci"] == COLOR_AZUL_ACI:
                k_azul = True
        if upper == "K2":
            tiene_k2  = True
            nombre_k2 = nombre
            color_k2  = info["color_nombre"]
            if info["color_aci"] == COLOR_VERDE_ACI:
                k2_verde = True

    if k_azul and k2_verde:
        estado = "ACTUALIZADA"
    elif not tiene_k and not tiene_k2:
        estado = "VIEJA"
    else:
        estado = "INCOMPLETA"

    return {
        "estado":    estado,
        "tiene_k":   tiene_k,
        "nombre_k":  nombre_k,
        "color_k":   color_k,
        "k_azul":    k_azul,
        "tiene_k2":  tiene_k2,
        "nombre_k2": nombre_k2,
        "color_k2":  color_k2,
        "k2_verde":  k2_verde,
    }


# ──────────────────────────────────────────────────────────
# RECOLECCIÓN DE DWGs
# ──────────────────────────────────────────────────────────
def recolectar_dwgs(ruta_artes):
    """
    Recolecta DWGs desde:
      1. Archivos sueltos en ruta_artes
      2. Subcarpetas cuyo nombre contenga "BN" (sin importar mayúsculas)
         - Si dentro de esa carpeta hay una subcarpeta con "OBSOLETO" en el nombre → la ignora
    """
    dwgs = []
    try:
        items = os.listdir(ruta_artes)
    except Exception as e:
        log.warn(f"No se puede leer {ruta_artes}: {e}")
        return dwgs

    for item in items:
        ruta_item = os.path.join(ruta_artes, item)

        if os.path.isfile(ruta_item) and item.lower().endswith(".dwg"):
            dwgs.append(("suelto", ruta_item))

        elif os.path.isdir(ruta_item) and "BN" in item.upper():
            if "OBSOLETO" in item.upper():
                log.info(f"      [SKIP] Carpeta OBSOLETOS ignorada: {item}")
                continue
            try:
                for sub in os.listdir(ruta_item):
                    ruta_sub = os.path.join(ruta_item, sub)
                    if os.path.isdir(ruta_sub) and "OBSOLETO" in sub.upper():
                        log.info(f"      [SKIP] Subcarpeta OBSOLETOS ignorada: {sub}")
                        continue
                    if os.path.isfile(ruta_sub) and sub.lower().endswith(".dwg"):
                        dwgs.append((f"BN/{item}", ruta_sub))
            except Exception as e:
                log.warn(f"Error leyendo carpeta BN {item}: {e}")

    return sorted(set(dwgs), key=lambda x: x[1])


# ──────────────────────────────────────────────────────────
# ESCANEO PRINCIPAL
# ──────────────────────────────────────────────────────────
def escanear(ruta_base, motor):
    """
    Navega: Vehículo > Modelo > Versión > ARTES
    Para cada DWG → abre en AutoCAD, inspecciona layers, evalúa estado.
    Retorna: {vehiculo: [filas]}
    """
    datos = defaultdict(list)

    if not os.path.exists(ruta_base):
        log.error(f"Ruta base no accesible: {ruta_base}")
        return datos

    vehiculos = sorted(
        d for d in os.listdir(ruta_base)
        if os.path.isdir(os.path.join(ruta_base, d))
    )
    log.info(f"Vehículos encontrados: {len(vehiculos)}")

    for vehiculo in vehiculos:
        ruta_vehiculo = os.path.join(ruta_base, vehiculo)
        log.info("=" * 70)
        log.info(f"VEHÍCULO: {vehiculo}")

        modelos = sorted(
            d for d in os.listdir(ruta_vehiculo)
            if os.path.isdir(os.path.join(ruta_vehiculo, d))
        )

        for modelo in modelos:
            ruta_modelo = os.path.join(ruta_vehiculo, modelo)

            versiones = sorted(
                d for d in os.listdir(ruta_modelo)
                if os.path.isdir(os.path.join(ruta_modelo, d))
            )

            for version in versiones:
                ruta_version = os.path.join(ruta_modelo, version)

                # Buscar carpeta ARTES (case insensitive)
                ruta_artes = None
                try:
                    for sub in os.listdir(ruta_version):
                        if sub.upper() == "ARTES" and \
                           os.path.isdir(os.path.join(ruta_version, sub)):
                            ruta_artes = os.path.join(ruta_version, sub)
                            break
                except Exception:
                    continue

                if not ruta_artes:
                    continue

                dwgs = recolectar_dwgs(ruta_artes)

                if not dwgs:
                    log.info(f"  {modelo}/{version}: sin DWGs en ARTES")
                    continue

                log.info(f"  {modelo}/{version}: {len(dwgs)} archivo(s)")

                for origen, dwg_path in dwgs:
                    nombre_archivo = os.path.basename(dwg_path)
                    log.info(f"    [{origen}] {nombre_archivo}")

                    fila = {
                        "vehiculo":      vehiculo,
                        "modelo":        modelo,
                        "version":       version,
                        "archivo":       nombre_archivo,
                        "origen":        origen,
                        "ruta":          dwg_path,
                        "estado":        "ERROR",
                        "tiene_k":       False,
                        "nombre_k":      "",
                        "color_k":       "",
                        "k_azul":        False,
                        "tiene_k2":      False,
                        "nombre_k2":     "",
                        "color_k2":      "",
                        "k2_verde":      False,
                        "total_layers":  0,
                        "lista_layers":  "",
                        "detalle_error": "",
                    }

                    doc = motor.abrir(dwg_path)
                    if doc is None:
                        fila["detalle_error"] = "No se pudo abrir el archivo"
                        log.warn(f"      No se pudo abrir: {nombre_archivo}")
                    else:
                        try:
                            layers_dict = motor.inspeccionar_layers(doc)
                            evaluacion  = evaluar_estado(layers_dict)

                            fila.update(evaluacion)
                            fila["total_layers"] = len(layers_dict)
                            fila["lista_layers"] = " | ".join(
                                f"{n}({info['color_nombre']})"
                                for n, info in sorted(layers_dict.items())
                            )

                            icono = {"ACTUALIZADA": "✓", "VIEJA": "✗", "INCOMPLETA": "!"}.get(
                                evaluacion["estado"], "?"
                            )
                            log.info(f"      {icono} {evaluacion['estado']}  "
                                     f"K={'SI(azul)' if evaluacion['k_azul'] else ('SI(no azul)' if evaluacion['tiene_k'] else 'NO')}  "
                                     f"K2={'SI(verde)' if evaluacion['k2_verde'] else ('SI(no verde)' if evaluacion['tiene_k2'] else 'NO')}")
                        except Exception as e:
                            fila["detalle_error"] = str(e)[:80]
                            log.error(f"      Error analizando: {e}")
                        finally:
                            motor.cerrar(doc)

                    time.sleep(0.2)
                    datos[vehiculo].append(fila)

    return datos


# ──────────────────────────────────────────────────────────
# GENERACIÓN DE EXCEL
# ──────────────────────────────────────────────────────────
ESTADO_CONFIG = {
    "ACTUALIZADA": {"bg": C["actualizada"], "label": "ACTUALIZADA"},
    "VIEJA":       {"bg": C["vieja"],       "label": "VIEJA / OBSOLETA"},
    "INCOMPLETA":  {"bg": C["incompleta"],  "label": "INCOMPLETA"},
    "ERROR":       {"bg": C["error"],       "label": "ERROR AL ABRIR"},
}

HEADERS_DETALLE = [
    "#", "VEHÍCULO", "MODELO", "VERSIÓN", "ARCHIVO", "ORIGEN",
    "ESTADO", "TIENE K", "NOMBRE K", "COLOR K", "K ES AZUL",
    "TIENE K2", "NOMBRE K2", "COLOR K2", "K2 ES VERDE",
    "TOTAL LAYERS", "LAYERS PRESENTES", "RUTA COMPLETA", "DETALLE ERROR"
]

ANCHOS_DETALLE = [5, 25, 25, 20, 35, 12, 18, 10, 12, 12, 12,
                   10, 12, 12, 12, 14, 60, 80, 40]


def _fila_titulo(ws, texto, cols, fila=1, bg=None):
    bg = bg or C["hdr"]
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=cols)
    c = ws.cell(fila, 1, texto)
    c.font      = Font(name="Arial", size=13, bold=True, color=C["white"])
    c.fill      = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[fila].height = 30


def _fila_meta(ws, texto, cols, fila=2):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=cols)
    c = ws.cell(fila, 1, texto)
    c.font      = Font(name="Arial", size=9, italic=True, color="555555")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[fila].height = 16


def escribir_headers(ws, headers, fila_inicio=3, bg=None):
    bg = bg or C["titulo"]
    for col, h in enumerate(headers, 1):
        _cel_header(ws.cell(fila_inicio, col, h), bg=bg)
    ws.row_dimensions[fila_inicio].height = 22


def escribir_fila_detalle(ws, fila_excel, num, f):
    alt  = (fila_excel % 2 == 0)
    cfg  = ESTADO_CONFIG.get(f["estado"], ESTADO_CONFIG["ERROR"])
    bg_estado = cfg["bg"]

    vals = [
        num,
        f["vehiculo"], f["modelo"], f["version"],
        f["archivo"],  f["origen"],
        cfg["label"],
        "SÍ" if f["tiene_k"]  else "NO",
        f["nombre_k"], f["color_k"],
        "SÍ ✓" if f["k_azul"]  else "NO ✗",
        "SÍ" if f["tiene_k2"] else "NO",
        f["nombre_k2"], f["color_k2"],
        "SÍ ✓" if f["k2_verde"] else "NO ✗",
        f["total_layers"],
        f["lista_layers"],
        f["ruta"],
        f["detalle_error"],
    ]

    for col, val in enumerate(vals, 1):
        c = ws.cell(fila_excel, col, val)
        if col == 7:  # columna ESTADO
            c.font      = Font(name="Arial", size=9, bold=True)
            c.fill      = PatternFill("solid", start_color=bg_estado)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border    = _border
        elif col in (8, 12):  # TIENE K / TIENE K2 (SÍ/NO existencia)
            tiene = val == "SÍ"
            _cel_dato(c, alt=alt, center=True,
                      bg="C6EFCE" if tiene else "FFCCCC" if not alt else None)
        elif col in (11, 15):  # K ES AZUL / K2 ES VERDE
            ok_ = "✓" in val
            _cel_dato(c, alt=alt, center=True,
                      bg="C6EFCE" if ok_ else "FFCCCC")
        elif col == 1:
            _cel_dato(c, alt=alt, center=True)
        elif col in (17,):  # ruta
            c.font      = Font(name="Courier New", size=8, color="0070C0")
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border    = _border
            if alt:
                c.fill = PatternFill("solid", start_color=C["alt"])
        else:
            _cel_dato(c, alt=alt)


def crear_excel(datos, ruta_salida):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    todos = [f for filas in datos.values() for f in filas]
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")

    # ── Conteos globales ──────────────────────────────────
    tot_act  = sum(1 for f in todos if f["estado"] == "ACTUALIZADA")
    tot_viej = sum(1 for f in todos if f["estado"] == "VIEJA")
    tot_inc  = sum(1 for f in todos if f["estado"] == "INCOMPLETA")
    tot_err  = sum(1 for f in todos if f["estado"] == "ERROR")

    # ══════════════════════════════════════════════════════
    # HOJA 1: RESUMEN GENERAL
    # ══════════════════════════════════════════════════════
    ws_res = wb.create_sheet("RESUMEN GENERAL")
    ws_res.sheet_view.showGridLines = False

    _fila_titulo(ws_res, "AUDITORÍA LAYERS K / K2 — AGP PLANOS TÉCNICOS", 8)
    _fila_meta(ws_res,
               f"Generado: {fecha}   |   "
               f"Total archivos DWG: {len(todos)}   |   "
               f"Actualizadas: {tot_act}  |  Viejas: {tot_viej}  |  "
               f"Incompletas: {tot_inc}  |  Errores: {tot_err}",
               8)

    HEADERS_RES = [
        "VEHÍCULO", "TOTAL DWG",
        "ACTUALIZADAS", "% ACT.",
        "VIEJAS", "% VIEJ.",
        "INCOMPLETAS", "ERRORES"
    ]
    escribir_headers(ws_res, HEADERS_RES)

    fila = 4
    for vehiculo, filas in sorted(datos.items()):
        act  = sum(1 for f in filas if f["estado"] == "ACTUALIZADA")
        viej = sum(1 for f in filas if f["estado"] == "VIEJA")
        inc  = sum(1 for f in filas if f["estado"] == "INCOMPLETA")
        err  = sum(1 for f in filas if f["estado"] == "ERROR")
        tot  = len(filas)
        alt  = (fila % 2 == 0)

        pct_act  = f"{act/tot*100:.1f}%" if tot else "0%"
        pct_viej = f"{viej/tot*100:.1f}%" if tot else "0%"

        _cel_dato(ws_res.cell(fila, 1, vehiculo), alt=alt, bold=True)
        _cel_dato(ws_res.cell(fila, 2, tot),  alt=alt, center=True)
        c_act = ws_res.cell(fila, 3, act)
        _cel_dato(c_act, center=True, bg=C["actualizada"] if act > 0 else None)
        _cel_dato(ws_res.cell(fila, 4, pct_act),  center=True, alt=alt)
        c_viej = ws_res.cell(fila, 5, viej)
        _cel_dato(c_viej, center=True, bg=C["vieja"] if viej > 0 else None)
        _cel_dato(ws_res.cell(fila, 6, pct_viej), center=True, alt=alt)
        c_inc = ws_res.cell(fila, 7, inc)
        _cel_dato(c_inc, center=True, bg=C["incompleta"] if inc > 0 else None)
        c_err = ws_res.cell(fila, 8, err)
        _cel_dato(c_err, center=True, bg=C["error"] if err > 0 else None)

        fila += 1

    # Fila totales
    _cel_dato(ws_res.cell(fila, 1, "TOTAL GENERAL"), bold=True)
    _cel_dato(ws_res.cell(fila, 2, len(todos)),  bold=True, center=True)
    _cel_dato(ws_res.cell(fila, 3, tot_act),  bold=True, center=True, bg=C["actualizada"])
    pct_g = f"{tot_act/len(todos)*100:.1f}%" if todos else "0%"
    _cel_dato(ws_res.cell(fila, 4, pct_g), bold=True, center=True)
    _cel_dato(ws_res.cell(fila, 5, tot_viej), bold=True, center=True, bg=C["vieja"])
    pct_gv = f"{tot_viej/len(todos)*100:.1f}%" if todos else "0%"
    _cel_dato(ws_res.cell(fila, 6, pct_gv), bold=True, center=True)
    _cel_dato(ws_res.cell(fila, 7, tot_inc),  bold=True, center=True, bg=C["incompleta"])
    _cel_dato(ws_res.cell(fila, 8, tot_err),  bold=True, center=True, bg=C["error"] if tot_err else None)

    anchos_res = [35, 12, 15, 10, 12, 10, 14, 10]
    for i, w in enumerate(anchos_res, 1):
        ws_res.column_dimensions[
            openpyxl.utils.get_column_letter(i)
        ].width = w
    ws_res.freeze_panes = "A4"

    # ══════════════════════════════════════════════════════
    # HOJA 2: TODOS LOS ARCHIVOS
    # ══════════════════════════════════════════════════════
    ws_todos = wb.create_sheet("TODOS LOS ARCHIVOS")
    ws_todos.sheet_view.showGridLines = False

    _fila_titulo(ws_todos, "DETALLE COMPLETO — TODOS LOS VEHÍCULOS", len(HEADERS_DETALLE))
    _fila_meta(ws_todos,
               f"Generado: {fecha}   |   Total: {len(todos)} archivos   |   "
               f"Ruta base: {RUTA_BASE}",
               len(HEADERS_DETALLE))
    escribir_headers(ws_todos, HEADERS_DETALLE)

    for i, f in enumerate(todos, 1):
        escribir_fila_detalle(ws_todos, i + 3, i, f)

    for i, w in enumerate(ANCHOS_DETALLE, 1):
        ws_todos.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws_todos.freeze_panes = "A4"

    # ══════════════════════════════════════════════════════
    # HOJA 3+: UNA POR VEHÍCULO
    # ══════════════════════════════════════════════════════
    for vehiculo, filas in sorted(datos.items()):
        nombre_hoja = (vehiculo[:31]
                       .replace("/", "-").replace("\\", "-")
                       .replace("*", "").replace("?", "")
                       .replace("[", "").replace("]", "").replace(":", ""))
        ws_v = wb.create_sheet(nombre_hoja)
        ws_v.sheet_view.showGridLines = False

        act  = sum(1 for f in filas if f["estado"] == "ACTUALIZADA")
        viej = sum(1 for f in filas if f["estado"] == "VIEJA")
        inc  = sum(1 for f in filas if f["estado"] == "INCOMPLETA")
        err  = sum(1 for f in filas if f["estado"] == "ERROR")

        _fila_titulo(ws_v, f"VEHÍCULO: {vehiculo}", len(HEADERS_DETALLE))
        _fila_meta(ws_v,
                   f"Total DWG: {len(filas)}   |   "
                   f"Actualizadas: {act}   Viejas: {viej}   "
                   f"Incompletas: {inc}   Errores: {err}   |   {fecha}",
                   len(HEADERS_DETALLE))
        escribir_headers(ws_v, HEADERS_DETALLE, bg=C["titulo"])

        num = 0
        for f in filas:
            num += 1
            escribir_fila_detalle(ws_v, num + 3, num, f)

        for i, w in enumerate(ANCHOS_DETALLE, 1):
            ws_v.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        ws_v.freeze_panes = "A4"

    # ══════════════════════════════════════════════════════
    # HOJA: SOLO PROBLEMÁTICAS (VIEJA + INCOMPLETA + ERROR)
    # ══════════════════════════════════════════════════════
    problematicas = [f for f in todos if f["estado"] != "ACTUALIZADA"]
    if problematicas:
        ws_p = wb.create_sheet("REQUIEREN ATENCIÓN")
        ws_p.sheet_view.showGridLines = False

        _fila_titulo(ws_p, f"ARTES QUE REQUIEREN ATENCIÓN ({len(problematicas)} archivos)",
                     len(HEADERS_DETALLE), bg="C00000")
        _fila_meta(ws_p,
                   f"Incluye: VIEJAS ({tot_viej})  |  INCOMPLETAS ({tot_inc})  |  "
                   f"ERRORES ({tot_err})   |   Generado: {fecha}",
                   len(HEADERS_DETALLE))
        escribir_headers(ws_p, HEADERS_DETALLE, bg="C00000")

        for i, f in enumerate(problematicas, 1):
            escribir_fila_detalle(ws_p, i + 3, i, f)

        for i, w in enumerate(ANCHOS_DETALLE, 1):
            ws_p.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        ws_p.freeze_panes = "A4"

    wb.save(ruta_salida)
    log.info(f"Excel guardado: {ruta_salida}")


# ──────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────
def main():
    log.info("=" * 70)
    log.info("  AUDITORÍA LAYERS K / K2 — AGP PLANOS TÉCNICOS")
    log.info("=" * 70)
    log.info(f"Ruta base : {RUTA_BASE}")
    log.info(f"Salida    : {ARCHIVO_EXCEL}")
    log.info("-" * 70)

    log.info("\n[1/3] Conectando a AutoCAD...")
    motor = AutoCADMotor()

    log.info("\n[2/3] Escaneando archivos DWG...")
    t0 = time.time()
    datos = escanear(RUTA_BASE, motor)
    duracion = time.time() - t0

    motor.quit()

    if not datos:
        log.warn("No se encontraron archivos DWG. Verifica conexión de red y ruta.")
        return

    todos = [f for filas in datos.values() for f in filas]
    act   = sum(1 for f in todos if f["estado"] == "ACTUALIZADA")
    viej  = sum(1 for f in todos if f["estado"] == "VIEJA")
    inc   = sum(1 for f in todos if f["estado"] == "INCOMPLETA")
    err   = sum(1 for f in todos if f["estado"] == "ERROR")

    log.info(f"\n[3/3] Generando Excel ({len(todos)} archivos)...")
    crear_excel(datos, ARCHIVO_EXCEL)

    log.info("\n" + "=" * 70)
    log.info("  RESUMEN FINAL")
    log.info("=" * 70)
    log.info(f"  Total archivos DWG analizados : {len(todos)}")
    log.info(f"  ACTUALIZADAS  (K + K2 verde)  : {act}")
    log.info(f"  VIEJAS        (sin K ni K2)   : {viej}")
    log.info(f"  INCOMPLETAS   (solo una cond.) : {inc}")
    log.info(f"  ERRORES al abrir              : {err}")
    log.info(f"  Tiempo total                  : {duracion:.1f} s")
    log.info(f"  Excel generado                : {ARCHIVO_EXCEL}")
    log.info("=" * 70)


if __name__ == "__main__":
    main()
