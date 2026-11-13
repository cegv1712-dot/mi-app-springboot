# -*- coding: utf-8 -*-
"""
Panel de Gastos 2026 — servidor local.

Lee el Excel "CONTROL DE GASTOS 2026 CLOUD.xlsx" y expone una web con:
  - Resumen mensual y anual (totales, disponible, ahorro)
  - Comparativa entre meses
  - Listado detallado con búsqueda y filtro por categoría
  - Añadir gasto (escribe en el Excel, con backup previo)

Uso:
    python servidor.py [puerto]

Por defecto usa el puerto 8720 y abre el navegador automáticamente.
"""

from __future__ import annotations

import ast
import json
import re
import shutil
import sys
import threading
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import openpyxl

VERSION = "1.0"
APP_DIR = Path(__file__).resolve().parent
WEB_DIR = APP_DIR / "web"
CONFIG_FILE = APP_DIR / "config.json"
PUERTO_DEFECTO = 8720

MESES_ES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
]

EXCEL_LOCK = threading.Lock()


# ============================================================
# UTILIDADES
# ============================================================

def _norm(texto: str) -> str:
    return " ".join(str(texto or "").strip().upper().split())


def _a_numero(valor):
    """Convierte un valor de celda a float o devuelve None."""
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace("€", "").replace(" ", "")
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _a_dia(valor) -> int:
    """Convierte el valor de la columna DIA a entero 1-31 (0 si no se puede)."""
    if isinstance(valor, bool):
        return 0
    if isinstance(valor, int):
        return valor if 1 <= valor <= 31 else 0
    if isinstance(valor, float):
        return int(valor) if valor.is_integer() and 1 <= valor <= 31 else 0
    if valor is None:
        return 0
    m = re.match(r"^\s*(\d{1,2})", str(valor))
    if not m:
        return 0
    d = int(m.group(1))
    return d if 1 <= d <= 31 else 0


def _eval_ast(node) -> float:
    """Evalúa un AST aritmético permitido (solo números y + - * /)."""
    if isinstance(node, ast.Expression):
        return _eval_ast(node.body)
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
        return node.value
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
        v = _eval_ast(node.operand)
        return v if isinstance(node.op, ast.UAdd) else -v
    if isinstance(node, ast.BinOp):
        left = _eval_ast(node.left)
        right = _eval_ast(node.right)
        if isinstance(node.op, ast.Add):
            return left + right
        if isinstance(node.op, ast.Sub):
            return left - right
        if isinstance(node.op, ast.Mult):
            return left * right
        if isinstance(node.op, ast.Div):
            return left / right
    raise ValueError("Expresión no soportada")


def _eval_formula_simple(formula):
    """Evalúa fórmulas aritméticas simples (ej: '=19.48+28.97'). Devuelve float o None."""
    if not isinstance(formula, str):
        return None
    s = formula.strip()
    if not s.startswith("="):
        return None
    s = s[1:].strip().replace(",", ".")
    if not s:
        return None
    try:
        tree = ast.parse(s, mode="eval")
    except SyntaxError:
        return None
    try:
        resultado = _eval_ast(tree)
    except Exception:
        return None
    if isinstance(resultado, (int, float)) and not isinstance(resultado, bool):
        return float(resultado)
    return None


# ============================================================
# LOCALIZACIÓN DEL EXCEL
# ============================================================

def _cargar_config() -> dict:
    if CONFIG_FILE.is_file():
        try:
            datos = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if isinstance(datos, dict):
                return datos
        except Exception:
            pass
    return {}


def _guardar_config(datos: dict) -> None:
    try:
        CONFIG_FILE.write_text(
            json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass


def _candidatos_excel() -> list[Path]:
    home = Path.home()
    bases = [
        home / "OneDrive" / "Desktop", home / "OneDrive" / "Escritorio",
        home / "Desktop", home / "Escritorio",
        home / "OneDrive" / "Documents", home / "OneDrive" / "Documentos",
        home / "Documents", home / "Documentos",
    ]
    candidatos: list[Path] = []
    vistos: set[str] = set()
    for base in bases:
        if not base.is_dir():
            continue
        try:
            for f in base.iterdir():
                if not f.is_file() or f.suffix.lower() != ".xlsx":
                    continue
                nombre = f.name.lower()
                if "gasto" not in nombre or "2026" not in nombre:
                    continue
                if "mousewithoutborders" in nombre or "screencaptures" in nombre:
                    continue
                clave = str(f.resolve())
                if clave in vistos:
                    continue
                vistos.add(clave)
                candidatos.append(f)
        except OSError:
            continue

    def _prioridad(p: Path) -> tuple:
        nombre = p.name.lower()
        return (
            0 if "cloud" in nombre else 1,
            0 if nombre == "control de gastos 2026 cloud.xlsx" else 1,
            -(p.stat().st_mtime if p.exists() else 0.0),
            nombre,
        )

    candidatos.sort(key=_prioridad)
    return candidatos


def encontrar_excel() -> Path | None:
    config = _cargar_config()
    ruta_txt = config.get("excel_path", "")
    if ruta_txt:
        p = Path(ruta_txt)
        if p.is_file() and p.suffix.lower() == ".xlsx":
            return p
    candidatos = _candidatos_excel()
    if candidatos:
        _guardar_config({"excel_path": str(candidatos[0])})
        return candidatos[0]
    return None


# ============================================================
# LECTURA (PARSEO) DEL EXCEL
# ============================================================

def _encontrar_hoja(wb, nombre_objetivo: str, mes: str) -> str | None:
    objetivo = _norm(nombre_objetivo)
    mes_norm = mes.strip().upper()
    for s in wb.sheetnames:
        if _norm(s) == objetivo:
            return s
    palabras = objetivo.split()
    for s in wb.sheetnames:
        sn = _norm(s)
        if all(p in sn for p in palabras):
            return s
    for s in wb.sheetnames:
        if _norm(s).startswith(mes_norm):
            return s
    return None


def _parsear_hoja_mes(ws_v, ws_f, mes: str, anio: int, nombre_hoja: str) -> dict:
    final_row = None
    total_row = None
    disponible_row = None
    diferencia_row = None
    ahorro_row = None

    for r in range(1, ws_v.max_row + 1):
        c = ws_v.cell(r, 3).value
        if c is None:
            continue
        cn = _norm(c)
        if final_row is None and cn in ("FIN", "FINAL"):
            final_row = r
        if total_row is None and "TOTAL GASTOS" in cn:
            total_row = r
        elif disponible_row is None and "TOTAL DISPONIBLE" in cn:
            disponible_row = r
        elif diferencia_row is None and "DIFERENCIA" in cn:
            diferencia_row = r
        elif ahorro_row is None and "AHORRO" in cn:
            ahorro_row = r

    limite = final_row if final_row is not None else (
        total_row if total_row is not None else ws_v.max_row + 1
    )

    gastos: list[dict] = []
    for r in range(4, limite):
        c = ws_v.cell(r, 3).value
        if c is None:
            continue
        cn = _norm(c)
        if cn in ("FIN", "FINAL"):
            continue
        importe = _a_numero(ws_v.cell(r, 4).value)
        if importe is None:
            # Algunos importes están como fórmula simple (ej: =19.48+28.97)
            importe = _eval_formula_simple(ws_f.cell(r, 4).value)
        if importe is None or importe <= 0:
            continue
        concepto = str(c).strip()
        if not concepto:
            continue
        dia = _a_dia(ws_v.cell(r, 2).value)
        categoria = str(ws_v.cell(r, 5).value or "").strip().upper() or "EXTRA"
        gastos.append({
            "dia": dia,
            "concepto": concepto,
            "importe": round(importe, 2),
            "categoria": categoria,
        })

    categorias: dict[str, float] = {}
    for g in gastos:
        categorias[g["categoria"]] = round(
            categorias.get(g["categoria"], 0.0) + g["importe"], 2
        )

    def _val(r, col=4):
        if r is None:
            return None
        v = _a_numero(ws_v.cell(r, col).value)
        return round(v, 2) if v is not None else None

    suma_detalle = round(sum(g["importe"] for g in gastos), 2)
    disponible = _val(disponible_row)
    # La diferencia se calcula (disponible - gastos); no depende de la fórmula
    # cacheada, que openpyxl pierde al guardar.
    diferencia = round(disponible - suma_detalle, 2) if disponible is not None else None
    # AHORRO FINAL: fórmula en la mayoría de hojas. Su valor cacheado se pierde
    # tras guardar con openpyxl (devuelve None) hasta que Excel recalcula.
    ahorro = _val(ahorro_row)

    return {
        "nombre_hoja": nombre_hoja.strip(),
        "mes": mes,
        "anio": anio,
        "idx": MESES_ES.index(mes) + 1,
        "gastos_total": suma_detalle,
        "disponible": disponible,
        "diferencia": diferencia,
        "ahorro": ahorro,
        "suma_detalle": suma_detalle,
        "n_expenses": len(gastos),
        "categorias": categorias,
        "gastos": gastos,
    }


def _parsear_anual(wb) -> dict:
    ws = None
    for n in wb.sheetnames:
        nn = _norm(n)
        if "ANUAL" in nn or "ANUALES" in nn:
            ws = wb[n]
            break
    if ws is None:
        return {"ingresos_mes": {}, "otro_gasto_mes": {}, "ahorro_mes": {}, "ahorro_total": None}

    # INGRESOS y OTRO GASTO por mes (columnas B y E, valores estáticos)
    ingresos_mes: dict[str, float | None] = {}
    otro_gasto_mes: dict[str, float | None] = {}
    for r in range(2, 14):
        a = ws.cell(r, 1).value
        if a is None:
            continue
        m = _norm(a)
        if m not in MESES_ES:
            continue
        ingresos = _a_numero(ws.cell(r, 2).value)
        ingresos_mes[m] = round(ingresos, 2) if ingresos is not None else None
        otro = _a_numero(ws.cell(r, 5).value)
        otro_gasto_mes[m] = round(otro, 2) if otro is not None else None

    # "AHORRO {MES}" (B17:B28): valores estáticos que sobreviven al guardado
    ahorro_mes: dict[str, float] = {}
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is None:
            continue
        aa = _norm(a)
        m = re.match(r"^AHORRO\s+(.+)$", aa)
        if not m:
            continue
        mes_nombre = m.group(1).strip()
        if mes_nombre not in MESES_ES:
            continue
        v = _a_numero(ws.cell(r, 2).value)
        if v is not None:
            ahorro_mes[mes_nombre] = round(v, 2)

    ahorro_total = round(sum(ahorro_mes.values()), 2) if ahorro_mes else None

    return {
        "ingresos_mes": ingresos_mes,
        "otro_gasto_mes": otro_gasto_mes,
        "ahorro_mes": ahorro_mes,
        "ahorro_total": ahorro_total,
    }


def parsear_libro(ruta: Path) -> dict:
    wb_v = openpyxl.load_workbook(ruta, data_only=True)
    wb_f = openpyxl.load_workbook(ruta, data_only=False)

    meses: list[dict] = []
    for nombre in wb_v.sheetnames:
        tokens = nombre.strip().upper().split()
        if not tokens:
            continue
        mes = tokens[0]
        if mes not in MESES_ES:
            continue
        anio = None
        for t in tokens:
            if re.fullmatch(r"\d{4}", t):
                anio = int(t)
        if anio is None:
            continue
        meses.append(_parsear_hoja_mes(wb_v[nombre], wb_f[nombre], mes, anio, nombre))

    meses.sort(key=lambda m: (m["anio"], m["idx"]))

    anual = _parsear_anual(wb_v)

    # Cruce de ingresos y ahorro (anual, estáticos) sobre cada mes
    for m in meses:
        m["ingresos"] = anual["ingresos_mes"].get(m["mes"])
        m["otro_gasto"] = anual["otro_gasto_mes"].get(m["mes"])
        m["ahorro_anual"] = anual["ahorro_mes"].get(m["mes"])

    # Totales anuales desde fuentes fiables (detalle y valores estáticos)
    gastos_total_anio = round(sum(m["gastos_total"] or 0 for m in meses), 2)
    disponible_total_anio = round(sum(m["disponible"] or 0 for m in meses), 2)
    ingresos_total_anio = round(
        sum(v for v in anual["ingresos_mes"].values() if v is not None), 2
    )

    categorias_disponibles = sorted({
        c for m in meses for c in m["categorias"].keys()
    })

    return {
        "excel_path": str(ruta),
        "excel_nombre": ruta.name,
        "excel_mtime": datetime.fromtimestamp(ruta.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "meses": meses,
        "categorias_disponibles": categorias_disponibles,
        "anual": {
            "ingresos_total": ingresos_total_anio,
            "gastos_total": gastos_total_anio,
            "disponible_total": disponible_total_anio,
            "ahorro_total": anual["ahorro_total"],
            "n_meses": len(meses),
        },
    }


# ============================================================
# ESCRITURA: AÑADIR GASTO
# ============================================================

def _shift_formula_refs(formula: str, from_row: int) -> str:
    def _repl(m):
        pre = m.group(1) or ""
        col = m.group(2)
        pre2 = m.group(3) or ""
        row = int(m.group(4))
        if row >= from_row:
            row += 1
        return f"{pre}{col}{pre2}{row}"
    return re.sub(r"(\$?)([A-Z]+)(\$?)(\d+)", _repl, formula)


def apuntar_gasto(ruta: Path, importe: float, concepto: str,
                  categoria: str, fecha: datetime) -> dict:
    importe = float(importe)
    if importe <= 0:
        raise ValueError("El importe debe ser mayor que cero.")
    concepto = (concepto or "").strip()
    if not concepto:
        raise ValueError("El concepto no puede estar vacío.")
    categoria = (categoria or "EXTRA").strip().upper() or "EXTRA"

    nombre_mes = MESES_ES[fecha.month - 1]
    nombre_hoja = f"{nombre_mes} {fecha.year}"

    with EXCEL_LOCK:
        try:
            wb = openpyxl.load_workbook(ruta)
        except PermissionError:
            raise RuntimeError(
                "El Excel está abierto en otro programa (Excel/OneDrive). "
                "Ciérralo y vuelve a intentarlo."
            )
        except Exception as error:
            raise RuntimeError(f"No pude abrir el Excel: {error}")

        hoja_real = _encontrar_hoja(wb, nombre_hoja, nombre_mes)
        if hoja_real is None:
            raise RuntimeError(
                f"No existe una hoja para '{nombre_hoja}'. "
                "Créala en Excel antes de añadir el gasto."
            )
        ws = wb[hoja_real]

        final_row = None
        total_row = None
        for r in range(1, ws.max_row + 1):
            c = ws.cell(r, 3).value
            if c is None:
                continue
            cn = _norm(c)
            if final_row is None and cn in ("FIN", "FINAL"):
                final_row = r
                break
            if total_row is None and "TOTAL GASTOS" in cn:
                total_row = r

        insertar = False
        if final_row is not None:
            ws.insert_rows(final_row)
            fila_escritura = final_row
            insertar = True
        else:
            if total_row is None:
                total_row = ws.max_row + 1
            ultima_con_datos = 3
            for r in range(4, total_row):
                c = ws.cell(r, 3).value
                d = ws.cell(r, 4).value
                if (c is not None and str(c).strip()) or d is not None:
                    ultima_con_datos = r
            fila_escritura = ultima_con_datos + 1
            if fila_escritura >= total_row:
                ws.insert_rows(total_row)
                fila_escritura = total_row
                insertar = True

        ws.cell(fila_escritura, 1, nombre_mes)
        ws.cell(fila_escritura, 2, fecha.day)
        ws.cell(fila_escritura, 3, concepto.upper())
        ws.cell(fila_escritura, 4, importe)
        ws.cell(fila_escritura, 5, categoria)

        if insertar:
            for r in range(1, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    celda = ws.cell(r, c)
                    if celda.value and isinstance(celda.value, str) and celda.value.startswith("="):
                        celda.value = _shift_formula_refs(celda.value, fila_escritura)

        # Recalcula la fórmula de TOTAL GASTOS para que incluya la nueva fila
        for r in range(1, ws.max_row + 1):
            c = ws.cell(r, 3).value
            if c is not None and "TOTAL GASTOS" in _norm(c):
                ws.cell(r, 4, f"=SUM(D4:D{fila_escritura})")
                break

        backup = ruta.with_suffix(".bak.xlsx")
        try:
            shutil.copy2(ruta, backup)
        except Exception:
            pass

        try:
            wb.save(ruta)
        except PermissionError:
            raise RuntimeError(
                "No se pudo guardar: el Excel está abierto en otro programa. "
                "Ciérralo y reintenta."
            )
        except Exception as error:
            raise RuntimeError(f"Error guardando el Excel: {error}")

    return {
        "hoja": hoja_real.strip(),
        "fila": fila_escritura,
        "concepto": concepto.upper(),
        "importe": round(importe, 2),
        "categoria": categoria,
        "backup": backup.name,
    }


# ============================================================
# SERVIDOR HTTP
# ============================================================

MIME = {
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".png": "image/png",
    ".ico": "image/x-icon",
}

STATIC_FILES = {
    "/": ("index.html", "text/html; charset=utf-8"),
    "/index.html": ("index.html", "text/html; charset=utf-8"),
    "/app.js": ("app.js", "application/javascript; charset=utf-8"),
    "/styles.css": ("styles.css", "text/css; charset=utf-8"),
}


class Handler(BaseHTTPRequestHandler):
    server_version = f"PanelGastos/{VERSION}"

    def _send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_static(self, rel: str, ctype: str):
        path = (WEB_DIR / rel)
        try:
            data = path.read_bytes()
        except OSError:
            self._send_json({"error": "No encontrado"}, 404)
            return
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        path = self.path.split("?", 1)[0]
        if path in STATIC_FILES:
            rel, ctype = STATIC_FILES[path]
            self._send_static(rel, ctype)
            return
        if path == "/api/datos":
            self._api_datos()
            return
        if path == "/api/excel":
            self._api_excel()
            return
        self._send_json({"error": "Ruta no encontrada"}, 404)

    def do_POST(self):
        path = self.path.split("?", 1)[0]
        if path == "/api/gasto":
            self._api_gasto()
            return
        self._send_json({"error": "Ruta no encontrada"}, 404)

    def _read_json_body(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
        except ValueError:
            length = 0
        raw = self.rfile.read(length) if length else b"{}"
        try:
            return json.loads(raw.decode("utf-8"))
        except Exception:
            return {}

    def _api_datos(self):
        global RUTA_EXCEL
        ruta = RUTA_EXCEL or encontrar_excel()
        if ruta is None or not ruta.is_file():
            self._send_json({
                "error": "No encuentro el Excel de gastos (CONTROL DE GASTOS 2026 CLOUD.xlsx)."
            }, 200)
            return
        try:
            self._send_json(parsear_libro(ruta))
        except Exception as error:
            self._send_json({"error": f"Error leyendo el Excel: {error}"}, 200)

    def _api_excel(self):
        global RUTA_EXCEL
        ruta = RUTA_EXCEL or encontrar_excel()
        if ruta is None or not ruta.is_file():
            self._send_json({"path": None, "ok": False})
            return
        self._send_json({
            "ok": True,
            "path": str(ruta),
            "nombre": ruta.name,
            "mtime": datetime.fromtimestamp(ruta.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
        })

    def _api_gasto(self):
        global RUTA_EXCEL
        datos = self._read_json_body()
        ruta = RUTA_EXCEL or encontrar_excel()
        if ruta is None or not ruta.is_file():
            self._send_json({"ok": False, "error": "No encuentro el Excel de gastos."})
            return

        try:
            importe = float(str(datos.get("importe", "")).replace(",", "."))
        except (ValueError, TypeError):
            self._send_json({"ok": False, "error": "Importe no válido."})
            return

        concepto = str(datos.get("concepto", "")).strip()
        categoria = str(datos.get("categoria", "")).strip()
        fecha_txt = str(datos.get("fecha", "")).strip()

        try:
            fecha = datetime.strptime(fecha_txt, "%Y-%m-%d")
        except ValueError:
            self._send_json({"ok": False, "error": "Fecha no válida."})
            return

        try:
            resultado = apuntar_gasto(ruta, importe, concepto, categoria, fecha)
        except Exception as error:
            self._send_json({"ok": False, "error": str(error)})
            return

        self._send_json({"ok": True, "resultado": resultado})

    def log_message(self, fmt, *args):
        pass


def main():
    global RUTA_EXCEL
    puerto = PUERTO_DEFECTO
    if len(sys.argv) > 1:
        try:
            puerto = int(sys.argv[1])
        except ValueError:
            pass

    RUTA_EXCEL = encontrar_excel()
    if RUTA_EXCEL is None:
        print("⚠️  No encontré el Excel de gastos automáticamente.")
        print("    Ponlo en el Escritorio con 'gasto' y '2026' en el nombre,")
        print("    o edita config.json con la ruta completa (excel_path).")
    else:
        print(f"📁 Excel: {RUTA_EXCEL}")

    servidor = ThreadingHTTPServer(("127.0.0.1", puerto), Handler)
    url = f"http://127.0.0.1:{puerto}"
    print(f"🚀 Panel de Gastos v{VERSION} en {url}")
    print("   Pulsa Ctrl+C para detener.")

    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\n👋 Detenido.")
    finally:
        servidor.server_close()


if __name__ == "__main__":
    main()
