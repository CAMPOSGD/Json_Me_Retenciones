# main.py
import os
import json
import pandas as pd
from typing import Any, Tuple
from pathlib import Path
import traceback
import sys

# Trabajar desde la carpeta donde está el ejecutable/script
BASE = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))  # si PyInstaller empaca usa _MEIPASS
HERE = Path.cwd()  # si preferís que use la carpeta donde está el exe, usar Path(__file__).parent

# Preferible: usar la carpeta del ejecutable para buscar la carpeta Json al lado del exe
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
CARPETA = APP_DIR / "Json"

# Si querés forzar usar la carpeta actual donde se hizo doble clic:
# CARPETA = Path.cwd() / "Json"

OUT_XLSX = APP_DIR / "Resultado de Jsons..xlsx"
LOGFILE = APP_DIR / "logs.txt"

def log(msg):
    with open(LOGFILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")
    print(msg)

def ensure_list(x):
    if x is None:
        return []
    if isinstance(x, list):
        return x
    return [x]

def find_value_by_key_ci(obj: Any, target_key: str) -> Tuple[Any, str]:
    target_lower = target_key.lower()
    def rec(o, p):
        if isinstance(o, dict):
            for k, v in o.items():
                if k.lower() == target_lower:
                    return v, " -> ".join(p + [k])
            for k, v in o.items():
                res = rec(v, p + [k])
                if res is not None:
                    return res
        elif isinstance(o, list):
            for i, item in enumerate(o):
                res = rec(item, p + [f"[{i}]"])
                if res is not None:
                    return res
        return None
    found = rec(obj, [])
    return found or (None, None)

def main():
    filas = []
    try:
        if not CARPETA.exists() or not CARPETA.is_dir():
            log(f"ERROR: carpeta Json no encontrada en: {CARPETA}")
            return

        for nombre_archivo in os.listdir(CARPETA):
            if not nombre_archivo.lower().endswith(".json"):
                continue
            ruta = CARPETA / nombre_archivo
            try:
                with open(ruta, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception as e:
                log(f"ERROR leyendo {nombre_archivo}: {e}")
                continue

            identificacion = data.get("identificacion", {}) or {}
            emisor = data.get("emisor", {}) or {}
            resumen = data.get("resumen", {}) or {}

            emisor_nombre = emisor.get("nombre")
            codigoGeneracion = identificacion.get("codigoGeneracion")
            numeroControl = identificacion.get("numeroControl")
            totalIVAretenido = resumen.get("totalIVAretenido")

            sello_val, sello_origen = find_value_by_key_ci(data, "selloRecibido")

            if sello_val is None:
                acuse = data.get("acuseMH")
                if isinstance(acuse, dict) and "firma" in acuse:
                    sello_val = acuse.get("firma")
                    sello_origen = "acuseMH -> firma"

            if sello_val is None:
                val, origin = find_value_by_key_ci(data, "firmaElectronica")
                if val is not None:
                    sello_val = val
                    sello_origen = origin

            if sello_val is None:
                # buscar claves que contengan 'sello' o 'firma'
                def find_contains(o, substr_lower):
                    if isinstance(o, dict):
                        for k, v in o.items():
                            if substr_lower in k.lower():
                                return v, " -> ".join([k])
                        for k, v in o.items():
                            res = find_contains(v, substr_lower)
                            if res is not None:
                                return res
                    elif isinstance(o, list):
                        for i, it in enumerate(o):
                            res = find_contains(it, substr_lower)
                            if res is not None:
                                return res
                    return None
                r = find_contains(data, "sello")
                if r is None:
                    r = find_contains(data, "firma")
                if r is not None:
                    if isinstance(r, tuple):
                        sello_val, sello_origen = r
                    else:
                        sello_val = r
                        sello_origen = "clave_contiene_sello_o_firma"

            items = data.get("cuerpoDocumento", []) or []
            if not items:
                filas.append({
                    "Nombre de emisor": emisor_nombre,
                    "Fecha de emisión": identificacion.get("fecEmi"),
                    "Código de generación": codigoGeneracion,
                    "Sello de recepción": sello_val,
                    "Origen sello": sello_origen,
                    "Número de control": numeroControl,
                    "Documento retenido": None,
                    "Total IVA retenido": totalIVAretenido,
                    "Nombre de Json": nombre_archivo
                })
                continue

            for item in items:
                fecha_emision_item = item.get("fechaEmision") or identificacion.get("fecEmi")
                raw_numdoc = item.get("numDocumento")
                numdocs = ensure_list(raw_numdoc)
                if not numdocs:
                    filas.append({
                        "Nombre de emisor": emisor_nombre,
                        "Fecha de emisión": fecha_emision_item,
                        "Código de generación": codigoGeneracion,
                        "Sello de recepción": sello_val,
                        "Origen sello": sello_origen,
                        "Número de control": numeroControl,
                        "Documento retenido": None,
                        "Total IVA retenido": totalIVAretenido,
                        "Nombre de Json": nombre_archivo
                    })
                else:
                    for nd in numdocs:
                        filas.append({
                            "Nombre de emisor": emisor_nombre,
                            "Fecha de emisión": fecha_emision_item,
                            "Código de generación": codigoGeneracion,
                            "Sello de recepción": sello_val,
                            "Origen sello": sello_origen,
                            "Número de control": numeroControl,
                            "Documento retenido": nd,
                            "Total IVA retenido": totalIVAretenido,
                            "Nombre de Json": nombre_archivo
                        })

        cols = [
            "Nombre de emisor",
            "Fecha de emisión",
            "Código de generación",
            "Sello de recepción",
            "Origen sello",
            "Número de control",
            "Documento retenido",
            "Total IVA retenido",
            "Nombre de Json"
        ]
        df = pd.DataFrame(filas, columns=cols)
        if "Total IVA retenido" in df.columns:
            df["Total IVA retenido"] = pd.to_numeric(df["Total IVA retenido"], errors="coerce")

        df.to_excel(OUT_XLSX, index=False)
        log(f"Generado: {OUT_XLSX} (filas: {len(df)})")
    except Exception as e:
        log("EXCEPCION en main:")
        log(traceback.format_exc())

if __name__ == "__main__":
    main()
    # opcional: pausa para ver el log si se ejecuta en consola
