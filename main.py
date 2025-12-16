import os
import json
import pandas as pd
from typing import Any, Tuple
from pathlib import Path
import traceback

# =========================
# CONFIGURACIÓN DE RUTAS
# =========================
BASE_DIR = Path(__file__).resolve().parent
JSON_DIR = BASE_DIR / "Json"
OUTPUT_XLSX = BASE_DIR / "resumen_final.xlsx"
LOG_FILE = BASE_DIR / "export_log.txt"


def log(msg: str):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


# =========================
# UTILIDADES
# =========================
def ensure_list(x):
    if x is None:
        return []
    if isinstance(x, list):
        return x
    return [x]


def find_value_by_key_ci(obj: Any, target_key: str) -> Tuple[Any, str]:
    """
    Busca una clave ignorando mayúsculas/minúsculas en cualquier parte del JSON.
    Retorna (valor, ruta) o (None, None)
    """
    target = target_key.lower()

    def rec(o, path):
        if isinstance(o, dict):
            for k, v in o.items():
                if k.lower() == target:
                    return v, " -> ".join(path + [k])
            for k, v in o.items():
                res = rec(v, path + [k])
                if res:
                    return res
        elif isinstance(o, list):
            for i, it in enumerate(o):
                res = rec(it, path + [f"[{i}]"])
                if res:
                    return res
        return None

    result = rec(obj, [])
    return result if result else (None, None)


# =========================
# PROCESO PRINCIPAL
# =========================
def main():
    filas = []

    try:
        log("=== INICIO PROCESO ===")

        if not JSON_DIR.exists():
            log(f"ERROR: Carpeta Json no encontrada en {JSON_DIR}")
            return

        for file in os.listdir(JSON_DIR):
            if not file.lower().endswith(".json"):
                continue

            json_path = JSON_DIR / file
            log(f"Procesando: {file}")

            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception as e:
                log(f"ERROR leyendo {file}: {e}")
                continue

            identificacion = data.get("identificacion", {})
            emisor = data.get("emisor", {})
            resumen = data.get("resumen", {})

            emisor_nombre = emisor.get("nombre")
            codigo_generacion = identificacion.get("codigoGeneracion")
            numero_control = identificacion.get("numeroControl")
            fecha_emision_doc = identificacion.get("fecEmi")
            total_iva_retenido = resumen.get("totalIVAretenido")

            # Buscar selloRecibido en cualquier parte
            sello, sello_origen = find_value_by_key_ci(data, "selloRecibido")

            # Fallback: acuseMH.firma
            if sello is None:
                acuse = data.get("acuseMH", {})
                if isinstance(acuse, dict):
                    sello = acuse.get("firma")
                    sello_origen = "acuseMH -> firma"

            items = data.get("cuerpoDocumento", [])

            if not items:
                filas.append({
                    "Nombre de emisor": emisor_nombre,
                    "Fecha de emisión": fecha_emision_doc,
                    "Código de generación": codigo_generacion,
                    "Sello de recepción": sello,
                    "Número de control": numero_control,
                    "Documento retenido": None,
                    "Total IVA retenido": total_iva_retenido,
                    "Nombre de Json": file
                })
                continue

            for item in items:
                fecha_item = item.get("fechaEmision") or fecha_emision_doc
                documentos = ensure_list(item.get("numDocumento"))

                if not documentos:
                    documentos = [None]

                for doc in documentos:
                    filas.append({
                        "Nombre de emisor": emisor_nombre,
                        "Fecha de emisión": fecha_item,
                        "Código de generación": codigo_generacion,
                        "Sello de recepción": sello,
                        "Número de control": numero_control,
                        "Documento retenido": doc,
                        "Total IVA retenido": total_iva_retenido,
                        "Nombre de Json": file
                    })

        df = pd.DataFrame(filas, columns=[
            "Nombre de emisor",
            "Fecha de emisión",
            "Código de generación",
            "Sello de recepción",
            "Número de control",
            "Documento retenido",
            "Total IVA retenido",
            "Nombre de Json"
        ])

        df["Total IVA retenido"] = pd.to_numeric(
            df["Total IVA retenido"], errors="coerce"
        )

        df.to_excel(OUTPUT_XLSX, index=False)
        log(f"Excel generado: {OUTPUT_XLSX}")
        log(f"Filas totales: {len(df)}")

    except Exception:
        log("ERROR FATAL:")
        log(traceback.format_exc())


if __name__ == "__main__":
    main()
