import os
import json
import pandas as pd
from typing import Any, Tuple

CARPETA = "Json"
filas = []

def ensure_list(x):
    if x is None:
        return []
    if isinstance(x, list):
        return x
    return [x]

def find_value_by_key_ci(obj: Any, target_key: str) -> Tuple[Any, str]:
    """
    Busca recursivamente la primera ocurrencia de una clave *ignorando mayúsc/minúsculas*.
    Devuelve (valor, ruta_str) o (None, None) si no la encuentra.
    Ruta ejemplo: "raiz -> acuseMH -> firma"
    """
    target_lower = target_key.lower()
    path = []

    def rec(o, p):
        if isinstance(o, dict):
            # comprobar claves del dict en este nivel (case-insensitive)
            for k, v in o.items():
                newp = p + [k]
                if k.lower() == target_lower:
                    return v, " -> ".join(newp)
            # si no se encontró en este nivel, buscar en valores
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
    if found is None:
        return None, None
    return found

for nombre_archivo in os.listdir(CARPETA):
    if not nombre_archivo.lower().endswith(".json"):
        continue
    ruta = os.path.join(CARPETA, nombre_archivo)

    try:
        with open(ruta, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"ERROR leyendo {nombre_archivo}: {e}")
        continue

    identificacion = data.get("identificacion", {}) or {}
    emisor = data.get("emisor", {}) or {}
    resumen = data.get("resumen", {}) or {}

    emisor_nombre = emisor.get("nombre")
    codigoGeneracion = identificacion.get("codigoGeneracion")
    numeroControl = identificacion.get("numeroControl")
    totalIVAretenido = resumen.get("totalIVAretenido")

    # 1) Buscar SelloRecibido (case-insensitive) en cualquier parte
    sello_val, sello_origen = find_value_by_key_ci(data, "selloRecibido")

    # 2) Si no se encontró, probar fallbacks comunes
    if sello_val is None:
        # acuseMH.firma
        acuse = data.get("acuseMH")
        if isinstance(acuse, dict) and "firma" in acuse:
            sello_val = acuse.get("firma")
            sello_origen = "acuseMH -> firma"
    if sello_val is None:
        # raiz.firmaElectronica (o alguna variante)
        val, origin = find_value_by_key_ci(data, "firmaElectronica")
        if val is not None:
            sello_val = val
            sello_origen = origin
    # último recurso: buscar cualquier clave que contenga 'firma' o 'sello' (case-insensitive)
    if sello_val is None:
        # recorrido simple para encontrar la primera clave que contenga 'firma' o 'sello'
        def find_contains(o, substr_lower="firma"):
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
            # r puede ser (value, route) o value; intentar normalizar
            if isinstance(r, tuple):
                sello_val, sello_origen = r
            else:
                sello_val = r
                sello_origen = "clave_contiene_sello_o_firma"
    # Si sigue None, queda None y origen None

    # Procesar items
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

# DataFrame en el orden pedido (añadí "Origen sello" después del sello)
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

out_xlsx = "ResultadoDeJsons.xlsx"
df.to_excel(out_xlsx, index=False)

print(f"Generado: {out_xlsx} (filas: {len(df)})")
