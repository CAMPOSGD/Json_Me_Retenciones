# -------------------------------------------------------------------------------------------
# Librerias

import os
import json
import pandas as pd
import logging


# -------------------------------------------------------------------------------------------
# Declaraciones

ruta = r"C:\Users\gcampos\OneDrive\Development\Json-Me\Json"


# -------------------------------------------------------------------------------------------
# Configuración del log

logging.basicConfig(
    filename="procesamiento_json.log",
    level=logging.INFO, # Captura INFO y errores
    format="%(asctime)s - %(levelname)s - %(message)s"
)
# -------------------------------------------------------------------------------------------
# Funciones

def fecha_emision(data):
    if data.get("identificacion", {}).get("fecEmi"):
        return data.get("identificacion", {}).get("fecEmi")

    if data.get("json", {}).get("identificacion", {}).get("fecEmi"):
        return data.get("json", {}).get("identificacion", {}).get("fecEmi")

def nombre_de_emisor(data):
    if data.get("emisor", {}).get("nombre"):
        return data.get("emisor", {}).get("nombre")

def codigo_de_generacion(data):
    if data.get("identificacion", {}).get("codigoGeneracion"):
        return data.get("identificacion", {}).get("codigoGeneracion")
    
def numero_de_control(data):
    if data.get("identificacion", {}).get("numeroControl"):
        return data.get("identificacion", {}).get("numeroControl")

def sello_recepcion(data):
    if data.get("selloRecibido"):
        return data.get("selloRecibido")
    
    if data.get("selloRecepcion"):
        return data.get("selloRecepcion")
        
    if data.get("respuestaHacienda",{}).get("selloRecibido"):
        return data.get("respuestaHacienda",{}).get("selloRecibido")   

    if data.get("respuesta",{}).get("selloRecepcion"):
        return data.get("respuesta",{}).get("selloRecepcion")
    
    if data.get("responseMH",{}).get("selloRecibido"):
        return data.get("responseMH",{}).get("selloRecibido")
    
    if data.get("acuseMH",{}).get("numValidacion"):
        return data.get("acuseMH",{}).get("numValidacion")
    
    else :
        return "No se encotró sello de recepción :v"
 
def items_detalle(data_json, fila_base):
    items = data_json.get("cuerpoDocumento", [])
    
    lista_de_filas = [] 
    
    if not items:
        fila_sin_items = fila_base.copy()
        fila_sin_items["Descripción"] = "El json no tiene detalles de items"
        lista_de_filas.append(fila_sin_items)
        
        return lista_de_filas


    for item in items:

        fila_producto = fila_base.copy()

        fila_producto.update({
            "Item #": item.get("numItem"),
            "Tipo DTE": item.get("tipoDte"),
            "Doc Relacionado": item.get("numDocumento"),
            "Monto Sujeto": item.get("montoSujetoGrav"),
            "IVA Retenido": item.get("ivaRetenido"),
            "Descripción": item.get("descripcion")
        })

        lista_de_filas.append(fila_producto)

    return lista_de_filas
        

# -------------------------------------------------------------------------------------------

filas = []

for nombre in os.listdir(ruta):
    ruta_completa = os.path.join(ruta, nombre)

    if os.path.isfile(ruta_completa) and nombre.endswith(".json"):
        try:
            # 1. Intentamos leer el archivo (Manejo de encoding)
            try:
                with open(ruta_completa, "r", encoding="utf-8-sig") as archivo:
                    data = json.load(archivo)
            except UnicodeDecodeError:
                # Si falla utf-8, intentamos con latin-1
                logging.warning(f"Encoding utf-8 el archivo {nombre} entró por latin-1")
                with open(ruta_completa, "r", encoding="latin-1") as archivo:
                    data = json.load(archivo)

            fila_base = {
                "Archivo": nombre,
                "Emisor": nombre_de_emisor(data),
                "Número de control": numero_de_control(data),
                "Sello de recepción": sello_recepcion(data)
            }
            
            filas.extend(items_detalle(data, fila_base))
            
            logging.info(f"Procesado exitosamente: {nombre}")

        except json.JSONDecodeError:
            logging.error(f"El archivo {nombre} no es un JSON válido o está vacío.")
            continue
            
        except KeyError as e:
            logging.error(f"Falta una clave estructura en {nombre}: {e}")
            continue

        except Exception as e:
            logging.exception(f"Error inesperado procesando el archivo {nombre}")
            continue


df = pd.DataFrame(filas)

columnas_ordenadas = [
    "Archivo", "Emisor", "Número de control", "Sello de recepción",
    "Item #", "Tipo DTE", "Doc Relacionado", "Monto Sujeto", "IVA Retenido", "Descripción"
]
df = df.reindex(columns=columnas_ordenadas)

print(df)
df.to_excel("emisor_manual.xlsx", index=False)
