import os
import json
import pandas as pd
import logging
import re

def sello_recepcion(data):
    rutas = [
        ("selloRecibido",),                      
        ("SelloRecibido",),                      
        ("SelloRecepcion",),                     
        ("selloMH",),
        ("selloRecepcion", "selloRecibido"),     
        ("respuestaHacienda", "selloRecibido"),
        ("respuesta_dgi", "selloRecibido"),      
        ("respuesta", "selloRecepcion"),
        ("responseMH", "selloRecibido"),
        ("acuseMH", "numValidacion"),
        ("respuestamh", "selloRecibido"),
        ("selloMH", "selloRecibido"),
        ("response", "selloRecibido")
    ]

    for ruta in rutas:
        valor = data
        for llave in ruta:
            if isinstance(valor, dict):
                valor = valor.get(llave)
            else:
                valor = None
                break
        
        if valor and isinstance(valor, str):
            return valor

    fallback = data.get("selloRecepcion")
    if isinstance(fallback, str):
        return fallback

    return "No se encontró sello de recepción :v"
 
def informacion_receptor(data):  
    receptor = data.get("receptor", {})
    
    return {
        "Correo Receptor": receptor.get("correo"),
    }
     
def identificacion_y_emisor(data, nombre_archivo):
    identificacion = data.get("identificacion") or {}
    
    if not identificacion.get("fecEmi") and (data.get("json") or {}).get("identificacion"):
        identificacion = (data.get("json") or {}).get("identificacion") or {}

    emisor = data.get("emisor") or {}

    tipo_dte = identificacion.get("tipoDte")
    nombres_dte = {
        "01": "Factura",
        "03": "Comprobante de Crédito Fiscal",
        "04": "Nota de remisión",
        "05": "Nota de Crédito",
        "06": "Nota de Débito",
        "07": "Comprobante de Retención",
        "08": "Comprobante de liquidación",
        "09": "Documento contable de liquidación",
        "11": "Factura de exportación",
        "14": "Factura de sujeto excluido",
        "15": "Comprobante de donación"
    }

    return {
        "Fecha": identificacion.get("fecEmi"),
        "Emisor": emisor.get("nombre"),
        "Nit de emisor": emisor.get("nit"),
        "NRC de emisor": emisor.get("nrc"),
        "Número de control": identificacion.get("numeroControl"),
        "Código de generación": identificacion.get("codigoGeneracion"),
        "Sello de recepción": sello_recepcion(data),
        "Nombre del archivo": nombre_archivo,
        "Tipo DTE": nombres_dte.get(tipo_dte, tipo_dte)
    }       

def items_detalle(data_json, fila_base):
    items = data_json.get("cuerpoDocumento") or []
    
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
            
            "Doc Relacionado": item.get("numDocumento"),
            "Monto Sujeto": item.get("montoSujetoGrav"),
            "IVA Retenido": item.get("ivaRetenido"),
            "Descripción": item.get("descripcion"),
                        
            "Cantidad CCF": item.get("cantidad"),
            "Precio Unitario CCF": item.get("precioUni"),
            "Venta gravada CCF": item.get("ventaGravada"),
        })

        lista_de_filas.append(fila_producto)

    return lista_de_filas
     
def resumen(data):
    resumen_data = data.get("resumen") or {}
    
    tributos = resumen_data.get("tributos") or []
    total_iva = sum(t.get("valor", 0) for t in tributos)

    return {
        "Sub Total CCF": resumen_data.get("subTotalVentas"),
        "Total IVA": total_iva,
        "Total CCF": resumen_data.get("totalPagar")
    }

def procesar_bytes_json(raw_data, nombre):
    """Procesa el contenido binario de un archivo JSON y devuelve una lista de filas."""
    filas = []
    try:
        if not raw_data:
            logging.error(f"El archivo {nombre} está realmente vacío (0 bytes).")
            return []

        data = None
        for encoding in ["utf-8-sig", "utf-16", "latin-1", "cp1252"]:
            try:
                decoded_str = raw_data.decode(encoding).strip()
                if not decoded_str:
                    continue

                decoder = json.JSONDecoder()
                objs = []
                text = decoded_str

                while text:
                    text = text.lstrip()
                    if not text:
                        break
                    try:
                        obj, index = decoder.raw_decode(text)
                        objs.append(obj)
                        text = text[index:].lstrip()
                        if text.startswith(','):
                            text = text[1:].lstrip()
                    except json.JSONDecodeError:
                        break
                
                if objs: 
                    data = objs[0] if len(objs) == 1 else objs
                    break
                    
            except UnicodeDecodeError:
                continue
        
        if data is None:
            msg = f"FALLO DECODIFICACION: {nombre}. Primeros bytes: {raw_data[:50]}"
            logging.error(msg)
            return []

        if isinstance(data, dict):
            data = [data]
            
        if not isinstance(data, list):
            msg = f"   ERROR: El archivo {nombre} no tiene un formato válido, es {type(data)}."
            logging.error(msg)
            return []
        
        respuesta_hacienda = {}

        for obj in data:

            if (
                isinstance(obj, dict)
                and obj.get("selloRecibido")
            ):

                respuesta_hacienda = obj
                break

        for elemento in data:
            if not isinstance(elemento, dict):
                continue

            if "identificacion" not in elemento and "json" in elemento and isinstance(elemento["json"], dict):
                elemento.update(elemento["json"])

            if "identificacion" not in elemento and "" in elemento and isinstance(elemento[""], dict):
                elemento.update(elemento[""])

            if "identificacion" not in elemento and "emisor" not in elemento:
                continue

            if respuesta_hacienda:
                elemento = elemento.copy()
                elemento.update(respuesta_hacienda)

            fila_base = identificacion_y_emisor(elemento, nombre)
            fila_base.update(informacion_receptor(elemento))
            fila_base.update(resumen(elemento))
            
            filas.extend(items_detalle(elemento, fila_base))
            
        if filas:
            logging.info(f"Procesado exitosamente: {nombre}")
        else:
            logging.warning(f"No se encontraron datos válidos de DTE en: {nombre}")
            
        return filas

    except json.JSONDecodeError as e:
        logging.error(f"JSON INVALIDO en {nombre}: {e}")
        return []
        
    except KeyError as e:
        logging.error(f"FALTA CAMPO en {nombre}: {e}")
        return []

    except Exception as e:
        logging.exception(f"ERROR INESPERADO en {nombre}: {e}")
        return []

def crear_dataframe_desde_filas(filas):
    df = pd.DataFrame(filas)

    columnas_ordenadas = [


        "Tipo DTE", 
        "Fecha", "Emisor", "Nit de emisor", "NRC de emisor", 
        "Número de control", "Código de generación", "Sello de recepción",
        "Item #", 
        "Doc Relacionado", "Monto Sujeto", "IVA Retenido", 
        "Cantidad CCF", "Precio Unitario CCF", "Venta gravada CCF", 
        "Sub Total CCF", "Total IVA", "Total CCF", 
        "Descripción","Nombre del archivo", "Correo Receptor"
    ]
    if not df.empty:
        df = df.reindex(columns=columnas_ordenadas)
    return df

def procesar_directorio(ruta_directorio, progreso=None):
    
    filas = []


    if not os.path.exists(ruta_directorio):
        logging.error(f"La ruta no existe: {ruta_directorio}")
        return pd.DataFrame()
    
    archivos = [
    nombre
    for nombre in os.listdir(ruta_directorio)
    if os.path.isfile(os.path.join(ruta_directorio, nombre))
    and nombre.lower().endswith(".json")
]

    total_archivos = len(archivos)

    for indice, nombre in enumerate(archivos, start=1):
        ruta_completa = os.path.join(ruta_directorio, nombre)

        if os.path.isfile(ruta_completa) and nombre.lower().endswith(".json"):
            logging.info(f"-> Archivo detectado: {nombre}")

            if progreso:
                progreso(indice, total_archivos, nombre)

            try:
                with open(ruta_completa, "rb") as f:
                    raw_data = f.read()

                filas.extend(procesar_bytes_json(raw_data, nombre))

            except Exception as e:
                logging.exception(f"Error leyendo archivo {ruta_completa}: {e}")
                continue

    return crear_dataframe_desde_filas(filas)

def generar_excel_separado(df, output_target):

    try:
        with pd.ExcelWriter(output_target) as writer:
            tipos_unicos = df["Tipo DTE"].fillna("Sin Tipo").unique()
            
            for tipo in tipos_unicos:
                df_filtrado = df[df["Tipo DTE"].fillna("Sin Tipo") == tipo]
                
                if tipo != "Comprobante de Retención":
                    cols_retencion = ["Doc Relacionado", "Monto Sujeto", "IVA Retenido"]
                    df_filtrado = df_filtrado.drop(columns=[c for c in cols_retencion if c in df_filtrado.columns])
                else:
                    cols_ccf = ["Cantidad CCF", "Precio Unitario CCF", "Venta gravada CCF", "Sub Total CCF", "Total IVA", "Total CCF"]
                    df_filtrado = df_filtrado.drop(columns=[c for c in cols_ccf if c in df_filtrado.columns])

                nombre_hoja = str(tipo)[:31].replace(":", "").replace("/", "-").replace("\\", "").replace("?", "").replace("*", "").replace("[", "").replace("]", "")
                
                df_filtrado.to_excel(writer, sheet_name=nombre_hoja, index=False)
        return True
    except Exception as e:
        logging.error(f"No se pudo crear el archivo separado por hojas: {e}")
        return False

if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    ruta = "."
    df = procesar_directorio(ruta)
    print(df)
    
    if not df.empty:
        df.to_excel("Todos_Archivos_procesados.xlsx", index=False)
        
        nombre_excel_separado = "Resumen_Por_Tipo_Documento.xlsx"
        if generar_excel_separado(df, nombre_excel_separado):
            print(f"Se generó exitosamente el archivo separado: {nombre_excel_separado}")
