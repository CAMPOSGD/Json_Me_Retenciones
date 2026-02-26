# -------------------------------------------------------------------------------------------
# Librerias
import os
import json
import pandas as pd
import logging


# -------------------------------------------------------------------------------------------
# Declaraciones


# -------------------------------------------------------------------------------------------
# Funciones

def sello_recepcion(data):
    # Las tuplas representan ["llave_padre", "llave_hija"]
    rutas = [
        ("selloRecibido",),                      
        ("SelloRecibido",),                      # raíz (Mayúscula)
        ("selloMH",),
        ("selloRecepcion", "selloRecibido"),     
        ("respuestaHacienda", "selloRecibido"),
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
            
            # estos son los datos para retencion
            "Doc Relacionado": item.get("numDocumento"),
            "Monto Sujeto": item.get("montoSujetoGrav"),
            "IVA Retenido": item.get("ivaRetenido"),
            "Descripción": item.get("descripcion"),
            
            # estos son los datos para FCF y CCF
            
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
                data = json.loads(raw_data.decode(encoding))
                break 
            except (UnicodeDecodeError, json.JSONDecodeError):
                continue
        
        if data is None:
            msg = f"FALLO DECODIFICACION: {nombre}. Primeros bytes: {raw_data[:50]}"
            logging.error(msg)
            return []

        if isinstance(data, list):
            if len(data) > 0:
                logging.warning(f"El archivo {nombre} es una lista. Usando el primer elemento.")
                data = data[0]
            else:
                msg = f"   ERROR: El archivo {nombre} es una lista vacía []."
                logging.error(msg)
                return []
        
        if not isinstance(data, dict):
            msg = f"   ERROR: El archivo {nombre} no es un objeto JSON, es {type(data)}."
            logging.error(msg)
            return []

        fila_base = identificacion_y_emisor(data, nombre)
        fila_base.update(informacion_receptor(data))
        fila_base.update(resumen(data))
        
        filas.extend(items_detalle(data, fila_base))
        
        logging.info(f"Procesado exitosamente: {nombre}")
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
    # Reindex solo si el dataframe no está vacío para evitar errores
    if not df.empty:
        df = df.reindex(columns=columnas_ordenadas)
    return df

def procesar_directorio(ruta_directorio):
    filas = []
    
    if not os.path.exists(ruta_directorio):
        logging.error(f"La ruta no existe: {ruta_directorio}")
        return pd.DataFrame()

    for nombre in os.listdir(ruta_directorio):
        ruta_completa = os.path.join(ruta_directorio, nombre)

        if os.path.isfile(ruta_completa) and nombre.lower().endswith(".json"):
            try:
                with open(ruta_completa, "rb") as f:
                    raw_data = f.read()
                
                # Usamos la nueva función refactorizada
                filas.extend(procesar_bytes_json(raw_data, nombre))

            except Exception as e:
                logging.exception(f"Error leyendo archivo {nombre}: {e}")
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
    # Configuración de log solo si se ejecuta directamente (para pruebas)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    # Ruta de prueba local (solo se usa si ejecutas este script directamente)
    ruta = "." 
    # Bloque de ejecución original
    df = procesar_directorio(ruta)
    print(df)
    
    if not df.empty:
        df.to_excel("Todos_Archivos_procesados.xlsx", index=False)
        
        nombre_excel_separado = "Resumen_Por_Tipo_Documento.xlsx"
        if generar_excel_separado(df, nombre_excel_separado):
            print(f"Se generó exitosamente el archivo separado: {nombre_excel_separado}")
