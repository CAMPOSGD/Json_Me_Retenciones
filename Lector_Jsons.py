# -------------------------------------------------------------------------------------------
# Librerias
import os
import json
import pandas as pd
import logging


# -------------------------------------------------------------------------------------------
# Declaraciones

#ruta = r"C:\Users\gcampos\OneDrive\Development\Json-Me\Json"
ruta = r"C:\Users\gabri\Downloads\Json_Me_Retenciones\Json"


# -------------------------------------------------------------------------------------------
# Configuración del log

logging.basicConfig(
    filename="log de archivos procesados.log",
    level=logging.INFO, # Captura INFO y errores
    format="%(asctime)s - %(levelname)s - %(message)s"
)
# -------------------------------------------------------------------------------------------
# Funciones

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
 
def informacion_receptor(data):  
    # if data.get("receptor", {}).get("correo"):
    #     return data.get("receptor", {}).get("correo")
    
    receptor = data.get("receptor", {})
    
    return {
        "Correo Receptor": receptor.get("correo"),
        #"Nit": receptor.get("nit"),
        #"Nombre": receptor.get("nombre")
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
            #"Tipo DTE": item.get("tipoDte"
            
            # estos son los datos para retencion
            "Doc Relacionado": item.get("numDocumento"),
            "Monto Sujeto": item.get("montoSujetoGrav"),
            "IVA Retenido": item.get("ivaRetenido"),
            "Descripción": item.get("descripcion"),
            
            # estos son los datos para FCF y CCF
            
            # detalle? , montos, iva , iva percibido, correos electrónicos
            
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

# 5596.166836

filas = []

for nombre in os.listdir(ruta):
    ruta_completa = os.path.join(ruta, nombre)

    if os.path.isfile(ruta_completa) and nombre.endswith(".json"):
        try:
            try:
                with open(ruta_completa, "r", encoding="utf-8-sig") as archivo:
                    data = json.load(archivo)
            except UnicodeDecodeError:
                logging.warning(f"Encoding utf-8 el archivo {nombre} entró por latin-1")
                with open(ruta_completa, "r", encoding="latin-1") as archivo:
                    data = json.load(archivo)

            fila_base = identificacion_y_emisor(data, nombre)
            fila_base.update(informacion_receptor(data))
            fila_base.update(resumen(data))
            
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
    "Tipo DTE", 
    "Fecha", "Emisor", "Nit de emisor", "NRC de emisor", 
    "Número de control", "Código de generación", "Sello de recepción",
    "Item #", 
    "Doc Relacionado", "Monto Sujeto", "IVA Retenido", 
    "Cantidad CCF", "Precio Unitario CCF", "Venta gravada CCF", 
    "Sub Total CCF", "Total IVA", "Total CCF", 
    "Descripción","Nombre del archivo", "Correo Receptor"
]
df = df.reindex(columns=columnas_ordenadas)

print(df)
df.to_excel("Todos_Archivos_procesados.xlsx", index=False)

# -------------------------------------------------------------------------------------------

nombre_excel_separado = "Resumen_Por_Tipo_Documento.xlsx"

try:
    with pd.ExcelWriter(nombre_excel_separado) as writer:
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
            
    print(f"Se generó exitosamente el archivo separado: {nombre_excel_separado}")

except Exception as e:
    logging.error(f"No se pudo crear el archivo separado por hojas: {e}")
    print(f"Error creando el archivo separado: {e}")
