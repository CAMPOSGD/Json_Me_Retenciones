import streamlit as st
import pandas as pd
import os
import logging
import io
import tkinter as tk
from tkinter import filedialog
from io import BytesIO
from Lector_Jsons import procesar_directorio, generar_excel_separado

# Configuración de la página
st.set_page_config(page_title="Procesador JSON a Excel", layout="wide")

st.title("Procesador de Archivos JSON")
st.markdown("""
Seleccione la carpeta local que contiene los archivos JSON para generar los reportes.
""")

# Inicializar el estado de la sesión para guardar los datos entre interacciones
if 'df_resultado' not in st.session_state:
    st.session_state.df_resultado = None

if 'log_proceso' not in st.session_state:
    st.session_state.log_proceso = ""

if 'ruta_carpeta' not in st.session_state:
    st.session_state.ruta_carpeta = os.getcwd()

# Layout para la selección de carpeta
col1, col2 = st.columns([4, 1])

with col2:
    # Espaciado para alinear el botón con el input
    st.write("") 
    st.write("")
    if st.button("📂 Buscar Carpeta"):
        # Crear ventana oculta de Tkinter
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes('-topmost', 1) # Forzar que aparezca encima
        
        # Abrir selector de carpetas
        carpeta_seleccionada = filedialog.askdirectory(master=root)
        root.destroy()
        
        if carpeta_seleccionada:
            st.session_state.ruta_carpeta = carpeta_seleccionada
            st.rerun()

with col1:
    ruta_input = st.text_input("Ruta seleccionada:", value=st.session_state.ruta_carpeta)
    # Actualizar estado si el usuario edita manualmente
    st.session_state.ruta_carpeta = ruta_input

if st.button("Procesar Carpeta", type="primary"):
    if not os.path.isdir(st.session_state.ruta_carpeta):
        st.error(f"❌ La ruta ingresada no existe o no es una carpeta válida: {st.session_state.ruta_carpeta}")
    else:
        with st.spinner(f"Leyendo y procesando archivos en: {st.session_state.ruta_carpeta}..."):
            # 1. Configurar captura de logs en memoria
            log_capture_string = io.StringIO()
            ch = logging.StreamHandler(log_capture_string)
            ch.setLevel(logging.INFO)
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            ch.setFormatter(formatter)
            
            root_logger = logging.getLogger()
            root_logger.setLevel(logging.INFO) # Forzamos nivel INFO para capturar mensajes de éxito
            root_logger.addHandler(ch)
            
            # Procesamos y guardamos el resultado en session_state
            df = procesar_directorio(st.session_state.ruta_carpeta)
            st.session_state.df_resultado = df
            
            # 2. Guardar el log capturado y limpiar
            st.session_state.log_proceso = log_capture_string.getvalue()
            root_logger.removeHandler(ch)
        
        if df.empty:
            st.warning("⚠️ No se encontraron datos o archivos JSON válidos en la carpeta seleccionada.")
        else:
            st.success(f"Proceso completado. Se encontraron {len(df)} registros.")

# Si hay datos en memoria, mostramos la tabla y los botones de descarga
if st.session_state.df_resultado is not None and not st.session_state.df_resultado.empty:
    df = st.session_state.df_resultado
    
    # Mostrar una vista previa de los datos
    st.subheader("Vista Previa (Primeros 50 registros)")
    st.dataframe(df.head(50))
    
    # Columnas para los botones de descarga
    col1, col2, col3 = st.columns(3)
    
    # 1. Generar Excel Completo en memoria
    buffer_completo = BytesIO()
    with pd.ExcelWriter(buffer_completo, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    
    col1.download_button(
        label="--> Descargar Excel General",
        data=buffer_completo.getvalue(),
        file_name="Todos_Archivos_procesados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # 2. Generar Excel Separado por Hojas en memoria
    buffer_separado = BytesIO()
    exito = generar_excel_separado(df, buffer_separado)
    
    if exito:
        col2.download_button(
            label="--> Descargar Resumen por Tipo",
            data=buffer_separado.getvalue(),
            file_name="Resumen_Por_Tipo_Documento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        col2.error("Hubo un error generando el reporte separado.")

    # 3. Descargar Log de Procesamiento
    if st.session_state.log_proceso:
        col3.download_button(
            label="📜 Descargar Log del Proceso",
            data=st.session_state.log_proceso,
            file_name="Log_Procesamiento.txt",
            mime="text/plain"
        )
