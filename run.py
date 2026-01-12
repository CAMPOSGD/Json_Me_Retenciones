import sys
import os
import logging
import traceback

def main():
    # 1. Configurar log de emergencia inmediatamente (junto al ejecutable)
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        
    log_path = os.path.join(base_dir, "debug_inicio.log")
    
    # Forzamos la configuración del log para asegurar que escriba
    logging.basicConfig(
        filename=log_path,
        level=logging.DEBUG,
        format="%(asctime)s - %(levelname)s - %(message)s",
        force=True
    )

    # Desactivar estadísticas para evitar el prompt de bienvenida que requiere stdin (teclado)
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

    # --- FIX: Simular stdin para evitar crash si Streamlit pide email ---
    class MockStdin:
        def read(self, size=-1):
            return ""
        def readline(self, size=-1):
            return "\n"
        def close(self):
            pass
            
    if sys.stdin is None:
        sys.stdin = MockStdin()

    # Clase para redirigir stdout (print) y stderr (errores) al log
    class StreamToLogger(object):
        def __init__(self, logger, log_level=logging.INFO):
            self.logger = logger
            self.log_level = log_level
            self.linebuf = ''
            self.encoding = 'utf-8'

        def write(self, buf):
            for line in buf.rstrip().splitlines():
                self.logger.log(self.log_level, line.rstrip())

        def flush(self):
            pass

    # Redirigir stdout y stderr
    sys.stdout = StreamToLogger(logging.getLogger('STDOUT'), logging.INFO)
    sys.stderr = StreamToLogger(logging.getLogger('STDERR'), logging.ERROR)

    try:
        logging.info("--- INICIANDO APLICACIÓN ---")
        
        # 2. Importaciones dentro del try para capturar errores de librerías faltantes
        logging.info("Importando librerías...")
        from streamlit.web import cli as stcli
        import Lector_Jsons # Esto asegura que PyInstaller incluya pandas/openpyxl
        import tkinter
        from tkinter import filedialog
        
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(__file__)

        app_path = os.path.join(base_path, "app.py")
        logging.info(f"Ruta del script principal: {app_path}")

        logging.info("Lanzando Streamlit...")
        # Quitamos headless para que abra el navegador y forzamos localhost para evitar firewall
        sys.argv = ["streamlit", "run", app_path, "--global.developmentMode=false", "--server.address=127.0.0.1"]
        
        sys.exit(stcli.main())
        
    except Exception as e:
        # 3. Captura robusta de errores
        err_msg = traceback.format_exc()
        logging.critical("ERROR FATAL NO CONTROLADO:")
        logging.critical(err_msg)
        
        # Intentamos mostrar una alerta visual en Windows
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, f"Error fatal al inicio.\nRevise el archivo: {log_path}", "Error de Ejecución", 0x10)
        except:
            pass

if __name__ == "__main__":
    main()