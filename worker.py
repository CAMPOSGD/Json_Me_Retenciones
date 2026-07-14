from PySide6.QtCore import QObject, Signal

from Lector_Jsons import procesar_directorio


class Worker(QObject):

    terminado = Signal(object)

    error = Signal(str)

    progreso = Signal(int, int, str)

    def procesar(self, ruta):

        try:

            df = procesar_directorio(
                ruta,
                self.enviar_progreso
            )

            self.terminado.emit(df)

        except Exception as e:

            self.error.emit(str(e))

    def enviar_progreso(self, actual, total, archivo):

        self.progreso.emit(
            actual,
            total,
            archivo
        )