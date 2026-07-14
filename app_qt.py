import sys

from PySide6.QtCore import QThread

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QPushButton,
    QLineEdit,
    QVBoxLayout,
    QHBoxLayout,
    QProgressBar,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QMessageBox,
    QFileDialog
)

from worker import Worker


class VentanaPrincipal(QWidget):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Procesador de JSON DTE")
        self.resize(1100, 700)

        self.df = None

        self.crear_interfaz()

    def crear_interfaz(self):

        layout = QVBoxLayout()

        ###################################################
        # Carpeta
        ###################################################

        layout.addWidget(QLabel("Carpeta de trabajo"))

        fila = QHBoxLayout()

        self.txtRuta = QLineEdit()

        self.btnExaminar = QPushButton("Examinar")

        fila.addWidget(self.txtRuta)
        fila.addWidget(self.btnExaminar)

        layout.addLayout(fila)

        ###################################################
        # Procesar
        ###################################################

        self.btnProcesar = QPushButton("Procesar Archivos")

        layout.addWidget(self.btnProcesar)

        ###################################################
        # Barra
        ###################################################

        self.barra = QProgressBar()
        self.barra.setValue(0)

        layout.addWidget(self.barra)

        self.lblEstado = QLabel("Esperando...")

        layout.addWidget(self.lblEstado)

        ###################################################
        # Tabla
        ###################################################

        layout.addWidget(QLabel("Vista previa"))

        self.tabla = QTableWidget()

        layout.addWidget(self.tabla)

        ###################################################
        # Log
        ###################################################

        layout.addWidget(QLabel("Registro del proceso"))

        self.log = QTextEdit()
        self.log.setReadOnly(True)

        layout.addWidget(self.log)

        ###################################################
        # Botones inferiores
        ###################################################

        botones = QHBoxLayout()

        self.btnExcel = QPushButton("Guardar Excel General")
        self.btnLog = QPushButton("Guardar Log")

        botones.addWidget(self.btnExcel)
        self.btnLog.clicked.connect(self.guardar_log)
        botones.addWidget(self.btnLog)

        layout.addLayout(botones)

        self.setLayout(layout)

        ###################################################
        # Eventos
        ###################################################

        self.btnExaminar.clicked.connect(self.buscar_carpeta)
        self.btnProcesar.clicked.connect(self.procesar_carpeta)
        self.btnExcel.clicked.connect(self.guardar_excel)


    def guardar_log(self):

        if not self.log.toPlainText():

            QMessageBox.warning(
                self,
                "Sin información",
                "No existe un log para guardar."
            )

            return

        archivo, _ = QFileDialog.getSaveFileName(

            self,

            "Guardar Log",

            "Log_Procesamiento.txt",

            "Archivos de texto (*.txt)"
        )

        if not archivo:
            return

        try:

            with open(archivo, "w", encoding="utf-8") as f:
                f.write(self.log.toPlainText())

            QMessageBox.information(
                self,
                "Éxito",
                "Log guardado correctamente."
            )

        except Exception as e:

            QMessageBox.critical(
                self,
                "Error",
                str(e)
            )

    ###################################################
    # Buscar carpeta
    ###################################################

    def buscar_carpeta(self):

        carpeta = QFileDialog.getExistingDirectory(
            self,
            "Seleccione una carpeta"
        )

        if carpeta:
            self.txtRuta.setText(carpeta)

    ###################################################
    # Procesar
    ###################################################

    def procesar_carpeta(self):

        ruta = self.txtRuta.text().strip()

        if not ruta:
            self.log.append("Debe seleccionar una carpeta.")
            return

        self.btnProcesar.setEnabled(False)

        self.log.clear()

        self.lblEstado.setText("Procesando...")

        # Barra en modo indeterminado
        self.barra.setRange(0, 0)

        self.thread = QThread()

        self.worker = Worker()

        self.worker.moveToThread(self.thread)

        self.thread.started.connect(
            lambda: self.worker.procesar(ruta)
        )

        self.worker.terminado.connect(self.proceso_terminado)

        self.worker.error.connect(self.proceso_error)

        self.worker.terminado.connect(self.thread.quit)
        self.worker.error.connect(self.thread.quit)
        self.worker.progreso.connect(self.actualizar_progreso)

        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    ###################################################
    # Cuando termina
    ###################################################

    def proceso_terminado(self, df):

        self.df = df

        self.btnProcesar.setEnabled(True)

        self.lblEstado.setText("Proceso terminado")

        self.barra.setRange(0, 100)
        self.barra.setValue(100)

        if self.df.empty:
            self.log.append("No se encontraron registros.")
            return

        self.log.append(
            f"Se encontraron {len(self.df)} registros."
        )

        self.cargar_tabla()

    ###################################################
    # Error
    ###################################################

    def proceso_error(self, mensaje):

        self.btnProcesar.setEnabled(True)

        self.barra.setRange(0, 100)
        self.barra.setValue(0)

        self.lblEstado.setText("Ocurrió un error")

        self.log.append(mensaje)

    def actualizar_progreso(self, actual, total, archivo):

        self.barra.setRange(0, total)

        self.barra.setValue(actual)

        self.lblEstado.setText(
            f"Procesando archivo {actual} de {total}"
        )

        self.log.append(f"Procesando: {archivo}")

    ###################################################
    # Tabla
    ###################################################

        ###################################################
    # Tabla
    ###################################################

    def cargar_tabla(self):

        if self.df is None:
            return

        self.tabla.clear()

        self.tabla.setRowCount(len(self.df))
        self.tabla.setColumnCount(len(self.df.columns))

        self.tabla.setHorizontalHeaderLabels(
            [str(c) for c in self.df.columns]
        )

        for fila in range(len(self.df)):
            for columna in range(len(self.df.columns)):

                valor = str(self.df.iat[fila, columna])

                self.tabla.setItem(
                    fila,
                    columna,
                    QTableWidgetItem(valor)
                )

        self.tabla.resizeColumnsToContents()

    ###################################################
    # Guardar Excel
    ###################################################

    def guardar_excel(self):

        if self.df is None or self.df.empty:

            QMessageBox.warning(
                self,
                "Sin datos",
                "Primero debe procesar una carpeta."
            )
            return

        archivo, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar Excel",
            "Todos_Archivos_Procesados.xlsx",
            "Excel (*.xlsx)"
        )

        if not archivo:
            return

        try:

            self.df.to_excel(
                archivo,
                index=False
            )

            QMessageBox.information(
                self,
                "Éxito",
                "Excel generado correctamente."
            )

        except Exception as e:

            QMessageBox.critical(
                self,
                "Error",
                str(e)
            )


if __name__ == "__main__":

    app = QApplication(sys.argv)

    ventana = VentanaPrincipal()
    ventana.show()

    sys.exit(app.exec())

