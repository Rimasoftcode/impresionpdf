import sys
import os
import csv
import re
from itertools import cycle
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QGridLayout, QWidget, QLabel, 
                             QComboBox, QLineEdit, QTextEdit, QPushButton, QMessageBox, QDateEdit, QSpacerItem, QSizePolicy, QMessageBox)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QCursor
from PyQt5.Qt import QClipboard
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, QFileDialog, QHBoxLayout, QDesktopWidget, 
                             QTextEdit, QMessageBox, QComboBox, QLabel, QLineEdit, QPushButton)
from PyQt5.QtCore import Qt
import fitz  # PyMuPDF
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QLineEdit, QComboBox, QGraphicsView, QGraphicsScene, QGraphicsTextItem, QLabel, QFormLayout, QPushButton, QMessageBox
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtCore import Qt
from fpdf import FPDF
from datetime import datetime


current_dir = os.path.dirname(os.path.abspath(__file__))
CSV_FILE_PATH = os.path.join(current_dir, 'Respuestas_tipos.csv')
CSV_REQUERIMIENTOS = os.path.join(current_dir, 'Datos_requerimientos.csv')
#CSV_SUB_CLASIFICACIONES = os.path.join(current_dir, 'Clasifica_respuestas_tipos.csv')
ICON_PATH = os.path.join(current_dir, 'EstrellaNsur.ico')

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Formulario Dinámico")
        self.setWindowIcon(QtGui.QIcon(ICON_PATH))
        self.setGeometry(100, 100, 600, 300)
        
        self.classes = {}
        self.fields = {}
        self.field_labels = {}
        self.load_classes_from_csv()
        
        self.pdf_templates = {
            "Entrega Descuidada/Equivocada/Fuera de Plazo": (
                "Talca, {Fecha Respuesta}\n\n"
                "Señor(a)\n{Nombre Solicitante}\nPresente\n\n"
                "Referencia: Contacto Nº {N° Contacto}\n"
                "Materia: {N° Reclamo}\n"
                "Nº Cliente: {ID.}\n\n"
                "De nuestra consideración:\n\n"
                "Junto con saludarle muy cordialmente y en respuesta a su solicitud de {N° Contacto}, registrado en nuestro sistema con el número de atención {ID.}, podemos informar lo siguiente:\n\n"
                "Con fecha {Fecha Respuesta} hemos procedido a efectuar {Nombre Solicitante} en nuestro sistema.\n\n"
                "Adicionalmente, le señalamos que en el caso de presentar algún inconveniente de cualquier tipo con nuestro servicio, puede comunicarse con nosotros a través del Centro de Ayuda 6003311000 - 6003724000, desde celulares al *3311, a través de la Oficina Virtual www.nuevosur.cl donde le atenderá un ejecutivo.\n\n"
                "Finalmente y de acuerdo a la exigencia que la autoridad nos impone, es nuestro deber informar que si considera insatisfactoria esta respuesta, también tiene el derecho de concurrir a la Superintendencia de Servicios Sanitarios, ubicada en calle 2 Norte Nº1103, Talca, Fono 800 381 800.\n\n"
                "Sin otro particular se despide atentamente de usted.\n\n"
                "Servicio al Cliente"
            ),
            # Agrega más cartas tipo aquí
        }


        self.text_structures = {
            "Entrega Descuidada/Equivocada/Fuera de Plazo": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "No entrega de Boletas": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Lectura incorrecta": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Medidor Defectuoso/Invertido": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Cambio Medidor Solicitado por Cliente": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, "
                "Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, Inspección o instalación: {Inspección o instalación}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Cambio de medidor ejecutado por cliente": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Bloqueo/desbloqueo de convenios/cheque a fecha/Cam": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Cambio de Dirección": (
                "{Tipo usuario} solicita cambio de dirección en su {Documento}, presenta certificado de número municipal que indica {Nueva dirección (certificado)}"
                ". Se adjunta respaldo. Nombre: {Nombre Solicitante}, Fono Contacto: {Fono Contacto}, "
                "Respuesta: {Respuesta}."
            ),
            "Incorporación PAT": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Eliminación PAT": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Emisión de Factura": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Eliminación de Factura": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Cambio de Razón Social": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Eliminación de facturación electrónica": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Incorporación a facturación electrónica": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Analisis de fuga invisible": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Modificación factura (llena mandato)": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, Sueldo: {Sueldo}, Fono Contacto: {Fono Contacto}, "
                "Fecha Ingreso: {Fecha Ingreso}, "
                "Respuesta: {Respuesta}."
            ),
            "Certificado de No-Deuda": (
                "Tipo usuario: {Tipo usuario}, Nombre Solicitante: {Nombre Solicitante}, RUT: {RUT}, "
                "Edad: {Edad}, "
                "Respuesta: {Respuesta}."
            )
        }
  
        self.initUI()
        
    def load_classes_from_csv(self):
        with open(CSV_FILE_PATH, newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                clase = row['Clase']
                subclase = row['Subclase']
                campo = row['Campo']
                vinculasap = row.get('Vinculasap', '')  # Obtener el campo Vinculasap
                if clase not in self.classes:
                    self.classes[clase] = {}
                if subclase not in self.classes[clase]:
                    self.classes[clase][subclase] = {'fields': [], 'vinculasap': vinculasap}
                self.classes[clase][subclase]['fields'].append(campo)
  
    def initUI(self):
        # Crear área de texto para visualizar PDF
        self.vistaPDF = QTextEdit(self)
        self.vistaPDF.setReadOnly(True)
        #self.pdfView = QGraphicsView(self)
        #self.pdfView.setMinimumWidth(600)
        main_layout = QHBoxLayout()
        left_layout = QVBoxLayout()
        top_layout = QHBoxLayout()
        middle_layout = QVBoxLayout()
        bottom_layout = QVBoxLayout()
        
        # Comboboxes
        self.combo1 = QComboBox()
        self.combo1.addItems(self.classes.keys())
        self.combo1.currentIndexChanged.connect(self.update_combo2)
        
        self.combo2 = QComboBox()
        self.combo2.currentIndexChanged.connect(self.update_fields)
        
        top_layout.addWidget(QLabel("Clase"))
        top_layout.addWidget(self.combo1)
        top_layout.addWidget(QLabel("Acción"))
        top_layout.addWidget(self.combo2)

        # Añadir espacio vertical
        middle_layout.addItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        
        # Dynamic fields
        self.fields_layout = QGridLayout()
        middle_layout.addLayout(self.fields_layout)
        
        # Añadir espacio vertical
        middle_layout.addItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        
        # Layout para campos dinámicos
        self.layoutCampos = QVBoxLayout()
        
        # Añadir otro espacio vertical
        middle_layout.addItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        for _ in range(5):
            middle_layout.addItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        
        # Botones para PDF
        self.botonCargarPDF = QPushButton('Cargar PDF', self)
        self.botonCargarPDF.clicked.connect(self.cargarPDF)
        self.botonGuardarPDF = QPushButton('Guardar PDF', self)
        self.botonGuardarPDF.clicked.connect(self.guardarPDF)
        self.botonActualizarPDF = QPushButton('Actualizar PDF', self)
        self.botonActualizarPDF.clicked.connect(self.actualizarPDF)
        
        # Layout de botones
        buttons_layout = QHBoxLayout()
        self.clean_button = QtWidgets.QPushButton("Limpiar Campos", self)
        self.clean_button.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DialogResetButton))
        self.clean_button.clicked.connect(self.update_fields)
        self.call_data_button = QtWidgets.QPushButton("Llamar Datos", self)
        self.call_data_button.clicked.connect(self.llama_datos)
        self.call_data_button.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DialogOpenButton))
        self.save_sap_button = QtWidgets.QPushButton("Grabar en SAP")
        self.save_sap_button.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DialogSaveButton))
        self.save_sap_button.clicked.connect(self.grabar_en_sap)
        buttons_layout.addWidget(self.clean_button)
        buttons_layout.addWidget(self.call_data_button)
        buttons_layout.addWidget(self.save_sap_button)
        buttons_layout.addWidget(self.botonCargarPDF)
        buttons_layout.addWidget(self.botonGuardarPDF)
        buttons_layout.addWidget(self.botonActualizarPDF)

        # Añadir layout de botones al layout inferior
        bottom_layout.addItem(QSpacerItem(20, 80, QSizePolicy.Minimum, QSizePolicy.Expanding))
        #self.summary_text = QTextEdit()
        #self.summary_text.setReadOnly(True)
        #self.summary_label = QLabel("Clasificación Asociada en SAP")
        #self.summary_label.setCursor(QCursor(Qt.PointingHandCursor))
        #self.summary_label.mousePressEvent = self.copy_summary_text
        #bottom_layout.addWidget(self.summary_label)
        #bottom_layout.addWidget(self.summary_text)
        bottom_layout.addLayout(buttons_layout)

        left_layout.addLayout(top_layout)
        left_layout.addLayout(middle_layout)
        left_layout.addLayout(bottom_layout)
        
        main_layout.addLayout(left_layout)
        main_layout.addWidget(self.vistaPDF, 6)
        
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        self.resize(900, 350)
        #self.setGeometry(100, 100, 600, 300)
        self.update_combo2()
       
   # def copy_summary_text(self, event):
   #     if self.summary_text.toPlainText().strip() == "":
   #         QMessageBox.warning(self, "Advertencia", "El cuadro de texto está vacío. No se puede copiar al portapapeles.")
   #     else:
   #         clipboard = QApplication.clipboard()
   #         clipboard.setText(self.summary_text.toPlainText())
   #         QMessageBox.information(self, "Copiado", "El contenido del resumen ha sido copiado al portapapeles.")

    def update_combo2(self):
        self.combo2.clear()
        clase = self.combo1.currentText()
        if clase in self.classes:
            self.combo2.addItems(self.classes[clase].keys())
        self.update_fields()
        self.update_pdf_template()  # Llamar a la carta tipo asociada


    def update_fields(self):
        # Limpiar el layout de los campos
        while self.fields_layout.count():
            child = self.fields_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        self.fields.clear()
        self.field_labels.clear()
        
        clase = self.combo1.currentText()
        subclase = self.combo2.currentText()
        if clase in self.classes and subclase in self.classes[clase]:
            row = 0
            for campo in self.classes[clase][subclase]['fields']:
                label = QLabel(campo)
                if campo in ["Fecha Respuesta", "Fecha Emisión", "Fecha Inspección", "Fecha Mod. SAP", "Fecha Requerimiento", "Fecha (cambio/Insp/informe)"]:
                    field = QDateEdit()
                    field.setCalendarPopup(True)
                    field.setDate(QDate.currentDate())
                elif campo == "Tipo Cliente":
                    field = QComboBox()
                    field.addItems(["Señor", "Señora", "Señorita", "Señor(a)"])
                    field.setStyleSheet("QComboBox { background-color: white; }")
                elif campo == "Documento":
                    field = QComboBox()
                    field.addItems(["boleta", "factura"])
                    field.setStyleSheet("QComboBox { background-color: white; }")
                elif campo == "Inspección o instalación":
                    field = QComboBox()
                    field.addItems(["Instalación MAP", "Instalación MAP y varal", "Instalación MAP y varales"])
                    field.setStyleSheet("QComboBox { background-color: white; }")
                elif campo == "Reclamo" or campo == "Solicitud":
                    if self.combo2.currentText() == "Reclamos asociados a medidores":
                        field = QComboBox()
                        field.addItems(["Reclamo por medidor defectuoso", "Reclamo por medidor defectuoso con inspección (no realizado)",
                                         "Reclamo por medidor defectuoso (no realizado) -Fuga Interna-", "Reclamo por medidor defectuoso (debe firmar mandato)",
                                         "Reclamo por medidor reventado (no realizado) sin contacto", "Reclamo por medidor defectuoso sin acceso (no realizado)",
                                           "Reclamo por medidor reventado"])
                        field.setStyleSheet("QComboBox { background-color: white; }")      
                    if self.combo2.currentText() == "Solicitudes asociados a medidores":
                        field = QComboBox()
                        field.addItems(["Solicitud de verificación metrológica", "Solicitud de reprogramación cambio de medidor",
                                         "Solicitud medidor instalado por cliente", "Cambio de medidor solicitado por cliente",
                                         "Cambio de medidor solicitado por cliente (debe firmar mandato)", "Cambio de medidor mal ejecutado",
                                           "Cambio de medidor por vida util"])
                        field.setStyleSheet("QComboBox { background-color: white; }")  
                    if self.combo2.currentText() == "Cobros indebidos":
                        field = QComboBox()
                        field.addItems(["Cobro indebido medidor", "Cobro indebido medidor y varales",
                                         "Cobro indebido varales", "Cobro indebido interéses",
                                         "Cobro indebido alcantarillado"])
                        field.setStyleSheet("QComboBox { background-color: white; }")  
                    if self.combo2.currentText() == "Certificado Comercial":
                        field = QComboBox()
                        field.addItems(["Certificado de No-Deuda", "Certificado de Cliente Activo",
                                         "Certificado de Calidad de AP"])
                        field.setStyleSheet("QComboBox { background-color: white; }")  
                    if self.combo2.currentText() == "Certificado Operacional":
                        field = QComboBox()
                        field.addItems(["Certificado de No Deuda ni Daños Redes (T.Pend)",
                                         "Certificado de No Deuda ni Daños Redes", "Certificado de Estado de Redes"])
                        field.setStyleSheet("QComboBox { background-color: white; }")  
                    if self.combo2.currentText() == "Solicitud de devolución de saldo a favor":
                        field = QComboBox()
                        field.addItems(["Solicita Saldo a Favor"])
                        field.setStyleSheet("QComboBox { background-color: white; }")  
                    if self.combo2.currentText() == "Asociados a PAT":
                        field = QComboBox()
                        field.addItems(["Incorporación a  PAT", "Eliminación de PAT"])
                        field.setStyleSheet("QComboBox { background-color: white; }")  
                    if self.combo2.currentText() == "Entrega descuidada/Equivocada/Fuera de Plazo":
                        field = QComboBox()
                        field.addItems(["Entrega Descuidada", "Entrega Equivocada","Entrega Fuera de Plazo"])
                        field.setStyleSheet("QComboBox { background-color: white; }")  
                    if self.combo2.currentText() == "No entrega de Boletas/Facturas":
                        field = QComboBox()
                        field.addItems(["No entrega de boletas", "No entrega de facturas"])
                        field.setStyleSheet("QComboBox { background-color: white; }")  
                    if self.combo2.currentText() == "No entrega de Facturación Electronica":
                        field = QComboBox()
                        field.addItems(["No entrega de facturación Electrónica"])
                        field.setStyleSheet("QComboBox { background-color: white; }")  
                    if self.combo2.currentText() == "Analisis de Fuga":
                        field = QComboBox()
                        field.addItems(["Análisis de fuga", "Análisis de fuga (no inspeccionado)",
                                         "Análisis de fuga (fuera de rango)", "Análisis de fuga (no reparada)",
                                         "Análisis de fuga (servicio no residencial)", "Análisis de fuga (fuga WC)",
                                           "Análisis de fuga (no cuenta con alcantarillado)"])
                        field.setStyleSheet("QComboBox { background-color: white; }")  
                    if self.combo2.currentText() == "Solicitud Comercial":
                        field = QComboBox()
                        field.addItems(["Cambio de dirección","Eliminación de arranque",
                                        "Bloqueo de convenios", "Desbloqueo de convenios",
                                         "Emisión de factura", "Eliminación de factura",
                                         "Eliminación de facturación electrónica", "Cambio de razón social",
                                         "Eliminación de aporte a bomberos", "Eliminación de Despacho Postal",
                                         "Cambio tipo de servicio (prorrateo)", "Incorporación a facturación electrónica",
                                         "Bloqueo cheque a fecha", "Desbloqueo de cheque a fecha",
                                         "Inhibición de cambio de nombre en boleta", "Bloqueo de convenios y compromisos de pagos"])
                        field.setStyleSheet("QComboBox { background-color: white; }")  
                else:
                    field = QLineEdit()
                    field.editingFinished.connect(lambda f=field, c=campo: self.validate_field(f, c))
                    field.textChanged.connect(self.update_pdf_template)
                    if campo == "RUT":
                        field.setPlaceholderText("Ingrese RUT con guión y sin puntos")
                    elif campo == "Fono Contacto":
                        prefijo = "+56 9"
                        field.setText(prefijo)
                    elif campo == "Tipo de Respuesta":
                        field.setPlaceholderText("Ingrese (P, T o correo electrónico)")
                        field.editingFinished.connect(self.cambiar_respuesta)
                    else:
                        field.setPlaceholderText(f"Ingresa {campo.lower()}")
                self.fields_layout.addWidget(label, row, 0)
                self.fields_layout.addWidget(field, row, 1)
                self.field_labels[campo] = label
                self.fields[campo] = field
                row += 1
            vinculasap = self.classes[clase][subclase]['vinculasap']
            #self.summary_label.setText(f'TEXTO INGRESO <span style="color:blue;">{vinculasap}</span>')      
        self.update_pdf_template()
        self.save_sap_button.setStyleSheet("background-color : none")
        self.save_sap_button.blockSignals(True)
        self.save_sap_button.setToolTip("Botón deshabilitado")
     
    def cambiar_respuesta(self):
        text = self.fields["Tipo de Respuesta"].text().strip()
        if text.lower() == "p":
            self.fields["Tipo de Respuesta"].setText("Presencial")
        elif text.lower() == "t":
            self.fields["Tipo de Respuesta"].setText("Telefónica")
        elif text.lower() == "telefonica" or text.lower() == "telefónica":
            self.fields["Tipo de Respuesta"].setText("Telefónica")
        elif text.lower() == "presencial" :
            self.fields["Tipo de Respuesta"].setText("Presencial")
        else: 
            len(text) >= 1
            self.validar_correo(text)

    def validar_correo(self, text):
        if not re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', text):
            QMessageBox.warning(self, "Error", "El correo electrónico ingresado no es válido.")
            self.fields["Tipo de Respuesta"].clear()

    def validate_field(self, field, campo):
        if isinstance(field, QLineEdit):
            text = field.text()
        elif isinstance(field, QComboBox):
            text = field.currentText()
            print(text)
        elif isinstance(field, QDateEdit):
            text = field.date().toString('dd/MM/yyyy')

        if campo in ["Tipo Cliente", "Documento", "Inspección o instalación"]:
            field.setCurrentText(text.title())
        if campo == "RUT" and text != "":
            if not self.validate_rut(text):
                QMessageBox.warning(self, "Error", f"El campo {campo} no es válido.")
                field.clear()
            else:
                field.setText(self.format_rut(text))
        elif campo == "Edad" and not (text.isdigit() and 17 < int(text) < 121):
            QMessageBox.warning(self, "Error", f"El campo {campo} debe ser un número entre 18 y 120.")
            field.clear()
        elif campo == "Nombre Solicitante":
            field.setText(text.title())
        elif campo in ["IC", "ID.", "Contacto", "ODS", "Lectura/N° informe", "Lectura Fijada", "M3 Rebajado","Lec. Manual"] and not (text.isdigit()) :
             QMessageBox.warning(self, "Error", f"El campo {campo} debe contener solo números.")
             field.clear()
              

        #self.update_dynamic_text()

        self.colorboton_grabar_en_sap()

    def validate_rut(self, rut):
        rut = rut.replace(".", "").replace("-", "")
        if len(rut) < 7:
            return False
        rut_body = rut[:-1]
        rut_verifier = rut[-1].upper() if len(rut) > 7 else self.calculate_verifier(rut)
        reverse_rut_body = map(int, reversed(rut_body))
        factors = cycle(range(2, 8))
        s = sum(d * f for d, f in zip(reverse_rut_body, factors))
        mod = (-s) % 11
        if mod == 10:
            mod = 'K'
        return str(mod) == rut_verifier

    def calculate_verifier(self, rut):
        rut_body = rut
        reverse_rut_body = map(int, reversed(rut_body))
        factors = cycle(range(2, 8))
        s = sum(d * f for d, f in zip(reverse_rut_body, factors))
        mod = (-s) % 11
        if mod == 10:
            mod = 'K'
        return str(mod)

    def format_rut(self, rut):
        rut = rut.replace(".", "").replace("-", "")
        rut_body = rut[:-1]
        rut_verifier = rut[-1].upper() if len(rut) > 7 else self.calculate_verifier(rut)
        formatted_rut = f"{int(rut_body):,}".replace(",", ".") + "-" + rut_verifier
        return formatted_rut


    def colorboton_grabar_en_sap(self):
        missing_fields = []
        for label, field in self.fields.items():
            if label == "Texto cuerpo de requerimiento":
                continue  # No hacer obligatorio este campo
            if isinstance(field, QLineEdit) and not field.text():
                missing_fields.append(label)
            elif isinstance(field, QComboBox) and not field.currentText():
                missing_fields.append(label)
            elif isinstance(field, QDateEdit) and not field.date().isValid():
                missing_fields.append(label)
        
        if not missing_fields:
            self.save_sap_button.setStyleSheet("background-color : yellow")
            self.save_sap_button.blockSignals(False)
            self.save_sap_button.setToolTip("")
        else:
            self.save_sap_button.setStyleSheet("background-color : none")
            self.save_sap_button.blockSignals(True)
            self.save_sap_button.setToolTip(f"Campos faltantes: {', '.join(missing_fields)}")

    
    #def update_dynamic_text(self):
    #    clase = self.combo1.currentText()
    #    subclase = self.combo2.currentText()
    #    if subclase in self.text_structures:
    #        text_structure = self.text_structures[subclase]
    #        field_values = {
    #            campo: self.get_field_value_with_conditions(campo, subclase)
    #            for campo in self.fields
    #        }
    #        # Proporcionar valores predeterminados para las claves que faltan
    #        for key in re.findall(r'{(.*?)}', text_structure):
    #            if key not in field_values:
    #                field_values[key] = "N/A"  # Valor predeterminado

            # Verificar si "Nombre Solicitante" contiene información
    #        nombre_solicitante = field_values.get("Nombre Solicitante", "").strip()
    #        if nombre_solicitante:
    #            texto_cuerpo_requerimiento = self.fields.get("Texto cuerpo de requerimiento", QLineEdit()).text()
    #            texto_cuerpo_requerimiento = texto_cuerpo_requerimiento.capitalize()
    #            fono_contacto = field_values.get("Fono Contacto", "N/A")
    #            respuesta = field_values.get("Respuesta", "N/A")
                
    #            if texto_cuerpo_requerimiento:
    #                text = (
    #                    f"{texto_cuerpo_requerimiento}. "
    #                    f"Nombre Solicitante: {nombre_solicitante}, "
    #                    f"Fono Contacto: {fono_contacto}, "
    #                    f"Respuesta: {respuesta}"
    #                )
    #            else:
    #                text = text_structure.format(**field_values)
                
            #    self.summary_text.setText(text)
    #        else:
            #    self.summary_text.clear()
    #    else:
        #    self.summary_text.clear()


    def get_field_value_with_conditions(self, campo, subclase):
        field = self.fields[campo]
        value = self.get_field_value(campo)
        
        # Añadir condiciones específicas para el campo "Edad"
        if campo == "Edad":
            if value.isdigit():
                edad = int(value)
                if 20 <= edad <= 25:
                    return f"{value} (cliente mayor de edad)"
                elif 18 <= edad <= 40:
                    return f"{value} (cliente mediana edad)"
                else:
                    return f"{value} (cliente ya de avanzada edad)"
        
        # Condiciones para "Sueldo"
        if campo == "Sueldo" and value:
            try:
                sueldo = int(value.replace(".", "").replace("$ ", ""))
                if sueldo < 500000:
                    return f"{value} (sueldo bajo)"
                elif 500000 <= sueldo <= 1000000:
                    return f"{value} (sueldo medio)"
                else:
                    return f"{value} (sueldo alto)"
            except ValueError:
                return value  # En caso de error, devolver el valor original
        
        return value

    def get_field_value(self, campo):
        field = self.fields[campo]
        if isinstance(field, QLineEdit):
            return field.text()
        elif isinstance(field, QComboBox):
            return field.currentText()
        elif isinstance(field, QDateEdit):
            return field.date().toString('dd/MM/yyyy')
        return ""
    
    def grabar_en_sap(self):
    # Validación de campos
        for label, field in self.fields.items():
            if label == "Texto cuerpo de requerimiento":
                continue  # No hacer obligatorio este campo
            if isinstance(field, QLineEdit) and not field.text():
                QMessageBox.warning(self, "Error", f"El campo {label} no puede estar vacío.")
                return
            elif isinstance(field, QComboBox) and not field.currentText():
                QMessageBox.warning(self, "Error", f"El campo {label} no puede estar vacío.")
                return
            elif isinstance(field, QDateEdit) and not field.date().isValid():
                QMessageBox.warning(self, "Error", f"El campo {label} no puede estar vacío.")
                return

        # Generar texto dinámico
       # self.update_dynamic_text()
     #   dynamic_text = self.summary_text.toPlainText()
     #   QMessageBox.information(self, "SAP", f"Datos grabados en SAP: {dynamic_text}")
        self.update_fields()
        # Bloquea botón de guardar
        self.save_sap_button.setStyleSheet("background-color : none")
        self.save_sap_button.blockSignals(True)
        self.save_sap_button.setToolTip("Botón deshabilitado")

    #def update_summary(self):
    #    summary = ""
    #    for campo, field in self.fields.items():
    #        if isinstance(field, QLineEdit):
    #            summary += f"{field.text()} "
    #        elif isinstance(field, QDateEdit):
    #            summary += f"{field.date().toString('dd/MM/yyyy')} "
    #    self.summary_text.setPlainText(summary)

    def call_data(self):
        QMessageBox.information(self, "Llamar Datos", "Datos llamados exitosamente.")
  
    def llama_datos(self):
        ic_solicitante = self.fields.get("ODS", QLineEdit()).text()
       
        if not ic_solicitante:
            QMessageBox.warning(self, "Error", "Debe ingresar datos de ODS a llamar.")
            return
        datos_interlocutor = self.load_interlocutor_data()

        if ic_solicitante:
            self.handle_ic_solicitante(ic_solicitante, datos_interlocutor)
      
    def load_interlocutor_data(self):
        datos_interlocutor = []
        with open(CSV_REQUERIMIENTOS, newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                datos_interlocutor.append(row)
        return datos_interlocutor

    def handle_ic_solicitante(self, ic_solicitante, datos_interlocutor):
        for row in datos_interlocutor:
            if 'ODS' in row and row['ODS'] == ic_solicitante:
                self.update_fields_from_row(row)
               # QMessageBox.information(self, "Datos", "Datos cargados correctamente.")
                return
        QMessageBox.warning(self, "Error", "No se encontraron datos de ODS para cargar.")

    def handle_both_fields(self, ic_solicitante, id_servicio, datos_interlocutor):
        found_ic = None
        found_id = None
        for row in datos_interlocutor:
            if row['IC.'] == ic_solicitante:
                found_ic = row
            if row['ID.'] == id_servicio:
                found_id = row

        if found_ic:
            self.update_fields_from_row(found_ic)
            QMessageBox.information(self, "Datos", "Datos cargados correctamente. Solo se encontró coincidencia con IC. Solicitante.")
        else:
            QMessageBox.warning(self, "Error", "No se encontraron datos para los campos ingresados.")

    def update_fields_from_row(self, row):
        if self.fields["Nombre Solicitante"].text() != "":
            choice = QMessageBox.question(self, "Sobrescribir datos", 
                                          "El campo Nombre Solicitante contiene datos. ¿Desea sobrescribirlos?",
                                          QMessageBox.Yes | QMessageBox.No)
            if choice == QMessageBox.No:
                return

        if self.fields["Tipo de Respuesta"].text() != "":
            choice = QMessageBox.question(self, "Sobrescribir datos", 
                                          "El campo Tipo de Respuesta contiene datos. ¿Desea sobrescribirlos?",
                                          QMessageBox.Yes | QMessageBox.No)
            if choice == QMessageBox.No:
                return

        if "IC" in self.fields:
            if self.fields["IC"].text():
                choice = QMessageBox.question(self, "Sobrescribir datos", 
                                            "El campo IC contiene datos. ¿Desea sobrescribirlos?",
                                            QMessageBox.Yes | QMessageBox.No)
                if choice == QMessageBox.No:
                    return
                else:
                    self.fields["IC"].setText(row.get("IC", ""))
            else:
                self.fields["IC"].setText(row.get("IC", ""))           

        if self.fields["ID."].text() == "":
           self.fields["ID."].setText(row.get("ID", ""))
        
        if self.fields["IC"].text() == "":
           self.fields["IC"].setText(row.get("IC", "")) 

        #if self.fields["ODS"].text() != "":
        #   self.fields["ODS"].setText(row.get("ODS", "")) 
        
        if self.fields["Tipo de Respuesta"].text() == "":
           self.fields["Tipo de Respuesta"].setText(row.get("Respuesta", ""))
           print(self.fields["Tipo de Respuesta"].text())
        self.fields["Nombre Solicitante"].setText(row.get('Nombre Solicitante', ''))
        self.fields["N° Contacto"].setText(row.get('N° Contacto', ''))
 

        
     #### carga datos de base ODS.

    def save_to_sap(self):
        QMessageBox.information(self, "Grabar en SAP", "Datos grabados en SAP exitosamente.")


    ## del tema PDF
    def cargarPDF(self):
        try:
            nombreArchivo, _ = QFileDialog.getOpenFileName(self, "Cargar PDF", "", "PDF Files (*.pdf);;All Files (*)")
            if nombreArchivo:
                self.doc = fitz.open(nombreArchivo)
                self.actualizarVistaPDF()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al cargar el PDF: {e}")

    def guardarPDF(self):
        if not self.validate_fields():
            return

        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getSaveFileName(self, "Guardar PDF", "", "Archivos PDF (*.pdf);;Todos los archivos (*)", options=options)
        if fileName:
            self.actualizarPDF(fileName)

    def validate_fields(self):
        missing_fields = []
        for label, field in self.fields.items():
            if isinstance(field, QLineEdit) and not field.text():
                missing_fields.append(label)
            elif isinstance(field, QComboBox) and not field.currentText():
                missing_fields.append(label)
            elif isinstance(field, QDateEdit) and not field.date().isValid():
                missing_fields.append(label)
        
        if missing_fields:
            QMessageBox.warning(self, "Error", f"Campos faltantes: {', '.join(missing_fields)}")
            return False
        return True
    
    def actualizarPDF(self, filePath=None):
        if not filePath:
            options = QFileDialog.Options()
            filePath, _ = QFileDialog.getSaveFileName(self, "Actualizar PDF", "", "Archivos PDF (*.pdf);;Todos los archivos (*)", options=options)
        if filePath:
            try:
                doc = fitz.open()
                page = doc.new_page(width=fitz.paper_size("a4")[0], height=fitz.paper_size("a4")[1])
                logo = fitz.Pixmap("logo_empresa.png")  # Asegúrate de tener el logo en la misma carpeta
                page.insert_image((72, 72, 72 + logo.width, 72 + logo.height), pixmap=logo)
                page.insert_text((72, 72 + logo.height + 20), self.generate_pdf_text())
                doc.save(filePath)
                QMessageBox.information(self, "Información", "PDF actualizado y guardado exitosamente.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Ocurrió un error al actualizar el PDF: {str(e)}")

    def generate_pdf_text(self):
        subclase = self.combo2.currentText()
        if subclase in self.pdf_templates:
            template = self.pdf_templates[subclase]
            field_values = {campo: self.get_field_value(campo) for campo in self.fields}
            return template.format(**field_values)
        return ""

    def update_pdf_template(self):
        subclase = self.combo2.currentText()
        if subclase in self.pdf_templates:
            template = self.pdf_templates[subclase]
            field_values = {campo: self.get_field_value(campo) for campo in self.fields}
            
            # Manejar campos faltantes
            for key in template.format_map({}).keys():
                if key not in field_values:
                    field_values[key] = "N/A"  # Valor predeterminado para campos faltantes
            
        #    self.summary_text.setPlainText(template.format(**field_values))

    def actualizarVistaPDF(self):
        if self.doc:
            try:
                pagina = self.doc.load_page(0)  # Cargar la primera página del PDF
                texto = pagina.get_text("text")
                self.vistaPDF.setPlainText(texto)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al actualizar la vista del PDF: {e}")

    def generarTextoPDF(self):
        texto = f"Talca, {datetime.now().strftime('%Y-%m-%d')}\n\n"
        texto += f"Señor(a)\n{self.obtenerValorCampo('Nombre Solicitante')}\nPresente\n\n"
        texto += f"Referencia: Contacto Nº {self.obtenerValorCampo('Id. contacto')}\n"
        texto += f"Materia: {self.obtenerValorCampo('Solicitud')}\n"
        texto += f"Nº Cliente: {self.obtenerValorCampo('N°')}\n\n"
        texto += f"De nuestra consideración:\n\n"
        texto += f"Junto con saludarle muy cordialmente y en respuesta a su solicitud de {self.obtenerValorCampo('Solicitud')}, registrado en nuestro sistema con el número de atención {self.obtenerValorCampo('Id. contacto')}, podemos informar lo siguiente:\n\n"
        texto += f"Con fecha {self.obtenerValorCampo('Fecha Respuesta')} hemos procedido a efectuar {self.obtenerValorCampo('Solicitud')} en nuestro sistema.\n\n"
        texto += f"Adicionalmente, le señalamos que en el caso de presentar algún inconveniente de cualquier tipo con nuestro servicio, puede comunicarse con nosotros a través del Centro de Ayuda 6003311000 - 6003724000, desde celulares al *3311, a través de la Oficina Virtual www.nuevosur.cl donde le atenderá un ejecutivo.\n\n"
        texto += f"Finalmente y de acuerdo a la exigencia que la autoridad nos impone, es nuestro deber informar que si considera insatisfactoria esta respuesta, también tiene el derecho de concurrir a la Superintendencia de Servicios Sanitarios, ubicada en calle 2 Norte Nº1103, Talca, Fono 800 381 800.\n\n"
        texto += f"Sin otro particular se despide atentamente de usted.\n\n"
        texto += f"Servicio al Cliente"
        return texto

   ## del tema PDF

    # revisar para despues eliminar
    def updateText(self):
        # Crear el texto de la carta tipo
        carta = self.comboBox.currentText()
        if carta == "No entrega de boletas":
            text = f"""
            <h1>Carta de Invitación</h1>
            <p>Estimado/a <b>{self.fields[0].text()}</b>,</p>
            <p>Nos complace invitarle al evento que se celebrará el próximo <b>{self.fields[1].text()}</b> a las <b>{self.fields[2].text()}</b> en <b>{self.fields[3].text()}</b>. El motivo de este evento es <b>{self.fields[4].text()}</b>, y esperamos contar con su presencia para compartir este momento especial.</p>
            <p>El evento será organizado por <b>{self.fields[5].text()}</b>. Para cualquier consulta, puede contactarnos al teléfono <b>{self.fields[6].text()}</b> o al correo electrónico <b>{self.fields[7].text()}</b>.</p>
            <p>Le recordamos que el código de vestimenta es <b>{self.fields[8].text()}</b>. Agradecemos confirmar su asistencia antes del <b>{self.fields[9].text()}</b>.</p>
            <p>Esperamos verle pronto.</p>
            <p>Atentamente,</p>
            <p><b>{self.fields[5].text()}</b></p>
            """
        elif carta == "Carta de Agradecimiento":
            text = f"""
            <h1>Carta de Agradecimiento</h1>
            <p>Estimado/a Sr./Sra. <b>{self.fields[0].text()}</b>,</p>
            <p>Me dirijo a usted para expresarle mi más sincero agradecimiento por su valiosa colaboración en <b>{self.fields[2].text()}</b>. Su apoyo ha sido fundamental para el éxito de nuestro proyecto.</p>
            <p>En nombre de <b>{self.fields[4].text()}</b>, le agradezco su dedicación y profesionalismo. Esperamos seguir contando con su colaboración en futuros proyectos.</p>
            <p>Para cualquier consulta, no dude en contactarme al teléfono <b>{self.fields[6].text()}</b> o al correo electrónico <b>{self.fields[7].text()}</b>.</p>
            <p>Atentamente,</p>
            <p><b>{self.fields[3].text()}</b></p>
            <p><b>{self.fields[4].text()}</b></p>
            """
        elif carta == "Carta de Recomendación":
            text = f"""
            <h1>Carta de Recomendación</h1>
            <p>A quien corresponda,</p>
            <p>Por la presente, recomiendo a <b>{self.fields[0].text()}</b> para el puesto de <b>{self.fields[2].text()}</b> en su empresa. He tenido el placer de supervisar a <b>{self.fields[0].text()}</b> durante su tiempo en <b>{self.fields[3].text()}</b>, donde ha demostrado ser un/a profesional dedicado/a y competente.</p>
            <p>Durante su tiempo con nosotros, <b>{self.fields[0].text()}</b> ha mostrado habilidades excepcionales en <b>{self.fields[4].text()}</b> y ha contribuido significativamente a <b>{self.fields[5].text()}</b>.</p>
            <p>Estoy seguro/a de que <b>{self.fields[0].text()}</b> será una valiosa adición a su equipo. Para cualquier consulta adicional, puede contactarme al teléfono <b>{self.fields[6].text()}</b> o al correo electrónico <b>{self.fields[7].text()}</b>.</p>
            <p>Atentamente,</p>
            <p><b>{self.fields[1].text()}</b></p>
            <p><b>{self.fields[3].text()}</b></p>
            """
        elif carta == "Carta de Renuncia":
            text = f"""
            <h1>Carta de Renuncia</h1>
            <p>Estimado/a <b>{self.fields[5].text()}</b>,</p>
            <p>Por medio de la presente, presento mi renuncia formal al puesto de <b>{self.fields[1].text()}</b> en <b>{self.fields[6].text()}</b>, con efecto a partir del <b>{self.fields[2].text()}</b>. Mi último día de trabajo será el <b>{self.fields[3].text()}</b>.</p>
            <p>La decisión de renunciar ha sido difícil, pero he decidido aceptar una nueva oportunidad que me permitirá crecer profesionalmente. Agradezco sinceramente la oportunidad de haber trabajado en <b>{self.fields[6].text()}</b> y todo el apoyo brindado durante mi tiempo aquí.</p>
            <p>Estoy dispuesto/a a colaborar en la transición de mis responsabilidades para asegurar una salida ordenada y sin contratiempos.</p>
            <p>Atentamente,</p>
            <p><b>{self.fields[0].text()}</b></p>
            <p><b>{self.fields[1].text()}</b></p>
            <p><b>{self.fields[6].text()}</b></p>
            """
        elif carta == "Carta de Solicitud":
            text = f"""
            <h1>Carta de Solicitud</h1>
            <p>Estimado/a <b>{self.fields[4].text()}</b>,</p>
            <p>Me dirijo a usted para solicitar información detallada sobre <b>{self.fields[3].text()}</b>. Estoy interesado/a en conocer más sobre <b>{self.fields[8].text()}</b> y cómo podría participar en <b>{self.fields[9].text()}</b>.</p>
            <p>Agradecería recibir la información solicitada a la mayor brevedad posible. Puede contactarme al teléfono <b>{self.fields[7].text()}</b> o al correo electrónico <b>{self.fields[8].text()}</b>.</p>
            <p>Quedo a la espera de su pronta respuesta.</p>
            <p>Atentamente,</p>
            <p><b>{self.fields[0].text()}</b></p>
            <p><b>{self.fields[1].text()}</b></p>
            <p><b>{self.fields[2].text()}</b></p>
            """
        self.textItem.setHtml(text)

    def savePDF(self):
        # Verificar si todos los campos están completos
        missing_fields = [field for field in self.fields if not field.text()]
        if missing_fields:
            QMessageBox.warning(self, "Campos incompletos", "Por favor, complete todos los campos antes de guardar el PDF.")
            return

        # Obtener el tipo de carta y el nombre del remitente
        carta = self.comboBox.currentText()
        remitente = self.fields[5].text() if carta == "Carta de Invitación" else self.fields[3].text()
        fecha_actual = datetime.now().strftime("%Y-%m-%d")
        nombre_archivo = f"{carta}_{remitente}_{fecha_actual}.pdf"
        ruta_archivo = os.path.join(os.path.dirname(os.path.abspath(__file__)), nombre_archivo)

        # Verificar si el archivo ya existe
        if os.path.exists(ruta_archivo):
            respuesta = QMessageBox.question(self, "Archivo existente", "El archivo ya existe. ¿Desea sobreescribirlo?", QMessageBox.Yes | QMessageBox.No)
            if respuesta == QMessageBox.No:
                return

        # Crear el PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=10)

        # Añadir el logo
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Logo_empresa.png')
        if os.path.exists(logo_path):
            pdf.image(logo_path, 10, 8, 33)

        # Añadir el texto de la carta
        pdf.ln(40)  # Espacio después del logo
        pdf.multi_cell(0, 10, self.textItem.toPlainText())

        # Guardar el PDF
        pdf.output(ruta_archivo)
        QMessageBox.information(self, "PDF guardado", f"El archivo PDF se ha guardado como {nombre_archivo}.")
## revisar para despues eliminar


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
