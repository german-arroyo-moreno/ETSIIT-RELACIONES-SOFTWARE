#!/usr/bin/env python3

"""
@autor : MIGUEL ÁNGEL PÉREZ DÍAZ

    SUBDIRECCIÓN DE RELACIONES INTERNACIONALES ETSIIT (GRANADA)

"""

# IMPORTAMOS LOS MÓDULOS NECESARIOS PARA REALIZAR LA TAREA
import csv
import os
import argparse
import sys
from PyQt5 import QtCore, QtGui, QtWidgets, QtPrintSupport

# from PyQt5.QtPrintSupport import QPrintDialog, QPrinter, QPrintPreviewDialog


# from plantilla import Ui_MainWindow

# Generamos el objeto ArgumentParser
parser = argparse.ArgumentParser()
# Añadimos los argumentos
parser.add_argument('--input', required=True, help='input folder that contains the files csv')
parser.add_argument('--output', required=True, help='output file where the data will be save')


class Tabla:
    def __init__(self):
        """
        Constructor de la clase, inicializa el diccionario necesario y el contador de alumnos, y obtiene por argumento
        el nombre del fichero de salida.
        :param output_file: nombre del fichero de salida
        """
        self.fields_form = {}
        self.cont_alumnos = 0

    def add_row(self, path, delimiter=';', encoding='utf-8-sig'):
        """
        Método que añade una nueva fila con los datos de los alumnos obtenidos a partir del path pasado por agumento.
        La forma de codificar y el delimitador son argumentos por defecto.
        :param path:ruta del archivo de entrada del que vamos a extraer los datos de los alumnos
        :param delimiter: delimitador del fichero csv de entrada
        :param encoding: formato de codificación de los caracteres
        """
        with open(path, encoding=encoding) as File:
            fields = []
            reader = csv.reader(File, delimiter=delimiter)
            # Para cada fila del fichero
            for row in reader:
                # Comprobamos que no esté vacía y que no sea comentarios
                if row and not row[0].strip() in (None, "") and row[0][0] != '#':
                    # Comprobamos si no hemos guardado dicho campo de forumulario y le asignamos una lista
                    if not row[0].strip() in self.fields_form:
                        data = []
                        self.fields_form[row[0].strip()] = data
                        self.add_spaces(0, row)

                    # Extraemos los datos introducidos por los alumnos para cada campo
                    self.fields_form[row[0].strip()].append(row[1].strip())
                    # Guardamos los campos del alumnos para luego comprobar el caso de tener campos vacios para un
                    # determinado alumno
                    fields.append(row[0].strip())

            self.add_spaces(1, None, fields)

            # Aumentamos el contador de alumnos registrados
            self.cont_alumnos += 1

    def save_table(self, output_file, delimiter=';', encoding='utf-8-sig'):
        """
        Método que guarda los datos de los alumnos extraidos anteriormente en un fichero con el nombre de salida
        indicado como argumento del programa
        :param delimiter: delimitador del fichero csv de salida
        :param encoding: formato de codificación de los caracteres
        """
        # Creamos/abrimos el fichero de salida con el nombre dado por el usuario
        output = open(output_file, 'w', encoding=encoding, newline='')
        with output:
            writer = csv.writer(output, delimiter=delimiter)
            # Escribimos en el fichero los campos del formulario como fila en lugar de columnas
            writer.writerow(self.fields_form.keys())

            for i in range(self.cont_alumnos):
                data = []
                for key in self.fields_form.keys():
                    data.append(self.fields_form[key][i])
                    # Escribimos en el fichero las respuestas de cada alumno en una fila en lugar de columnas

                writer.writerow(data)

    def add_spaces(self, opcion, row=None, fields=None):
        """
        Método para añadir valor al vacío para un determinado alumno si no tiene valor para un campo ya guardado o en
        caso que el nuevo alumno tenga un campo nuevo se añaden valores vácios a los demás alumnos anteriormente
        guardados
        :param opcion: Variable que indica si se va a realizar la primera o segunda funcion indicada anteriormente en la
        descripcion del método
        :param row: fila que nos indica el campo nuevo generado por el alumno
        :param fields: lista que se utiliza para comprobar si falta algún campo ya existente al alumno, guarda los
        campos rellenos por el alumno
        """

        # Comprobamos las opciones
        if opcion == 0:
            # Recorremos los distintos alumnos guardados
            for i in range(self.cont_alumnos):
                # Añadimos valores vácios para ese determinado campo
                self.fields_form[row[0].strip()].append(" ")
        elif opcion == 1:
            # Recorremos los campos ya guardados
            for key in self.fields_form.keys():
                # Si el alumno no posee algun campo en su formulario, le añadimos un valor vacio para dicho alumno en
                # ese campo
                if not key in fields:
                    self.fields_form[key].append(" ")

        # Posible caso de error
        else:
            print("\n ERROR : Opción inválida ( opciones disponibles : 0 y 1")
            exit(3)

    def read_CSV(self, path, delimiter=';', encoding='utf-8-sig'):

        with open(path, encoding=encoding) as File:
            reader = csv.reader(File, delimiter=delimiter)
            # Para cada fila del fichero
            for index, row in enumerate(reader):

                # Comprobamos que no esté vacía y que no sea comentarios
                if row:
                    for col_index, col in enumerate(row):
                        if index == 0:
                            first_row = row

                            # Comprobamos si no hemos guardado dicho campo de forumulario y le asignamos una lista
                            if not col.strip() in self.fields_form:
                                data = []
                                self.fields_form[col.strip()] = data
                            # self.add_spaces(0, row)
                        else:
                            self.fields_form[first_row[col_index]].append(col.strip())

                    if index != 0:
                        self.cont_alumnos += 1



    def clear_table(self):
        self.cont_alumnos = 0
        self.fields_form = {}


def ls1(path):
    """Función que obtiene todos los archivos del directorio pasado como argumento a la función , comprueba si son
    ficheros csv y finalmente devolvemos una lista con el nombre de todos los archivos encontrados.

    :param path: directorio del que vamos a obtener todos los archivos

    :return: lista con los nombres de los diferentes archivos encontrados en el directorio
    """
    return [obj for obj in os.listdir(path) if os.path.isfile(os.path.join(path, obj)) and obj[-3:] in ['csv']]


def check_dir(dir):
    """ Función que comprueba si existe el directorio pasado por argumento

    :param dir:  directorio que vamos a comprobar si existe
    """
    if not os.path.isdir(dir):
        print('\n ERROR : El directorio no existe')
        exit(2)


def check_input():
    """Función que comprueba los argumentos pasados por línea de órdenes. Comprobamos si obtenemos solo
    dos argumentos (sin contar el nombre del programa), en caso contrario pintamos por pantalla el error
    producido con el formato válido de entrado y finalizamos la ejecución.

    :return: lista con los nombres de los dos argumentos pasados por línea de órdenes
    """

    args = parser.parse_args()
    check_dir(args.input)

    return args.input, args.output


"""
# Comprobamos y obtenemos los argumentos de entrada
folder, output_file = check_input()
# Obtenemos los archivos contenidos en el directorio pasado como entrada al programa
files = ls1(folder)

# Creamos un objeto de la clase Tabla pasandole al constructor el nombre del fichero de salida
table = Tabla(output_file)
# Recorremos todos los archivos del directorio
for file in files:
    # Guardamos los datos del alumno en una nueva fila del nuevo fichero
    table.add_row(os.path.join(folder, file))

# Guardamos la tabla final con toda la información de los alumnos
table.save_table()
"""


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")

        MainWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(0, 1, 771, 551))
        self.tableWidget.setRowCount(30)
        self.tableWidget.setColumnCount(30)

        self.tableWidget.setObjectName("tableWidget")
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setItem(0, 0, item)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.menuBar = QtWidgets.QMenuBar(MainWindow)
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 800, 26))
        self.menuBar.setObjectName("menuBar")
        self.menuArchivo = QtWidgets.QMenu(self.menuBar)
        self.menuArchivo.setObjectName("menuArchivo")
        self.actionGuardarCSV = QtWidgets.QAction(MainWindow)
        self.actionGuardarCSV.setObjectName("actionGuardarCSV")
        MainWindow.setMenuBar(self.menuBar)
        self.actionNuevo = QtWidgets.QAction(MainWindow)
        self.actionNuevo.setObjectName("actionNuevo")
        self.actionAbrir = QtWidgets.QAction(MainWindow)
        self.actionAbrir.setObjectName("actionAbrir")
        self.actionImportarAlumnos = QtWidgets.QAction(MainWindow)
        self.actionImportarAlumnos.setObjectName("actionImportarAlumnos")
        self.actionSalir = QtWidgets.QAction(MainWindow)
        self.actionSalir.setObjectName("actionSalir")

        self.menuArchivo.addSeparator()
        self.menuArchivo.addAction(self.actionNuevo)
        self.menuArchivo.addAction(self.actionAbrir)
        self.menuArchivo.addAction(self.actionImportarAlumnos)
        self.menuArchivo.addAction(self.actionGuardarCSV)
        self.menuArchivo.addAction(self.actionSalir)

        self.menuBar.addAction(self.menuArchivo.menuAction())

        self.menuEditar = QtWidgets.QMenu(self.menuBar)
        self.menuEditar.setObjectName("menuEditar")
        self.actionA_adir_filas = QtWidgets.QAction(MainWindow)
        self.actionA_adir_filas.setObjectName("actionA_adir_filas")
        self.actionA_adir_columnas = QtWidgets.QAction(MainWindow)
        self.actionA_adir_columnas.setObjectName("actionA_adir_columnas")

        self.menuEditar.addAction(self.actionA_adir_filas)
        self.menuEditar.addAction(self.actionA_adir_columnas)
        self.menuBar.addAction(self.menuEditar.menuAction())

        self.Tabla = Tabla()
        self.confirm_clean_before_open = False

        self.diccionario_keys = {}

        self.retranslateUi(MainWindow)
        self.actionNuevo.triggered.connect(self.clearTable)
        self.actionImportarAlumnos.triggered.connect(self.importarAlumnos)
        self.actionAbrir.triggered.connect(self.abrirCSV)
        self.actionGuardarCSV.triggered.connect(self.saveFile)
        self.actionA_adir_filas.triggered.connect(self.addRows)
        self.actionA_adir_columnas.triggered.connect(self.addCols)
        self.actionSalir.triggered.connect(sys.exit)

        self.tableWidget.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)

        self.h_layout = QtWidgets.QHBoxLayout()
        self.h_layout.addWidget(self.tableWidget)
        self.centralwidget.setLayout(self.h_layout)

        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    ########################################################################################################################
    ########################################################################################################################

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Subdirección de Internacionalización ETSIIT (Granada)"))
        __sortingEnabled = self.tableWidget.isSortingEnabled()
        self.tableWidget.setSortingEnabled(False)
        self.tableWidget.setSortingEnabled(__sortingEnabled)
        self.menuArchivo.setTitle(_translate("MainWindow", "Archivo"))
        self.actionGuardarCSV.setText(_translate("MainWindow", "Guardar como CSV"))
        self.actionNuevo.setText(_translate("MainWindow", "Nuevo"))
        self.actionAbrir.setText(_translate("MainWindow", "Abrir"))
        self.actionSalir.setText(_translate("MainWindow", "Salir"))
        self.actionImportarAlumnos.setText(_translate("MainWindow", "Importar alumnos..."))
        self.menuEditar.setTitle(_translate("MainWindow", "Editar"))
        self.actionA_adir_filas.setText(_translate("MainWindow", "Añadir filas"))
        self.actionA_adir_columnas.setText(_translate("MainWindow", "Añadir columnas"))

    ########################################################################################################################
    ########################################################################################################################

    def clearTable(self):

        if not self.confirm_clean_before_open:
            buttonReply = QtWidgets.QMessageBox()
            buttonReply.setWindowTitle("Ventana de confirmación")
            buttonReply.setText("¿Estás seguro que desea crear un nuevo documento?")
            buttonReply.setIcon(QtWidgets.QMessageBox.Question)
            buttonReply.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            buttonReply.setDefaultButton(QtWidgets.QMessageBox.Yes)

            result = buttonReply.exec()
        if self.confirm_clean_before_open or result == QtWidgets.QMessageBox.Yes:
            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(30)
            self.tableWidget.setColumnCount(30)
            self.Tabla.clear_table()
            self.diccionario_keys.clear()
            self.confirm_clean_before_open = False

    ########################################################################################################################
    ########################################################################################################################

    def importarAlumnos(self):

        añadir = False

        buttonReply = QtWidgets.QMessageBox()
        buttonReply.setWindowTitle("Ventana de confirmación")
        buttonReply.setText("¿Desea borrar el contenido actual o añadir nuevos datos a él?")
        buttonReply.setIcon(QtWidgets.QMessageBox.Question)
        buttonReply.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        buttonReply.setDefaultButton(QtWidgets.QMessageBox.Yes)
        buttonYES = buttonReply.button(QtWidgets.QMessageBox.Yes)
        buttonYES.setText("Borrar contenido")
        buttonNO = buttonReply.button(QtWidgets.QMessageBox.No)
        buttonNO.setText("Añadir nuevos datos")

        buttonReply.exec()

        if buttonReply.clickedButton() == buttonYES:
            self.confirm_clean_before_open = True
            self.clearTable()

        else:
            añadir = True
            self.Tabla.clear_table()



        fileName = QtWidgets.QFileDialog()
        folder = fileName.getExistingDirectory()

        if folder:

            files = ls1(folder)

            for file in files:
                # Guardamos los datos del alumno en una nueva fila del nuevo fichero
                self.Tabla.add_row(os.path.join(folder, file))

            self.saveData(añadir)

    ########################################################################################################################
    ########################################################################################################################

    def saveData(self, añadir):



        if añadir:
            prev_rows = self.tableWidget.rowCount()
            prev_cols = self.tableWidget.columnCount()
            self.tableWidget.setRowCount(self.Tabla.cont_alumnos + prev_rows)
            inicio = prev_rows-1


        else:
            self.tableWidget.setRowCount(self.Tabla.cont_alumnos + 1)
            self.tableWidget.setColumnCount(len(self.Tabla.fields_form.keys()) + 1)
            inicio=0



        for i in range(self.Tabla.cont_alumnos):
            for col in range(len(self.Tabla.fields_form.keys())):
                key = list(self.Tabla.fields_form.keys())[col]

                if col == 0:
                    chkBoxItem = QtWidgets.QTableWidgetItem()
                    chkBoxItem.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                    chkBoxItem.setCheckState(QtCore.Qt.Unchecked)
                    self.tableWidget.setItem(inicio + i + 1, col, chkBoxItem)

                if i == 0:
                    item = QtWidgets.QTableWidgetItem(key)
                    item.setBackground(QtGui.QColor(255, 128, 128))

                    if not key in self.diccionario_keys:
                        if añadir:
                            self.tableWidget.setColumnCount(prev_cols+1)
                            self.diccionario_keys[key] = prev_cols
                            self.tableWidget.setItem(i, prev_cols, item)
                            prev_cols+=1
                        else:
                            self.diccionario_keys[key] = col +1
                            self.tableWidget.setItem(i, col + 1, item)




                data = self.Tabla.fields_form[key][i]
                self.tableWidget.setItem(inicio + i + 1, self.diccionario_keys[key], QtWidgets.QTableWidgetItem(data))



    ########################################################################################################################
    ########################################################################################################################

    def saveFile(self):
        filename = QtWidgets.QFileDialog()
        name = filename.getSaveFileName(filter="Archivo CSV (*.csv)")
        if name[1] == 'Archivo CSV (*.csv)':
            self.writeCsv(name[0])

            confirm = QtWidgets.QMessageBox()
            confirm.setWindowTitle("Exportar a CSV")
            confirm.setText("Datos exportados con éxito.   ")
            confirm.exec()
        else:
            pass

        """
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(name[0])
            self.documento.print_(printer)



            confirm = QtWidgets.QMessageBox()
            confirm.setWindowTitle("Exportar a PDF")
            confirm.setText("Datos exportados con éxito.   ")
            confirm.exec()
        """

    ########################################################################################################################
    ########################################################################################################################

    def writeCsv(self, output_file):
        output = open(output_file, 'w', encoding='utf-8-sig', newline='')
        with output:
            writer = csv.writer(output, delimiter=';')
            # Escribimos en el fichero los campos del formulario como fila en lugar de columnas

            for row in range(self.tableWidget.rowCount()):
                rowdata = []
                for column in range(self.tableWidget.columnCount()):
                    item = self.tableWidget.item(row, column+1)
                    if item is not None:
                        rowdata.append(item.text())

                writer.writerow(rowdata)

    ########################################################################################################################
    ########################################################################################################################

    def addRows(self):
        prev_rows = self.tableWidget.rowCount()
        self.tableWidget.setRowCount(self.tableWidget.rowCount() + 5)
        for fil in range(prev_rows, self.tableWidget.rowCount()):
            chkBoxItem = QtWidgets.QTableWidgetItem()
            chkBoxItem.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            chkBoxItem.setCheckState(QtCore.Qt.Unchecked)
            self.tableWidget.setItem(fil, 0, chkBoxItem)

    def addCols(self):
        self.tableWidget.setColumnCount(self.tableWidget.columnCount() + 5)

    ########################################################################################################################
    ########################################################################################################################

    def emptyTable(self):

        for fil in range(self.tableWidget.rowCount()):
            for col in range(self.tableWidget.columnCount()):
                item = self.tableWidget.item(fil, col)
                if item is not None and item.text():
                    return False

        return True

    ########################################################################################################################
    ########################################################################################################################
    def abrirCSV(self):
        tableEmpty = True

        if not self.emptyTable():
            buttonReply = QtWidgets.QMessageBox()
            buttonReply.setWindowTitle("Ventana de confirmación")
            buttonReply.setText("¿Estás seguro? Se perderá el contenido actual")
            buttonReply.setIcon(QtWidgets.QMessageBox.Question)
            buttonReply.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            buttonReply.setDefaultButton(QtWidgets.QMessageBox.Yes)

            result = buttonReply.exec()

            if result == QtWidgets.QMessageBox.Yes:
                self.confirm_clean_before_open = True
                self.clearTable()
            else:
                tableEmpty = False

        if tableEmpty:
            filename = QtWidgets.QFileDialog()
            name = filename.getOpenFileName(filter="Archivo CSV (*.csv)")
            self.Tabla.read_CSV(name[0])
            self.saveData(False)

            confirm = QtWidgets.QMessageBox()
            confirm.setWindowTitle("Importar de CSV")
            confirm.setText("Datos importados con éxito.   ")
            confirm.exec()


########################################################################################################################
########################################################################################################################



app = QtWidgets.QApplication([])  # Creamos app y le pasamos una lista de argumentos vacíos
ventana = Ui_MainWindow()
main_window = QtWidgets.QMainWindow()
ventana.setupUi(main_window)
main_window.show()

sys.exit(app.exec())
