#!/usr/bin/env/python3

"""
This file is part of ToR Conversor.

    ToR Conversor is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    ToR Conversor is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with ToR Conversor.  If not, see <https://www.gnu.org/licenses/>.
"""

import csvh
import tor
import csv
import sys
import os
from subprocess import call
from PyQt5.QtWidgets import QTabWidget,QPushButton,QWidget,QTableWidget,QMainWindow,QMessageBox,QMenu,QMenuBar,QVBoxLayout,QHeaderView,QHBoxLayout,QStatusBar,QAction,QComboBox,QLabel,QFileDialog,QTableWidgetItem,QApplication,QProgressBar,QToolButton,QDialog
from PyQt5.QtGui import QColor,QRegion
from PyQt5.QtCore import QMetaObject,QRect,Qt,QCoreApplication
from configparser import ConfigParser



HOME = "España"
UNIV_COLUMN = "Código VICERRECTORADO donde se han cursado los estudios:"

HOME = HOME.strip().upper()
UNIV_COLUMN = UNIV_COLUMN.strip()


def exportCSVToR(personalData, ToR, fileName):
    csv_data = []
    for d in personalData:
        csv_data.append([d, personalData[d]])
    csv_data.append([])
    idv = 1
    for d in ToR:
        csv_data.append(["Bloque:", idv])
        csv_data.append(
            ["", "Asignatura Destino", "Créditos", "Nota Destino", "Sugerencia Origen", "Min. Origen", "Máx. Origen",
             "Min. Destino", "Máx. Destino", "Alias"])
        for subject in ToR[d][0]:
            csv_data.append([""] + subject)
        csv_data.append([])
        csv_data.append(["", "Asignatura Origen", "Créditos"])
        for subject in ToR[d][1]:
            csv_data.append([""] + subject)
        csv_data.append([])
        idv += 1

    csvh.exportRawCSVData(fileName, csv_data)


def readData(fileName, origin, destination):
    eq_data = csvh.importRawCSVData(fileName)

    data = {}
    for d in eq_data:
        d[0] = d[0].strip().upper()
        if str(d[0])[0] == "#":
            pass
        else:

            data[d[0]] = d[1:]

    r = data[destination]
    raw_destination = []
    for d in r:
        raw_destination.append(str(d))
    raw_origin = data[origin]
    return (raw_destination, raw_origin)


def readToR(fileName):
    ToR = csvh.importRawCSVData(fileName)

    subjectData = []
    readSubjects = False
    personalData = {}
    for d in ToR:
        d[0] = d[0].replace("\n", " ").replace("\r", " ").strip()
        if readSubjects:
            subjectData.append(d)
        else:
            if d[0] == "Asignatura":
                readSubjects = True
            elif (d[0] != "" and str(d[0])[0] == "#") or d[0] == "":
                pass
            else:
                personalData[d[0]] = str(d[1]).strip()

    return (personalData, subjectData)


def ls1(path, option):
    """Función que obtiene todos los archivos del directorio pasado como argumento a la función , comprueba si son
    ficheros csv y finalmente devolvemos una lista con el nombre de todos los archivos encontrados.

    :param path: directorio del que vamos a obtener todos los archivos

    :return: lista con los nombres de los diferentes archivos encontrados en el directorio
    """
    if option:
        return [obj for obj in os.listdir(path) if os.path.isfile(os.path.join(path, obj)) and obj[-3:] in ['png']]
    else:
        return [obj for obj in os.listdir(path) if os.path.isfile(os.path.join(path, obj)) and obj[-3:] in ['ods']]

class HoverButton(QToolButton):
    def __init__(self, parent=None):
        super(HoverButton, self).__init__(parent)
        self.setStyleSheet('''border-image: url("config/images/back.png")''')

    def resizeEvent(self, event):
        self.setMask(QRegion(self.rect(), QRegion.Ellipse))
        QToolButton.resizeEvent(self, event)

class BackButton(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        self.layout = QHBoxLayout(parent)

        self.back_buttom = HoverButton()
        self.back_buttom.setMinimumWidth(50)
        self.back_buttom.setMinimumHeight(50)
        self.layout.addWidget(self.back_buttom)
        self.layout.addSpacing(0)
        self.setLayout(self.layout)
        self.layout.setAlignment(Qt.AlignLeft)
        self.layout.setContentsMargins(0,0,0,0)
        self.layout.setSpacing(0)

class SingleConversor(QDialog):
    def __init__(self, parent=None):
        super(QDialog,self).__init__(parent)

        self.setWindowTitle("Conversor Manual")
        self.resize(950, 500)


        layout = QVBoxLayout(parent)
        self.config = {}


        self.pushButton = QPushButton("Generar")

        self.pushButton.setMaximumHeight(31)
        self.pushButton.setStyleSheet("background-color:rgb(220,128,128)")
        self.pushButton.clicked.connect(self.generate_manual)

        self.tableWidget = QTableWidget(self)
        self.tableWidget.setGeometry(QRect(0, 1, 771, 551))
        self.tableWidget.setRowCount(30)
        self.tableWidget.setColumnCount(30)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        self.select_country = QComboBox(self)
        self.select_country.setObjectName("selectCountry")
        self.texto_select = QLabel("Selecciona el país de destino :")

        layout.setSpacing(25)
        layout.addWidget(self.texto_select)
        layout.addWidget(self.select_country)
        layout.addWidget(self.tableWidget)
        layout.addWidget(self.pushButton)

        self.setLayout(layout)

    def getCountries(self):

        with open(self.config['conversion_table'], encoding='utf-8-sig') as File:
            reader = csv.reader(File, delimiter=';', quotechar='"')
            fields_form = []
            for row in reader:
                # Comprobamos que no esté vacía y que no sea comentarios
                if row and not row[0].strip() in (None, "") and row[0][0] != '#':
                    # Comprobamos si no hemos guardado dicho campo de forumulario y le asignamos una lista
                    if not row[0].strip() in fields_form:
                        fields_form.append(row[0].strip())

        self.select_country.addItems(fields_form)


    def generate_manual(self):

        data = []

        bloques = []
        for i in range(1, self.tableWidget.rowCount()):
            row = []
            for p in range(self.tableWidget.columnCount() - 2):
                item = self.tableWidget.item(i, p)
                if item is not None and item.text():
                    if p == 8:
                        bloques.append(i)
                    row.append(item.text())
                else:
                    row.append("")

            data.append(row)

        american, ToR = tor.parseToR(data)
        raw_destination, raw_origin = readData(self.config['conversion_table'], HOME, self.select_country.currentText().upper())
        x, aliasx, y, aliasy = tor.expandScores(raw_origin, raw_destination, american)


        # 3. Expand the table to score suggestions for each destination subject
        ToR = tor.extendToR(ToR, x, aliasx, y, aliasy, american)

        for block in ToR:
            maxCr = 0
            fail = -1
            score = 0
            n = len(ToR[block][0])
            for i in range(n):
                maxCr += ToR[block][0][i][1]

            for i in range(n):
                if ToR[block][0][i][3] < 5:
                    fail = ToR[block][0][i][3]
                    score = fail
                    break
                else:
                    score += ToR[block][0][i][3] * (ToR[block][0][i][1] / maxCr)
            for i in range(max(n, len(ToR[block][1]))):

                score = float("{:.1f}".format(score))
                if score < 0:
                    crLabel = "NO PRESENTADO"
                elif score < 5:
                    crLabel = "(SUSPENSO)"
                elif score < 7:
                    crLabel = "(APROBADO)"
                elif score < 9:
                    crLabel = "(NOTABLE)"
                elif score < 9.5:
                    crLabel = "(SOBRESALIENTE)"
                else:
                    crLabel = "(MATRÍCULA)"

            cont = bloques.pop(0)
            self.tableWidget.setItem(cont, 9, QTableWidgetItem(str(score)))
            self.tableWidget.setItem(cont, 10, QTableWidgetItem(crLabel))

    def convertSingle(self):


        switcher = {
            0: "Asignatura destino",
            1: "ECTS",
            2: "otro",
            3: "ID_BLOQUE",
            4: "Calificación Destino",
            5: "Asignatura origen",
            6: "ECTS",
            7: "LRU",
            8: "ID_BLOQUE",
            9: "Calificación Origen",
            10: "CRLabel"
        }

        self.tableWidget.setColumnCount(11)
        for i in range(self.tableWidget.rowCount()):
            for p in range(self.tableWidget.columnCount()):
                if i == 0:
                    item = QTableWidgetItem(switcher.get(p))
                    if p == 0 or p == 5:
                        item.setBackground(QColor(123, 193, 233))
                    else:
                        item.setBackground(QColor(255, 128, 128))
                    self.tableWidget.setItem(i, p, item)


class MyTableWidget(QWidget):

    def __init__(self,parent):
        super(QWidget, self).__init__(parent)
        self.layout = QVBoxLayout(self)
        # Inicializar pantalla de pestañas
        self.tabs = QTabWidget()
        self.tabs.tabCloseRequested.connect(lambda index: self.tabs.removeTab(index))
        self.tabs.setTabsClosable(True)
        self.datos_alumno = {}
        self.config = {}

        # Agregar pestañas al widget
        self.generateALL = QPushButton("Generar PDF todos los alumnos")
        self.generateALL.setStyleSheet("background-color:rgb(220,128,128)")
        self.generateALL.clicked.connect(self.exportPDF_ALL)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.hide()
        self.layout.addSpacing(0)
        self.layout.addWidget(self.tabs)
        self.layout.addWidget(self.progress_bar)
        self.layout.addWidget(self.generateALL)
        self.setLayout(self.layout)


    def crear_tabs(self,files):
        self.tabs.clear()
        self.files = files
        self.datos_alumno.fromkeys(files)
        for file in files:
            tab = QWidget()
            self.datos_alumno[file] = []
            tableWidget = QTableWidget()
            tableWidget.setGeometry(QRect(0, 1, 771, 551))
            tableWidget.setRowCount(30)
            tableWidget.setColumnCount(30)
            tableWidget.setObjectName("tableWidget")
            tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
            self.datos_alumno[file].append(tableWidget)

            tab.layout = QVBoxLayout(self)
            botonesLayout = QHBoxLayout(self)
            tab.layout.setAlignment(Qt.AlignCenter)
            tab.layout.setSpacing(25)
            pushButton1 = QPushButton("Generar PDF alumno")
            pushButton2 = QPushButton("Exportar CSV alumno")
            pushButton1.setStyleSheet("background-color:rgb(228,195,195)")
            pushButton2.setStyleSheet("background-color:rgb(191,243,191)")
            tab.layout.addWidget(tableWidget)
            tab.layout.addLayout(botonesLayout)
            botonesLayout.addWidget(pushButton2)
            botonesLayout.addWidget(pushButton1)
            tab.setLayout(tab.layout)

            pushButton2.clicked.connect(self.exportCSV)
            pushButton1.clicked.connect(self.exportPDF)

            self.tabs.addTab(tab,file)

    def exportCSV(self):
        if self.tabs.count()>0:
            file = self.tabs.tabText(self.tabs.currentIndex())
            self.check_info_show(self.tabs.tabText(self.tabs.currentIndex()))
            exportCSVToR(self.datos_alumno[self.tabs.tabText(self.tabs.currentIndex())][2], self.datos_alumno[self.tabs.tabText(self.tabs.currentIndex())][3], os.path.join(self.datos_alumno[self.tabs.tabText(self.tabs.currentIndex())][1], self.tabs.tabText(self.tabs.currentIndex())[:-4] + "--debug_mode.csv"))

            confirm = QMessageBox()
            confirm.setIcon(QMessageBox.Information)
            confirm.setWindowTitle("Exportar CSV alumno")
            confirm.setText("Alumno " + file[:-4] + " exportado a CSV con éxito.")
            confirm.exec()
        else:
            confirm = QMessageBox()
            confirm.setIcon(QMessageBox.Critical)
            confirm.setWindowTitle("Exportar CSV alumno")
            confirm.setText("No hay alumnos actualmente")
            confirm.exec()
    def exportPDF(self):
        if self.tabs.count()>0:
            self.check_info_show(self.tabs.tabText(self.tabs.currentIndex()))
            file = self.tabs.tabText(self.tabs.currentIndex())
            personalData = self.datos_alumno[file][2]
            Tor = self.datos_alumno[file][3]
            folder_path = self.datos_alumno[self.tabs.tabText(self.tabs.currentIndex())][1]
            f = open(os.path.join(os.path.abspath(os.getcwd()), "config/tex/template01.tex"), "r", encoding='utf8')
            text = f.read()
            f.close()

            for d in personalData:
                personalData[d] = personalData[d].replace("\\", "/").replace("_", "\_").replace("$","\$")
                text = text.replace("[[" + str(d) + "]]", personalData[d].upper())
                text = text.replace("{{" + str(d) + "}}", personalData[d])

            # 5. Add the final calification:
            table = ""

            for block in Tor:
                table += "\\hline \\hline \n"
                maxCr = 0
                fail = -1
                score = 0
                n = len(Tor[block][0])

                for i in range(n):
                    maxCr += Tor[block][0][i][1]

                for i in range(n):
                    if Tor[block][0][i][3] < 5:
                        fail = Tor[block][0][i][3]
                        score = fail
                        break
                    else:
                        score += Tor[block][0][i][3] * (Tor[block][0][i][1] / maxCr)
                for i in range(max(n, len(Tor[block][1]))):
                    try:
                        sOrig = Tor[block][1][i][0]
                    except:
                        sOrig = ""
                    try:
                        sDst = Tor[block][0][i][0]
                    except:
                        sDst = ""
                    try:
                        crOrig = Tor[block][1][i][1]
                    except:
                        crOrig = ""
                    try:
                        crDst = Tor[block][0][i][1]
                    except:
                        crDst = ""
                    try:
                        scoreDst = Tor[block][0][i][2]
                    except:
                        scoreDst = ""

                    score = float("{:.1f}".format(score))
                    if score < 0:
                        crLabel = "NO PRESENTADO"
                    elif score < 5:
                        crLabel = "(SUSPENSO)"
                    elif score < 7:
                        crLabel = "(APROBADO)"
                    elif score < 9:
                        crLabel = "(NOTABLE)"
                    elif score < 9.5:
                        crLabel = "(SOBRESALIENTE)"
                    else:
                        crLabel = "(MATRÍCULA)"

                    if score < 0 and sOrig != "":
                        table += " {} & {} &  & {} & {} & {} &  & {} \\\\ \\hline \n".format(sDst, crDst,
                                                                                             scoreDst,
                                                                                             sOrig,
                                                                                             crOrig,
                                                                                             crLabel)
                    elif sOrig == "":
                        table += " {} & {} &  & {} & {} & {} &  & \\\\ \\hline \n".format(sDst, crDst,
                                                                                          scoreDst,
                                                                                          sOrig, crOrig)
                    else:
                        table += " {} & {} &  & {} & {} & {} &  & {:.1f} {} \\\\ \\hline \n".format(sDst,
                                                                                                    crDst,
                                                                                                    scoreDst,
                                                                                                    sOrig,
                                                                                                    crOrig,
                                                                                                    score,
                                                                                                    crLabel)

            text = text.replace("{{:SUBJECT-LIST:}}", table)
            f = open(os.path.join(folder_path, file[:-3] + "tex"), "w", encoding='utf8')
            f.write(text)
            f.close()

            cmd = '"%s"  -output-directory="%s" "%s"' % (
                self.config['pdflatex'], folder_path, os.path.join(folder_path, file[:-3] + "tex"))
            call(cmd)


            confirm = QMessageBox()
            confirm.setIcon(QMessageBox.Information)
            confirm.setWindowTitle("Generar calificaciones finales")
            confirm.setText("Calificaciones de " + file[:-4]+ " generadas con éxito.")
            confirm.exec()

        else:
            confirm = QMessageBox()
            confirm.setIcon(QMessageBox.Critical)
            confirm.setWindowTitle("Generar calificaciones finales")
            confirm.setText("No hay alumnos actualmente")
            confirm.exec()

    def check_tabs(self):
        self.progress_bar.show()
        self.progress_bar.setValue(0)
        num_tabs = self.tabs.count()
        self.files.clear()
        for i in range(num_tabs):
            self.files.append(self.tabs.tabText(i))


    def exportPDF_ALL(self):

        self.check_tabs()

        if self.tabs.count()> 0:
            for i, file in enumerate(self.files):
                self.check_info_show(file)

                personalData = self.datos_alumno[file][2]
                Tor = self.datos_alumno[file][3]
                folder_path = self.datos_alumno[file][1]
                f = open(os.path.join(os.path.abspath(os.getcwd()), "config/tex/template01.tex"), "r", encoding='utf8')
                text = f.read()
                f.close()

                for d in personalData:
                    personalData[d] = personalData[d].replace("\\", "/").replace("_", "\_").replace("$",
                                                                                                              "\$")
                    text = text.replace("[[" + str(d) + "]]", personalData[d].upper())
                    text = text.replace("{{" + str(d) + "}}", personalData[d])

                # 5. Add the final calification:
                table = ""

                for block in Tor:
                    table += "\\hline \\hline \n"
                    maxCr = 0
                    fail = -1
                    score = 0
                    n = len(Tor[block][0])

                    for i in range(n):
                        maxCr += Tor[block][0][i][1]

                    for i in range(n):
                        if Tor[block][0][i][3] < 5:
                            fail = Tor[block][0][i][3]
                            score = fail
                            break
                        else:
                            score += Tor[block][0][i][3] * (Tor[block][0][i][1] / maxCr)
                    for i in range(max(n, len(Tor[block][1]))):
                        try:
                            sOrig = Tor[block][1][i][0]
                        except:
                            sOrig = ""
                        try:
                            sDst = Tor[block][0][i][0]
                        except:
                            sDst = ""
                        try:
                            crOrig = Tor[block][1][i][1]
                        except:
                            crOrig = ""
                        try:
                            crDst = Tor[block][0][i][1]
                        except:
                            crDst = ""
                        try:
                            scoreDst = Tor[block][0][i][2]
                        except:
                            scoreDst = ""

                        score = float("{:.1f}".format(score))
                        if score < 0:
                            crLabel = "NO PRESENTADO"
                        elif score < 5:
                            crLabel = "(SUSPENSO)"
                        elif score < 7:
                            crLabel = "(APROBADO)"
                        elif score < 9:
                            crLabel = "(NOTABLE)"
                        elif score < 9.5:
                            crLabel = "(SOBRESALIENTE)"
                        else:
                            crLabel = "(MATRÍCULA)"

                        if score < 0 and sOrig != "":
                            table += " {} & {} &  & {} & {} & {} &  & {} \\\\ \\hline \n".format(sDst, crDst,
                                                                                                 scoreDst,
                                                                                                 sOrig,
                                                                                                 crOrig,
                                                                                                 crLabel)
                        elif sOrig == "":
                            table += " {} & {} &  & {} & {} & {} &  & \\\\ \\hline \n".format(sDst, crDst,
                                                                                              scoreDst,
                                                                                              sOrig, crOrig)
                        else:
                            table += " {} & {} &  & {} & {} & {} &  & {:.1f} {} \\\\ \\hline \n".format(sDst,
                                                                                                        crDst,
                                                                                                        scoreDst,
                                                                                                        sOrig,
                                                                                                        crOrig,
                                                                                                        score,
                                                                                                        crLabel)

                text = text.replace("{{:SUBJECT-LIST:}}", table)
                f = open(os.path.join(folder_path, file[:-3] + "tex"), "w", encoding='utf8')
                f.write(text)
                f.close()

                cmd = '"%s"  -output-directory="%s" "%s"' % (
                    self.config['pdflatex'], folder_path, os.path.join(folder_path, file[:-3] + "tex"))
                call(cmd)

                if i == len(self.files):
                    self.progress_bar.setValue(100)
                else:
                    self.progress_bar.setValue(self.progress_bar.value() + int(100 / len(self.files)))



            confirm = QMessageBox()
            confirm.setIcon(QMessageBox.Information)
            confirm.setWindowTitle("Generar calificaciones finales")
            confirm.setText("Calificaciones generadas con éxito.")
            confirm.exec()
            self.progress_bar.hide()
        else:
            confirm = QMessageBox()
            confirm.setIcon(QMessageBox.Critical)
            confirm.setWindowTitle("Generar calificaciones finales")
            confirm.setText("No hay alumnos actualmente")
            confirm.exec()
            self.progress_bar.hide()



    def show_info_check(self, personalData, ToR,file):

        tableWidget = self.datos_alumno[file][0]
        tableWidget.setRowCount(0)

        for d in personalData:
            rowPosition = tableWidget.rowCount()
            tableWidget.insertRow(rowPosition)
            item = QTableWidgetItem(d)
            item.setBackground(QColor(228, 195, 195))
            tableWidget.setItem(rowPosition, 0, item)
            tableWidget.setItem(rowPosition, 1, QTableWidgetItem(personalData[d]))

        tableWidget.setRowCount(tableWidget.rowCount() + 1)
        idv = 1
        for d in ToR:
            rowPosition = tableWidget.rowCount()
            tableWidget.insertRow(rowPosition)

            item = QTableWidgetItem("Bloque")
            item.setBackground(QColor(123, 193, 233))
            tableWidget.setItem(rowPosition, 0, item)
            item = QTableWidgetItem(str(idv))
            item.setBackground(QColor(123, 193, 233))
            tableWidget.setItem(rowPosition, 1, item)
            for blue in range(2, 10):
                item = QTableWidgetItem("")
                item.setBackground(QColor(123, 193, 233))
                tableWidget.setItem(rowPosition, blue, item)

            rowPosition = tableWidget.rowCount()
            tableWidget.insertRow(rowPosition)
            item = QTableWidgetItem("Asignatura Destino")
            item.setBackground(QColor(228, 195, 195))
            tableWidget.setItem(rowPosition, 1, item)

            item = QTableWidgetItem("Créditos")
            item.setBackground(QColor(228, 195, 195))
            tableWidget.setItem(rowPosition, 2, item)

            item = QTableWidgetItem("Nota Destino")
            item.setBackground(QColor(228, 195, 195))
            tableWidget.setItem(rowPosition, 3, item)

            item = QTableWidgetItem("Sugerencia Origen")
            item.setBackground(QColor(228, 195, 195))
            tableWidget.setItem(rowPosition, 4, item)

            item = QTableWidgetItem("Min. Origen")
            item.setBackground(QColor(228, 195, 195))
            tableWidget.setItem(rowPosition, 5, item)

            item = QTableWidgetItem("Max. Origen")
            item.setBackground(QColor(228, 195, 195))
            tableWidget.setItem(rowPosition, 6, item)

            item = QTableWidgetItem("min. Destino")
            item.setBackground(QColor(228, 195, 195))
            tableWidget.setItem(rowPosition, 7, item)

            item = QTableWidgetItem("Max. Destino")
            item.setBackground(QColor(228, 195, 195))
            tableWidget.setItem(rowPosition, 8, item)

            item = QTableWidgetItem("Alias")
            item.setBackground(QColor(228, 195, 195))
            tableWidget.setItem(rowPosition, 9, item)

            for subject in ToR[d][0]:
                rowPosition = tableWidget.rowCount()
                tableWidget.insertRow(rowPosition)
                for col in range(len(subject)):
                    tableWidget.setItem(rowPosition, col + 1, QTableWidgetItem(str(subject[col])))

            tableWidget.setRowCount(tableWidget.rowCount() + 1)
            rowPosition = tableWidget.rowCount()
            tableWidget.insertRow(rowPosition)

            item = QTableWidgetItem("Asignatura Origen")
            item.setBackground(QColor(228, 195, 195))
            tableWidget.setItem(rowPosition, 1, item)

            item = QTableWidgetItem("Créditos")
            item.setBackground(QColor(228, 195, 195))
            tableWidget.setItem(rowPosition, 2, item)

            for subject in ToR[d][1]:
                rowPosition = tableWidget.rowCount()
                tableWidget.insertRow(rowPosition)
                for col in range(len(subject)):
                    tableWidget.setItem(rowPosition, col + 1, QTableWidgetItem(str(subject[col])))

            tableWidget.setRowCount(tableWidget.rowCount() + 1)
            idv += 1

    def check_info_show(self,file):

        idv = -1
        personalData = self.datos_alumno[file][2]
        Tor = self.datos_alumno[file][3]
        tableWidget = self.datos_alumno[file][0]

        for d in personalData:
            idv += 1
            item = tableWidget.item(idv, 1)
            if item is not None and item.text():
                personalData[d] = item.text()
            else:
                personalData[d] = ""

        for d in Tor:
            idv += 3

            for subject in range(len(Tor[d][0])):

                idv += 1

                for col in range(len(Tor[d][0][subject])):
                    item = tableWidget.item(idv, col + 1)
                    if item is not None and item.text():

                        if col == 1:

                            Tor[d][0][subject][col] = int(item.text())
                        elif col == 8 or col == 0 or not item.text()[0].isdigit():

                            Tor[d][0][subject][col] = item.text()
                        else:
                            Tor[d][0][subject][col] = float(item.text())
                    else:
                        Tor[d][0][subject][col] = ""

            idv += 2

            for subject in range(len(Tor[d][1])):
                idv += 1
                for col in range(len(Tor[d][1][subject])):
                    item = tableWidget.item(idv, col + 1)
                    if item is not None and item.text():
                        if col == 1:
                            Tor[d][1][subject][col] = int(item.text())
                        else:
                            Tor[d][1][subject][col] = item.text()
                    else:
                        Tor[d][1][subject][col] = ""



class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1050, 600)

        self.MainWindow = MainWindow

        self.MyTable = MyTableWidget(self.MainWindow)
        self.MyTable.hide()


        # CENTRAL WIDGET
        self.centralwidget = QWidget(self.MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.MainWindow.setCentralWidget(self.centralwidget)

        # TABLE WIDGET
        self.tableWidget = QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QRect(0, 1, 771, 551))
        self.tableWidget.setRowCount(30)
        self.tableWidget.setColumnCount(30)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        # BOTONES
        self.pushButton = QPushButton(self.centralwidget)
        self.pushButton.setObjectName("pushButton")
        self.pushButton.setMaximumHeight(31)
        self.pushButton.setStyleSheet("background-color:rgb(220,128,128)")


        # MENÚ BAR
        self.menubar = QMenuBar(self.MainWindow)
        self.menubar.setGeometry(QRect(0, 0, 435, 26))
        self.menubar.setObjectName("menubar")
        self.MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(self.MainWindow)
        self.statusbar.setObjectName("statusbar")
        self.MainWindow.setStatusBar(self.statusbar)

        self.helpButton = QAction(self.MainWindow)
        self.helpButton.setObjectName("Ayuda")

        # MENÚ ARCHIVO
        self.menuArchivo = QMenu(self.menubar)
        self.menuArchivo.setObjectName("menuaArchivo")
        self.actionImportarAlumnos = QAction(self.MainWindow)
        self.actionImportarAlumnos.setObjectName("Importar alumnos")

        self.menuArchivo.addAction(self.actionImportarAlumnos)


        # MENÚ OPCIONES
        self.menuOpciones = QMenu(self.menubar)
        self.menuOpciones.setObjectName("menuOpciones")


        self.actionA_adir_ruta_LibreOffice = QAction(self.MainWindow)
        self.actionA_adir_ruta_LibreOffice.setObjectName("actionA_adir_ruta_LibreOffice")

        self.actionA_adir_ruta_Latex = QAction(self.MainWindow)
        self.actionA_adir_ruta_Latex.setObjectName("actionA_adir_ruta_Latex")

        self.actionAñadirTablaEquivalencias = QAction(self.MainWindow)

        self.menuOpciones.addAction(self.actionA_adir_ruta_LibreOffice)
        self.menuOpciones.addAction(self.actionA_adir_ruta_Latex)
        self.menuOpciones.addAction(self.actionAñadirTablaEquivalencias)



        # CONFIG FILES
        self.FolderAlumnos = False
        self.nameFolderAlumnos = ""
        self.config = {}
        self.fileConfig = 'config/config.ini'

        # PARSER
        self.parser = ConfigParser()
        self.initialize_config()

        # PROGRESS BAR
        self.progress_bar = QProgressBar(self.centralwidget)
        self.progress_bar.setGeometry(0, 0, 300, 25)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.hide()

        # BACK BUTTON
        self.back_button = BackButton(self.centralwidget)
        self.back_button.hide()


        # MODE MANUAL
        self.conversor_individual = QAction(self.MainWindow)
        self.conversor_individual.setObjectName("Conversor calificaciones manual")

        self.menubar.addAction(self.menuArchivo.menuAction())
        self.menubar.addAction(self.menuOpciones.menuAction())
        self.menubar.addAction(self.conversor_individual)
        self.menubar.addAction(self.helpButton)

        # ACTIONS
        self.actionA_adir_ruta_LibreOffice.triggered.connect(self.addOfficeRoute)
        self.actionA_adir_ruta_Latex.triggered.connect(self.addLatexRoute)
        self.actionImportarAlumnos.triggered.connect(self.getAlumnos)
        self.actionAñadirTablaEquivalencias.triggered.connect(self.addTable)
        self.pushButton.clicked.connect(self.generate)
        self.helpButton.triggered.connect(self.showHelp)
        self.conversor_individual.triggered.connect(self.controller)
        self.back_button.back_buttom.clicked.connect(self.back)

        self.retranslateUi(MainWindow)

        # LAYOUT
        self.h_layout = QVBoxLayout()
        self.h_layout.setAlignment(Qt.AlignCenter)
        self.h_layout.setSpacing(25)
        self.h_layout.addSpacing(0)
        self.h_layout.addWidget(self.back_button)


        self.h_layout.addWidget(self.MyTable)

        self.h_layout.addWidget(self.tableWidget)

        self.h_layout.addWidget(self.progress_bar)

        self.h_layout.addWidget(self.pushButton, Qt.AlignHCenter)

        self.centralwidget.setLayout(self.h_layout)


        self.centralwidget.setWindowTitle("ToR Conversor")





        QMetaObject.connectSlotsByName(self.MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "ToR Conversor"))
        self.pushButton.setText(_translate("MainWindow", "Generar"))
        self.menuOpciones.setTitle(_translate("MainWindow", "Opciones"))
        self.actionA_adir_ruta_LibreOffice.setText(_translate("MainWindow", "Añadir ruta Soffice"))
        self.actionA_adir_ruta_Latex.setText(_translate("MainWindow", "Añadir ruta generador de documentos"))
        self.actionImportarAlumnos.setText((_translate("MainWindow", "Importar alumnos")))
        self.helpButton.setText((_translate("MainWindow", "Ayuda")))
        self.conversor_individual.setText((_translate("MainWindow", "Conversor calificaciones manual")))
        self.actionAñadirTablaEquivalencias.setText("Añadir tabla de equivalencias")
        self.menuArchivo.setTitle("Archivo")

    def initialize_config(self):
        self.parser.read(self.fileConfig)
        self.config['soffice'] = self.parser.get('config','soffice')
        self.config['pdflatex'] = self.parser.get('config','pdflatex')
        self.config['conversion_table'] = self.parser.get('config','conversion_table')
        self.MyTable.config = self.config


    def back(self):
        self.MyTable.hide()
        self.tableWidget.show()
        self.tableWidget.clearContents()
        self.tableWidget.setColumnCount(30)
        self.h_layout.insertSpacing(2, 25)
        self.pushButton.show()
        self.back_button.hide()


    def controller(self):


        single_conversor = SingleConversor()
        single_conversor.config = self.config
        single_conversor.getCountries()
        single_conversor.convertSingle()
        single_conversor.exec()


    def showHelp(self):

        confirm = QMessageBox()
        confirm.setIcon(QMessageBox.Information)
        confirm.setWindowTitle("Ayuda")
        confirm.setIcon(QMessageBox.Information)
        """
        confirm.setText(
            "Para obtener las calificaciones finales se deben realizar los siguientes pasos:\n\n 1. Se selecciona la ruta de Soffice "
            "(Opciones -> Añadir ruta Soffice) y se selecciona el archivo .bin o .com (Windows)\n\n 2. (OPCIONAL) Se selecciona la ruta del "
            "generador de documentos (Opciones -> Añadir ruta generador de documentos) y se selecciona el archivo ejecutable .exe"
            "\n\n 3.  Se selecciona la carpeta que contiene los ficheros de los alumnos (Opciones -> importar Alumnos) \n\n"
            "4. (OPCIONAL) Se selecciona la tabla de equivalencias (Opciones -> Añadir tabla de equivalencias) y se selecciona un fichero .csv\n\n"
            "5. Se pulsa GENERAR, se pedirá la tabla de equivalencia de calificaciones y se ejecutará el software")"""
        confirm.exec()

    def addOfficeRoute(self):
        filename = QFileDialog()
        filename.setWindowTitle("Añadir ruta Soffice")

        name = filename.getOpenFileName(filter="Todos los archivos(*.*)")
        if name[0]:
            self.OfficeRoute = True
            self.parser.set('config','soffice',name[0])
            with open(self.fileConfig, "w") as f:
                self.parser.write(f)
            self.initialize_config()
            confirm = QMessageBox()
            confirm.setIcon(QMessageBox.Information)
            confirm.setWindowTitle("Añadir ruta Soffice")
            confirm.setText("Convetidor de formato cargado con éxito.")
            confirm.exec()
        else:
            pass

    def addLatexRoute(self):
        filename = QFileDialog()
        filename.setWindowTitle("Añadir ruta PdfLatex")
        name = filename.getOpenFileName(filter="Todos los archivos(*.*)")
        if name[0]:

            self.parser.set('config', 'pdflatex', name[0])
            with open(self.fileConfig, "w") as f:
                self.parser.write(f)
            self.initialize_config()
            confirm = QMessageBox()
            confirm.setIcon(QMessageBox.Information)
            confirm.setWindowTitle("Añadir ruta PdfLatex")
            confirm.setText("Generador de documentos cargado con éxito.")
            confirm.exec()
        else:
            pass

    def getAlumnos(self):
        fileName = QFileDialog()
        fileName.setWindowTitle("Seleccionar carpeta con datos de los alumnos")
        folder = fileName.getExistingDirectory()

        if folder:
            self.nameFolderAlumnos = folder
            self.FolderAlumnos = True
            confirm = QMessageBox()
            confirm.setIcon(QMessageBox.Information)
            confirm.setWindowTitle("Importar alumnos")
            confirm.setText("Alumnos importados con éxito.")
            confirm.exec()
        else:
            self.FolderAlumnos = False

    def addTable(self):
        filename = QFileDialog()
        filename.setWindowTitle("Añadir tabla de equivalencia de calificaciones")
        name = filename.getOpenFileName(filter="Archivo CSV (*.csv)")

        if name[0]:
            self.parser.set('config', 'conversion_table', name[0])
            with open(self.fileConfig, "w") as f:
                self.parser.write(f)
            self.initialize_config()
            confirm = QMessageBox()
            confirm.setIcon(QMessageBox.Information)
            confirm.setWindowTitle("Añadir tabla de equivalencias")
            confirm.setText("Tabla de equivalencias cargada con éxito.")
            confirm.exec()
        else:
            pass

    def generate(self):


            if self.FolderAlumnos == False:
                confirm = QMessageBox()
                confirm.setIcon(QMessageBox.Critical)
                confirm.setWindowTitle("Generar calificaciones finales")
                confirm.setText("ERROR : No se ha añadido la carpeta con los datos de los alumnos")
                confirm.exec()

            elif not self.config.get('soffice'):
                confirm = QMessageBox()
                confirm.setIcon(QMessageBox.Critical)
                confirm.setWindowTitle("Generar calificaciones finales")
                confirm.setText("ERROR : Se debe añadir la ruta de Soffice")
                confirm.exec()

            elif not self.config.get('pdflatex'):
                confirm = QMessageBox()
                confirm.setIcon(QMessageBox.Critical)
                confirm.setWindowTitle("Generar calificaciones finales")
                confirm.setText("ERROR : Se debe añadir la ruta del generador de documentos")
                confirm.exec()

            else:
                self.progress_bar.show()
                self.progress_bar.setValue(0)
                files = ls1(self.nameFolderAlumnos, False)
                self.MyTable.crear_tabs(files)

                folder_primary = "ALUMNOS"
                primary_path = os.path.join(os.path.abspath(os.getcwd()), folder_primary)
                try:
                    os.stat(primary_path)
                except:
                    os.mkdir(primary_path)

                for i, file in enumerate(files):
                    try:
                        os.stat(os.path.join(primary_path, file[:-4]))
                    except:
                        os.mkdir(os.path.join(primary_path, file[:-4]))

                    FNULL = open(os.devnull, 'w')
                    folder_path = os.path.join(primary_path, file[:-4])
                    cmd = '"%s" --headless --convert-to csv:"Text - txt - csv (StarCalc)":"59,34,76,1" --outdir "%s" "%s"' % (
                        self.config['soffice'], folder_path, os.path.join(self.nameFolderAlumnos, file))
                    call(cmd)


                    csv_file = os.path.join(folder_path, file[:-3] + "csv")
                    personalData, ToR = readToR(csv_file)

                    destination = personalData[UNIV_COLUMN].upper()

                    american, ToR = tor.parseToR(ToR)

                    # 2. Parse and prepare equivalences table.
                    raw_destination, raw_origin = readData(self.config['conversion_table'], HOME, destination)

                    x, aliasx, y, aliasy = tor.expandScores(raw_origin, raw_destination, american)

                    # 3. Expand the table to score suggestions for each destination subject
                    ToR = tor.extendToR(ToR, x, aliasx, y, aliasy, american)

                    # Generate debug informatio

                    self.MyTable.datos_alumno[file].extend([folder_path, personalData, ToR])
                    self.MyTable.show_info_check(personalData, ToR,file)

                    if i == len(files)-1:
                        self.progress_bar.setValue(100)
                    else:
                        self.progress_bar.setValue(self.progress_bar.value()+int(100/len(files)))

                self.progress_bar.hide()
                self.progress_bar.setValue(0)
                self.pushButton.hide()
                self.tableWidget.hide()
                self.back_button.show()
                self.h_layout.insertSpacing(2, -25)
                self.MyTable.show()




########################################################################################################################
########################################################################################################################


app = QApplication([])
ventana = Ui_MainWindow()
main_window = QMainWindow()
ventana.setupUi(main_window)
main_window.show()

sys.exit(app.exec())
