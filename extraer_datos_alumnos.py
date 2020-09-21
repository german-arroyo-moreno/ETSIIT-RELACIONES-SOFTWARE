#!/usr/bin/env python3

"""
@autor : MIGUEL ÁNGEL PÉREZ DÍAZ

    SUBDIRECCIÓN DE RELACIONES INTERNACIONALES ETSIIT (GRANADA)

"""

# IMPORTAMOS LOS MÓDULOS NECESARIOS PARA REALIZAR LA TAREA
import csv

from PyQt5 import QtCore, QtGui, QtWidgets

from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

from email.message import EmailMessage

import mimetypes
from Crypto.Cipher import _DES
import smtplib
from configparser import ConfigParser
import os
import sys
import uuid

FONT_SIZES = [7, 8, 9, 10, 11, 12, 13, 14, 18, 24, 36, 48, 64, 72, 96, 144, 288]
IMAGE_EXTENSIONS = ['.jpg', '.png', '.bmp']
HTML_EXTENSIONS = ['.htm', '.html']


def hexuuid():
    return uuid.uuid4().hex


def splitext(p):
    return os.path.splitext(p)[1].lower()


class TextEdit(QTextEdit):

    def canInsertFromMimeData(self, source):

        if source.hasImage():
            return True
        else:
            return super(TextEdit, self).canInsertFromMimeData(source)

    def insertFromMimeData(self, source):

        cursor = self.textCursor()
        document = self.document()

        if source.hasUrls():

            for u in source.urls():
                file_ext = splitext(str(u.toLocalFile()))
                if u.isLocalFile() and file_ext in IMAGE_EXTENSIONS:
                    image = QImage(u.toLocalFile())
                    document.addResource(QTextDocument.ImageResource, u, image)
                    cursor.insertImage(u.toLocalFile())

                else:
                    # If we hit a non-image or non-local URL break the loop and fall out
                    # to the super call & let Qt handle it
                    break

            else:
                # If all were valid images, finish here.
                return


        elif source.hasImage():
            image = source.imageData()
            uuid = hexuuid()
            document.addResource(QTextDocument.ImageResource, uuid, image)
            cursor.insertImage(uuid)
            return

        super(TextEdit, self).insertFromMimeData(source)


class InputForm(QDialog):
    def __init__(self, parent=None):
        super(QDialog, self).__init__(parent)

        self.setWindowTitle("Ajustes Email")
        self.resize(600, 300)

        layout_dialog = QVBoxLayout(parent)

        self.pushButton = QPushButton("Enviar")

        self.pushButton.setMaximumHeight(31)
        self.pushButton.setStyleSheet("background-color:rgb(220,128,128)")


        self.helpButton = QPushButton("Ayuda")
        self.helpButton.setMaximumHeight(31)
        self.helpButton.setMaximumWidth(75)
        self.helpButton.clicked.connect(self.showHelp)

        self.layout_user = QHBoxLayout()
        self.user = QLineEdit()
        user_info = QLabel("Usuario :")
        self.layout_user.addWidget(user_info, Qt.AlignLeft)
        self.layout_user.addWidget(self.user, Qt.AlignRight)

        self.config = {}

        self.layout_passwd = QHBoxLayout()
        self.passwd = QLineEdit()
        self.passwd.setEchoMode(QLineEdit.Password)
        passwd_info = QLabel("Contraseña :")
        self.layout_passwd.addWidget(passwd_info, Qt.AlignLeft)
        self.layout_passwd.addWidget(self.passwd, Qt.AlignRight)
        self.layout_passwd.setSpacing(10)

        self.layout_autenticacion = QHBoxLayout()
        self.autenticacion = QComboBox(self)
        self.autenticacion.addItems(['STARTTLS', 'SSL/TLS'])
        info = QLabel("Seguridad de la conexión")
        self.layout_autenticacion.addWidget(info, Qt.AlignLeft)
        self.layout_autenticacion.addWidget(self.autenticacion, Qt.AlignRight)
        self.layout_autenticacion.setSpacing(10)

        self.layout_puerto = QHBoxLayout()
        self.puerto = QSpinBox()
        self.puerto.setRange(0, 65535)
        info_puerto = QLabel("Puerto :")
        self.layout_puerto.addWidget(info_puerto, Qt.AlignLeft)
        self.layout_puerto.addWidget(self.puerto, Qt.AlignRight)
        self.layout_puerto.setSpacing(10)

        self.layout_server = QHBoxLayout()
        self.servidor = QLineEdit('smtp.ugr.es')
        info_server = QLabel("Nombre del servidor :")
        self.layout_server.addWidget(info_server, Qt.AlignLeft)
        self.layout_server.addWidget(self.servidor, Qt.AlignRight)
        self.layout_server.setSpacing(10)

        self.layout_buttons = QHBoxLayout()
        self.layout_buttons.setSpacing(250)
        self.layout_buttons.addWidget(self.helpButton, Qt.AlignLeft)
        self.layout_buttons.addWidget(self.pushButton, Qt.AlignRight)

        layout_dialog.addLayout(self.layout_user)
        layout_dialog.addLayout(self.layout_passwd)
        layout_dialog.addLayout(self.layout_autenticacion)
        layout_dialog.addLayout(self.layout_server)
        layout_dialog.addLayout(self.layout_puerto)
        layout_dialog.addLayout(self.layout_buttons)

        self.setLayout(layout_dialog)


    def showHelp(self):
        confirm = QMessageBox()
        confirm.setIcon(QMessageBox.Information)
        confirm.setWindowTitle("Ayuda")
        confirm.setIcon(QMessageBox.Information)

        confirm.setText(
            "Es necesario especificar los diferentes elementos de la conexión, para ello:\n"
            "1. Escribir el correo que se va a usar para mandar emails.\n"
            "2. Escribir la contraseña correspondiente a dicho correo.\n"
            "3. Se define el protocolo de seguridad : STARTTSL o SSL/TSL, en este caso se recomienda usar STARTTSL.\n"
            "4. Se especifica el servidor de correo asociado al correo a utilizar, en este caso si se utiliza una cuenta "
            "gmail el servidor sería smtp.gmail.com, cuenta hotmail : smtp.live.com, correo ugr : smtp.ugr.es.\n"
            "5. Se indica el puerto al que se realizará la conexión, en este caso están definidos el puerto 465 para SSL/TSL y 587 para STARTTSL.\n"
            "\nUna vez se pulsa el botón de enviar la configuración queda guardada.")

        confirm.exec()


    def readValues(self):
        self.user.setText(self.config['user'])
        self.passwd.setText(self.config['password'])
        self.puerto.setValue(int(self.config['port']))
        self.servidor.setText(self.config['server'])
        self.autenticacion.setCurrentText(self.config['security'])




class Editor(QMainWindow):

    def __init__(self, parent=None):
        super(QMainWindow, self).__init__(parent)

        layout = QVBoxLayout(parent)
        self.editor = TextEdit()
        # Setup the QTextEdit editor configuration
        self.editor.setAutoFormatting(QTextEdit.AutoAll)
        self.editor.selectionChanged.connect(self.update_format)

        # Initialize default font size.
        font = QFont('Times', 12)
        self.editor.setFont(font)
        # We need to repeat the size to init the current format.
        self.editor.setFontPointSize(12)

        self.sendEmail = QPushButton("Enviar correo")
        self.sendEmail.clicked.connect(self.send_email)

        layout_adjuntos = QHBoxLayout(parent)
        self.view_adjuntos = QPushButton("Ver archivos adjuntos")
        self.view_adjuntos.clicked.connect(self.view_attachments)
        layout_adjuntos.addWidget(self.sendEmail)
        layout_adjuntos.addWidget(self.view_adjuntos)

        self.ruta = None
        self.config = {}
        self.form = None
        self.fileConfig = 'config/config.ini'
        self.parser = ConfigParser()
        self.initialize_config()



        layout.addWidget(self.editor)
        layout.addLayout(layout_adjuntos)

        layout.setSpacing(25)
        self.sendEmail.setStyleSheet("background-color:rgb(228,195,195)")

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.archivos_adjuntos = []

        self.status = QStatusBar()
        self.setStatusBar(self.status)

        toolbar = QToolBar()
        toolbar.setIconSize(QSize(30, 30))
        self.addToolBar(toolbar)
        file_menu = self.menuBar().addMenu("Archivo")

        open_file_action = QAction(QIcon(os.path.join('images', 'blue-folder-open-document.png')), "Abrir fichero...",
                                   self)
        open_file_action.setStatusTip("Abrir Fichero")
        open_file_action.triggered.connect(self.openFile)
        file_menu.addAction(open_file_action)
        toolbar.addAction(open_file_action)

        save_file_action = QAction(QIcon(os.path.join('images', 'disk.png')), "Guardar", self)
        save_file_action.setStatusTip("Guardar")
        save_file_action.triggered.connect(self.saveFile)
        file_menu.addAction(save_file_action)
        toolbar.addAction(save_file_action)

        saveas_file_action = QAction(QIcon(os.path.join('images', 'disk--pencil.png')), "Guardar como...", self)
        saveas_file_action.setStatusTip("Guardar como...")
        saveas_file_action.triggered.connect(self.guardarComo)
        file_menu.addAction(saveas_file_action)
        toolbar.addAction(saveas_file_action)

        edit_menu = self.menuBar().addMenu("Editar")

        undo_action = QAction(QIcon(os.path.join('images', 'arrow-curve-180-left.png')), "Deshacer", self)
        undo_action.setStatusTip("Deshacer último cambio")
        undo_action.triggered.connect(self.editor.undo)
        toolbar.addAction(undo_action)
        edit_menu.addAction(undo_action)

        redo_action = QAction(QIcon(os.path.join('images', 'arrow-curve.png')), "Rehacer", self)
        redo_action.setStatusTip("Rehacer último cambio")
        redo_action.triggered.connect(self.editor.redo)
        toolbar.addAction(redo_action)
        edit_menu.addAction(redo_action)

        cut_action = QAction(QIcon(os.path.join('images', 'scissors.png')), "Cortar", self)
        cut_action.setStatusTip("Cortar")
        cut_action.setShortcut(QKeySequence.Cut)
        cut_action.triggered.connect(self.editor.cut)
        toolbar.addAction(cut_action)
        edit_menu.addAction(cut_action)

        copy_action = QAction(QIcon(os.path.join('images', 'document-copy.png')), "Copiar", self)
        copy_action.setStatusTip("Copiar")
        cut_action.setShortcut(QKeySequence.Copy)
        copy_action.triggered.connect(self.editor.copy)
        toolbar.addAction(copy_action)
        edit_menu.addAction(copy_action)

        paste_action = QAction(QIcon(os.path.join('images', 'clipboard-paste-document-text.png')), "Pegar", self)
        paste_action.setStatusTip("Pegar")
        cut_action.setShortcut(QKeySequence.Paste)
        paste_action.triggered.connect(self.editor.paste)
        toolbar.addAction(paste_action)
        edit_menu.addAction(paste_action)

        select_action = QAction(QIcon(os.path.join('images', 'selection-input.png')), "Seleccionar todo", self)
        select_action.setStatusTip("Seleccionar todo")
        cut_action.setShortcut(QKeySequence.SelectAll)
        select_action.triggered.connect(self.editor.selectAll)
        toolbar.addAction(select_action)
        edit_menu.addAction(select_action)

        attach_file = QAction(QIcon(os.path.join('images', 'attach-file.png')), "Adjuntar archivo", self)
        attach_file.setStatusTip("Adjuntar archivo")
        attach_file.triggered.connect(self.attach_file)
        toolbar.addAction(attach_file)

        format_menu = self.menuBar().addMenu("Formato")

        self.fonts = QFontComboBox()
        self.fonts.currentFontChanged.connect(self.editor.setCurrentFont)
        toolbar.addWidget(self.fonts)

        self.fontsize = QComboBox()
        self.fontsize.addItems([str(s) for s in FONT_SIZES])

        self.fontsize.currentIndexChanged[str].connect(lambda s: self.editor.setFontPointSize(float(s)))
        toolbar.addWidget(self.fontsize)

        fontColor = QAction(QIcon(os.path.join('images', 'font-color.png')), "Cambiar color de la fuente", self)
        fontColor.triggered.connect(self.change_color_font)
        fontColor.setStatusTip("Cambiar color de la fuente")
        toolbar.addAction(fontColor)
        format_menu.addAction(fontColor)

        self.bold_action = QAction(QIcon(os.path.join('images', 'edit-bold.png')), "Negrita", self)
        self.bold_action.setStatusTip("Negrita")
        self.bold_action.setShortcut(QKeySequence.Bold)
        self.bold_action.setCheckable(True)
        self.bold_action.toggled.connect(lambda x: self.editor.setFontWeight(QFont.Bold if x else QFont.Normal))
        toolbar.addAction(self.bold_action)
        format_menu.addAction(self.bold_action)

        self.italic_action = QAction(QIcon(os.path.join('images', 'edit-italic.png')), "Cursiva", self)
        self.italic_action.setStatusTip("Cursiva")
        self.italic_action.setShortcut(QKeySequence.Italic)
        self.italic_action.setCheckable(True)
        self.italic_action.toggled.connect(self.editor.setFontItalic)
        toolbar.addAction(self.italic_action)
        format_menu.addAction(self.italic_action)

        self.underline_action = QAction(QIcon(os.path.join('images', 'edit-underline.png')), "Subrayado", self)
        self.underline_action.setStatusTip("Subrayado")
        self.underline_action.setShortcut(QKeySequence.Underline)
        self.underline_action.setCheckable(True)
        self.underline_action.toggled.connect(self.editor.setFontUnderline)
        toolbar.addAction(self.underline_action)
        format_menu.addAction(self.underline_action)

        self.alignl_action = QAction(QIcon(os.path.join('images', 'edit-alignment.png')), "Alinear izquierda", self)
        self.alignl_action.setStatusTip("Alinear a la izquierda")
        self.alignl_action.setCheckable(True)
        self.alignl_action.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignLeft))
        toolbar.addAction(self.alignl_action)
        format_menu.addAction(self.alignl_action)

        self.alignc_action = QAction(QIcon(os.path.join('images', 'edit-alignment-center.png')), "Alinear centro", self)
        self.alignc_action.setStatusTip("Alinear en el centro")
        self.alignc_action.setCheckable(True)
        self.alignc_action.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignCenter))
        toolbar.addAction(self.alignc_action)
        format_menu.addAction(self.alignc_action)

        self.alignr_action = QAction(QIcon(os.path.join('images', 'edit-alignment-right.png')), "Alinear derecha", self)
        self.alignr_action.setStatusTip("Alinear a la derecha")
        self.alignr_action.setCheckable(True)
        self.alignr_action.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignRight))
        toolbar.addAction(self.alignr_action)
        format_menu.addAction(self.alignr_action)

        web_link = QAction(QIcon(os.path.join('images', 'web-link.png')), "Convertir en enlace", self)
        web_link.setStatusTip("Convertir en enlace URL")
        web_link.triggered.connect(self.convert_to_url)
        toolbar.addAction(web_link)

        self.datos_alumnos = None
        self.dicc_fields = {}

        format_group = QActionGroup(self)
        format_group.setExclusive(True)
        format_group.addAction(self.alignl_action)
        format_group.addAction(self.alignc_action)
        format_group.addAction(self.alignr_action)


        configuracion_email = self.menuBar().addAction("Ajustes Email")
        configuracion_email.triggered.connect(self.changeInfoEmail)

        help = self.menuBar().addAction("Ayuda")
        help.triggered.connect(self.showHelp)

        self._format_actions = [
            self.fonts,
            self.fontsize,
            self.bold_action,
            self.italic_action,
            self.underline_action,
        ]

        self.update_format()
        self.update_title()
        self.show()

    def block_signals(self, objects, b):
        for o in objects:
            o.blockSignals(b)

    def showHelp(self):
        confirm = QMessageBox()
        confirm.setIcon(QMessageBox.Information)

        confirm.setWindowTitle("Ayuda")
        confirm.setIcon(QMessageBox.Information)

        confirm.setText(
            "Una vez seleccionados los alumnos deseados, para enviarles un email es necesario :\n"
            "1. Escribir el texto pertinente que se desea enviar con el formato deseado que puede utilizar todas las herramientas disponibles.\n"
            "\t1.1. Para enviar emails personalizados se puede utilizar diferentes patrones que capte los datos del alumno correspondiente y se envie un email "
            "personaizado. Para ello una vez cargados los alumnos, es necesario fijarse en las columnas (atributos) : Nombre, apellidos, DNI, etc\n"
            "\t1.2. Una vez observados estos atributos, se pueden utilizar dentro del texto del mensaje de la siguiente manera: {{nombre}} o {{apellidos}}, "
            "así con el campo que desee. De esta forma si se seleccionan los alumnos Pedro y Lucía y se pone como texto : hola {{nombre}}, cada uno de ellos"
            " recibirá un mensaje personalizado en función de su nombre, así Pedro recibirá hola Pedro, y Lucía recibirá hola Lucía.\n"
            "\t1.3. Estos patrones se pueden utilizar tanto en el texto del mensaje como en el asunto.\n"
            "\n2. También se pueden adjuntar ficheros al email que se va a enviar y se pueden consultar los ficheros que se van a enviar en la opción 'Ver ficheros adjuntos'.\n"
            "\n3. Una vez pulsado el botón de enviar:\n"
            "\t3.1. Si es la primera vez que se utiliza pedirá que se introduzcan la configuración del correo a utilizar, una vez enviado este formulario se guardarán"
            "estos datos y se cargarán automáticamente cada vez que se utiliza. Estos ajustes se pueden cambiar en 'Ajustes Email'\n"
            "\n4. Se da la opción de enviar el mensaje al correo UGR o personal de los alumnos seleccionados\n"
            "\n5. Se pedirá el asunto del mensaje\n"
            "\n6. Se enviará un mensaje a cada alumno seleccionado y se mostrará un mensaje de confirmación tras cada envío.")

        confirm.exec()

    def dialog_critical(self, s):
        dlg = QMessageBox(self)
        dlg.setText(s)
        dlg.setIcon(QMessageBox.Critical)
        dlg.exec()

    def changeInfoEmail(self):
        cypher = _DES.new('etsiit-5')
        try:
            if os.path.isfile('credenciales.enc'):

                try:
                    archivo = open('credenciales.enc', "rb")

                    if  os.stat("credenciales.enc").st_size != 0:
                        info = ""
                        while True:
                            data = archivo.read(8)
                            n = len(data)
                            if n == 0:
                                break
                            decd = cypher.decrypt(data)
                            info += decd.decode('utf-8')

                        info.replace('\n', " ")
                        self.config['user'],  self.config['password'] = info.split()
                    else:
                        self.config['user'] = ""
                        self.config['password'] = ""




                    archivo.close()
                except BaseException as e:
                    self.dialog_critical(str(e))
                else:

                 self.form = InputForm()


                 self.form.config = self.config.copy()
                 self.form.readValues()
                 self.form.pushButton.clicked.connect(self.getValues)
                 self.form.exec()

            else:
                self.form = InputForm()
                self.form.pushButton.clicked.connect(self.getValues)

                self.form.exec()
        except BaseException as e:
            self.dialog_critical(str(e))


        try:
            archivo_encriptado = open("credenciales.enc", "wb")
            usuario = self.config['user']

            while True:
                if len(usuario) % 8 != 0:
                    usuario += '\n'
                else:
                    break

            enc = cypher.encrypt(usuario)
            archivo_encriptado.write(enc)

            passwd = self.config['password']
            while True:
                if len(passwd) % 8 != 0:
                    passwd += '\n'
                else:
                    break

            enc = cypher.encrypt(passwd)
            archivo_encriptado.write(enc)
            archivo_encriptado.close()
        except BaseException as e:
            self.dialog_critical(str(e))


    def initialize_config(self):
        self.parser.read(self.fileConfig)
        self.config['security'] = self.parser.get('config', 'seguridad')
        self.config['port'] = self.parser.get('config', 'puerto')
        self.config['server'] = self.parser.get('config', 'servidor')


    def update_format(self):
        """
        Update the font format toolbar/actions when a new text selection is made. This is neccessary to keep
        toolbars/etc. in sync with the current edit state.
        :return:
        """
        # Disable signals for all format widgets, so changing values here does not trigger further formatting.
        self.block_signals(self._format_actions, True)

        self.fonts.setCurrentFont(self.editor.currentFont())
        # Nasty, but we get the font-size as a float but want it was an int
        self.fontsize.setCurrentText(str(int(self.editor.fontPointSize())))

        self.italic_action.setChecked(self.editor.fontItalic())
        self.underline_action.setChecked(self.editor.fontUnderline())
        self.bold_action.setChecked(self.editor.fontWeight() == QFont.Bold)

        self.alignl_action.setChecked(self.editor.alignment() == Qt.AlignLeft)
        self.alignc_action.setChecked(self.editor.alignment() == Qt.AlignCenter)
        self.alignr_action.setChecked(self.editor.alignment() == Qt.AlignRight)

        self.block_signals(self._format_actions, False)

    def dialog_critical(self, s):
        dlg = QMessageBox(self)
        dlg.setText(s)
        dlg.setIcon(QMessageBox.Critical)
        dlg.exec()

    def openFile(self):
        path, _ = QFileDialog.getOpenFileName(self, "Abrir fichero", "",
                                              "Documentos HTML (*.html);;Archivos de texto (*.txt);;Todos los archivos(*.*)")

        try:
            with open(path, 'rU') as f:
                text = f.read()

        except Exception as e:
            self.dialog_critical(str(e))

        else:
            self.ruta = path
            # Qt will automatically try and guess the format as txt/html
            self.editor.setText(text)
            self.update_title()

    def saveFile(self):
        if self.ruta is None:
            # If we do not have a path, we need to use Save As.
            return self.guardarComo()

        text = self.editor.toHtml() if splitext(self.ruta) in HTML_EXTENSIONS else self.editor.toPlainText()

        try:
            with open(self.path, 'w') as f:
                f.write(text)

        except Exception as e:
            self.dialog_critical(str(e))

    def guardarComo(self):
        path, _ = QFileDialog.getSaveFileName(self, "Guardar fichero", "",
                                              "Documentos HTML (*.html);;Archivos de texto (*.txt);;Todos los archivos(*.*)")

        if not path:
            # If dialog is cancelled, will return ''
            return

        text = self.editor.toHtml() if splitext(path) in HTML_EXTENSIONS else self.editor.toPlainText()

        try:
            with open(path, 'w') as f:
                f.write(text)

        except Exception as e:
            self.dialog_critical(str(e))

        else:
            self.ruta = path

    def update_title(self):
        self.setWindowTitle("%s" % (os.path.basename(self.ruta) if self.ruta else "Untitled"))

    def change_color_font(self):
        color = QColorDialog.getColor()
        self.editor.setTextColor(color)

    def view_attachments(self):

        if self.archivos_adjuntos:
            confirm = QMessageBox()
            confirm.setIcon(QMessageBox.Information)
            confirm.setWindowTitle("Archivos adjuntos")
            confirm.setIcon(QMessageBox.Information)

            text = ""

            for file in self.archivos_adjuntos:
                text += '· ' + os.path.basename(os.path.normpath(file) + '\n')

            confirm.setText(text)
            confirm.exec()
        else:
            confirm = QMessageBox()
            confirm.setIcon(QMessageBox.Warning)
            confirm.setWindowTitle("Archivos adjuntos")
            confirm.setIcon(QMessageBox.Information)
            confirm.setText("No se han encontrado archivos adjuntos")
            confirm.exec()

    def convert_to_url(self):
        cursor = self.editor.textCursor()

        cursor.insertHtml("<a href={} </a>".format(cursor.selectedText()))

        format = cursor.charFormat()
        format.setForeground(Qt.blue)
        cursor.setCharFormat(format)

    def attach_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Abrir archivo", "", "All files (*.*)")

        if path:
            self.archivos_adjuntos.append(path)

    def checked_alumnos(self):
        for i in range(1, self.datos_alumnos.rowCount()):
            item = self.datos_alumnos.item(i, 0)
            if item is not None and item.checkState() == 2:
                return True

        return False

    def getValues(self):
        self.config = dict(zip(['user', 'password', 'security', 'port', 'server'],
                               [self.form.user.text(), self.form.passwd.text(), self.form.autenticacion.currentText(),
                                self.form.puerto.text(), self.form.servidor.text()]))

        try:
            self.parser.set('config', 'puerto', self.config['port'])
            self.parser.set('config', 'servidor', self.config['server'])
            self.parser.set('config', 'seguridad', self.config['security'])

            with open(self.fileConfig, "w") as f:
                self.parser.write(f)

            self.initialize_config()
        except BaseException as e:
            self.dialog_critical(str(e))


        self.form.close()




    def send_email(self):

        cypher = _DES.new('etsiit-5')

        if os.path.isfile('credenciales.enc'):
            pass
        else:
            self.form = InputForm()
            self.form.pushButton.clicked.connect(self.getValues)
            self.form.exec()

            try:
                archivo_encriptado = open("credenciales.enc", "wb")
                usuario = self.config['user']
                while True:
                    if len(usuario) % 8 != 0:
                        usuario += '\n'
                    else:
                        break

                enc = cypher.encrypt(usuario)
                archivo_encriptado.write(enc)

                passwd = self.config['password']
                while True:
                    if len(passwd) % 8 != 0:
                        passwd += '\n'
                    else:
                        break

                enc = cypher.encrypt(passwd)
                archivo_encriptado.write(enc)
                archivo_encriptado.close()
            except BaseException as e:
                self.dialog_critical(str(e))

        if self.checked_alumnos():

            buttonReply = QMessageBox()
            buttonReply.setWindowTitle("Enviar correo")
            buttonReply.setText("¿Desea enviar el email al correo de la UGR o al correo personal de los alumnos?")
            buttonReply.setIcon(QMessageBox.Question)
            buttonReply.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            buttonReply.setDefaultButton(QMessageBox.Yes)
            buttonYES = buttonReply.button(QMessageBox.Yes)
            buttonYES.setText("Correo UGR")
            buttonNO = buttonReply.button(QMessageBox.No)
            buttonNO.setText("Correo personal")

            buttonReply.setWindowFlags(QtCore.Qt.CustomizeWindowHint)

            buttonReply.exec()

            if buttonReply.clickedButton() == buttonYES:
                email = "Correo UGR:"

            else:
                email = "Correo personal:"

            try:
                archivo = open('credenciales.enc', "rb")

                info = ""
                while True:
                    data = archivo.read(8)
                    n = len(data)
                    if n == 0:
                        break
                    decd = cypher.decrypt(data)
                    info += decd.decode('utf-8')

                info.replace('\n', " ")
                usuario, passwd = info.split()

                archivo.close()


            except BaseException as e:
                self.dialog_critical(str(e))

            else:

                text, ok = QInputDialog.getText(self, 'Enviar correo', 'Introduce el asunto del mensaje:')
                if ok:

                    for i in range(1, self.datos_alumnos.rowCount()):
                        item = self.datos_alumnos.item(i, 0)
                        if item is not None and item.checkState() == 2:

                            correo = self.datos_alumnos.item(i, self.dicc_fields[email]).text()
                            body = self.editor.toHtml()
                            asunto = text

                            for d in self.dicc_fields.keys():
                                item = self.datos_alumnos.item(i, self.dicc_fields[d])
                                if item is not None and item.text():
                                    body = body.replace("{{" + str(d).upper()[:-1] + "}}", item.text())
                                    body = body.replace("{{" + str(d.lower()[:-1]) + "}}", item.text())
                                    body = body.replace("{{" + str(d)[:-1] + "}}", item.text())

                                    asunto = asunto.replace("{{" + str(d).upper()[:-1] + "}}", item.text())
                                    asunto = asunto.replace("{{" + str(d).lower()[:-1] + "}}", item.text())
                                    asunto = asunto.replace("{{" + str(d)[:-1] + "}}", item.text())


                            try:
                                msg = EmailMessage()

                                msg['From'] = usuario
                                msg['To'] = correo
                                msg['Subject'] = asunto

                                # attach image to message body
                                # msg.attach(MIMEImage(file("google.jpg").read()))
                                msg.add_header('Content-Type', 'text/html; charset=UTF-8')
                                msg.set_payload(body)
                            except BaseException as e:
                                self.dialog_critical(str(e))
                            else:

                                try:
                                    for file in self.archivos_adjuntos:

                                        ctype, encoding = mimetypes.guess_type(file)
                                        if ctype is None or encoding is not None:

                                            ctype = 'application/octet-stream'
                                        maintype, subtype = ctype.split('/', 1)
                                        with open(file, 'rb') as fp:
                                            msg.add_attachment(fp.read(),
                                                               maintype=maintype,
                                                               subtype=subtype,
                                                               filename=os.path.basename(os.path.normpath(file)))

                                except BaseException as e:
                                    self.dialog_critical(str(e))


                            try:
                                if(self.config['security'] == 'STARTTLS'):

                                    server = smtplib.SMTP(self.config['server'], self.config['port'])
                                else:

                                    server = smtplib.SMTP_SSL(self.config['server'], self.config['port'])

                            except BaseException as e:

                                self.dialog_critical(str(e))
                            else:

                                server.ehlo()
                                if (self.config['security'] == 'STARTTLS'):
                                    server.starttls()
                                    server.ehlo()

                                try:
                                    server.login(usuario, passwd)
                                except BaseException as e:
                                    self.dialog_critical(str(e))
                                else:

                                    try:
                                        server.sendmail(msg['From'], msg['To'], msg.as_string().encode('utf-8'), )
                                    except BaseException as e:
                                        self.dialog_critical(str(e))
                                    else:

                                        server.quit()

                                        confirm = QMessageBox()
                                        confirm.setWindowTitle("Enviar correo")
                                        confirm.setIcon(QMessageBox.Information)
                                        confirm.setText("Email enviado correctamente a : <b> %s </b>" % (msg['To']))
                                        confirm.exec()

                    self.archivos_adjuntos.clear()


        else:
            confirm = QMessageBox()
            confirm.setWindowTitle("Enviar correo")
            confirm.setIcon(QMessageBox.Warning)
            confirm.setText("No se han seleccionado alumnos para el envío de correo")
            confirm.exec()


########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################


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
            self.dialog_critical("\n ERROR : Opción inválida ( opciones disponibles : 0 y 1")
            exit(3)

    def dialog_critical(self, s):
        dlg = QMessageBox(self)
        dlg.setText(s)
        dlg.setIcon(QMessageBox.Critical)
        dlg.exec()

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

        self.marcarTodo = QtWidgets.QAction(MainWindow)
        self.desmarcarTodo = QtWidgets.QAction(MainWindow)
        self.invertirMarcado = QtWidgets.QAction(MainWindow)

        self.marcarTodo.setShortcut("Ctrl+M")
        self.desmarcarTodo.setShortcut("Ctrl+D")
        self.invertirMarcado.setShortcut("Ctrl+I")

        self.menuEditar.addAction(self.actionA_adir_filas)
        self.menuEditar.addAction(self.actionA_adir_columnas)
        self.menuEditar.addAction(self.marcarTodo)
        self.menuEditar.addAction(self.desmarcarTodo)
        self.menuEditar.addAction(self.invertirMarcado)

        self.menuBar.addAction(self.menuEditar.menuAction())



        self.menuOrdenar = QtWidgets.QMenu(self.menuBar)
        self.menuOrdenar.setObjectName("menuOrdenar")
        self.actionApellido = QtWidgets.QAction(MainWindow)
        self.actionApellido.setObjectName("actionApellido")
        self.actionUniversidad_Destino = QtWidgets.QAction(MainWindow)
        self.actionUniversidad_Destino.setObjectName("actionUniversidad_Destino")
        self.actionPa_s_destino = QtWidgets.QAction(MainWindow)
        self.actionPa_s_destino.setObjectName("actionPa_s_destino")
        self.menuOrdenar.addAction(self.actionApellido)
        self.menuOrdenar.addAction(self.actionUniversidad_Destino)
        self.menuOrdenar.addAction(self.actionPa_s_destino)
        self.menuBar.addAction(self.menuOrdenar.menuAction())

        self.menuEnviarCorreo = QtWidgets.QAction("Enviar correo")
        self.menuBar.addAction(self.menuEnviarCorreo)

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
        self.actionApellido.triggered.connect(self.orderBySurname)
        self.actionUniversidad_Destino.triggered.connect(self.orderByDestUniversity)
        self.actionPa_s_destino.triggered.connect(self.orderByDestCountry)
        self.marcarTodo.triggered.connect(self.markALL)
        self.desmarcarTodo.triggered.connect(self.unmarkALL)
        self.invertirMarcado.triggered.connect(self.invertALL)

        help = self.menuBar.addAction("Ayuda")
        help.triggered.connect(self.showHelp)

        self.menuEnviarCorreo.triggered.connect(self.CreateEmail)

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
        self.menuOrdenar.setTitle(_translate("MainWindow", "Ordenar"))
        self.actionApellido.setText(_translate("MainWindow", "Apellido"))
        self.actionUniversidad_Destino.setText(_translate("MainWindow", "Universidad destino"))
        self.actionPa_s_destino.setText(_translate("MainWindow", "País destino"))
        self.marcarTodo.setText("Marcar todo")
        self.desmarcarTodo.setText("Desmarcar todo")
        self.invertirMarcado.setText("Invertir lo marcado")

    ########################################################################################################################
    ########################################################################################################################

    def CreateEmail(self):
        try:
            if self.emptyTable():
                confirm = QMessageBox()
                confirm.setIcon(QMessageBox.Information)
                confirm.setWindowTitle("Enviar Correo")
                confirm.setIcon(QMessageBox.Critical)
                confirm.setText("No se han cargado datos de los alumnos")
                confirm.exec()
            else:
                try:
                    dialog = Editor()
                    dialog.setWindowModality(Qt.WindowModal)
                    dialog.datos_alumnos = self.tableWidget
                    dialog.dicc_fields = self.diccionario_keys
                except BaseException as e:
                    self.dialog_critical(str(e))
                else:
                    dialog.show()


        except BaseException as e:
            self.dialog_critical(str(e))

    def dialog_critical(self, s):
        dlg = QMessageBox(self)
        dlg.setText(s)
        dlg.setIcon(QMessageBox.Critical)
        dlg.exec()

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

    def showHelp(self):
        confirm = QMessageBox()
        confirm.setIcon(QMessageBox.Information)

        confirm.setWindowTitle("Ayuda")
        confirm.setIcon(QMessageBox.Information)

        confirm.setText(
            "<ul><li><b>Archivo :</b><ul>  \n"
            "<li><i>Nuevo :</i> se crea una nueva página y se borran los contenidos.</li>\n"
            "<li><i>Abrir :</i> se abre un fichero CSV con los datos de los alumnos. Normalmente se abre el fichero CSV generado tras cargar una serie de alumnos.</li>\n"
            "<li><i>Importar Alumnos :</i> se selecciona la carpeta con los todos los ficheros CSV de los alumnos y se cargan estos datos en un único fichero que se muestra por pantalla.</li>\n"
            "<li><i>Guardar CSV :</i> una vez cargados los datos de los alumnos, se permite guardar todo ello como un único CSV, que posteriormente se puede abrir con la opcíon 'Abrir'</li>\n"
            "<li><i>Salir :</i> termina la ejecución del programa.</li></ul></li>\n\n"
            "<li><b>Editar :</b> ofrece una serie de opciones de cara a la edición de la tabla de alumnos.</li>\n"
            "<li><b>Ordenar : </b> permite ordenar los alumnos dependiendo de la opción que se desee.</li>\n"
            "<li><b>Enviar Correo : </b> una vez se han cargado los alumnos, se permite el envío de emails a los alumnos seleccionados.</li>\n"
            "</ul>"
            )

        confirm.exec()

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
            inicio = prev_rows - 1


        else:
            self.tableWidget.setRowCount(self.Tabla.cont_alumnos + 1)
            self.tableWidget.setColumnCount(len(self.Tabla.fields_form.keys()) + 1)
            inicio = 0

        for i in range(self.Tabla.cont_alumnos):
            for col in range(len(self.Tabla.fields_form.keys())):
                key = list(self.Tabla.fields_form.keys())[col]

                if col == 0:
                    chkBoxItem = TableWidgetItem()
                    chkBoxItem.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                    chkBoxItem.setCheckState(QtCore.Qt.Unchecked)
                    self.tableWidget.setItem(inicio + i + 1, col, chkBoxItem)

                if i == 0:
                    item = TableWidgetItem(key)
                    item.setBackground(QtGui.QColor(255, 128, 128))
                    item.setFlags(Qt.ItemIsEnabled)

                    if not key in self.diccionario_keys:
                        if añadir:
                            self.tableWidget.setColumnCount(prev_cols + 1)
                            self.diccionario_keys[key] = prev_cols
                            self.tableWidget.setItem(i, prev_cols, item)
                            prev_cols += 1
                        else:
                            self.diccionario_keys[key] = col + 1
                            self.tableWidget.setItem(i, col + 1, item)

                data = self.Tabla.fields_form[key][i]
                self.tableWidget.setItem(inicio + i + 1, self.diccionario_keys[key], TableWidgetItem(data))

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
                    item = self.tableWidget.item(row, column + 1)
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
            if name[1] == 'Archivo CSV (*.csv)':
                self.Tabla.read_CSV(name[0])
                self.saveData(False)

                confirm = QtWidgets.QMessageBox()
                confirm.setWindowTitle("Importar de CSV")
                confirm.setText("Datos importados con éxito.   ")
                confirm.exec()
            else:
                pass

    ########################################################################################################################
    ########################################################################################################################

    def orderBySurname(self):

        if ("APELLIDOS:") in self.diccionario_keys.keys():
            self.tableWidget.sortItems(self.diccionario_keys["APELLIDOS:"], QtCore.Qt.AscendingOrder)
            self.tableWidget.show()

            confirm = QtWidgets.QMessageBox()
            confirm.setWindowTitle("Ordenar")
            confirm.setText("Datos ordenados por apellidos. ")
            confirm.exec()

    def orderByDestUniversity(self):

        if ("Universidad de destino:") in self.diccionario_keys.keys():
            self.tableWidget.sortItems(self.diccionario_keys["Universidad de destino:"], QtCore.Qt.AscendingOrder)
            self.tableWidget.show()

            confirm = QtWidgets.QMessageBox()
            confirm.setWindowTitle("Ordenar")
            confirm.setText("Datos ordenados por universidad de destino. ")
            confirm.exec()

    def orderByDestCountry(self):

        if ("País de destino:") in self.diccionario_keys.keys():
            self.tableWidget.sortItems(self.diccionario_keys["País de destino:"], QtCore.Qt.AscendingOrder)
            self.tableWidget.show()

            confirm = QtWidgets.QMessageBox()
            confirm.setWindowTitle("Ordenar")
            confirm.setText("Datos ordenados por país de destino. ")
            confirm.exec()

    def markALL(self):
        for fil in range(1, self.tableWidget.rowCount()):
            item = self.tableWidget.item(fil, 0)
            if item is not None:
                item.setCheckState(QtCore.Qt.Checked)

    def unmarkALL(self):
        for fil in range(1, self.tableWidget.rowCount()):
            item = self.tableWidget.item(fil, 0)
            if item is not None:
                item.setCheckState(QtCore.Qt.Unchecked)

    def invertALL(self):
        for fil in range(1, self.tableWidget.rowCount()):
            item = self.tableWidget.item(fil, 0)
            if item is not None:
                if item.checkState() == 0:
                    item.setCheckState(QtCore.Qt.Checked)
                elif item.checkState() == 2:
                    item.setCheckState(QtCore.Qt.Unchecked)


class TableWidgetItem(QtWidgets.QTableWidgetItem):

    def __lt__(self, other):
        if self.row() != 0 and other.row() != 0:
            return (self.text() <
                    other.text())
        else:
            return False


########################################################################################################################
########################################################################################################################


app = QtWidgets.QApplication([])  # Creamos app y le pasamos una lista de argumentos vacíos
ventana = Ui_MainWindow()
main_window = QtWidgets.QMainWindow()
ventana.setupUi(main_window)
main_window.show()

sys.exit(app.exec())
