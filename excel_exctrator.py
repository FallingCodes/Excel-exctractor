
#This program is free software: you can redistribute it and/or modify 
#it under the terms of the GNU General Public License as published by
#the Free Software Foundation, either version 3 of the License, or any later version.
#
#This program is distributed in the hope that it will be useful, but 
#WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY
#or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License 
#for more details.
#
#You should have received a copy of the GNU General Public License 
#along with this program. 
#If not, see http://www.gnu.org/licenses/.

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QPushButton, QMainWindow
import sys
import xlrd
import matplotlib.pyplot as plt




class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(316, 281)
        MainWindow.setAutoFillBackground(True)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(70, 140, 161, 41))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.openFileNameDialog)
        self.pushButton.setStyleSheet("QPushButton{ background-color: rgb(190, 48, 48); border-color: rgb(190, 48, 48); border: none; border-radius: 5px; color: rgb(255, 255, 255);} QPushButton:hover{ border: 1px solid white; }")

        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(150, 80, 113, 28))
        self.lineEdit.setObjectName("lineEdit")

        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(40, 80, 84, 28))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.showNomeFoglio)
        self.pushButton_2.setStyleSheet("QPushButton{ background-color: rgb(190, 48, 48); border-color: rgb(190, 48, 48); border: none; border-radius: 5px; color: rgb(255, 255, 255);} QPushButton:hover{ border: 1px solid white; }")


        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 50, 291, 20))
        self.label.setObjectName("label")

        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(80, 120, 141, 20))
        self.label_2.setObjectName("label_2")

        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(80, 10, 141, 31))
        self.label_3.setObjectName("label_3")

        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(20, 210, 281, 28))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.clicked.connect(self.elabora)
        self.pushButton_3.setStyleSheet("QPushButton{ background-color: rgb(190, 48, 48); border-color: rgb(190, 48, 48); border: none; border-radius: 5px; color: rgb(255, 255, 255);} QPushButton:hover{ border: 1px solid white; }")

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.show()

    def openFileNameDialog(self, fileName):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, "Apri File", "", "xlsx file (*.xlsx)", options=options)
        global nomeFile
        nomeFile = fileName

    def showNomeFoglio(self):
        text, result = QInputDialog.getText(self, "Nome Foglio", "Metti il nome del foglio")

        if result == True:
            self.lineEdit.setText(str(text))
            global nomeFoglio
            nomeFoglio = text

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton.setText(_translate("MainWindow", "Apri File"))
        self.pushButton_2.setText(_translate("MainWindow", "Scrivi"))
        self.label.setText(_translate("MainWindow", "Metti il nome del foglio excel da analizzare:"))
        self.label_2.setText(_translate("MainWindow", "Poi apri il file .xlsx :"))
        self.label_3.setText(_translate("MainWindow", "EXCEL EXCTRACTOR"))
        self.pushButton_3.setText(_translate("MainWindow", "Analizza e crea il grafico!"))

    def elabora(self):
        workbook = xlrd.open_workbook(nomeFile)
        worksheet = workbook.sheet_by_name(nomeFoglio)

        global table
        table = list()
        record = list()
        global total_cols
        total_rows = worksheet.nrows
        total_cols = worksheet.ncols

        for x in range(total_rows):
          for y in range(total_cols):
            record.append(worksheet.cell(x,y).value)
            table.append(record)
          record = []
          x += 1
        print(table)
        self.crea()

    def crea(self):
        n = total_cols
        labels = table[n-1]
        sizes = table[n]
        patches, texts = plt.pie(sizes, shadow=True, startangle=90)
        plt.legend(patches, labels, loc="best")
        plt.axis('equal')
        plt.tight_layout()
        plt.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    MainWindow = Window()
    sys.exit(app.exec_())
