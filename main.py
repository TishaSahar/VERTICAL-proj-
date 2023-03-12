from PyQt5 import uic, QtCore
from PyQt5.QtWidgets import QDialog, QApplication, QInputDialog, QLineEdit, QFileDialog, QDateTimeEdit, QFontDialog
from PyQt5.QtGui import *
import sys

from VzletParser import VzletParser
from VKTParser import VKTParser
from MKTSParser import MKTSParser
from SPTParser import SPTParser
from TV7Parser import TV7Parser
from TMKParser import TMKParser


Form, _ = uic.loadUiType("my.ui")

class MyApp(QDialog, Form):
    def __init__(self):
        super(MyApp, self).__init__()

        # Extracting xlsx data
        self.__name_list = []
        self.__data_list = {'ВЗЛЕТ': [], 'ВКТ': [], 'МКТС': [], 'СПТ': [], 'ТВ-7': [], 'ТМК': []}

        # Saving directory and local path
        self._saving_directory = ''

        # My widgets styles
        self.title = 'Обработка отчетов'
        self.left = 10
        self.top = 10
        self.width = 640
        self.height = 480

        self.setupUi(self)
        self.setAttribute(QtCore.Qt.WA_StyledBackground, True)
        self.setStyleSheet('background-color: rgb(210, 221, 239);')
        self.label.setStyleSheet('color: rgb(255, 68, 68)')
        self.label.setFont(QFont('Arial', 11))
        self.pushButton_2.setStyleSheet('color: rgb(95, 68, 68); background-color: rgb(212, 210, 239);')
        self.pushButton_2.setFont(QFont('Arial', 14))
        self.pushButton.setStyleSheet('color: rgb(95, 68, 68); background-color: rgb(212, 210, 239);')
        self.pushButton.setFont(QFont('Arial', 14))
        self.pushButton_3.setStyleSheet('color: rgb(95, 68, 68); background-color: rgb(212, 210, 239);')
        self.pushButton_3.setFont(QFont('Arial', 14))
        self.groupBox.setStyleSheet('background-color: rgb(212, 210, 239);')

        # Change check boxes style
        self.checkVzlet.setFont(QFont('Arial', 12))
        self.checkVKT.setFont(QFont('Arial', 12))
        self.checkMKTS.setFont(QFont('Arial', 12))
        self.checkSPT.setFont(QFont('Arial', 12))
        self.checkTV7.setFont(QFont('Arial', 12))
        self.checkTMK.setFont(QFont('Arial', 12))

        self.textBrowser_2.append('<b>Файлы для обработки не выбраны</b>' + '\n')

        # Connecting functions with bottoms
        self.pushButton_2.clicked.connect(self.openFileNamesDialog)
        self.pushButton.clicked.connect(self.parse_files)
        self.pushButton_3.clicked.connect(self.saveFileDialog)
        self.pushButton_4.clicked.connect(self.clearMyFileList)
        
    # Getters and setters for my touple of Exel data 
    def get_data(self):
        return self.__data_list


    def set_data(self, data):
        self.__data_list.append(data)
    data = property(get_data, set_data)

    # Set my UI
    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.openFileNameDialog()
        self.saveFileDialog()
        self.show()


    def openFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"Выберите необходимый файл", "","Exel 2003 (*.xls);; Exel Files (*.xlsx);; Text files (*.txt)", options=options)
        if fileName:
            print(type(fileName))
    

    def openFileNamesDialog(self):
        self.textBrowser_2.clear()
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        files, _ = QFileDialog.getOpenFileNames(self,"Выберите необходимые файлы", "","Exel 2003 (*.xls);; Exel Files (*.xlsx);; Text files (*.txt)", options=options)
        if files:
            for f in files:
                self.__name_list.append(f)
                self.textBrowser_2.append(str(f) + '\n')
        self.printCurrData()


    def saveFileDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        self._saving_directory = QFileDialog.getExistingDirectory(self,"Выберите путь для сохранения", self.getCurrPath(), options=options)
        self.label.setText(str(self._saving_directory))
        self.label.setStyleSheet('color: rgb(28, 181, 74)')


    def printCurrData(self):
        for key in self.__data_list.keys():
            if len(self.__data_list[key]) == 0:
                continue
            self.textBrowser_2.append('\t' + key + '\n')
            for file in self.__data_list[key]:
                self.textBrowser_2.append(file + '\n')


    def clearMyFileList(self):
        self.__name_list.clear()
        self.__data_list.clear()
        self.__data_list = {'ВЗЛЕТ': [], 'ВКТ': [], 'МКТС': [], 'СПТ': [], 'ТВ-7': [], 'ТМК': []}
        self.textBrowser_2.clear()
        self.printCurrData()


    def getCurrPath(self):
        return QtCore.QDir.currentPath()
    

    def read_files(self):
        for file in self.__name_list:
            for key in self.__data_list.keys():
                if key in file.replace(' ', '') and file not in self.__data_list[key]:
                    self.__data_list[key].append(file)

        return self.__data_list


    def parse_files(self):
        self.textBrowser.clear()
        self.textBrowser.append('<b>\tОбработано:</b>' + '\n')
        self.__data_list = self.read_files()

        self.textBrowser_2.clear()
        self.printCurrData()

        if self.checkVzlet.isChecked():
            vzlet = VzletParser(self.__data_list, self.getCurrPath(), self._saving_directory)
            self.textBrowser.append(vzlet())
        if self.checkVKT.isChecked():
            vkt = VKTParser(self.__data_list, self.getCurrPath(), self._saving_directory)
            self.textBrowser.append(vkt())
        if self.checkMKTS.isChecked():
            mkts = MKTSParser(self.__data_list, self.getCurrPath(), self._saving_directory)
            self.textBrowser.append(mkts())
        if self.checkSPT.isChecked():
            spt = SPTParser(self.__data_list, self.getCurrPath(), self._saving_directory)
            self.textBrowser.append(spt())
        if self.checkTV7.isChecked():
            tv7 = TV7Parser(self.__data_list, self.getCurrPath(), self._saving_directory)
            self.textBrowser.append(tv7())
        if self.checkTMK.isChecked():
            tmk = TMKParser(self.__data_list, self.getCurrPath(), self._saving_directory)
            self.textBrowser.append(tmk())


def main():
    app = QApplication(sys.argv)
    w = MyApp()
    w.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()