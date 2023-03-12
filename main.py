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
        self.groupBox_2.setStyleSheet('background-color: rgb(212, 210, 239);')

        # Change check boxes style
        self.checkVzlet.setFont(QFont('Arial', 12))
        self.checkVKT.setFont(QFont('Arial', 12))
        self.checkMKTS.setFont(QFont('Arial', 12))
        self.checkSPT.setFont(QFont('Arial', 12))
        self.checkTV7.setFont(QFont('Arial', 12))
        self.checkTMK.setFont(QFont('Arial', 12))

        self.textBrowser.append('Выберите файлы для обработки' + '\n')

        # Connecting functions with bottoms
        self.pushButton_2.clicked.connect(self.openFileNamesDialog)
        self.pushButton.clicked.connect(self.parse_files)
        self.pushButton_3.clicked.connect(self.saveFileDialog)
        
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
        #self.textBrowser.clear()
        self.textBrowser.append('Выбранные файлы:\n')
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        files, _ = QFileDialog.getOpenFileNames(self,"Выберите необходимые файлы", "","Exel 2003 (*.xls);; Exel Files (*.xlsx);; Text files (*.txt)", options=options)
        if files:
            for f in files:
                self.__name_list.append(f)
                self.textBrowser.append(str(f) + '\n')


    def saveFileDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        self._saving_directory = QFileDialog.getExistingDirectory(self,"Выберите путь для сохранения", self.getCurrPath(), options=options)
        self.label.setText(str(self._saving_directory))
        self.label.setStyleSheet('color: rgb(28, 181, 74)')


    def getCurrPath(self):
        return QtCore.QDir.currentPath()


    def get_dates(self):
        return [self.dateEdit.dateTime().toString('dd-MM-yyyy'), self.dateEdit_2.dateTime().toString('dd-MM-yyyy')]
    

    def read_files(self):
        for file in self.__name_list:
            for key in self.__data_list.keys():
                if key in file.replace(' ', '') and file not in self.__data_list[key]:
                    self.__data_list[key].append(file)

        return self.__data_list


    def parse_files(self):
        self.textBrowser.clear()
        self.textBrowser.append('Обработано:\n')
        self.__data_list = self.read_files()
        dates = self.get_dates()
        if self.checkVzlet.isChecked():
            vzlet = VzletParser(self.__data_list, self.getCurrPath(), self._saving_directory)
            self.textBrowser.append(vzlet(dates[0], dates[1]))
        if self.checkVKT.isChecked():
            vkt = VKTParser(self.__data_list, self.getCurrPath(), self._saving_directory)
            self.textBrowser.append(vkt(dates[0], dates[1]))
        if self.checkMKTS.isChecked():
            mkts = MKTSParser(self.__data_list, self.getCurrPath(), self._saving_directory)
            self.textBrowser.append(mkts(dates[0], dates[1]))
        if self.checkSPT.isChecked():
            spt = SPTParser(self.__data_list, self.getCurrPath(), self._saving_directory)
            self.textBrowser.append(spt(dates[0], dates[1]))
        if self.checkTV7.isChecked():
            tv7 = TV7Parser(self.__data_list, self.getCurrPath(), self._saving_directory)
            self.textBrowser.append(tv7(dates[0], dates[1]))
        if self.checkTMK.isChecked():
            tmk = TMKParser(self.__data_list, self.getCurrPath(), self._saving_directory)
            self.textBrowser.append(tmk(dates[0], dates[1]))


def main():
    app = QApplication(sys.argv)
    w = MyApp()
    w.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()