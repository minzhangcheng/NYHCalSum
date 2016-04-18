import sys
from PyQt5 import QtCore, QtWidgets, uic


qtCreatorFile = "selectFile.ui"
Ui_Widget, QtBaseClass = uic.loadUiType(qtCreatorFile)

class InputFile(QtWidgets.QWidget, Ui_Widget):
    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        Ui_Widget.__init__(self)
        self.setupUi(self)
        self.browse.clicked.connect(self.browseFile)
        self.next.clicked.connect(self.nextStep)

    def browseFile(self):
        filename = QtWidgets.QFileDialog.\
            getOpenFileName(None,'原始文件', '', 'Excel文件 (*.xls *.xlsx)')[0]
        self.file.setText(filename)

    def nextStep(self):
        if self.file.text():
            self.filename = self.file.text()
            # self.hide()



if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = InputFile()
    window.show()
    sys.exit(app.exec_())