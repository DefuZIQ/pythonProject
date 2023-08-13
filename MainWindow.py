from PyQt5.QtWidgets import QMainWindow
from PyQt5 import uic, QtGui
from AuthWindow import AuthWindow


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.window = None
        uic.loadUi('main.ui', self)
        self.setWindowIcon(QtGui.QIcon('static/images/icon.png'))
        self.action_2.triggered.connect(self.create_window)

    def create_window(self):
        self.window = AuthWindow()
        self.window.show()



