from PyQt5.QtWidgets import QMainWindow
from PyQt5 import uic, QtGui


class AuthWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('auth.ui', self)
        self.setWindowIcon(QtGui.QIcon('static/images/credentials.png'))
