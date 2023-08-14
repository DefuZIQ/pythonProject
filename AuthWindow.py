from PyQt5.QtWidgets import QDialog

from PyQt5 import uic


class AuthWindow(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        uic.loadUi('auth.ui', self)


