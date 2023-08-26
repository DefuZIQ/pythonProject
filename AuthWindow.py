from PyQt5.QtWidgets import QDialog
import requests
import json
import sys
import os
from PyQt5 import uic


class AuthWindow(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        auth_path = self.resource_path('static/auth.ui')
        uic.loadUi(auth_path, self)
        self.pushButton.clicked.connect(self.authorization)

    @staticmethod
    def resource_path(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def vpn_on(self):
        try:
            requests.get('https://order-backoffice-apigateway.samokat.ru/swagger-ui/index.html?configUrl=%2Fv3%2Fapi-docs%2Fswagger-config&urls.primaryName=backoffice-public-api')
            return True
        except:
            self.save_log('Включите vpn')
            return False

    def authorization(self):
        vpn = self.vpn_on()
        if vpn is True:
            username = self.lineEdit.text()
            password = self.lineEdit_2.text()
            data = {"username": username , "password": password}
            response = requests.post('https://order-backoffice-apigateway.samokat.ru/oauth/tokenByPassword', data=data)
            token = response.json()
            with open('token.json', 'w', encoding="utf-8") as f:
                f.write(json.dumps(token, indent=4, ensure_ascii=False))



