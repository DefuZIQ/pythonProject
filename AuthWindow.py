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
        except Exception:
            self.save_log('Включите vpn')
            return False

    def save_log(self, text):
        self.logs.setText(text)

    def authorization(self):
        vpn = self.vpn_on()
        if vpn is True:
            username = self.lineEdit.text()
            password = self.lineEdit_2.text()
            data = {"username": username, "password": password}
            response = requests.post('https://order-backoffice-apigateway.samokat.ru/'
                                     'oauth/tokenByPassword', data=data)
            token = response.json()
            if token == {'code': 'INVALID_CREDENTIALS', 'message': 'Invalid credentials'}:
                print(token)
                self.logs.setStyleSheet("color: red;")
                self.save_log('Вы ввели неправильный логин или пароль')
            elif token == {'code': 'INTERNAL_SERVER_ERROR', 'message': 'Read timed out executing POST https://idm-auth-employee.samokat.ru/oauth/tokenByPassword'}:
                print(token)
                self.logs.setStyleSheet("color: red;")
                self.save_log('Не успешно, попробуйте еще раз')
            else:
                print(token)
                self.logs.setStyleSheet("color: green;")
                self.save_log('Вы успешно авторизовались')
                with open('token.json', 'w', encoding="utf-8") as f:
                    f.write(json.dumps(token, indent=4, ensure_ascii=False))
                with open('cred.json', 'w', encoding="utf-8") as f:
                    f.write(json.dumps(data, indent=4, ensure_ascii=False))

