from PyQt5.QtWidgets import QDialog
import requests
import json
import sys
import os
from PyQt5 import uic
from sqlitedict import SqliteDict

class AuthWindow(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        auth_path = self.resource_path('static/auth.ui')
        uic.loadUi(auth_path, self)
        self.pushButton.clicked.connect(self.authorization)
        login = Cache.load("login")
        password = Cache.load("password")
        self.lineEdit.setText(login)
        self.lineEdit_2.setText(password)

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
                self.logs.setStyleSheet("color: red;")
                self.save_log('Вы ввели неправильный логин или пароль')
            elif token == {'code': 'INTERNAL_SERVER_ERROR', 'message': 'Read timed out executing POST https://idm-auth-employee.samokat.ru/oauth/tokenByPassword'}:
                self.logs.setStyleSheet("color: red;")
                self.save_log('Не успешно, попробуйте еще раз')
            else:
                if Cache.load("login") is None:
                    Cache.save("login", username)
                    Cache.save("password", password)
                self.logs.setStyleSheet("color: green;")
                self.save_log('Вы успешно авторизовались')
                token = token['access_token']
                Cache.save("token", token)


class Cache():
    def __init__(self, param):
        self.param = param

    def save(key, value, cache_file="cache.sqlite3"):
        try:
            with SqliteDict(cache_file) as mydict:
                mydict[key] = value  # Using dict[key] to store
                mydict.commit()  # Need to commit() to actually flush the data
        except Exception as ex:
            print("Error during storing data (Possibly unsupported):", ex)

    def load(key, cache_file="cache.sqlite3"):
        try:
            with SqliteDict(cache_file) as mydict:
                value = mydict[key]  # No need to use commit(), since we are only loading data!
            return value
        except Exception as ex:
            return None