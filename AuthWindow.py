from PyQt5.QtWidgets import QDialog
import requests
import json
from PyQt5 import uic


class AuthWindow(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        uic.loadUi('C:/Users/defuziq/PycharmProjects/pythonProject/static/auth.ui', self)
        self.pushButton.clicked.connect(self.authorization)

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



