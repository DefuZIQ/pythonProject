import requests
from PyQt5.QtWidgets import QMainWindow, QFileDialog
from PyQt5 import uic
from AuthWindow import AuthWindow
import json
import os
from datetime import datetime, date
from openpyxl import load_workbook
import pandas as pd
from dateutil.relativedelta import relativedelta
import csv
import sys
from decimal import Decimal
import psycopg2
from psycopg2 import Error


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        main_path = self.resource_path('static/main.ui')
        uic.loadUi(main_path, self)
        self.action_2.triggered.connect(self.create_window)
        self.pushButton.clicked.connect(self.upload_excel)
        self.pushButton_2.clicked.connect(self.upload_excel2)
        self.pushButton_3.clicked.connect(self.upload_excel3)
        self.pushButton_4.clicked.connect(self.upload_excel4)
        self.pushButton_5.clicked.connect(self.upload_excel5)
        self.pushButton_6.clicked.connect(self.upload_excel6)
        self.pushButton_7.clicked.connect(self.upload_excel7)

        self.start_1.clicked.connect(self.start_app1)
        self.start_2.clicked.connect(self.start_app2)
        self.start_3.clicked.connect(self.start_app3)
        self.start_4.clicked.connect(self.start_app4)
        self.start_5.clicked.connect(self.start_app5)
        self.start_6.clicked.connect(self.start_app6)
        self.start_7.clicked.connect(self.start_app7)
        self.start_8.clicked.connect(self.start_test)
        # Надо использовать QtWidget.setToolTip('text')

    def create_window(self):
        window = AuthWindow(self)
        window.show()

    @staticmethod
    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def save_log(self, text):
        self.logs.setReadOnly(False)
        self.logs.appendPlainText(text)
        self.logs.setReadOnly(True)

    def path_file(self, label, label2=None, label3=None, filetype=0):
        self.logs.clear()
        self.save_log('Идёт чтение файла')
        filetypes = [("Excel (*.xlsx *.xls *.xlsm)"), ("txt (*.txt)"), ("csv (*.csv)")]
        path = QFileDialog.getOpenFileName(self, 'Выбрать файл', '',filetypes[filetype])[0]
        file_name = os.path.basename(path)
        if path == "":
            self.logs.clear()
            self.save_log(text='Вы не выбрали файл')
            label.setText('')
            if label2 is None:
                pass
            else:
                label2.setText('')
                label3.setText('')
        else:
            label.setText(path)
            self.logs.clear()
            self.save_log(text='Вы выбрали файл: ' + file_name)
        return path

    def open_excel(self, label, label2, label3):
        path = self.path_file(label, label2, label3)
        if path == "":
            return
        else:
            try:
                wb = load_workbook(path)
                sheets = wb.sheetnames
                sheet_row = []
                for sheet in sheets:
                    df = pd.read_excel(io=path, sheet_name=sheet, header=None)
                    count_columns = len(df.axes[1])
                    array = []
                    for i in range(count_columns):
                        count_rows = df[df.columns[i]].count()
                        array.append({i: count_rows})
                    sheet_row.append({sheet: array})
                label2.setText(str(sheets))
                label3.setText(str(sheet_row))
            except Exception:
                self.logs.clear()
                label.setText('')
                label2.setText('')
                label3.setText('')
                self.save_log('Выберите корректный файл')

    def upload_excel(self):
        self.open_excel(label=self.label_9, label2=self.label_10, label3=self.label_11)

    def upload_excel2(self):
        self.open_excel(label=self.label_20, label2=self.label_21, label3=self.label_22)

    def upload_excel3(self):
        self.open_excel(label=self.label_31, label2=self.label_32, label3=self.label_33)

    def upload_excel4(self):
        self.open_excel(label=self.label_40, label2=self.label_41, label3=self.label_42)

    def upload_excel5(self):
        self.path_file(self.label_44, filetype=1)

    def upload_excel6(self):
        self.path_file(self.label_46, filetype=1)

    def upload_excel7(self):
        self.path_file(self.label_48)

    def validate_integer(self, value):
        try:
            value = int(value)
            return value
        except Exception:
            return 'Invalid'

    @staticmethod
    def show_input(label):
        result = label.text()
        return result

    def sheets_excel(self, label):
        path = self.show_input(label)
        wb = load_workbook(path)
        sheets = wb.sheetnames
        return sheets

    def validate_input_sheet(self, label, line):
        sheets = self.sheets_excel(label)
        input_sheet_null = self.show_input(line)
        input_sheet = self.validate_integer(self.show_input(line))
        if input_sheet_null == '':
            self.save_log('Вы не выбрали лист')
            return False
        elif input_sheet is False:
            self.save_log('Вы ввели некорректное значение')
            return False
        elif input_sheet >= len(sheets):
            self.save_log('Вы выбрали не существующий лист')
            return False
        else:
            return sheets[self.validate_integer(input_sheet)]

    def validate_input_columns(self, label, line, line2):
        columns = self.show_input(line2)
        if columns == '':
            return columns
        else:
            columns = [self.validate_integer(column) for column in columns.split(',')]
            if 'Invalid' in columns:
                self.save_log('Вы ввели некорректное значение в поле "Выбрать столбцы"')
                return False
            else:
                path = self.show_input(label)
                sheets = self.sheets_excel(label)
                sheet = sheets[int(self.show_input(line))]
                df = pd.read_excel(path, sheet_name=sheet, header=None)
                all_columns = [i for i in range(len(df.axes[1]))]
                bool_columns = tuple(x in all_columns for x in columns)
                if False in bool_columns:
                    self.save_log('Вы выбрали несуществующий столбец в поле "Выбрать столбцы"')
                    return False
                else:
                    return columns

    def validate_input_slice(self, line):
        if self.validate_integer(self.show_input(label=line)) == 'Invalid' or self.validate_integer(
                self.show_input(label=line)) == 0:
            self.save_log('Вы не выбрали по сколько разделить или ввели некорректное значение')
            return False
        else:
            return self.validate_integer(self.show_input(label=line))

    def validate_input(self, label, line, line2, line3):
        path = self.show_input(label=label)
        if path == '':
            self.save_log('Вы не выбрали файл')
        else:
            input_sheet = self.validate_input_sheet(label, line)
            input_columns = self.validate_input_columns(label, line, line2)
            input_slice = self.validate_input_slice(line3)
            if input_sheet is False:
                pass
            elif input_columns is False:
                pass
            elif input_slice is False:
                pass
            else:
                self.save_log('Вы выбрали лист: ' + str(input_sheet))
                if input_columns == '':
                    self.save_log('Вы выбрали все столбцы')
                else:
                    self.save_log('Вы выбрали столбцы: ' + str(', '.join(map(str, input_columns))))
                self.save_log('Вы выбрали разделить по: ' + str(input_slice))
                return True

    def vpn_on(self):
        try:
            requests.get('https://order-backoffice-apigateway.samokat.ru/swagger-ui/index.html?'
                         'configUrl=%2Fv3%2Fapi-docs%2Fswagger-config&urls.primaryName=backoffice-public-api')
            return True
        except Exception:
            self.save_log('Включите vpn')
            return False

    def get_dataframe(self, label, line, line2):
        sheets = self.sheets_excel(label)
        sheet = sheets[int(self.show_input(line))]
        path = self.show_input(label)
        dfs = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str)
        df = pd.DataFrame([])
        columns = self.show_input(line2)
        if columns == '':
            for i in range(len(dfs.axes[1])):
                df = pd.concat([df, dfs[i]], ignore_index=True)
            df = df.dropna(axis=0, how='any')
            df = df.reset_index(drop=True)
            return df
        else:
            columns = [int(column) for column in columns.split(',')]
            for i in range(len(dfs.axes[1])):
                for column in columns:
                    if i == column:
                        df = pd.concat([df, dfs[i]], ignore_index=True)
            df = df.dropna(axis=0, how='any')
            df = df.reset_index(drop=True)
            return df

    def validate_df(self, df):
        symbols = [',', ';', ':', ' ', '\\.', '\\(', '\\)']
        for symbol in symbols:
            df = df.replace(symbol, '', regex=True)
        self.save_log('Файл провалидирован')
        return df

    def find_duplicates(self, df):
        duplicates = df[df.duplicated()]
        duplicates = duplicates.drop_duplicates()
        duplicates = duplicates[0].values.astype(str).tolist()
        return duplicates

    def drop_duplicates(self, df):
        df = df.drop_duplicates()
        self.save_log('Дубликаты удалены')
        return df

    def df_slice(self, df, line, vtype):
        result = []
        _max = len(df.axes[0])
        n = int(self.show_input(line))
        for start in range(0, _max, n):
            stop = start + n
            slice_object = slice(start, stop)
            result.append(df[slice_object][0].values.astype(vtype).tolist())
        return result

    def create_csv(self, checkbox, result, vtype):
        if checkbox.isChecked() is True:
            str_current_datetime = str(datetime.now()).replace(':', '-')
            if vtype == str:
                _name = "promocodes"
                delimiter = '","'
            elif vtype == int:
                _name = "users"
                delimiter = ','
            file_name = "postman(" + _name + ")" + str_current_datetime + ".csv"
            with open(file_name, "w", encoding='utf-8', newline="") as f:
                writer = csv.writer(f)
                writer.writerow([_name])
                for i in result:
                    i = delimiter.join([str(n) for n in i])
                    writer.writerow([i])
                self.save_log('Готово, создан файл: ' + file_name)

    def start_app1(self):
        try:
            valid = self.validate_input(label=self.label_9, line=self.lineEdit, line2=self.lineEdit_2,
                                        line3=self.lineEdit_3)
            if valid is True:
                df = self.validate_df(self.get_dataframe(label=self.label_9, line=self.lineEdit, line2=self.lineEdit_2))
                duplicates = self.find_duplicates(df)
                df = self.drop_duplicates(df)
                result = self.df_slice(df, self.lineEdit_3, str)
                str_current_datetime = str(datetime.now()).replace(':', '-')
                file_name = "promotion(promocodes) " + str_current_datetime + ".json"
                with open(file_name, 'w', encoding='utf-8') as file:
                    for i in result:
                        file.write(
                            f'{{"promotionId": "{self.lineEdit_4.text()}", "promocodes": {i}}}\n')
                    file.write(f'Дубликаты {duplicates}\n')
                with open(file_name, 'r', encoding='utf-8') as f:
                    old_data = f.read()
                new_data = old_data.replace("'", '"')
                new_data = new_data.replace('\\xa0', '')
                with open(file_name, 'w', encoding='utf-8') as f:
                    f.write(new_data)
                    self.save_log('Готово, создан файл: ' + file_name)
                self.create_csv(self.checkBox, result, str)
        except Exception:
            self.save_log('Что-то пошло не так')

    def start_app2(self):
        try:
            valid = self.validate_input(label=self.label_20, line=self.lineEdit_5, line2=self.lineEdit_6,
                                        line3=self.lineEdit_7)
            if valid is True:
                df = self.validate_df(self.get_dataframe(label=self.label_20, line=self.lineEdit_5, line2=self.lineEdit_6))
                duplicates = self.find_duplicates(df)
                df = self.drop_duplicates(df)
                result = self.df_slice(df, self.lineEdit_7, int)
                str_current_datetime = str(datetime.now()).replace(':', '-')
                file_name = "promotion(users) " + str_current_datetime + ".json"
                with open(file_name, 'w', encoding='utf-8') as file:
                    for i in result:
                        file.write(
                            f'{{"promotionId": "{self.show_input(self.lineEdit_8)}", "userIds": {i}, '
                            f'"userType": "SAMOKAT", "disableNotifications": true}}\n')
                    file.write(f'Дубликаты {duplicates}\n')
                with open(file_name, 'r', encoding='utf-8') as f:
                    old_data = f.read()
                new_data = old_data.replace('\\xa0', '')
                with open(file_name, 'w', encoding='utf-8') as f:
                    f.write(new_data)
                    self.save_log('Готово, создан файл: ' + file_name)
                self.create_csv(self.checkBox_2, result, int)
        except Exception:
            self.save_log('Что-то пошло не так')

    def start_app3(self):
        try:
            valid = self.validate_input(label=self.label_31, line=self.lineEdit_9, line2=self.lineEdit_10,
                                        line3=self.lineEdit_11)
            if valid is True:
                df = self.validate_df(self.get_dataframe(label=self.label_31, line=self.lineEdit_9,
                                                         line2=self.lineEdit_10))
                duplicates = self.find_duplicates(df)
                df = self.drop_duplicates(df)
                result = self.df_slice(df, self.lineEdit_11, int)
                str_current_datetime = str(datetime.now()).replace(':', '-')
                file_name = "banner(users) " + str_current_datetime + ".json"
                with open(file_name, 'w', encoding='utf-8') as file:
                    for i in result:
                        file.write(
                            f'{{"userIds": {i}, "userType": "SAMOKAT", "bannerId": '
                            f'"{self.show_input(self.lineEdit_12)}"}}\n')
                    file.write(f'Дубликаты {duplicates}\n')
                print(1)
                with open(file_name, 'r', encoding='utf-8') as f:
                    old_data = f.read()
                new_data = old_data.replace('\\xa0', '')
                with open(file_name, 'w', encoding='utf-8') as f:
                    f.write(new_data)
                    self.save_log('Готово, создан файл: ' + file_name)
                self.create_csv(self.checkBox_3, result, int)
        except Exception:
            self.save_log('Что-то пошло не так')

    def start_app4(self):
        try:
            valid = self.validate_input(label=self.label_40, line=self.lineEdit_13, line2=self.lineEdit_14,
                                        line3=self.lineEdit_15)
            if valid is True:
                df = self.validate_df(self.get_dataframe(label=self.label_40, line=self.lineEdit_13,
                                                         line2=self.lineEdit_14))
                duplicates = self.find_duplicates(df)
                df = self.drop_duplicates(df)
                result = self.df_slice(df, self.lineEdit_15, int)
                str_current_datetime = str(datetime.now()).replace(':', '-')
                file_name = "newfile " + str_current_datetime + ".json"
                with open(file_name, 'w', encoding='utf-8') as file:
                    for i in result:
                        file.write(f'{i}\n')
                    file.write(f'Дубликаты {duplicates}\n')
                with open(file_name, 'r', encoding='utf-8') as f:
                    old_data = f.read()
                new_data = old_data.replace("'", '')
                new_data = new_data.replace('\\xa0', '')
                with open(file_name, 'w', encoding='utf-8') as f:
                    f.write(new_data)
                    self.save_log('Готово, создан файл: ' + file_name)
        except Exception:
            self.save_log('Что-то пошло не так')


    def start_app5(self):
        path = self.show_input(self.label_44)
        vpn = self.vpn_on()
        if vpn:
            if path == '':
                self.save_log('Вы не выбрали файл')
            else:
                try:
                    text = open(path, "r", encoding="utf-8")
                    text = text.read().split('\n')
                    result = []
                    for js in text:
                        jsons = json.loads(js)
                        for shipment in jsons['SHIPMENTS']:
                            items, temp, temp_arr = [], [], {}
                            names = [item['ITEM'] for item in shipment['DETAIL']]
                            res_slice, responses = [], {'data': []}
                            for start in range(0, len(names), 2):
                                stop = start + 200
                                slice_object = slice(start, stop)
                                res_slice.append(names[slice_object])
                            for i in res_slice:
                                search = {"filter": {"nomenclatureCodes": i}, "limit": len(i)}
                                response = requests.post('https://ds-metadata.samokat.ru/products/by-filter',
                                                         json=search)
                                response_json = response.json()
                                responses['data'].extend(response_json['data'])
                            for item, name in zip(shipment['DETAIL'], names):
                                if names.count(name) > 1:
                                    print(item, name)
                                    if name in [key for key, value in temp_arr.items()]:
                                        temp_arr[name].append(item)
                                    else:
                                        temp_arr.update({name: [item]})
                                else:
                                    print(item, name)
                                    valid_packages = []
                                    packages = [i['packages'] for i in responses['data'] if i['nomenclatureCode'] == name][0]
                                    for package in packages:
                                        if package['packageType'] == item['OP_QTY_UM']:
                                            quantity = Decimal(item['QUANTITY'].split('.')[0]).quantize(Decimal("1.00"))
                                            op_qty = Decimal(item['OP_QTY'].split('.')[0]).quantize(Decimal("1.00"))
                                            coefficient = package['coefficient']
                                            if coefficient * op_qty == quantity:
                                                valid_packages.append(True)
                                            else:
                                                valid_packages.append(False)

                                    if True not in valid_packages:
                                        pack = {
                                            'packages': [package['name'].replace('\xa0', '') for package in packages]}
                                        item.update(pack)
                                        items.append(item)

                            for e in temp_arr:
                                quantity = 0
                                op_qty = 0
                                valid_packages = []
                                packages = [i['packages'] for i in responses['data'] if i['nomenclatureCode'] == e][0]
                                for i in temp_arr[e]:
                                    quantity += Decimal(i['QUANTITY']).quantize(Decimal("1.00"))
                                    op_qty += Decimal(i['OP_QTY']).quantize(Decimal("1.00"))
                                    for package in packages:
                                        if package['packageType'] == i['OP_QTY_UM']:
                                            coefficient = package['coefficient']
                                            if coefficient * op_qty == quantity:
                                                valid_packages.append(True)
                                            else:
                                                valid_packages.append(False)
                                if True not in valid_packages:
                                    pack = {
                                        'packages': [package['name'].replace('\xa0', '') for package in
                                                     packages]}
                                    temp_arr[e].append(pack)
                                    items.append(temp_arr[e])
                            for x in items:
                                if x not in temp:
                                    temp.append(x)
                            items = temp
                            result.append({shipment['SHIPMENT_ID']: items})
                    for res in result:
                        res = res[list(res.keys())[0]]
                        print(res)
                        print(len(res))
                        res.insert(0,{'count': len(res)})
                        print(res)
                    print(result)
                    str_current_datetime = str(datetime.now()).replace(':', '-')
                    file_name = "shipment(YT) " + str_current_datetime + ".sql"
                    with open(file_name, 'w', encoding="utf-8") as f:
                        f.write(json.dumps(result, indent=4, ensure_ascii=False))
                        self.save_log('Готово, создан файл: ' + file_name)
                except Exception:
                    self.save_log('Некорректный JSON или что-то пошло не так')
    def start_app6(self):
        path = self.show_input(self.label_46)
        if path == '':
            self.save_log('Вы не выбрали файл')
        else:
            try:
                path = self.show_input(self.label_46)
                jsons = open(path, "r", encoding="utf-8")
                jsons = json.loads(jsons.read())
                result = []
                data = date.today()
                if 'RECEIPTS' in jsons:
                    confirm = 'RECEIPTS'
                elif 'SHIPMENTS' in jsons:
                    confirm = 'SHIPMENTS'
                length2 = len(jsons[confirm][0]['DETAIL'])
                for n in range(0, length2 - 1):
                    if jsons[confirm][0]['DETAIL'][n]['MAN_DATE'] == '0001-01-01':
                        result.append(jsons[confirm][0]['DETAIL'][n])
                    elif jsons[confirm][0]['DETAIL'][n]['MAN_DATE'] > str(data):
                        result.append(jsons[confirm][0]['DETAIL'][n])
                    elif jsons[confirm][0]['DETAIL'][n]['EXP_DATE'] < str(data):
                        result.append(jsons[confirm][0]['DETAIL'][n])
                    elif jsons[confirm][0]['DETAIL'][n]['MAN_DATE'] < str(data - relativedelta(years=10)):
                        result.append(jsons[confirm][0]['DETAIL'][n])
                    elif jsons[confirm][0]['DETAIL'][n]['EXP_DATE'] > str(data + relativedelta(years=15)):
                        result.append(jsons[confirm][0]['DETAIL'][n])
                str_current_datetime = str(datetime.now()).replace(':', '-')
                file_name = "shipment(date) " + str_current_datetime + ".sql"
                with open(file_name, 'w', encoding='utf-8') as file:
                    file.write(f'{result}\n')
                with open(file_name, 'r', encoding='utf-8') as f:
                    old_data = f.read()
                new_data = old_data.replace(',', ', \n')
                with open(file_name, 'w', encoding='utf-8') as f:
                    f.write(new_data)
                    self.save_log('Готово, создан файл: ' + file_name)
            except Exception:
                self.save_log('Некорректный JSON')


    def start_app7(self):
        path = self.show_input(self.label_48)
        if path == '':
            self.save_log('Вы не выбрали файл')
        else:
            vpn = self.vpn_on()
            if vpn is True:
                try:
                    df = pd.read_excel(path)
                    str_current_datetime = str(datetime.now()).replace(':', '-')
                    file_name = 'manual_shipments ' + str_current_datetime + '.json'
                    shipment_json = [{"shipmentId": "Вставить id перемещения",
                                      "documentNumber": "Вставить номер перемещения", "products": []}]
                    products = [product for product in df['productId']]
                    search_json = {"productIds": products}
                    response = requests.post('https://ds-metadata.samokat.ru/products/by-ids', json=search_json)
                    response_json = response.json()
                    count_YT = 0
                    for product, quantity, date_1, date_2 in zip(df['productId'], df['totalProductQuantity'],
                                                                 df['productionDate'], df['bestBeforeDate']):
                        date_1 = str(date_1).split(' ')[0] + "T00:00:00.00Z"
                        date_2 = str(date_2).split(' ')[0] + "T00:00:00.00Z"
                        try:
                            product_yt = [p['nomenclatureCode'] for p in response_json if p['productId'] == product][0]
                        except Exception:
                            count_YT += 1
                            product_yt = 'Not Found'
                        product_json = {
                            "productId": product,
                            "productCode": product_yt,
                            "totalProductQuantity": quantity,
                            "packages": [
                                {
                                    "productQuantity": quantity,
                                    "productsPerPackageCoefficient": 1
                                }
                            ],
                            "productionDate": date_1,
                            "bestBeforeDate": date_2,
                            "isDamaged": False
                        }
                        shipment_json[0]['products'].append(product_json)
                    with open(file_name, 'w', encoding="utf-8") as f:
                        f.write(f'Кол-во не найденных УТ: {count_YT}\n')
                        f.write(json.dumps(shipment_json, indent=4, ensure_ascii=False))
                    self.save_log('Готово, создан файл: ' + file_name)
                except Exception:
                    self.save_log('Некорректный excel файл')

    def start_app8(self):
        self.logs.clear()
        vpn = self.vpn_on()
        if vpn is True:
            try:
                with open('token.json') as f:
                    token = json.load(f)
                token = token['access_token']
                guids = self.plainTextEdit.toPlainText().split('\n')
                if guids == ['']:
                    self.save_log('Вы не ввели guid ЦФЗ')
                else:
                    self.save_log('Вы ввели ' + str(len(guids)) + ' guid ЦФЗ')
                    invalid_guids = []
                    result,datas = [], {}
                    count_showcases = 0
                    count_receipts = 0
                    for guid in guids:
                        if len(guid) == 36:
                            cfz_setting = []
                            showcases_search = {"storeId": guid}
                            receipts_search = {"providerId": 0, "warehouseId": guid}
                            header = {'Authorization': 'Bearer ' + token}
                            url_showcases = 'https://order-backoffice-apigateway.samokat.ru/showcases/getBy'
                            url_receipts = 'https://smk-supportpaymentgw.samokat.ru/receipt/cash-registers/find'
                            showcase = requests.post(url_showcases, headers=header, json=showcases_search)
                            receipt = requests.get(url_receipts, headers=header, params=receipts_search)
                            showcase = showcase.json()
                            if showcase == {"items": []}:
                                pass
                            else:
                                count_showcases += 1
                            receipt = receipt.json()
                            if receipt == {"error": "NOT_FOUND", "value": None}:
                                receipts_search = {"providerId": 2, "warehouseId": guid}
                                receipt = requests.get(url_receipts, headers=header, params=receipts_search)
                                receipt = receipt.json()
                                if receipt == {"error": "NOT_FOUND", "value": None}:
                                    pass
                                else:
                                    count_receipts += 1
                            else:
                                count_receipts += 1
                            cfz_setting.append(showcase)
                            cfz_setting.append(receipt)
                            result.append({guid: cfz_setting})

                            if self.checkBox_4.isChecked() is False:
                                pass
                            else:
                                cfz = requests.get(f'https://ds-warehouse.samokat.ru/warehouses/load/by-id/{guid}')
                                cfz = cfz.json()
                                cfz = cfz['value']
                                cityId = cfz['cityId']
                                city = requests.get(f'https://ds-warehouse.samokat.ru/city/{cityId}')
                                city = city.json()
                                city = city['value']
                                show, group_code, login, password = '', '', '', ''
                                if showcase == {"items": []}:
                                    pass
                                else:
                                    show = 'Экспресс'
                                if receipt == {"error": "NOT_FOUND", "value": None}:
                                    receipts_search = {"providerId": 2, "warehouseId": guid}
                                    receipt = requests.get(url_receipts, headers=header, params=receipts_search)
                                    receipt = receipt.json()
                                    if receipt == {"error": "NOT_FOUND", "value": None}:
                                        pass
                                    else:
                                        group_code = receipt['value']['cashRegisterGroup']
                                        login = receipt['value']['login']
                                        password = receipt['value']['password']
                                else:
                                    group_code = receipt['value']['cashRegisterGroup']
                                    login = receipt['value']['login']
                                    password = receipt['value']['password']

                                data = [cfz['name'], cfz['startedToOperate'].split('T')[0], cfz['warehouseId'],
                                        cfz['email'], city['code'], cfz['address']['lat'], cfz['address']['lon'],
                                        cfz['address']['fullAddress'], show, group_code, login, password]
                                datas[guid] = data
                        else:
                            invalid_guids.append(guid)
                    if invalid_guids == [] or invalid_guids == ['']:
                        pass
                    else:
                        self.save_log(
                            'В списке присутсвуют некорректные guid: ' + str(', '.join(map(str, invalid_guids))))
                    str_current_datetime = str(datetime.now()).replace(':', '-')
                    file_name = 'cfz_settings ' + str_current_datetime + '.json'
                    with open(file_name, 'w', encoding="utf-8") as f:
                        f.write(f'Кол-во витрин: {count_showcases}\n')
                        f.write(f'Кол-во касс: {count_receipts}\n')
                        f.write(json.dumps(result, indent=4, ensure_ascii=False))
                    self.save_log('Готово, создан файл: ' + file_name)
                    if self.checkBox_4.isChecked() is True:
                        df = pd.DataFrame(datas)
                        print(df)
                        file_name2 = 'cfz_settings ' + str_current_datetime + '.xlsx'
                        writer = pd.ExcelWriter(file_name2)
                        df.to_excel(writer, index=False, header=False)
                        writer.close()
                        self.save_log('Готово, создан файл: ' + file_name2)

            except Exception:
                self.save_log('Вы не авторизовались')
    def start_test(self):
        try:
            order_id = self.plainTextEdit.toPlainText().split('\n')
            if order_id == ['']:
                self.save_log('Вы не ввели order id')
            else:
                order_id.remove('')
                order_id = tuple(order_id)
                with open('cred.json') as f:
                    cred = json.load(f)
                login = cred['username']
                password = cred['password']
                connection = psycopg2.connect(user=login,
                                              password=password,
                                              host="patroni-17.samokat.io",
                                              port="5434",
                                              dbname="order_history")
                cursor = connection.cursor()
                cursor.execute(f'WITH orders AS (SELECT order_id, order_line_changed  FROM order_history WHERE order_id in {order_id}), date_created AS (SELECT change_date as created_date, order_id FROM order_status WHERE order_id in (SELECT order_id FROM orders) and order_status_id = 3), date_picking AS (SELECT change_date as picking_date, order_id FROM order_status WHERE order_id in (SELECT order_id FROM orders) and order_status_id = 4), date_picking_hub AS (SELECT change_date as picking_date_hub, order_id FROM order_status WHERE order_id in (SELECT order_id FROM orders) and order_status_id = 9), date_picked AS (SELECT change_date as picked_date,order_id FROM order_status WHERE order_id in (SELECT order_id FROM orders) and order_status_id = 5), date_picked_hub AS (SELECT change_date as picked_date_hub,order_id FROM order_status WHERE order_id in (SELECT order_id FROM orders) and order_status_id = 10), date_change AS (SELECT created_date_time as change_date, order_id FROM order_change WHERE order_id in (SELECT order_id FROM orders)), hub_picker AS (SELECT order_id, picker_id FROM distribution_center_picking_info WHERE order_id in (SELECT order_id FROM orders)), cfz_picker AS (SELECT order_id, picker_uuid FROM picking_info WHERE order_id in (SELECT order_id FROM orders)), cfz_deliveryman AS (SELECT order_id, deliveryman_uuid FROM delivery_info WHERE order_id in (SELECT order_id FROM orders)), order_number AS (SELECT order_id, display_number FROM order_history WHERE order_id in (SELECT order_id FROM orders)), product AS (SELECT oc.created_date_time as change_dates, olc.product_id FROM order_change oc JOIN order_line_change olc ON olc.order_change_id = oc.id WHERE oc.order_id in (SELECT order_id FROM orders)) SELECT order_number.display_number AS "Номер заказа", date_created.order_id, CASE WHEN change_date < picking_date_hub THEN true WHEN change_date < picking_date and picking_date_hub is NULL THEN true WHEN change_date is NULL and order_line_changed = true THEN true WHEN change_date is not NULL and order_line_changed = true and picking_date_hub is NULL and picking_date is NULL THEN true ELSE false END "Автокорректировка", CASE WHEN change_date > picking_date_hub and change_date < picked_date_hub THEN true ELSE false END "Сборка на Хабе", hub_picker.picker_id, CASE WHEN change_date > picking_date and change_date < picked_date and picking_date_hub is NULL THEN true ELSE false END "Сборка на ЦФЗ", cfz_picker.picker_uuid, CASE WHEN change_date > picked_date THEN true ELSE false END "Доставка", cfz_deliveryman.deliveryman_uuid, product.product_id AS "Продукт" FROM date_created LEFT JOIN date_picking ON date_created.order_id = date_picking.order_id LEFT JOIN date_picking_hub ON date_picking_hub.order_id = date_created.order_id LEFT JOIN date_picked ON date_created.order_id = date_picked.order_id LEFT JOIN date_picked_hub ON date_picked_hub.order_id = date_created.order_id LEFT JOIN date_change ON date_change.order_id = date_created.order_id LEFT JOIN orders ON orders.order_id = date_created.order_id LEFT JOIN hub_picker ON hub_picker.order_id = date_created.order_id LEFT JOIN cfz_picker ON cfz_picker.order_id = date_created.order_id LEFT JOIN cfz_deliveryman ON cfz_deliveryman.order_id = date_created.order_id LEFT JOIN product ON product.change_dates = date_change.change_date LEFT JOIN order_number ON order_number.order_id = date_created.order_id GROUP BY date_created.order_id, "Автокорректировка", "Сборка на Хабе","Сборка на ЦФЗ", "Доставка", hub_picker.picker_id, cfz_picker.picker_uuid,cfz_deliveryman.deliveryman_uuid, "Продукт", "Номер заказа" ORDER BY date_created.order_id ASC')
                result = []
                result.extend(cursor.fetchall())
                df = pd.DataFrame(result)
                type_update = []
                for a,b,c,d in zip(df[2].tolist(),df[3].tolist(),df[5].tolist(),df[7].tolist()):
                    if a is True:
                        type_update.append('Автокорректировка')
                    if b is True:
                        type_update.append('Ручная(Сборка на ХАБе)')
                    if c is True:
                        type_update.append('Ручная(Сборка на ЦФЗ)')
                    if d is True:
                        type_update.append('Ручная(На этапе доставки)')

                result = {'Номер заказа': df[0].tolist(), 'order_id': df[1].tolist(), 'Тип корректировки': type_update, 'product_id': df[9].tolist()}
                res = pd.DataFrame(result)
                who = []
                for type, uuid1, uuid2, uuid3 in zip(res['Тип корректировки'].tolist(), df[4].tolist(), df[6].tolist(), df[8].tolist()):
                    if type == 'Автокорректировка':
                        who.append('')
                    if type == 'Ручная(Сборка на ХАБе)':
                        who.append(uuid1)
                    if type == 'Ручная(Сборка на ЦФЗ)':
                        who.append(uuid2)
                    if type == 'Ручная(На этапе доставки)':
                        who.append(uuid3)
                result = {'Номер заказа': df[0].tolist(), 'order_id': df[1].tolist(), 'Тип корректировки': type_update,
                          'Кто скорректировал': who, 'product_id': df[9].tolist()}
                res = pd.DataFrame(result)
                connection1 = psycopg2.connect(user=login,
                                              password=password,
                                              host="patroni-06.samokat.io",
                                              port="5434",
                                              dbname="employee_profiles_backend")
                cursor1 = connection1.cursor()
                profile_id = tuple(list(filter(None, res['Кто скорректировал'].values)))
                employees = []
                if profile_id == ():
                    pass
                if len(profile_id) == 1:
                    profile_id = profile_id[0]
                    cursor1.execute(f"SELECT profile_id, full_name FROM profile WHERE profile_id = '{profile_id}'")
                    employees.extend(cursor1.fetchall())
                else:
                    cursor1.execute(f'SELECT profile_id, full_name FROM profile WHERE profile_id in {profile_id}')
                    employees.extend(cursor1.fetchall())
                who = res['Кто скорректировал'].to_list()
                print(who)
                who_update = []
                for id in who:
                    for employee in employees:
                        if id == employee[0]:
                            who_update.append(employee[1])
                        if id == '':
                            who_update.append('')
                products = [i for i in df[9].values]
                search_json = {"productIds": products}
                response = requests.post('https://ds-metadata.samokat.ru/products/by-ids', json=search_json)
                response_json = response.json()
                products_name = []
                for id in df[9].tolist():
                    for product in response_json:
                        if id == product['productId']:
                            products_name.append(product['administrativeName'])
                print(products_name)
                print(who_update)
                print(type_update)
                result = {'Номер заказа': df[0].tolist(), 'order_id': df[1].tolist(), 'Тип корректировки': type_update,
                          'Кто скорректировал': who_update, 'product_id': df[9].tolist(), 'Продукт': products_name}
                print(result)
                res = pd.DataFrame(result)
                str_current_datetime = str(datetime.now()).replace(':', '-')
                file_name = 'Отчет по корректировкам ' + str_current_datetime + '.xlsx'
                writer = pd.ExcelWriter(file_name)
                res.to_excel(writer, index=False)
                writer.close()
                self.save_log('Готово, создан файл: ' + file_name)
        except (Exception, Error) as error:
            print("Ошибка при работе с PostgreSQL", error)