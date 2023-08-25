import requests
from PyQt5.QtWidgets import QMainWindow
from PyQt5 import uic
from AuthWindow import AuthWindow
import json
import os
from datetime import datetime, date
from tkinter import filedialog
from openpyxl import load_workbook
import pandas as pd
from dateutil.relativedelta import relativedelta
import csv


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('C:/Users/defuziq/PycharmProjects/pythonProject/static/main.ui', self)
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
        self.start_8.clicked.connect(self.start_app8)
        # Надо использовать QtWidget.setToolTip('text')

    def create_window(self):
        window = AuthWindow(self)
        window.show()

    def save_log(self, text):
        self.logs.setReadOnly(False)
        self.logs.appendPlainText(text)
        self.logs.setReadOnly(True)

    def path_file(self, label, label2=None, label3=None, filetype=0):
        self.logs.clear()
        self.save_log(text='Идёт чтение файла')
        filetypes = [(('Excel', '*.xlsx'), ('Excel', '*.xls'), ('Excel', '*.xlsm')),
                     (('txt', '*.txt'), ('csv', '*.csv'))]
        path = filedialog.askopenfilename(title='Выбрать файл', initialdir='', filetypes=filetypes[filetype])
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
            except:
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
        self.path_file(self.label_48, filetype=1)

    def validate_integer(self, value):
        try:
            value = int(value)
            return value
        except:
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
            requests.get('https://order-backoffice-apigateway.samokat.ru/swagger-ui/index.html?configUrl=%2Fv3%2Fapi-docs%2Fswagger-config&urls.primaryName=backoffice-public-api')
            return True
        except:
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
        symbols = [',', ';', ':', ' ', '\.', '\(', '\)']
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

    def start_app1(self):
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
            if self.checkBox.isChecked() is True:
                file_name2 = "postman(promocodes) " + str_current_datetime + ".csv"
                with open(file_name2, "w", encoding='utf-8', newline="") as f:
                    writer = csv.writer(f, quoting=csv.QUOTE_NONNUMERIC)
                    writer.writerow(["promocodes"])
                    for i in result:
                        i = [str(n) for n in i]
                        writer.writerow(i)
                    self.save_log('Готово, создан файл: ' + file_name2)

    def start_app2(self):
        valid = self.validate_input(label=self.label_20, line=self.lineEdit_5, line2=self.lineEdit_6,
                                    line3=self.lineEdit_7)
        if valid is True:
            df = self.validate_df(self.get_dataframe(label=self.label_20, line=self.lineEdit_5, line2=self.lineEdit_6))
            duplicates = self.find_duplicates(df)
            df = self.drop_duplicates(df)
            result = self.df_slice(df, self.lineEdit_7, int)
            str_current_datetime = str(datetime.now()).replace(':', '-')
            file_name = "promotion(users) " + str_current_datetime + ".json"
            file_name2 = "postman(users) " + str_current_datetime + ".csv"
            with open(file_name, 'w', encoding='utf-8') as file:
                for i in result:
                    file.write(
                        f'{{"promotionId": "{self.show_input(self.lineEdit_8)}", "userIds": {i}, "userType": "SAMOKAT", "disableNotifications": true}}\n')
                file.write(f'Дубликаты {duplicates}\n')
            with open(file_name, 'r', encoding='utf-8') as f:
                old_data = f.read()
            new_data = old_data.replace('\\xa0', '')
            with open(file_name, 'w', encoding='utf-8') as f:
                f.write(new_data)
                self.save_log('Готово, создан файл: ' + file_name)
            if self.checkBox.isChecked() is True:
                with open(file_name2, "w", encoding='utf-8', newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["users"])
                    for i in result:
                        i = ','.join([str(n) for n in i])
                        writer.writerow([i])
                    self.save_log('Готово, создан файл: ' + file_name2)


    def start_app3(self):
        valid = self.validate_input(label=self.label_31, line=self.lineEdit_9, line2=self.lineEdit_10,
                                    line3=self.lineEdit_11)
        if valid is True:
            df = self.validate_df(self.get_dataframe(label=self.label_31, line=self.lineEdit_9, line2=self.lineEdit_10))
            duplicates = self.find_duplicates(df)
            df = self.drop_duplicates(df)
            result = self.df_slice(df, self.lineEdit_11, int)
            str_current_datetime = str(datetime.now()).replace(':', '-')
            file_name = "banner(users) " + str_current_datetime + ".json"
            file_name2 = "postman(users) " + str_current_datetime + ".csv"
            with open(file_name, 'w', encoding='utf-8') as file:
                for i in result:
                    file.write(
                        file.write(f'{{"userIds": {i}, "userType": "SAMOKAT", "bannerId": "{self.show_input(self.lineEdit_13)}"}}\n'))
                file.write(f'Дубликаты {duplicates}\n')
            with open(file_name, 'r', encoding='utf-8') as f:
                old_data = f.read()
            new_data = old_data.replace('\\xa0', '')
            with open(file_name, 'w', encoding='utf-8') as f:
                f.write(new_data)
                self.save_log('Готово, создан файл: ' + file_name)
            if self.checkBox_3.isChecked() is True:
                with open(file_name2, "w", encoding='utf-8', newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["users"])
                    for i in result:
                        i = ','.join([str(n) for n in i])
                        writer.writerow([i])
                    self.save_log('Готово, создан файл: ' + file_name2)

    def start_app4(self):
        valid = self.validate_input(label=self.label_40, line=self.lineEdit_13, line2=self.lineEdit_14,
                                    line3=self.lineEdit_15)
        if valid is True:
            df = self.validate_df(self.get_dataframe(label=self.label_40, line=self.lineEdit_13, line2=self.lineEdit_14))
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

    def start_app5(self):
        jsons = open(self.show_input(self.label_44), "r", encoding="utf-8")
        jsons = jsons.read()
        jsons = jsons.split(', enriched requests document numbers = [')
        jsons = jsons[1]
        jsons = jsons.split(', document')
        jmax = len(jsons)
        for i in range(0, jmax):
            jsons[i] = jsons[i].split(', products=[')
        jsons.pop(0)
        jmax = len(jsons)
        for i in range(0, jmax):
            jsons[i][1] = jsons[i][1].split('ConfirmShipmentProduct(')
        for i in range(0, jmax):
            jsons[i][1] = [i for i in jsons[i][1] if 'packageId=null' in i]
            pmax = len(jsons[i][1])
            for n in range(0, pmax):
                jsons[i][1][n] = str(jsons[i][1][n]).split(', ')
            print(jsons[i])
        result = []
        for i in range(0, jmax):
            pmax = len(jsons[i][1])
            for n in range(0, pmax):
                for n in jsons[i][1][n]:
                    if 'УТ' in n:
                        result.append(jsons[i])
                        break
        jmax = len(result)
        for i in range(0, jmax):
            pmax = len(result[i][1])
            if pmax == 1:
                result[i][1][0].pop()
                result[i][1][0].pop()
                result[i][1][0].pop(3)
                result[i][1][0][5] = str(result[i][1][0][5]).replace(')]', '')
                break
            for n in range(0, pmax):
                result[i][1][n].pop()
                result[i][1][n].pop()
                result[i][1][n].pop(3)
                result[i][1][n][5] = str(result[i][1][n][5]).replace(')]', '')
        str_current_datetime = str(datetime.now()).replace(':', '-')
        file_name = "shipment " + str_current_datetime + ".sql"
        with open(file_name, 'w', encoding='utf-8') as file:
            file.write(f'{result}\n')
        with open(file_name, 'r', encoding='utf-8') as f:
            old_data = f.read()
        new_data = old_data.replace(',', '\n')
        new_data = new_data.replace("'", '')
        with open(file_name, 'w', encoding='utf-8') as f:
            f.write(new_data)
            print('Готово, создан файл: ' + file_name)

    def start_app6(self):
        path = self.show_input(self.label_46)
        jsons = open(path, "r", encoding="utf-8")
        jsons = json.loads(jsons.read())
        result = []
        data = date.today()
        if 'RECEIPTS' in jsons:
            confirm = 'RECEIPTS'
        elif 'SHIPMENTS' in jsons:
            confirm = 'SHIPMENTS'
        length = len(jsons[confirm])
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
        file_name = "shipment " + str_current_datetime + ".sql"
        with open(file_name, 'w', encoding='utf-8') as file:
            file.write(f'{result}\n')
        with open(file_name, 'r', encoding='utf-8') as f:
            old_data = f.read()
        new_data = old_data.replace(',', ', \n')
        with open(file_name, 'w', encoding='utf-8') as f:
            f.write(new_data)
            self.save_log('Готово, создан файл: ' + file_name)

    def start_app7(self):
        vpn = self.vpn_on()
        if vpn is True:
            pass

    def start_app8(self):
        vpn = self.vpn_on()
        if vpn is True:
            pass

