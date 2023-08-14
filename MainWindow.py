from PyQt5.QtWidgets import QMainWindow
from PyQt5 import uic
from AuthWindow import AuthWindow
import json
from pathlib import Path
from datetime import datetime, date
from tkinter import filedialog
from openpyxl import load_workbook
import pandas as pd
from dateutil.relativedelta import relativedelta



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('main.ui', self)
        self.action_2.triggered.connect(self.create_window)
        self.pushButton.clicked.connect(self.upload_excel1)
        self.pushButton_2.clicked.connect(self.upload_excel2)
        self.pushButton_3.clicked.connect(self.upload_excel3)
        self.pushButton_4.clicked.connect(self.upload_excel4)
        self.pushButton_5.clicked.connect(self.upload_excel5)
        self.pushButton_6.clicked.connect(self.upload_excel6)
        self.pushButton_7.clicked.connect(self.upload_excel6)
        self.start_1.clicked.connect(self.start_app1)
        self.start_2.clicked.connect(self.start_app2)
        self.start_3.clicked.connect(self.start_app3)
        self.start_4.clicked.connect(self.start_app4)
        self.start_5.clicked.connect(self.start_app5)
        self.start_6.clicked.connect(self.start_app6)
        self.start_7.clicked.connect(self.start_app7)
        self.start_8.clicked.connect(self.start_app8)

    def create_window(self):
        self.exit_window()
        self.window = AuthWindow(self)
        self.window.show()

    def exit_window(self):
        AuthWindow(self).close()


    def pathfile(self, label):
        filetypes = (("Excel", "*.xlsx"), ("Excel", "*.xls"), ("Excel", "*.xlsm"), ("csv", "*.csv"), ("txt", "*.txt"))
        path = filedialog.askopenfilename(title="Выбрать файлы", initialdir="", filetypes=filetypes)
        label.setText(path)
        print(path)
        return path

    def upload_excel1(self):
        path = self.pathfile(self.label_10)
        if path == "":
            "Вы не выбрали файл"
        else:
            wb = load_workbook(path)
            sheets = wb.sheetnames
            sheet_row = []
            for sheet in sheets:
                sheet = wb[sheet]
                sheet_row.append(sheet.max_row)
                print(sheet.max_row)
            print(sheet_row)
            self.label_11.setText(str(sheets))
            self.label_12.setText(str(sheet_row))

    def upload_excel2(self):
        wb = load_workbook(self.pathfile(self.label_20))
        sheets = wb.sheetnames
        print(sheets)
        sheets = wb.sheetnames
        sheet_row = []
        for sheet in sheets:
            sheet = wb[sheet]
            sheet_row.append(sheet.max_row)
            print(sheet.max_row)
        print(sheet_row)
        self.label_21.setText(str(sheets))
        self.label_22.setText(str(sheet_row))

    def upload_excel3(self):
        wb = load_workbook(self.pathfile(self.label_23))
        sheets = wb.sheetnames
        print(sheets)
        sheets = wb.sheetnames
        sheet_row = []
        for sheet in sheets:
            sheet = wb[sheet]
            sheet_row.append(sheet.max_row)
            print(sheet.max_row)
        print(sheet_row)
        self.label_24.setText(str(sheets))
        self.label_25.setText(str(sheet_row))

    def upload_excel4(self):
        wb = load_workbook(self.pathfile(self.label_28))
        sheets = wb.sheetnames
        print(sheets)
        sheets = wb.sheetnames
        sheet_row = []
        for sheet in sheets:
            sheet = wb[sheet]
            sheet_row.append(sheet.max_row)
            print(sheet.max_row)
        print(sheet_row)
        self.label_27.setText(str(sheets))
        self.label_26.setText(str(sheet_row))

    def upload_excel5(self):
        self.pathfiles(self.label_6)

    def upload_excel6(self):
        self.pathfiles(self.label_29)

    def show_data1(self):
        print(self.lineEdit.text())
        return self.lineEdit.text()

    def show_data2(self):
        print(self.lineEdit_2.text())
        return self.lineEdit_2.text()

    def show_data3(self):
        print(self.lineEdit_3.text())
        return self.lineEdit_3.text()

    def show_data4(self):
        print(self.lineEdit_4.text())
        return self.lineEdit_4.text()

    def show_data5(self):
        print(self.lineEdit_5.text())
        return self.lineEdit_5.text()

    def show_data6(self):
        print(self.lineEdit_6.text())
        return self.lineEdit_6.text()

    def show_data7(self):
        print(self.lineEdit_7.text())
        return self.lineEdit_7.text()

    def show_data8(self):
        print(self.lineEdit_8.text())
        return self.lineEdit_8.text()

    def show_data9(self):
        print(self.lineEdit_9.text())
        return self.lineEdit_9.text()

    def show_data10(self):
        print(self.lineEdit_10.text())
        return self.lineEdit_10.text()

    def show_data11(self):
        print(self.lineEdit_11.text())
        return self.lineEdit_11.text()

    def show_data12(self):
        print(self.lineEdit_12.text())
        return self.lineEdit_12.text()

    def show_data13(self):
        print(self.lineEdit_13.text())
        return self.lineEdit_13.text()

    def start_app1(self):
        path = self.label_17.text()
        wb = load_workbook(path)
        sheets = wb.sheetnames
        for sheet in sheets:
            sheet = wb[sheet]
            print(sheet)
        x = int(self.show_data1())

        # Валидация excel файла
        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(',', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(';', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(':', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(' ', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[df.duplicated()]
        df = df.drop_duplicates()
        prdf = df[x].values.astype(str).tolist()
        print(prdf)

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df.drop_duplicates()
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        wb = load_workbook(self.label_17.text())
        sheets = wb.sheetnames
        for sheet in sheets:
            sheet = wb[sheet]

        df = pd.DataFrame(sheet.values).dropna(axis=0, how='any')[0].astype(str).tolist()

        result = []
        nmax = sheet.max_row
        n = int(self.show_data2())

        for start in range(0, nmax, n):
            stop = start + n
            slice_object = slice(start, stop)
            result.append(df[slice_object])
        # print(result)

        str_current_datetime = str(datetime.now()).replace(':', '-')
        file_name = "promocodes " + str_current_datetime + ".json"

        with open(file_name, 'w', encoding='utf-8') as file:
            for i in result:
                file.write(
                    f'{{"promotionId": "{self.lineEdit_3.text()}", "promocodes": {i}, "usageLimit": {self.lineEdit_4.text()} }}\n')
            file.write(f'Дубликаты {prdf}\n')

        with open(file_name, 'r') as f:
            old_data = f.read()
        # Валидация итогового файла

        new_data = old_data.replace("'", '"')
        new_data = new_data.replace(' ', '')
        new_data = new_data.replace('.', '')
        new_data = new_data.replace(';', '')
        new_data = new_data.replace('\\xa0', '')

        with open(file_name, 'w') as f:
            f.write(new_data)
            print('Готово, создан файл: ' + file_name)

    def start_app2(self):
        path = self.label_20.text()
        wb = load_workbook(path)
        sheets = wb.sheetnames
        for sheet in sheets:
            sheet = wb[sheet]
            print(sheet)
        x = int(self.show_data6())

        # Валидация excel файла
        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(',', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(';', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(':', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(' ', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[df.duplicated()]
        df = df.drop_duplicates()
        prdf = df[x].values.astype(str).tolist()
        print(prdf)

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df.drop_duplicates()
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        wb = load_workbook(path)
        sheets = wb.sheetnames
        for sheet in sheets:
            sheet = wb[sheet]

        df = pd.DataFrame(sheet.values).dropna(axis=0, how='any')[0].astype(str).tolist()

        result = []
        nmax = sheet.max_row
        n = int(self.show_data7())

        for start in range(0, nmax, n):
            stop = start + n
            slice_object = slice(start, stop)
            result.append(df[slice_object])
        # print(result)

        str_current_datetime = str(datetime.now()).replace(':', '-')
        file_name = "users " + str_current_datetime + ".json"

        with open(file_name, 'w', encoding='utf-8') as file:
            for i in result:
                file.write(
                    f'{{"promotionId": "{self.show_data8()}", "userIds": {i}, "userType": "SAMOKAT", "disableNotifications": true}}\n')
            file.write(f'Дубликаты {prdf}\n')

        with open(file_name, 'r') as f:
            old_data = f.read()
        # Валидация итогового файла

        new_data = old_data.replace("'", '')
        new_data = new_data.replace(' ', '')
        new_data = new_data.replace('.', '')
        new_data = new_data.replace(';', '')
        new_data = new_data.replace('\\xa0', '')

        with open(file_name, 'w') as f:
            f.write(new_data)
            print('Готово, создан файл: ' + file_name)

    def start_app3(self):
        path = self.label_23.text()
        wb = load_workbook(path)
        sheets = wb.sheetnames
        for sheet in sheets:
            sheet = wb[sheet]
            print(sheet)
        x = int(self.show_data9())

        # Валидация excel файла
        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(',', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(';', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(':', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(' ', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[df.duplicated()]
        df = df.drop_duplicates()
        prdf = df[x].values.astype(str).tolist()
        print(prdf)

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df.drop_duplicates()
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        wb = load_workbook(path)
        sheets = wb.sheetnames
        for sheet in sheets:
            sheet = wb[sheet]

        df = pd.DataFrame(sheet.values).dropna(axis=0, how='any')[0].astype(str).tolist()

        result = []
        nmax = sheet.max_row
        n = int(self.show_data10())

        for start in range(0, nmax, n):
            stop = start + n
            slice_object = slice(start, stop)
            result.append(df[slice_object])
        # print(result)

        str_current_datetime = str(datetime.now()).replace(':', '-')
        file_name = "banner(users) " + str_current_datetime + ".json"

        with open(file_name, 'w', encoding='utf-8') as file:
            for i in result:
                file.write(f'{{"userIds": {i}, "userType": "SAMOKAT", "bannerId": "{self.show_data11()}"}}\n')
            file.write(f'Дубликаты {prdf}\n')

        with open(file_name, 'r') as f:
            old_data = f.read()
        # Валидация итогового файла

        new_data = old_data.replace("'", '')
        new_data = new_data.replace(' ', '')
        new_data = new_data.replace('.', '')
        new_data = new_data.replace(';', '')
        new_data = new_data.replace('\\xa0', '')

        with open(file_name, 'w') as f:
            f.write(new_data)
            print('Готово, создан файл: ' + file_name)

    def start_app4(self):
        path = self.label_28.text()
        wb = load_workbook(path)
        sheets = wb.sheetnames
        for sheet in sheets:
            sheet = wb[sheet]
            print(sheet)
        x = int(self.show_data12())

        # Валидация excel файла
        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(',', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(';', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(':', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[0].replace(' ', '', regex=True)
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df[df.duplicated()]
        df = df.drop_duplicates()
        prdf = df[x].values.astype(str).tolist()
        print(prdf)

        df = pd.read_excel(path, sheet_name=sheets[x], header=None)
        df = df.drop_duplicates()
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, index=False, sheet_name=sheets[x], header=False)
        writer.close()

        wb = load_workbook(path)
        sheets = wb.sheetnames
        for sheet in sheets:
            sheet = wb[sheet]

        df = pd.DataFrame(sheet.values).dropna(axis=0, how='any')[0].astype(str).tolist()

        result = []
        nmax = sheet.max_row
        n = int(self.show_data13())

        for start in range(0, nmax, n):
            stop = start + n
            slice_object = slice(start, stop)
            result.append(df[slice_object])
        # print(result)

        str_current_datetime = str(datetime.now()).replace(':', '-')
        file_name = "newfile " + str_current_datetime + ".json"

        with open(file_name, 'w', encoding='utf-8') as file:
            for i in result:
                file.write(f'{i}\n')
            file.write(f'Дубликаты {prdf}\n')

        with open(file_name, 'r') as f:
            old_data = f.read()
        # Валидация итогового файла

        new_data = old_data.replace("'", '')
        new_data = new_data.replace(' ', '')
        new_data = new_data.replace('.', '')
        new_data = new_data.replace(';', '')
        new_data = new_data.replace('\\xa0', '')

        with open(file_name, 'w') as f:
            f.write(new_data)
            print('Готово, создан файл: ' + file_name)

    def start_app5(self):
        jsons = open(self.label_6.text(), "r", encoding="utf-8")
        jsons = jsons.read()

        jsons = jsons.split(', enriched requests document numbers = [')
        print('Этап первый')
        jsons = jsons[1]
        print('Этап второй')
        jsons = jsons.split(', document')
        jmax = len(jsons)
        for i in range(0, jmax):
            jsons[i] = jsons[i].split(', products=[')

        print(jsons)
        jsons.pop(0)

        print('Этап третий')

        jmax = len(jsons)
        for i in range(0, jmax):
            jsons[i][1] = jsons[i][1].split('ConfirmShipmentProduct(')
            # jsons[i][1].remove('')

        print(jsons)
        print('Этап четвертый')

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
            print('pmax: ' + str(pmax))
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

        print(result)
        print('Этап пятый')

        str_current_datetime = str(datetime.now()).replace(':', '-')
        file_name = "shipment " + str_current_datetime + ".sql"

        with open(file_name, 'w', encoding='utf-8') as file:
            file.write(f'{result}\n')

        with open(file_name, 'r') as f:
            old_data = f.read()
        # Валидация итогового файла

        new_data = old_data.replace(',', '\n')
        new_data = new_data.replace("'", '')

        with open(file_name, 'w') as f:
            f.write(new_data)
            print('Готово, создан файл: ' + file_name)

    def start_app6(self):
        jsons = open(self.label_29.text(), "r", encoding="utf-8")
        jsons = jsons.read()
        jsons = json.loads(jsons)

        length = len(jsons['RECEIPTS'])
        length2 = len(jsons['RECEIPTS'][0]['DETAIL'])

        print(length2)

        result = []

        data = date.today()
        for n in range(0, length2 - 1):
            if jsons['RECEIPTS'][0]['DETAIL'][n]['MAN_DATE'] == '0001-01-01':
                result.append(jsons['RECEIPTS'][0]['DETAIL'][n])
            else:
                print('Пусто 1')

            if jsons['RECEIPTS'][0]['DETAIL'][n]['MAN_DATE'] > str(data):
                result.append(jsons['RECEIPTS'][0]['DETAIL'][n])
            else:
                print(str(date.today()))
                print('Пусто 2')

            if jsons['RECEIPTS'][0]['DETAIL'][n]['EXP_DATE'] < str(data):
                result.append(jsons['RECEIPTS'][0]['DETAIL'][n])
            else:
                print('Пусто 3')

            if jsons['RECEIPTS'][0]['DETAIL'][n]['MAN_DATE'] < str(data - relativedelta(years=10)):
                result.append(jsons['RECEIPTS'][0]['DETAIL'][n])
            else:
                print('Пусто 4')

            if jsons['RECEIPTS'][0]['DETAIL'][n]['EXP_DATE'] > str(data + relativedelta(years=15)):
                result.append(jsons['RECEIPTS'][0]['DETAIL'][n])
            else:
                print('Пусто 5')

        str_current_datetime = str(datetime.now()).replace(':', '-')
        file_name = "shipment " + str_current_datetime + ".sql"

        with open(file_name, 'w', encoding='utf-8') as file:
            file.write(f'{result}\n')

        with open(file_name, 'r') as f:
            old_data = f.read()
        # Валидация итогового файла

        new_data = old_data.replace(',', ', \n')
        # new_data = new_data.replace("'", '"')

        with open(file_name, 'w') as f:
            f.write(new_data)
            print('Готово, создан файл: ' + file_name)

    def start_app7(self):
        pass

    def start_app8(self):
        pass




