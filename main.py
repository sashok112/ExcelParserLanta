import pandas as pd
import os
import sys
from PyQt6 import uic
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog
import threading

LIST_PRICES = [" ", "E2E4", "ТД Булат", "ZipZip"]
FILE_OUTPUT = "./output.xlsx"


class MyWidget(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('mainWindow.ui', self)  # Загружаем дизайн
        self.status_bar = self.statusBar()
        self.status_bar.showMessage('Ready')
        self.SelectFile.clicked.connect(self.open_file)
        self.RunScript.clicked.connect(self.process_parse)
        self.comboBox.addItems(LIST_PRICES)

    def changeColourBar(self, color=("(255,0,0,255)")):
        self.status_bar.setStyleSheet(
            "QStatusBar{padding-left:8px;background:rgba" + color + ";color:black;font-weight:bold;}")

    def open_file(self):
        self.file_name = QFileDialog.getOpenFileName(None, "Open", "")
        if self.file_name[0] != '':
            self.filePath.setText(self.file_name[0])

    def process_parse(self):
        if self.filePath.text() == '':
            self.changeColourBar("(255,0,0,255)")
            self.status_bar.showMessage('Не указан путь к файлу')
        elif self.KursEdit.text() == '':
            self.changeColourBar("(255,0,0,255)")
            self.status_bar.showMessage('Не указан курс')
        elif self.comboBox.currentIndex() == 0:
            self.changeColourBar("(255,0,0,255)")
            self.status_bar.showMessage('Не указан поставщик')
        else:
            self.changeColourBar("(255,255,0,255)")
            self.status_bar.showMessage('Выполняется...')
            if self.comboBox.currentIndex() == 1:
                self.t1 = threading.Thread(target=self.parse_E2E4, args=(self.filePath.text(), FILE_OUTPUT,),
                                           daemon=True)
                self.t1.start()
            elif self.comboBox.currentIndex() == 2:
                self.t1 = threading.Thread(target=self.parse_Bulat, args=(self.filePath.text(), FILE_OUTPUT,
                                                                          float(self.KursEdit.text())), daemon=True)
                self.t1.start()
            elif self.comboBox.currentIndex() == 3:
                self.t1 = threading.Thread(target=self.parse_ZipZip, args=(self.filePath.text(), FILE_OUTPUT,
                                                                           float(self.KursEdit.text())), daemon=True)
                self.t1.start()

    def parse_E2E4(self, file_path_input, file_path_output):
        process_file(file_path_input).to_excel(file_path_output, index=False)
        self.changeColourBar("(0,255,0,255)")
        self.status_bar.showMessage(f"Данные сохранены в '{file_path_output}'.")

    def parse_Bulat(self, file_path_input, file_path_output, kurs, start_pos=10):
        manufacturer = "ТД Булат"
        ids = {"Поставщик": manufacturer,
               "Вендор": None,
               "Артикул": 0,
               "Наименование": 1,
               "Стоимость": 2,
               "Ресурс печати": None,
               "Количество на складе": "888",
               "Склад": "Москва"}

        outputData = {"Поставщик": [],
                      "Вендор": [],
                      "Артикул": [],
                      "Наименование": [],
                      "Стоимость": [],
                      "Ресурс печати": [],
                      "Количество на складе": [],
                      "Склад": []}
        df_inp = pd.read_excel(file_path_input)

        counter = 1
        for row in df_inp.iterrows():
            if counter >= start_pos:
                if str(row[1].iloc[0]) == "nan":
                    ids["Вендор"] = str(row[1].iloc[ids["Наименование"]])
                    continue
                for i in ids:
                    if isinstance(ids[i], int):
                        if i == "Стоимость":
                            outputData[i].append(float(row[1].iloc[ids[i]]) * kurs)
                        else:
                            outputData[i].append(row[1].iloc[ids[i]])
                    else:
                        if i == "Вендор":
                            outputData[i].append(str(ids[i]).split()[0])
                        elif i == "Количество на складе":
                            outputData[i].append(int(ids[i]))
                        else:
                            outputData[i].append(ids[i])
            counter += 1
        df_out = pd.DataFrame.from_dict(outputData)
        df_out.to_excel(file_path_output, index=False)
        self.changeColourBar("(0,255,0,255)")
        self.status_bar.showMessage(f"Данные сохранены в '{file_path_output}'.")

    def parse_ZipZip(self, file_path_input, file_path_output, kurs, start_pos=3):
        manufacturer = "ZipZip"
        ids = {"Поставщик": manufacturer,
               "Вендор": 0,
               "Артикул": 1,
               "Наименование": 3,
               "Стоимость": 4,
               "Ресурс печати": None,
               "Количество на складе": 9,
               "Склад": "Москва"}

        outputData = {"Поставщик": [],
                      "Вендор": [],
                      "Артикул": [],
                      "Наименование": [],
                      "Стоимость": [],
                      "Ресурс печати": [],
                      "Количество на складе": [],
                      "Склад": []}
        df_inp = pd.read_excel(file_path_input)

        counter = 1
        for row in df_inp.iterrows():
            if counter >= start_pos:
                for i in ids:
                    if isinstance(ids[i], int):
                        if i == "Стоимость":
                            outputData[i].append(float(str(row[1].iloc[ids[i]])))
                        elif i == "Наименование":
                            outputData[i].append(str(row[1].iloc[ids[i]]).replace("=", "-"))
                        elif i == "Количество на складе":
                            if row[1].iloc[ids[i]] == "+":
                                outputData[i].append(1)
                            else:
                                outputData[i].append(0)
                        else:
                            outputData[i].append(row[1].iloc[ids[i]])
                    else:
                        outputData[i].append(ids[i])
            counter += 1
        df_out = pd.DataFrame.from_dict(outputData)
        df_out.to_excel(file_path_output, index=False)
        self.changeColourBar("(0,255,0,255)")
        self.status_bar.showMessage(f"Данные сохранены в '{file_path_output}'.")


def get_column(df, possible_names, default_value=None):
    """
    Возвращает значение столбца, если он существует, или заполняет столбец значением по умолчанию.
    """
    for name in possible_names:
        if name in df.columns:
            return df[name]
    return pd.Series([default_value] * len(df))


def get_price(df):
    return get_column(df, ['оптовая цена, руб.', 'ценв rub', 'цена, руб', 'розница руб.', 'цена партнера', 'price'],
                      None)


def get_vendor(df):
    return get_column(df, ['марка', 'производитель', 'бренд'], 'Не указан')


def get_sklad(df):
    return get_column(df, ['город', 'склад'], 'Москва')


def process_file(file_path):
    try:
        # Загружаем данные
        df = pd.read_excel(file_path)
        df.columns = [col.strip().lower() for col in df.columns]
        # Получаем название файла для столбца "Поставщик"
        manufacturer = os.path.splitext(os.path.basename(file_path))[0]

        # Заполняем необходимые столбцы
        df['Стоимость'] = get_price(df)
        df['Поставщик'] = manufacturer
        df['Вендор'] = get_vendor(df)
        df['Артикул'] = get_column(df, ['артикул', 'арт.'], 'Не указан')
        df['Наименование'] = get_column(df, ['наименование', 'описание'], 'Не указано')
        df['Ресурс печати'] = get_column(df, ['макс кол-во отпечатков'], 0)
        df['Количество на складе'] = get_column(df, ['кол-во', 'наличие'], 0)
        df['Склад'] = get_sklad(df)

        # Выбираем только нужные столбцы
        result_df = df[['Поставщик', 'Вендор', 'Артикул', 'Наименование', 'Стоимость',
                        'Ресурс печати', 'Количество на складе', 'Склад']]
        return result_df
    except Exception as e:
        print(f"Ошибка при обработке файла {file_path}: {e}")
        return pd.DataFrame()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = MyWidget()
    ex.show()
    sys.exit(app.exec())
