import pandas as pd
import os
import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog
import threading
import numpy as np
from mainWindow import Ui_MainWindow

# Список который отображается в выпадающем списке, те которые мы обрабатываем
LIST_PRICES = [" ", "E2E4", "ТД Булат", "ZipZip", "Новые Айти-Решения"]
FILE_OUTPUT = "./output.xlsx"


class MyWidget(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.status_bar = self.statusBar()
        self.status_bar.showMessage('Ready')  # Меняем статус бар для отображения готовности
        self.SelectFile.clicked.connect(self.open_file)
        self.RunScript.clicked.connect(self.process_parse)
        self.comboBox.addItems(LIST_PRICES)

    def changeColourBar(self, color=("(255,0,0,255)")):  # Функция для смена цвета статус бара
        self.status_bar.setStyleSheet(
            "QStatusBar{padding-left:8px;background:rgba" + color + ";color:black;font-weight:bold;}")

    def open_file(self):  # Функция для открытия файла и записи пути к нему
        self.file_name = QFileDialog.getOpenFileName(None, "Open", "")
        if self.file_name[0] != '':
            self.filePath.setText(self.file_name[0])

    def process_parse(self):  # Основная функция для обработки файлов по алогиртмам
        """
        Проверяет наличие необходимой информации перед обработкой.
        В случае отсутствия данных отображает ошибки в статус-баре.
        """
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
            # Если все данные введены корректно, запускается обработка файла
            self.changeColourBar("(255,255,0,255)")
            self.status_bar.showMessage('Выполняется...')
            """
            ComboBox выпадающий список проеряем айди и на каждый элемент запускаем соответсвующий скрипт
            """
            if self.comboBox.currentIndex() == 1:  # E2E4
                self.t1 = threading.Thread(target=self.parse_E2E4, args=(self.filePath.text(), "./outputE2E4.xlsx",),
                                           daemon=True)
                self.t1.start()
            elif self.comboBox.currentIndex() == 2:  # Bulat
                self.t1 = threading.Thread(target=self.parse_Bulat, args=(self.filePath.text(), "./outputBulat.xlsx",
                                                                          float(self.KursEdit.text())), daemon=True)
                self.t1.start()
            elif self.comboBox.currentIndex() == 3:  # ZipZip
                self.t1 = threading.Thread(target=self.parse_ZipZip, args=(self.filePath.text(),
                                                                           "./outputZipZip.xlsx",
                                                                           float(self.KursEdit.text())), daemon=True)
                self.t1.start()
            elif self.comboBox.currentIndex() == 4:  # NewItSolutions
                self.t1 = threading.Thread(target=self.parse_It_Solutions, args=(self.filePath.text(),
                                                                                 "./outputNewItSolutions.xlsx",
                                                                                 float(self.KursEdit.text())),
                                           daemon=True)
                self.t1.start()

    def parse_E2E4(self, file_path_input, file_path_output):
        # Специфическая обработка для поставщика Е2Е4
        process_file(file_path_input).to_excel(file_path_output, index=False)
        self.changeColourBar("(0,255,0,255)")
        self.status_bar.showMessage(f"Данные сохранены в '{file_path_output}'.")

    def parse_It_Solutions(self, file_path_input, file_path_output, kurs, start_pos=1):
        # Специфическая обработка для поставщика Новые Айти-решения
        manufacturer = "Новые Айти-решения"
        # Словарь где хранятся айди каждого из столбцов
        ids = {"Поставщик": manufacturer,
               "Вендор": 2,
               "Артикул": 0,
               "Наименование": 1,
               "Стоимость": 3,
               "Ресурс печати": "0",
               "Количество на складе": 4,
               "Склад": "Москва"}
        # Словарь где будут хранится данные выгруженные из файла
        outputData = {"Поставщик": [],
                      "Вендор": [],
                      "Артикул": [],
                      "Наименование": [],
                      "Стоимость": [],
                      "Ресурс печати": [],
                      "Количество на складе": [],
                      "Склад": []}
        # читаем файл в датафрейм
        df_inp = pd.read_excel(file_path_input)
        counter = 1
        # читаем его построчно
        for row in df_inp.iterrows():
            if counter >= start_pos:  # проверяем на начало не с первой строки
                for i in ids:  # выбираем данные из столбцов, обрабатываем и добавляем в словарь
                    if isinstance(ids[i], int):
                        if i == "Стоимость":
                            outputData[i].append(float(row[1].iloc[ids[i]]) * kurs)
                        elif i == "Количество на складе":
                            try:
                                outputData[i].append(int(row[1].iloc[ids[i]]))
                            except:
                                outputData[i].append(888)
                        else:
                            outputData[i].append(row[1].iloc[ids[i]])
                    else:
                        if i == "Ресурс печати":
                            outputData[i].append(int(ids[i]))
                        else:
                            outputData[i].append(ids[i])
            counter += 1
        df_out = pd.DataFrame.from_dict(outputData)  # Записываем словарь в файл
        df_out.to_excel(file_path_output, index=False)
        self.changeColourBar("(0,255,0,255)")
        self.status_bar.showMessage(f"Данные сохранены в '{file_path_output}'.")  # Меняем статус бар на обработаный

    def parse_Bulat(self, file_path_input, file_path_output, kurs, start_pos=10):
        # Специфическая обработка для поставщика ТД булат
        manufacturer = "ТД Булат"
        # Словарь где хранятся айди каждого из столбцов
        ids = {"Поставщик": manufacturer,
               "Вендор": None,
               "Артикул": 0,
               "Наименование": 1,
               "Стоимость": 2,
               "Ресурс печати": "0",
               "Количество на складе": "888",
               "Склад": "Москва"}
        # Словарь где будут хранится данные выгруженные из файла
        outputData = {"Поставщик": [],
                      "Вендор": [],
                      "Артикул": [],
                      "Наименование": [],
                      "Стоимость": [],
                      "Ресурс печати": [],
                      "Количество на складе": [],
                      "Склад": []}
        # читаем файл в датафрейм
        df_inp = pd.read_excel(file_path_input)
        counter = 1
        # читаем его построчно
        for row in df_inp.iterrows():
            if counter >= start_pos:  # проверяем на начало не с первой строки
                """
                Если мы нашли пустой артикул значит это объединенная строчка
                От сюда мы можем записать вендор и использовать его до следующей объединенной строчки
                """
                if str(row[1].iloc[0]) == "nan":
                    ids["Вендор"] = str(row[1].iloc[ids["Наименование"]])
                    continue
                for i in ids:  # выбираем данные  из столбцов, обрабатываем и добавляем в словарь
                    if isinstance(ids[i], int):
                        if i == "Стоимость":
                            outputData[i].append(float(row[1].iloc[ids[i]]) * kurs)
                        else:
                            outputData[i].append(row[1].iloc[ids[i]])
                    else:
                        if i == "Вендор":
                            outputData[i].append(str(ids[i]).split()[0])
                        elif i == "Количество на складе" or i == "Ресурс печати":
                            outputData[i].append(int(ids[i]))
                        else:
                            outputData[i].append(ids[i])
            counter += 1
        df_out = pd.DataFrame.from_dict(outputData)  # Записываем словарь в файл
        df_out.to_excel(file_path_output, index=False)
        self.changeColourBar("(0,255,0,255)")
        self.status_bar.showMessage(f"Данные сохранены в '{file_path_output}'.")

    def parse_ZipZip(self, file_path_input, file_path_output, kurs, start_pos=3):
        # Специфическая обработка для поставщика ZipZip
        manufacturer = "ZipZip"
        # Словарь где хранятся айди каждого из столбцов
        ids = {"Поставщик": manufacturer,
               "Вендор": 0,
               "Артикул": 1,
               "Наименование": 3,
               "Стоимость": 4,
               "Ресурс печати": "0",
               "Количество на складе": 9,
               "Склад": "Москва"}
        # Словарь где будут хранится данные выгруженные из файла
        outputData = {"Поставщик": [],
                      "Вендор": [],
                      "Артикул": [],
                      "Наименование": [],
                      "Стоимость": [],
                      "Ресурс печати": [],
                      "Количество на складе": [],
                      "Склад": []}
        df_inp = pd.read_excel(file_path_input)
        # читаем файл в датафрейм
        counter = 1
        # читаем его построчно
        for row in df_inp.iterrows():
            if counter >= start_pos:
                for i in ids:  # выбираем данные из столбцов, обрабатываем и добавляем в словарь
                    if isinstance(ids[i], int):
                        if i == "Стоимость":
                            try:
                                outputData[i].append(float(row[1].iloc[ids[i]]))
                            except:
                                outputData[i].append(0)
                            # Преобразуем цену в численный тип данных, для корректного добавления в БД
                        elif i == "Наименование":
                            outputData[i].append(str(row[1].iloc[ids[i]]).replace("=", "-"))
                            # Заменим = на - во избежании ошибок в выходном эксель файле
                        elif i == "Количество на складе":
                            if row[1].iloc[ids[i]] == "+":
                                outputData[i].append(1)
                            else:
                                outputData[i].append(0)
                        else:
                            outputData[i].append(row[1].iloc[ids[i]])
                    else:
                        if i == "Ресурс печати":
                            outputData[i].append(int(ids[i]))
                        else:
                            outputData[i].append(ids[i])
            counter += 1
        df_out = pd.DataFrame.from_dict(outputData)  # Записываем словарь в файл
        df_out.to_excel(file_path_output, index=False)
        self.changeColourBar("(0,255,0,255)")
        self.status_bar.showMessage(f"Данные сохранены в '{file_path_output}'.")


def get_column(df, possible_names, default_value=None):
    """
    Возвращает значение столбца, если он существует, или заполняет столбец значением по умолчанию.
    """
    for name in possible_names:  # Проверяем имя в списке имен
        if name in df.columns:
            return df[name]
    return pd.Series([default_value] * len(df))


def process_file(file_path):
    try:
        # Загружаем данные
        df = pd.read_excel(file_path)
        df.columns = [col.strip().lower() for col in df.columns]
        # Получаем название файла для столбца "Поставщик"
        manufacturer = os.path.splitext(os.path.basename(file_path))[0]

        # Заполняем необходимые столбцы
        df['Стоимость'] = get_column(df, ['оптовая цена, руб.', 'ценв rub', 'цена, руб',
                                          'розница руб.', 'цена партнера', 'price'], None)
        df['Поставщик'] = manufacturer
        df['Вендор'] = get_column(df, ['марка', 'производитель', 'бренд'], 'Не указан')
        df['Артикул'] = get_column(df, ['артикул', 'арт.'], 'Не указан')
        df['Наименование'] = get_column(df, ['наименование', 'описание'], 'Не указано')
        df['Ресурс печати'] = get_column(df, ['макс кол-во отпечатков'], 0)
        df['Количество на складе'] = get_column(df, ['кол-во', 'наличие'], 0)
        df['Количество на складе'] =  df['Количество на складе'].fillna(888)
        df['Склад'] = get_column(df, ['город', 'склад'], 'Москва')
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
