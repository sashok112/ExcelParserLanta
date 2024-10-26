# from openpyxl import load_workbook
# import openpyxl
#
# def parserE2E4(file_path, start_pos=2):
#
#     ids = {"Поставщик": "E2E4",
#                 "Вендор": 2,
#                 "Артикул": 0,
#                 "Наименование": 3,
#                 "Стоимость": 6,
#                 "Ресурс печати": None,
#                 "Количество на складе": None,
#                 "Склад": "Москва"}
#     wb = load_workbook(filename=file_path, read_only=True)
#     ws = wb.active
#     headers = None
#     counter = 1
#     for row in ws.iter_rows(values_only=True):
#         if counter >= start_pos and row[1] == "Оргтехника и расходные материалы|ЗИП печатной техники|Прочий ЗИП для печатной техники":
#             print("---------------")
#             for i in ids:
#                 try:
#                     if isinstance(ids[i], int):
#                         print(i + "-" + str(row[ids[i]]))
#                     else:
#                         print(i + "-" + ids[i])
#                 except:
#                         print(i + "-" + "NONE")
#             print("---------------")
#         counter += 1
#
# def parserZipZip(file_path, start_pos=3):
#     file_path1 = 'Шаблон.xlsx'  # путь к файлу
#     workbook = openpyxl.load_workbook(file_path1)
#     sheet = workbook.active
#     ids = {"Поставщик": "ZipZip",
#                 "Вендор": 0,
#                 "Артикул": 1,
#                 "Наименование": 3,
#                 "Стоимость": 4,
#                 "Ресурс печати": None,
#                 "Количество на складе": 9,
#                 "Склад": "Москва"}
#     wb = load_workbook(filename=file_path, read_only=True)
#     ws = wb.active
#     counter = 1
#     for row in ws.iter_rows(values_only=True):
#         if counter > start_pos:
#             temp = []
#             for i in ids:
#                 try:
#                     if isinstance(ids[i], int):
#                         if i == "Количество на складе" and row[ids[i]] == "+":
#                             temp.append(1)
#                         else:
#                             temp.append(str(row[ids[i]]))
#                     else:
#                         temp.append(ids[i])
#                 except:
#                     temp.append("Nonne")
#             sheet.append(temp)
#         counter += 1
#     workbook.save(file_path1)
#
#
# parserZipZip("/vendorData/Прайс ZipZip Оригинад.xlsx")
#
#
#
#
# import pandas as pd
# import glob
# import os
#
# data_folder = './vendorData'
# files = glob.glob(os.path.join(data_folder, '*.xlsx'))
#
# def get_column(df, possible_names, default_value=None):
#     """
#     Возвращает значение столбца, если он существует, или заполняет столбец значением по умолчанию.
#     """
#     for name in possible_names:
#         if name in df.columns:
#             return df[name]
#     return pd.Series([default_value] * len(df))
#
# def get_price(df):
#     return get_column(df, ['оптовая цена, руб.', 'ценв rub', 'цена, руб', 'розница руб.', 'цена партнера', 'price'], None)
#
# def get_vendor(df):
#     return get_column(df, ['марка', 'производитель', 'бренд'], 'Не указан')
#
# def get_sklad(df):
#     return get_column(df, ['город', 'склад'], 'Москва')
#
# def process_file(file_path):
#     try:
#         # Загружаем данные
#         df = pd.read_excel(file_path)
#         df.columns = [col.strip().lower() for col in df.columns]
#
#         # Получаем название файла для столбца "Поставщик"
#         manufacturer = os.path.splitext(os.path.basename(file_path))[0]
#
#         # Заполняем необходимые столбцы
#         df['Стоимость'] = get_price(df)
#         df['Поставщик'] = manufacturer
#         df['Вендор'] = get_vendor(df)
#         df['Артикул'] = get_column(df, ['артикул'], 'Не указан')
#         df['Наименование'] = get_column(df, ['наименование'], 'Не указано')
#         df['Ресурс печати'] = get_column(df, ['макс кол-во отпечатков'], 0)
#         df['Количество на складе'] = get_column(df, ['кол-во', 'наличие'], 0)
#         df['Склад'] = get_sklad(df)
#
#         # Выбираем только нужные столбцы
#         result_df = df[['Поставщик', 'Вендор', 'Артикул', 'Наименование', 'Стоимость',
#                         'Ресурс печати', 'Количество на складе', 'Склад']]
#         return result_df
#     except Exception as e:
#         print(f"Ошибка при обработке файла {file_path}: {e}")
#         return pd.DataFrame()
#
# # Основной процесс обработки всех файлов
# final_df = pd.DataFrame()
#
# for file in files:
#     processed_df = process_file(file)
#     final_df = pd.concat([final_df, processed_df], ignore_index=True)
#
# # Сохраняем результат в Excel
# output_file = './обработанные_данные.xlsx'
# final_df.to_excel(output_file, index=False)
# print(f"Обработка завершена. Данные сохранены в '{output_file}'.")

import pandas as pd

df1 = pd.DataFrame([['a', 'b'], ['c', 'd']],
                   index=['row 1', 'row 2'],
                   columns=['col 1', 'col 2'])
df1.to_excel("output.xlsx")

df = pd.read_excel('vendorData/Прайс ZipZip Оригинад.xlsx', index_col=0,
              dtype={'Name': str, 'Value': float})

newdata = {'1':[], '2':[], '3':[], '4':[]}
for i in df.values:
    newdata['1'].append(str(i[0]))
    newdata['2'].append(str(i[1]))
    newdata['3'].append(str(i[2]))
    newdata['4'].append(str(i[3]))
df2 = pd.DataFrame(newdata)
df2.to_excel('writer.xlsx', index=False)

