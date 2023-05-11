import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
import openpyxl

wb = openpyxl.load_workbook('excel.xlsx')
sheet = wb.active
doc = DocxTemplate("word1.docx")


path = 'excel.xlsx'  # Свой путь к файлу

# while True:

    # search_text = input(' Введите фамилию работника: ')
search_text = (input('Введите имя сотрудника: '))
print('Ищем:', search_text)

for sheet in pd.ExcelFile(path).sheet_names:

    df = pd.read_excel(path, sheet_name=sheet)
        # если на листе нет заголовков столбцов, то нужно
        # df = pd.read_excel(path, sheet_name=sheet, header=None)

    df_find = df[df.apply(lambda row: row.astype(str).str.contains(search_text).any(), axis=1)]
    if not df_find.empty:
        print(df_find)


df = pd.read_excel('excel.xlsx')

for index, row in df.iterrows():
    r = row._values
    context = {
            'send': r[0],
            'date': r[1],
            'dateHolliday': r[2],
            'count': r[3],
            }
    if r[0] == search_text:
        doc.render(context)
        doc.save(f"generated_doc_{index}.docx")