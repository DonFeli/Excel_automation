from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


import pandas as pd

pd.set_option("display.max_columns", 500)
pd.set_option("display.width", 0)

workbook = Workbook()
sheet = workbook.active


def print_rows():
    for row in sheet.iter_rows(values_only=True):
        print(row)


data = pd.read_csv('data/admissions.csv')
print(type(data))

for r in dataframe_to_rows(data, index=True, header=True):
    sheet.append(r)


# print_rows()

# workbook.save(filename="admissions.xlsx")
