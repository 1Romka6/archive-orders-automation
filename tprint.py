import pandas as pd
import numpy as np
from openpyxl import load_workbook
from tabulate import tabulate


def options():
    pd.set_option('display.max_rows', None)  # Установите None, чтобы показать все строки
    pd.set_option('display.max_colwidth', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 1000)

def tprint(df):
    print(tabulate(df, headers='keys', tablefmt='psql', stralign='center'))