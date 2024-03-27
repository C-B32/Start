import pandas as pd
import xlwings as xw
import sqlite3

path = '2023TN_TAX_DATA.xlsx'
path2 = '1' + path

wingsbook = xw.Book(path)
wingsapp = xw.apps.active
wingsbook.save(path2)
wingsapp.quit()
df = pd.read_excel( path2 , engine='openpyxl')

path3 = "C:\\Users\\Cameron\\GPT Work\\TN_TAX_DATA.db"
conn = sqlite3.connect(path3)
df.to_sql('TaxData', conn, if_exists='replace', index=False)
conn.close()
