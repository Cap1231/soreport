import os
import sys
import datetime
from dateutil.relativedelta import relativedelta
import xlwings as xw
import re
import pprint

last_month_date = datetime.date.today() - relativedelta(months=1)
last_month = last_month_date.strftime("%Y%m")

# output するフォルダを作成
cwd = os.getcwd()
output_dir_path = os.path.join(cwd, 'output', '201910')
output_dir = os.path.join(cwd, 'output', last_month)
output_file = os.path.join(output_dir, 'output.xlsx')

wb = xw.Book(output_file)
sheet = wb.sheets['total']
sheet.range('B2', 'T21').columns.rng.column_width = 13
sheet.range('B2').api.Borders(1).LineStyle = 1
sheet.range('B2').api.Borders(2).LineStyle = 1
sheet.range('B2').api.Borders(3).LineStyle = 1
sheet.range('B2').api.Borders(4).LineStyle = 1

print(sheet.range('B2'))
print(sheet.range('B2').columns)
