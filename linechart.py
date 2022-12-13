import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows


wb = Workbook()
ws = wb.active

gs = wb.create_sheet('グラフ')

df = pd.read_csv('./data/matome.csv', encoding='shift-jis')

today_df = df.query('発生日 >= "2022/12/12"')


today_df['発生日'] = pd.to_datetime(
    today_df['発生日'] + ' ' + today_df['発生時刻'])

select_df = today_df[['発生日', '射出最前進位置[mm]', 'V-P切換圧[MPa]']]


for row in dataframe_to_rows(select_df, index=None, header=True):
    ws.append(row)


# 射出最前進位置のグラフ作成
line = LineChart()
line.style = 13
line.height = 17
line.width = 50
data = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row)
labels = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)

line.add_data(data, titles_from_data=True)
line.set_categories(labels)

line.x_axis.title = '射出最前進位置のしきい値は「1.89mm」　これ以上はショートリスク有り！'


gs.add_chart(line, 'A1')

# VP切替圧のグラフ作成

line2 = LineChart()
line2.style = 13
line2.height = 17
line2.width = 50
data = Reference(ws, min_col=3, min_row=1, max_row=ws.max_row)
labels = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)

line2.add_data(data, titles_from_data=True)
line2.set_categories(labels)

line2.x_axis.title = 'VP切替圧のしきい値は「47.3Mpa」　これ以下はショートリスク有り！'


gs.add_chart(line2, 'A33')


wb.save('H:\マイドライブ\成形機モニターデータ\横浜モニターデータ\横浜モニターデータ.xlsx')
