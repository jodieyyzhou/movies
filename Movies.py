import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import openpyxl
import xlsxwriter
from datetime import date

#load file and join separate sheets into one table
mv = pd.concat(pd.read_excel('movies.xls',sheet_name = None),ignore_index = True)
#print(mv.info())
#print(mv.isnull().sum())
#find duplicated rows
#duplicatedRowsMV = mv[mv.duplicated()]
#print(duplicatedRowsMV)
#remove duplicated columns
mv = mv.drop_duplicates()

#check if there's any non numeric value in these two columns, convert them to null value
mv['Gross Earnings'] = pd.to_numeric(mv['Gross Earnings'],errors = 'coerce')
mv['Budget'] = pd.to_numeric(mv['Budget'],errors = 'coerce')

#create new column 'net earnings' and save the file with timestamp
mv['Net Earnings'] = mv['Gross Earnings'] - mv['Budget']
mv['Net Earnings in millions'] = (mv['Net Earnings'].astype(float)/1000000).round(0).astype(float) 
mv = mv.sort_values(['Net Earnings in millions'],ascending=False)
writer1 = pd.ExcelWriter('output_{}.xlsx'.format(date.today().strftime('%Y-%m-%d')),engine='xlsxwriter')
mv.to_excel(writer1,sheet_name = 'Sheet1', index = False)
writer1.save()

#create bar chart and embed into excel table with timestamp filename
topten = mv.head(10)
#print(topten)
topten_title = topten['Title']
topten_netearnings = topten['Net Earnings in millions']
plt.barh(topten_title,topten_netearnings,align='center')
plt.xlabel('Net Earnings(Million$)')
plt.title('Top 10 successful commercial Movies')
#prevent the image from popping up in non-iterative server
matplotlib.use('Agg')
#plt.show()
plt.savefig('Figure_1.png',dpi = 300)
wb = openpyxl.load_workbook('output_{}.xlsx'.format(date.today().strftime('%Y-%m-%d')))
ws = wb.active
img = openpyxl.drawing.image.Image('Figure_1.png')
img.anchor = 'AC2'
ws.add_image(img)
wb.save('output_top10_bar_chat_{}.xlsx'.format(date.today().strftime('%Y-%m-%d')))

#create a dataframe with null values in 'Net Earnings' and export with timestamp filename
mv_null = mv[mv['Net Earnings'].isnull()]
mv_null = mv_null.loc[:,['Title','Year']]
mv_null = mv_null.sort_values(['Year'],ascending=True)
writer2 = pd.ExcelWriter('null_output_{}.xlsx'.format(date.today().strftime('%Y-%m-%d')),engine='xlsxwriter')
mv_null.to_excel(writer2,sheet_name = 'Sheet1', index = False)
writer2.save()

