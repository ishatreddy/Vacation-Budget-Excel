# Install xlsxwriter using the below command
# pip install xlsxwriter
# Explore more Options with the PIL Library using the below documentation
# https://xlsxwriter.readthedocs.io/tutorial01.html

import xlsxwriter
from datetime import datetime

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('VacationBudget.xlsx')
worksheet = workbook.add_worksheet('Vacation Budget')

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#####0'})
money.set_bg_color('#C0C0C0')

# Add an Excel date format.
date_format = workbook.add_format({'num_format': 'dd/mm/yy', 'align': 'left'})

# Adjust the column width.
worksheet.set_column('B:B', 10)
worksheet.set_column('A:A', 15)

# Write some data headers.
cell_format = workbook.add_format({'bold': True, 'font_color': 'Blue'})
cell_format.set_font_name('Arial Black')
cell_format.set_font_size(12)
cell_format.set_underline()
worksheet.write('A1', 'Item', cell_format)
worksheet.write('B1', 'Date', cell_format)
worksheet.write('C1', 'Cost', cell_format)

# Sample data we want to write to the worksheet
budget = (['Flight', '2019-3-17', 2500],
    ['Hotel', '2019-3-18', 1000],
    ['Food', '2019-3-18', 300],
    ['Sight Seeing', '2019-3-19', 1000],
    ['Local Transport', '2019-3-19', 300],
    ['Shopping', '2019-3-20', 1000])

# First Row and First Columns start at Index of 0
row = 1
col = 0

# Iterate over the budget dictionary and write it to the Excel File row by row.
for item, date_str, cost in budget:
    date = datetime.strptime(date_str, "%Y-%m-%d")
    cell_format2 = workbook.add_format()
    cell_format2.set_italic()
    worksheet.write_string(row, col, item, cell_format2)
    worksheet.write_datetime(row, col + 1, date, date_format)
    worksheet.write_number(row, col + 2, cost, money)
    row += 1

# Write a total at the bottom of the Budget Data using an Excel Formula
cell_format1 = workbook.add_format({'bold': True, 'font_color': 'red'})
cell_format1.set_font_name('Arial Black')
cell_format1.set_font_size(12)
cell_format1.set_underline()
worksheet.write(row, 0, 'Total', cell_format1)
worksheet.write(row, 1, 'N/A')
worksheet.write(row, 2, '=SUM(B2:B5)', money)


worksheet1 = workbook.add_worksheet()

# Add the worksheet data to be plotted.
data = [10, 40, 50, 20, 10, 50]
worksheet1.write_column('A1', data)

# Create a new chart object.
chart = workbook.add_chart({'type': 'line'})

# Add a series to the chart.
chart.add_series({
    'values': '=Sheet2!$A$1:$A$6',
    'marker': {'type': 'circle',
    'size': 6, 'border': {'color': 'blue'},
    'fill':{'color': 'blue'}},
    'trendline': {'type': 'linear', 'line': {'color': 'red', 'width': 2,}}
})

# Insert the chart into the worksheet.
worksheet1.insert_chart('E3', chart)


workbook.close()
