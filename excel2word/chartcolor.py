from openpyxl import load_workbook

workbook = load_workbook(filename="Survey.xlsx") 
#workbook.sheetnames
#workbook.worksheets
import docx

from openpyxl.chart import BarChart, Series, Reference
from openpyxl.chart.label import DataLabelList
wb = load_workbook(filename="Survey.xlsx") 
ws = wb.active

for ws in wb.worksheets:
    # Data for plotting
    values = Reference(ws, min_col= 2, min_row=3, max_col = 2, max_row = 7)
    cats = Reference(ws, min_col=1, max_col=1, min_row=4, max_row=8)

    # Create object of BarChart class
    chart = BarChart()
    chart.height = 10
    chart.width = 15
    chart.add_data(values, titles_from_data=True)
    chart.set_categories(cats)
    # set the title of the chart
    chart.title = ws["A2"].value

    # the top-left corner of the chart
    # is anchored to cell F2 .
    chart.varyColors = "000F0FFF"
    ws.add_chart(chart,"H2")

# save the file 
wb.save("test.xlsx")   ###check the test result at new created "test.xlsx" file 