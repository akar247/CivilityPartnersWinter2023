from openpyxl import load_workbook
workbook = load_workbook(filename="Survey.xlsx") 

multiple_choices={}
for sheet in workbook.worksheets:
    for i in range(1,50):
        s=sheet[f"A{i}"].value
        if s =='Answer Choices': # if the question is multiple choice
            res=[]
            #for cell in sheet[f'C{i+1}:C1048576']:
            for value in sheet.iter_rows(min_row=i+1,min_col=3, max_col=3,values_only=True):
                if (value[0]!=None):
                    res.append(value[0]) # add res to that single question 
            #print(res)
            multiple_choices[sheet.title]=res  #add list of res to one question to the big list
            break

import docx

doc = docx.Document()
for key,values in multiple_choices.items():
    doc.add_paragraph(key)
    for res in values:
        doc.add_paragraph(res,style='List Bullet 2')
doc.save('output.docx')

from docx import Document
from pptx.util import Pt, Inches
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_DATA_LABEL_POSITION

document = Document()
chart_data = CategoryChartData()
chart_data.categories = ['Strongly Disagree', 'Disagree', 'Agree', 'Strongly Agree']
chart_data.add_series('Series 1', (19.2, 21.4, 16.7))
x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
chart = document.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)

chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False

plot = chart.plots[0]
plot.has_data_labels = True
data_labels = plot.data_labels
data_labels.font.size = Pt(13)
data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END

chart.has_title = True
chart_title = chart.chart_title
text_frame = chart_title.text_frame
text_frame.text = 'Title'
paragraphs = text_frame.paragraphs
paragraph = paragraphs[0]
paragraph.font.size = Pt(18)

category_axis = chart.category_axis
category_axis.tick_labels.font.size = Pt(14)

document.save('test.docx')
