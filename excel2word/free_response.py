from openpyxl import load_workbook
workbook = load_workbook(filename="Survey.xlsx") 

free_responses={}
for sheet in workbook.worksheets:
    for i in range(1,20):
        s=sheet[f"A{i}"].value
        if s =='Respondent ID':
            #print(sheet.title)
            res=[]
            for cell in sheet['C']:
                if (cell.value!=None) and (cell.value!='Responses'):
                    res.append(cell.value) # add res to that single question 
            #print(res)
            free_responses[sheet.title]=res  #add list of res to one question to the big list
            
            break

import docx

doc = docx.Document()
for key,values in free_responses.items():
    doc.add_paragraph(key)
    for res in values:
        doc.add_paragraph(res,style='List Bullet 2')
doc.save('output.docx')