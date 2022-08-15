from win32com.client import Dispatch

def chart2():
    app = Dispatch("Excel.Application")
    workbook_file_name = r"C:\Users\damon\OneDrive\Documents\GitHub\CPxTCG\excel2word\test.xlsx"
    workbook = app.Workbooks.Open(Filename=workbook_file_name)

    app.DisplayAlerts = False

    i = 1
    for sheet in workbook.Worksheets:
        for chartObject in sheet.ChartObjects():
            print(sheet.Name + ':' + chartObject.Name)
            chartObject.Chart.Export(r"C:\Users\damon\OneDrive\Documents\GitHub\CPxTCG\excel2word\Images\Survey" + sheet.Name + ".png")
            i += 1

    workbook.Close(SaveChanges=False, Filename=workbook_file_name)

chart2()
        

