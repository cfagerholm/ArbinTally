import win32com.client

# Grab the Active Instance of Excel.
ExcelApp = win32com.client.GetActiveObject("Excel.Application")
ExcelApp.Visible = True

# Grab the workbook with the charts.
wb = ExcelApp.Workbooks.Open(r'C:\\Users\\cara.fagerholm\\Documents\\DataMgmt_Code\\book1.xlsx')

# Create a new instance of PowerPoint and make sure it's visible.
PPTApp = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
PPTApp.Visible = True

# Add a presentation to the PowerPoint Application, returns a Presentation Object.
PPTPresentation = PPTApp.Presentations.Add()

# Loop through each Worksheet.
for xlWorksheet in wb.Worksheets:

    # Grab the ChartObjects Collection for each sheet.
    xlCharts = xlWorksheet.ChartObjects()

    # Loop through each Chart in the ChartObjects Collection.
    for index, xlChart in enumerate(xlCharts):
        # Each chart needs to be on it's own slide, so at this point create a new slide.
        PPTSlide = PPTPresentation.Slides.Add(Index=index + 1, Layout=12)  # 12 is a blank layout

        # Display something to the user.
        print('Exporting Chart {} from Worksheet {}'.format(xlChart.Name, xlWorksheet.Name))

        # Copy the chart.
        xlChart.Copy()

        # Paste the Object to the Slide
        PPTSlide.Shapes.Paste

    # Save the presentation.
PPTPresentation.SaveAs(r"C:\\Users\\cara.fagerholm\\Documents\\DataMgmt_Code\\outppt")
