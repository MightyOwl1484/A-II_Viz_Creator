# A-II_Viz_Creator
Automates Visualizations from A-II data

VBA Script: Create A-II Visualization
This VBA script automates the process of creating visualizations from data in a source Excel workbook and copying them to a destination workbook. The script loops through each worksheet in the source workbook, excluding specific sheets, and generates a custom combination chart for each sheet.

Prerequisites
Before running the script, ensure you have the following:

Microsoft Excel with VBA enabled.
The source workbook named Current A-II.xlsx.
The destination workbook named viz.xlsx.
Script Overview
The script performs the following steps:

Opens the source workbook (Current A-II.xlsx) and the destination workbook (viz.xlsx).
Loops through each worksheet in the source workbook, excluding the "Cover Page" and "Consolidated Report" sheets.
Copies and transposes specific data ranges from each source worksheet to a corresponding new worksheet in the destination workbook.
Creates a custom combination chart in the destination worksheet with specific series and formats.
Sets the chart title based on the source worksheet name.
Cleans up objects to free memory.
Usage
Open both the source and destination workbooks in Excel.
Open the VBA editor (Alt + F11).
Insert a new module and paste the script into the module.
Run the script by pressing F5 or by using the "Run" button in the VBA editor.
Detailed Function Description
Sub CreateAIIVisualization()
This is the main subroutine that handles the entire process.

Variables and Objects
srcWB: Workbook object for the source workbook.
destWB: Workbook object for the destination workbook.
srcWS: Worksheet object for the current source worksheet being processed.
destWS: Worksheet object for the new worksheet created in the destination workbook.
wsName: String variable to store the worksheet name.
lastRow: Long variable to store the last row number (not used in this script).
i: Integer variable for loop iterations.
chartObj: ChartObject variable for the chart created in the destination worksheet.
chartTitle: String variable to store the chart title.
series: Series variable for the chart series (not used directly in the script).
Steps
Define the Source and Destination Workbooks:

vba
Copy code
Set srcWB = Workbooks("Current A-II.xlsx")
Set destWB = Workbooks("viz.xlsx")
Loop Through Each Worksheet in the Source Workbook:

Exclude the "Cover Page" and "Consolidated Report" sheets.

vba
Copy code
For Each srcWS In srcWB.Worksheets
    If srcWS.Name <> "Cover Page" And srcWS.Name <> "Consolidated Report" Then
Add a New Worksheet to the Destination Workbook:

Create a new worksheet with the same name as the source worksheet.

vba
Copy code
Set destWS = destWB.Sheets.Add(After:=destWB.Sheets(destWB.Sheets.Count))
destWS.Name = srcWS.Name
Set the Headers in the Destination Worksheet:

vba
Copy code
destWS.Range("A2").Value = "FY"
destWS.Range("B2").Value = "Total PAA"
destWS.Range("C2").Value = "Total Requirement"
destWS.Range("D2").Value = "TAI"
destWS.Range("E2").Value = "TII"
Copy and Transpose the Specified Ranges:

vba
Copy code
srcWS.Range("G2:AJ2").Copy
destWS.Range("A3").PasteSpecial Paste:=xlPasteAll, Transpose:=True

srcWS.Range("G7:AJ7").Copy
destWS.Range("B3").PasteSpecial Paste:=xlPasteAll, Transpose:=True

srcWS.Range("G10:AJ10").Copy
destWS.Range("C3").PasteSpecial Paste:=xlPasteAll, Transpose:=True

srcWS.Range("G40:AJ40").Copy
destWS.Range("D3").PasteSpecial Paste:=xlPasteAll, Transpose:=True

srcWS.Range("G49:AJ49").Copy
destWS.Range("E3").PasteSpecial Paste:=xlPasteAll, Transpose:=True
Create the Custom Combination Chart:

vba
Copy code
Set chartObj = destWS.ChartObjects.Add(Left:=100, Width:=500, Top:=50, Height:=300)
With chartObj.Chart
    .SetSourceData Source:=destWS.Range("A2:E32")
    .ChartType = xlColumnStacked
Set series for TAI, TII, Total PAA, and Total Requirement with specific colors and types.

vba
Copy code
.SeriesCollection.NewSeries
.SeriesCollection(1).Name = "TAI"
.SeriesCollection(1).Values = destWS.Range("D3:D32")
.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)

.SeriesCollection.NewSeries
.SeriesCollection(2).Name = "TII"
.SeriesCollection(2).Values = destWS.Range("E3:E32")
.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)

.SeriesCollection.NewSeries
.SeriesCollection(3).Name = "Total PAA"
.SeriesCollection(3).Values = destWS.Range("B3:B32")
.SeriesCollection(3).ChartType = xlLine
.SeriesCollection(3).AxisGroup = xlSecondary
.SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(0, 0, 255)

.SeriesCollection.NewSeries
.SeriesCollection(4).Name = "Total Requirement"
.SeriesCollection(4).Values = destWS.Range("C3:C32")
.SeriesCollection(4).ChartType = xlLine
.SeriesCollection(4).AxisGroup = xlSecondary
.SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
Set the Chart Title and Remove Secondary Axis:

vba
Copy code
chartTitle = srcWS.Name & " Total PAA and Total RQMT (PAA + BAA) over TAI + TII"
.HasTitle = True
.ChartTitle.Text = chartTitle
.Axes(xlValue, xlSecondary).Delete
Clean Up:

vba
Copy code
Application.CutCopyMode = False
Set srcWB = Nothing
Set destWB = Nothing
Set srcWS = Nothing
Set destWS = Nothing
Set chartObj = Nothing
Notes
Ensure the source and destination workbooks are open before running the script.
Modify the script to suit any specific requirements or additional customization.
The script assumes the data ranges are consistent across all source worksheets.
