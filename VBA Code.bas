Attribute VB_Name = "Module1"
Sub bookDemo() ' Manipulate collections of multiple workbooks

    Dim wb As Workbook
    Dim x As Integer

    ' Add a new workbook
    Application.Workbooks.Add
    
    ' Print the count of open workbooks
    Debug.Print Application.Workbooks.Count 'Shows how many workbooks you have open in Excel right now. If you have a current macro workbook running in the background, it will count that as well

    ' Open a specific workbook
    Set wb = Application.Workbooks.Open(ThisWorkbook.Path & "\SalaryReductionProposal.xlsx")
    
    ' Add a worksheet ot the opened workbook
    wb.Worksheets.Add
    
    ' Close the workbook with saving changes
    wb.Close True
    
    ' Update the workbook count
    x = Application.Workbooks.Count
    
    ' Loop through open workbooks and close them without saving
    Do Until x < 2
        Debug.Print x, Application.Workbooks(x).Name
        Application.Workbooks(x).Close False
        x = x - 1
        DoEvents
    Loop

End Sub
Sub sheetDemo() ' Manipulate collections of multiple worksheets

Dim x As Integer
Dim ws As Object 'Object variable

'Add a worksheet after the first worksheet in this workbook
Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(1))

'Activate the first worksheet and print the name of the active sheet
ThisWorkbook.Worksheets(1).Activate
Debug.Print ActiveSheet.Name

' Activate a specific worksheet by name
ThisWorkbook.Worksheets("Charts").Activate

' Print the name of the newly added worksheet
Debug.Print ws.Name

' Print the count of worksheets in the active workbook
Debug.Print ThisWorkbook.Worksheets.Count


x = 1

' Loop through the collection of worksheets and print their names
Do Until x > ThisWorkbook.Worksheets.Count ' Execute code on each worksheet in the collection until there are no worksheets left
    Debug.Print x, ThisWorkbook.Worksheets(x).Name
    x = x + 1
 DoEvents
Loop

End Sub
Sub chartDemo() ' Move all charts on an excel sheet using a loop

Dim shp As Shape
Dim s As Worksheet
Dim x As Integer

' Set the worksheet ot eh one named "Charts" in the current workbook
Set s = ThisWorkbook.Worksheets("Charts")

' Set the initial shape to the first shape on the sheet
Set shp = s.Shapes(1) 's.shapes represents the shapes (charts in this case) on the sheet

' Loop through all the shapes ont he sheet
x = 1
Do Until x > s.Shapes.Count
    ' Move each shape to the right by 100 units
    s.Shapes(x).Left = s.Shapes(x).Left + 100
    x = x + 1
 ' Allow events to occur during the loop
 DoEvents
Loop


End Sub

Sub changeWBvisible() 'Change the visibility of a workbook to make it visible

Dim objExcel As Object
Set objExcel = CreateObject("Excel.Application")

' Open the workbook and set it as the active workbook
Dim wb As Object
Set wb = objExcel.Workbooks.Open(ThisWorkbook.Path & "\SalaryReductionProposal.xlsx")  'objExcel is a top-level view of Excel
wb.Activate

' Make only the activated workbook visible
wb.Windows(1).Visible = True

' Perform operations on the active workbook
objExcel.Workbooks(1).Worksheets(1).Cells(1).Value = "Hi there. It's " & Now

End Sub
