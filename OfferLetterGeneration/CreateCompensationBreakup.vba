 '  Program Name: Offer Letter Generation
 '  Developer: Ramesh Raj
 '  Contact: mr_anusiram@yahoo.com
 '  Date: March 30, 2023
 '  Version: 1.0
 '  Description: This code creates Compensation Breakup in pdf format. It takes input from current sheet and inserts the values into excel templates and creates the required pdfs.
 ' No empty rows should be present in the excel sheet

Sub CreateCompensationBreakup()
    
    ' Define variables
    Dim wbTemplate As Workbook
    Dim wbData As Workbook
    Dim wsTemplate As Worksheet
    Dim wsData As Worksheet
    Dim filePath As String
    Dim empName As String
    Dim offerNo As String
    Dim grossCTC As Double
    Dim perBonus As Double
    
    ' Set file paths
    filePath = "C:\Users\INRARAJ\Automation\HR\Compensation Breakup Sheet Template.xltx"
    dataFilePath = "C:\Users\INRARAJ\Automation\HR\Offer_Data_Nos.xlsm"
    
    ' Open template workbook and get reference to worksheet
    Set wbTemplate = Workbooks.Open(filePath)
    Set wsTemplate = wbTemplate.Worksheets("Email")
    
    ' Open data workbook and get reference to worksheet
    Set wbData = Workbooks.Open(dataFilePath)
    Set wsData = wbData.Worksheets("Sheet1")
    
    ' Get values from last row of data sheet
    With wsData
        empName = .Range("A" & .Rows.Count).End(xlUp).Value
        offerNo = .Range("J" & .Rows.Count).End(xlUp).Value
        grossCTC = .Range("H" & .Rows.Count).End(xlUp).Value
        perBonus = .Range("I" & .Rows.Count).End(xlUp).Value
    End With
    
    ' Insert values into template sheet
    With wsTemplate
        .Range("A7").Value = empName
        .Range("A8").Value = offerNo
        .Range("F12").Value = grossCTC
        .Range("F13").Value = perBonus
    End With
    
    ' Replace spaces with underscore to avoid issues with file names
    empName = Replace(empName, " ", "_")
    ' Save and close template workbook
    wsTemplate.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "C:\Users\INRARAJ\Automation\HR\" & empName & "_Compensation_Breakup.pdf", _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    wbTemplate.Close savechanges:=False    
    
End Sub
