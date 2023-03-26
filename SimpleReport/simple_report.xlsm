/******************************************************
 *  Program Name: Simple Excel Report
 *  Developer: Ramesh Raj
 *  Contact: mr_anusiram@yahoo.com
 *  Date: March 18, 2023 
 *  Version: 1.0
 *  Description: This code creates a new worksheet called 'Revenue Report'
 *  from the master worksheet and copy A to U and AI to AL 
 *  project Z for company ABC.
 *  Acknowledgments: Thanks to Jane Doe for providing 
 *  valuable feedback during code review.
 ******************************************************/

Sub CreateRevenueReport()
    Dim wsMaster As Worksheet
    Dim wsRevenueReport As Worksheet
    Dim lastRow As Long
    
    'Get reference to the Master worksheet
    Set wsMaster = ThisWorkbook.Worksheets("Master")
    
    'Create a new worksheet called "Revenue Report"
    Set wsRevenueReport = ThisWorkbook.Worksheets.Add
    wsRevenueReport.Name = "Revenue Report"
    
    'Copy columns A to U and AI to AL from the Master worksheet to the new worksheet
    lastRow = wsMaster.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    wsMaster.Range("A1:U" & lastRow & ", AI1:AL" & lastRow).Copy Destination:=wsRevenueReport.Range("A1")  

    wsRevenueReport.Columns.AutoFit    
    wsRevenueReport.Activate
End Sub
