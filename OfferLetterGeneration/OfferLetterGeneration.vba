 '  Program Name: Offer Letter Generation
 '  Developer: Ramesh Raj
 '  Contact: ramesh.raj@fimer.com
 '  Date: March 29, 2023
 '  Version: 1.0
 '  Description: This code creates Offer Letter, Medical Annexure, Compensation Breakup in pdf format. It takes input from current sheet and insert the values into word and excel templates and create the required pdfs. Then, draft the email from template.
 '  No empty rows should be present in the excel sheet

Sub DraftOfferLetterMail()
    ' Declare necessary variables
    
    Dim wApp As Word.Application
    Dim wdoc As Word.Document
    Dim wdoc2 As Word.Document
    Dim empName, path As String
    Dim r As Long
    Dim OutApp As Object
    Dim OutMail As Object
    Dim RecipientEmail As String
    
    
    ' Find the last row of data in the Excel sheet
    r = Sheet1.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Create a new Word application
    Set wApp = CreateObject("Word.Application")
    wApp.Visible = True
    
    ' Open the Offer Letter template
    path = ThisWorkbook.path
    Set wdoc = wApp.Documents.Add(template:=path & "\Offer_Letter_Template.dotx", NewTemplate:=False, DocumentType:=0)
    
    ' Replace placeholders with values from Excel sheet
    With wdoc
        ' Replace «Offer No» with value from column J in the last row
        .Application.Selection.Find.Text = "«Offer No»"
        .Application.Selection.Find.Execute
        .Application.Selection = Sheet1.Cells(r, 10).Value
        .Application.Selection.EndOf
    
        ' Replace «Salutation» with value from column B in the last row
        .Application.Selection.Find.Text = "«Salutation»"
        .Application.Selection.Find.Execute
        .Application.Selection = Sheet1.Cells(r, 2).Value
        .Application.Selection.EndOf
    
        ' Replace «Emp Name» with value from column A in the last row
        .Application.Selection.Find.Text = "«Emp Name»"
        .Application.Selection.Find.Execute
        .Application.Selection = Sheet1.Cells(r, 1).Value
        .Application.Selection.EndOf
        
        ' Replace «Address» with value from column C in the last row
        .Application.Selection.Find.Text = "«Address»"
        .Application.Selection.Find.Execute
        .Application.Selection = Sheet1.Cells(r, 3).Value
        .Application.Selection.EndOf
        
        ' Replace «Mobile» with value from column D in the last row
        .Application.Selection.Find.Text = "«Mobile»"
        .Application.Selection.Find.Execute
        .Application.Selection = Sheet1.Cells(r, 4).Value
        .Application.Selection.EndOf
        
        ' Replace «Salutation» with value from column B in the last row
        .Application.Selection.Find.Text = "«Salutation»"
        .Application.Selection.Find.Execute
        .Application.Selection = Sheet1.Cells(r, 2).Value
        .Application.Selection.EndOf
        
        ' Replace «Emp Name» with value from column A in the last row
        .Application.Selection.Find.Text = "«Emp Name»"
        .Application.Selection.Find.Execute
        .Application.Selection = Sheet1.Cells(r, 1).Value
        .Application.Selection.EndOf
        
        ' Replace «Position» with value from column E in the last row
        .Application.Selection.Find.Text = "«Position»"
        .Application.Selection.Find.Execute
        .Application.Selection = Sheet1.Cells(r, 5).Value
        .Application.Selection.EndOf
        
        ' Replace «Function» with value from column K in the last row
        .Application.Selection.Find.Text = "«Function»"
        .Application.Selection.Find.Execute
        .Application.Selection = Sheet1.Cells(r, 11).Value
        .Application.Selection.EndOf
        
        ' Replace «Location of Posting» with value from column F in the last row
        .Application.Selection.Find.Text = "«Location of Posting»"
        .Application.Selection.Find.Execute
        .Application.Selection = Sheet1.Cells(r, 6).Value
        .Application.Selection.EndOf
        
        ' Replace «DOJ» with value from column L in the last row
        .Application.Selection.Find.Text = "«DOJ»"
        .Application.Selection.Find.Execute
        .Application.Selection = Sheet1.Cells(r, 12).Value
        .Application.Selection.EndOf
        
        ' Replace «Joining Place» with value from column G in the last row
        .Application.Selection.Find.Text = "«Joining Place»"
        .Application.Selection.Find.Execute
        .Application.Selection = Sheet1.Cells(r, 7).Value
        .Application.Selection.EndOf
        
        ' Get the employee name from column A in the last row
        empName = Sheet1.Cells(r, 1).Value
        ' Replace spaces with underscore to avoid issues with file names
        empName = Replace(empName, " ", "_")
        
        ' Save the document with the employee name as the file name
        .ExportAsFixedFormat OutputFileName:=path & empName & "_Offer_Letter.pdf", _
            ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
            Range:=wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, IncludeDocProps:=True, _
            KeepIRM:=True, CreateBookmarks:=wdExportCreateNoBookmarks, _
            DocStructureTags:=True, BitmapMissingFonts:=True, UseISO19005_1:=False
    End With
    
    
    ' Open the Medical Annexure 2 template
Set wdoc2 = wApp.Documents.Add(template:=path & "\Medical_Annexure_2_template.dotx", NewTemplate:=False, DocumentType:=0)

' Replace placeholders with values from Excel sheet
With wdoc2
    ' Replace «Salutation» with value from column B in the last row
    .Application.Selection.Find.Text = "«Salutation»"
    .Application.Selection.Find.Execute
    .Application.Selection = Sheet1.Cells(r, 2).Value
    .Application.Selection.EndOf

    ' Replace «Emp Name» with value from column A in the last row
    .Application.Selection.Find.Text = "«Emp Name»"
    .Application.Selection.Find.Execute
    .Application.Selection = Sheet1.Cells(r, 1).Value
    .Application.Selection.EndOf

    ' Save the document with the employee name as the file name
    .ExportAsFixedFormat OutputFileName:=path & empName & "_Medical_Annexure_2.pdf", _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, IncludeDocProps:=True, _
        KeepIRM:=True, CreateBookmarks:=wdExportCreateNoBookmarks, _
        DocStructureTags:=True, BitmapMissingFonts:=True, UseISO19005_1:=False
End With
    
    ' Close the Word document
    wdoc2.Close savechanges:=False
    wdoc.Close savechanges:=False
       
   'At the end of OfferLetterCode, call CreateCompensationBreakup
    Call CreateCompensationBreakup
    
    
    ' Create an Outlook mail item and fill in the details
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate(ThisWorkbook.path & "\OfferLetterMailTemplate.oft")
    
    ' Get the employee email address from column M in the last row
    RecipientEmail = Sheet1.Cells(r, 13).Value
        
    ' Get the employee name, position and data of joining from column A in the last row
    empName = Sheet1.Cells(r, 1).Value
    Position = Sheet1.Cells(r, 5).Value
    DOJ = Sheet1.Cells(r, 12).Value
    'Replace the placeholder text with the recipient name and email
    OutMail.HTMLBody = Replace(OutMail.HTMLBody, "{Emp Name}", empName)
    OutMail.HTMLBody = Replace(OutMail.HTMLBody, "{Position}", Position)
    OutMail.HTMLBody = Replace(OutMail.HTMLBody, "{DOJ}", DOJ)
    OutMail.To = RecipientEmail
    
    ' Replace spaces with underscore to avoid issues with file names
    empName = Replace(empName, " ", "_")
    With OutMail
        .Subject = "Offer Letter - FIMER India Private Limited"
        .Attachments.Add path & empName & "_Offer_Letter.pdf"
        .Attachments.Add path & empName & "_Medical_Annexure_2.pdf"
        .Attachments.Add path & empName & "_Compensation_Breakup.pdf"
        .To = RecipientEmail
        .Display
    End With

End Sub
