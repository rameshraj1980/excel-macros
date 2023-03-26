# excel-macros
**Overview:**

The CreateRevenueReport macro automates the process of creating a new worksheet called "Revenue Report" and copying selected columns from the "Master" worksheet to the new worksheet. 

**Requirements:**

This macro can be used in Microsoft Excel 2007 or later versions.

The Excel workbook must contain a worksheet named "Master" with revenue data in columns A to U and AI to AL.

The steps to add the Developer ribbon in Excel:

Open Excel and click on the File tab in the top left corner.

Click on Options in the left menu to open the Excel Options dialog box.

Click on Customize Ribbon in the left menu.

In the right pane, under the Customize the Ribbon section, check the box next to Developer.

Click on OK to save the changes and close the Excel Options dialog box.

The Developer ribbon should now be visible in the Excel ribbon at the top of the screen.

Note: If you cannot see the Developer tab after following the above steps, it may be because it is hidden. You can try unhiding it by going to File > Options > Customize Ribbon > Customize the Ribbon section > Reset > Reset all customizations. This will reset the ribbon to its default settings, which should include the Developer tab.

**How to use the macro:**

Open the Excel workbook containing the "Master" worksheet.

Press ALT + F11 to open the VBA editor.

In the VBA editor, insert a new module or open an existing one.

Copy and paste the code for the CreateRevenueReport macro into the module.

Save the VBA module and close the VBA editor.

Return to the Excel workbook and press ALT + F8 to open the Macros dialog box.

Select the CreateRevenueReport macro from the list and click Run.

The macro will create a new worksheet called "Revenue Report" and copy selected columns from the "Master" worksheet to the new worksheet.
The new worksheet will be activated and the columns will be auto-fit to fit the data.

Note: It is important to save the Excel workbook after running the macro to ensure that the new worksheet is saved along with the workbook.

**Code Explanation:**

The CreateRevenueReport macro uses the following variables and steps:

Dim wsMaster As Worksheet: creates a worksheet variable for the "Master" worksheet.

Dim wsRevenueReport As Worksheet: creates a worksheet variable for the new "Revenue Report" worksheet.

Dim lastRow As Long: creates a variable for the last row of data in the "Master" worksheet.

Set wsMaster = ThisWorkbook.Worksheets("Master"): sets the wsMaster variable to reference the "Master" worksheet.

Set wsRevenueReport = ThisWorkbook.Worksheets.Add: creates a new worksheet and sets the wsRevenueReport variable to reference it.

wsRevenueReport.Name = "Revenue Report": renames the new worksheet to "Revenue Report".

lastRow = wsMaster.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row: finds the last row of data in the "Master" worksheet and sets the lastRow variable to its row number.

wsMaster.Range("A1:U" & lastRow & ", AI1:AL" & lastRow).Copy Destination:=wsRevenueReport.Range("A1"): copies selected columns from the "Master" worksheet to the new "Revenue Report" worksheet.

wsRevenueReport.Columns.AutoFit: auto-fits the columns in the new "Revenue Report" worksheet.

wsRevenueReport.Activate: activates the new "Revenue Report" worksheet.
