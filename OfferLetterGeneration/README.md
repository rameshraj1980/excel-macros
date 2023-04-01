The code is written in VBA and is intended to automate the process of sending offer letter, Compensation Breakup letter and medical annexure to new employees. It uses a Word template for offer letter and medical annexure, excel template for Compensation Breakup letter and an Outlook email template to send the offer letter as an email attachment to the new employee. The new employee details entered in the 'Offer_Data_Nos' will be populated in the Word, Excel and Outlook email templates and displayed to the user for verifying and sending the mail.

**How to create placeholders & Word template?**

Here are the steps to create a placeholder using MERGEFIELD:

•	Open the Word document where you want to insert the placeholder.

•	Place the cursor where you want to insert the placeholder.

•	Go to the "Mailings" tab in the ribbon.

•	Click on the "Insert Merge Field" button in the "Write & Insert Fields" group.

•	In the "Insert Merge Field" dialog box, type a name for the placeholder (e.g. "Name") in the "Field name" field.

•	Click "OK" to insert the placeholder into the document.

•	Repeat these steps for each placeholder you want to create in the document, using a unique name for each placeholder.

•	Click on the "File" tab in the ribbon and choose "Save As".

•	In the "Save As" dialog box, select "Word Template" in the "Save as type" dropdown menu.

•	Choose a name for your template and select a location to save it.

•	Click "Save" to save the template.

Note: Please check "Offer_Letter_Template.dotx" for reference.

**How to create placeholders & mail template?**

•	Open a new email message in Outlook.

•	Customize the message with the desired subject, body, formatting, and attachments.

•	Add any merge fields or placeholders you want to use in your VBA code, such as {RecipientName} or {Date}.

•	Click on "File" in the ribbon.

•	Click on "Save As".

•	Select "Outlook Template" as the file type from the dropdown list.

•	Choose a location to save the template, and give it a name (for example, "My Email Template").

•	Click on "Save".

Note: Please check "OfferLetterMailTemplate.oft" for reference.

**Here's what the code "DraftOfferLetterMail" does:**

1.	Declares necessary variables

2.	Finds the last row of data in the Excel sheet

3.	Creates a new Word application

4.	Opens the offer letter template and replaces placeholders with values from the Excel sheet

5.	Saves the offer letter document as a PDF file with the employee name as the file name

6.	Opens the medical annexure template and replaces placeholders with values from the Excel sheet

7.	Saves the medical annexure document as a PDF file with the employee name as the file name

8.	Closes the Word documents

9.	Calls the "CreateCompensationBreakup" subroutine

10.	Creates an Outlook mail item and fills in the details

11.	Gets the employee email address from the Excel sheet

12.	Gets the employee name, position, and date of joining from the Excel sheet

13.	Replaces placeholders in the email template with the employee details

14.	Attaches the offer letter and medical annexure PDF files to the email

15.	Sends the email

The code assumes that there is a Word template named "Offer_Letter_Template.dotx" and "Medical_Annexure_2_template.dotx" in the same folder as the Excel workbook and an Outlook email template named "OfferLetterMailTemplate.oft" in the same folder as the Excel workbook.

**Here's what the code "DraftOfferLetterMail" does:**

The CreateCompensationBreakup() subroutine creates a PDF file of a compensation breakup sheet for the latest employee in the data sheet.

The subroutine first defines several variables, including two workbooks and their respective worksheets, a file path to the template file and the data file, and several variables that will hold the latest employee's data (employee name, offer number, gross CTC, and percentage bonus).

Next, the file paths to the template and data files are set, and the template workbook is opened, and a reference to the worksheet is obtained. The data workbook is then opened, and a reference to its worksheet is obtained.

The latest values of the employee's data are obtained by using the xlUp method to find the last row with data in each of the four columns that hold the employee data. These values are then inserted into the template worksheet.

The employee's name is then modified to replace spaces with underscores to avoid issues with file names. The template worksheet is then saved as a PDF file using the ExportAsFixedFormat method with the employee name included in the file name.

Finally, the template workbook is closed without saving any changes.

Note: This code is just an example and may need to be modified based on your specific requirements. It is also important to thoroughly test the code before using it in a production environment.
