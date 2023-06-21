1. Download the ‘PERSONAL.XLSB’ file and copy it to: C:\Users\(USER_NAME)\AppData\Roaming\Microsoft\Excel\XLSTART

2. Download Ortho Audit Template and place it in the Desktop

3. Open excel and click the customize quick access toolbar

4. Click on More Commands…

5. Click on Popular commands

6. Click on Macros

7. Add All Macros

8. Click Modify for each Macro so you can differentiate between them

9. Click OK

Notes:

These macros are designed to only work for the Appointment Detail Report xlsx file from Denticon reports. It will not work with other files.

“FilterORTHOC” reformats and filters for orthochair and orthoconsults then filters again from smallest to largest chart number. It then copies the ‘Template’ worksheet from your desktop and puts it into the active workbook.

“AddSheet” duplicates the template worksheet in the workbook and renames it for each WCD chart number available in the filtered table.

“Next_day_Audit_macro” reformats and filters for orthochair and orthoconsults then adds vlookup for checking accounts with remaining balance in WCD and checking for audits that were recently added. Copying the chart numbers into the new sheets is still needed. Chart numbers column still need to be copied from contract collections and apptdetails and into the designated worksheets
	
