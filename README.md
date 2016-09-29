# Hello-World
My First Respository
I will add some changes,and these changes will commit to the master branch!!!!!!!!


Parameters
Name
Required/Optional
Data Type
Description
TransferType
Optional
AcDataTransferType
The type of transfer you want to make. The default value is acImport.
SpreadsheetType
Optional
AcSpreadSheetType
The type of spreadsheet to import from, export to, or link to.
TableName
Optional
Variant
A string expression that is the name of the Microsoft Office Access table you want to import spreadsheet data into, export spreadsheet data from, or link spreadsheet data to, or the Access select query whose results you want to export to a spreadsheet.
FileName
Optional
Variant
A string expression that's the file name and path of the spreadsheet you want to import from, export to, or link to.
HasFieldNames
Optional
Variant
Use True (?1) to use the first row of the spreadsheet as field names when importing or linking. Use False (0) to treat the first row of the spreadsheet as normal data. If you leave this argument blank, the default (False) is assumed. When you export Access table or select query data to a spreadsheet, the field names are inserted into the first row of the spreadsheet no matter what you enter for this argument.
Range
Optional
Variant
A string expression that's a valid range of cells or the name of a range in the spreadsheet. This argument applies only to importing. Leave this argument blank to import the entire spreadsheet. When you export to a spreadsheet, you must leave this argument blank. If you enter a range, the export will fail.
UseOA
Optional
Variant
This argument is not supported.
Remarks
You can use the TransferSpreadsheet method to import or export data between the current Access database or Access project (.adp) and a spreadsheet file. You can also link the data in a Microsoft Excel spreadsheet to the current Access database. With a linked spreadsheet, you can view and edit the spreadsheet data with Access while still allowing complete access to the data from your Excel spreadsheet program. You can also link to data in a Lotus 1-2-3 spreadsheet file, but this data is read-only in Access.
