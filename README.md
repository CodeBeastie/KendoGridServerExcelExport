# Kendo grid server side Excel Export


To import excel files the openXML can be used. It will work with latest office docs such as .xlsx but not the older .xls files.
Can use NPOI if you need to support the older formats.

The openXML 2.5 can be found in NuGet as DocumentFormat.OpenXml
Created by Microsoft but  described as unofficial SDK V2.5



Structure notes:
Spreadsheet document
  Workbookpart
    Workbook
      Sheets
        Sheet -- holds relationship to worksheetpart
  Worksheetpart -- hooks on to workbookpart
    worksheet



References:

Structure of Excel document:
https://msdn.microsoft.com/EN-US/library/gg278316.aspx

https://msdn.microsoft.com/en-us/library/bb448854.aspx

http://www.codeproject.com/Articles/670141/Read-and-Write-Microsoft-Excel-with-Open-XML-SDK
https://msdn.microsoft.com/en-us/library/hh298534.aspx
http://www.codeproject.com/Articles/371203/Creating-basic-Excel-workbook-with-Open-XML

Need to add WindowsBase reference as well for the excelreader/writer code in use. For the System.IO.MemoryStream in use



Requirements:
Visual Studio 2015
