VerjIO_Excel_Report_Generator

Installation
------------
1. Install JExcel API
   a. Use the provided jxl.jar or
   b. Download  the latest version from https://sourceforge.net/projects/jexcelapi/files/jexcelapi/ extract the zip file and         locate jxl.jar
2. Copy jxl.jar to VerjIO\UfsServer\tomcat\webapps\ufs\WEB-INF\lib
3. Restart Verj IO
4. Create shared script 
5. Add these imports at the top of the script:
   importPackage(java.io);
   importPackage(Packages.jxl);
   importPackage(Packages.jxl.*);
   importPackage(Packages.jxl.write);
6. Paste the latest version of the createExcelReport function replacing the pervious version

Usage
-----
In order to create an Excel report, you need to build a JSON object to tell the application how to create the Excel document.

Example 1:

 var excelReportObj = {
          FileName: "Sales Report",
          Sheets: [{ 
               SheetName: "Sales Report", 
               SheetIndex: 0,
               ColumnSize: [10,null,15,null,10,35], 
               ColumnHeaders: "Order Number,Order Line,Status,Customer,Part Num,Part Description",
               Columns: [["OrderNum","INTEGER"],["OrderLine","INTEGER"],["StatusDescription","CHAR"],["Customer","CHAR"],["PartNum","CHAR"],["PartDescription","CHAR"]],
               SQL: "SELECT OrderNum,OrderLine,Status,PartNum,PartDescription FROM Sales Orders",
               DBConnection: "PRODUCTION",
          },
          ]
     };
     
     var result=createExcelReport(excelReportObj);

     if (result[0]=="ERROR") {
	        return result[1];
	        event.stopExecution();
     }

In the example above, the generated Excel file will be named "Sales Report as of 08-01-2018-09-20-10.xls". The date will automatically be added to the file name to make sure that each report always has a unique name.

This Excel file will have 1 sheet named "Sales Report" with 6 columns. The Columns property specify the database column names and their type. Valid types are BOOLEAN, CHAR, CURRENCY, DATE, DATETIME, INTEGER and NUMERIC

ColumnSize is optional and can be used to specify the column width. If you do not provide ColumnSize, all columns will be set to autosize the width automatically. If you provide an array with null for any values like in the example above, that column will also be set to autosize the width.

The data source in the example above uses an SQL statement. When an SQL statement is provided as a data source, you must also provide the name of the database connection. 

createExcelReport() returns an array

result[0] is the status OK or ERROR
result[1] is the filename if result[0] is OK or the error message if result[0] returns ERROR



