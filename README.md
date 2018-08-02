This is an Verj IO function that recieves a JSON object as a parameter and will generate an Excel report based on the provided parameters. 

Features
--------
1. Create multiple sheets using an SQL statement or a table resource as the data source
2. Specify Formulas to add to a sheet including the option to add a formula for each row of a table
3. Add hyperlinks in a sheet
4. Add custom text anywhere that you specify

Installation
------------
1. Download the latest version of JExcel API from https://sourceforge.net/projects/jexcelapi/files/jexcelapi/ , extract the zip file and locate jxl.jar or use the provided jxl.jar.
2. Copy jxl.jar to VerjIO\UfsServer\tomcat\webapps\ufs\WEB-INF\lib
3. Restart Verj IO
4. Create shared JavaScript script
5. Add the following imports at the top of the script:

        importPackage(java.io);
	
        importPackage(Packages.jxl);
	
        importPackage(Packages.jxl.*);
	
        importPackage(Packages.jxl.write);
	
6. Paste the contents of createExcelReport.js into the script replacing the previous version code if it is there.

Usage
-----
In order to create an Excel report, you need to build a JSON object to tell the application how to create the Excel document.

Example:

     var excelReportObj = {
          FileName: "Sales Report",
          Sheets: [{ 
               SheetName: "Sales Report", 
               SheetIndex: 0,
               ColumnSize: [10,null,15,null,10,35], 
               ColumnHeaders: "Order Number,Order Line,Status,Customer,Part Num,Part Description",
               Columns: [["OrderNum","INTEGER"],["OrderLine","INTEGER"],["StatusDescription","CHAR"],["Customer","CHAR"],["PartNum","CHAR"],["PartDescription","CHAR"]],
               SQL: "SELECT OrderNum,OrderLine,Status,PartNum,PartDescription FROM SalesOrders",
               DBConnection: "PRODUCTION",
          },{ 
               SheetName: "Shipment Report", 
               SheetIndex: 1,
               ColumnSize: [20,10,12,null,20,15], 
               ColumnHeaders: "Shipment Num,Shipment Date,Shipment Qty",
               Columns: [["ShipNum","INTEGER"],["ShipDate","DATE"],["ShipQty","INTEGER"]],
               TableData: "ShipmentS",
          },
          ]
     };
     
     var result=createExcelReport(excelReportObj);

     if (result[0]=="ERROR") {
	        return result[1];
	        event.stopExecution();
     }
     
     print ("The file name is " + result[1]);

In the example above, the generated Excel file will be named "Sales Report as of 08-01-2018-09-20-10.xls". The date will automatically be added to the file name to make sure that each report always has a unique name.

The first sheet will be named Sales Report with 6 columns. The Columns property specify the database column names and their type. This sheet uses an SQL statement as the data source

The second sheet will be named Shipment Report and have the 3 columns specified. This sheet uses a table resource as the data source.

createExcelReport() returns an array

result[0] is the status OK or ERROR

result[1] is the filename if result[0] is OK or the error message if result[0] returns ERROR


General tips:

See the provided JSON template for all of the options

If the number of columns specified in ColumnHeaders doesn't match the number of arrays in Columns, the application will return an error. They must have the same number of columns.

When an SQL statement is provided as a data source, you must also provide the name of the database connection. 

ColumnSize is optional and can be used to specify the column width. If you do not provide ColumnSize, all columns will be set to autosize the width automatically. If you provide an array with null for any values like in the example above, that column will also be set to autosize the width.

Valid data types are BOOLEAN, CHAR, CURRENCY, DATE, DATETIME, INTEGER and NUMERIC
