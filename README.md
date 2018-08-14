VerjIO_Excel_Report_Generator is a Verj IO function that recieves a JSON object as a parameter and will generate an Excel report based on the provided JSON "recipe" to build the report. A sample JSON object is provided below. The report is saved as an XLS that should be  backwards compatible with Excel 2010 and newer.

Features
--------
1. Create multiple sheets with a table in each sheet using an SQL statement or a table resource as the data source for each sheet
2. Add formulas to each sheet including the option to specify that a formula should be written for each row of a table
3. Specify cells to merge
4. Add hyperlinks in each sheet
5. Add custom text anywhere that you specify
6. Add heading text
7. Add custom styling
8. Re-useable named style definitions

Installation
------------
1. Download the latest version of JExcel API from https://sourceforge.net/projects/jexcelapi/files/jexcelapi/ , extract the zip file and locate jxl.jar or use the provided jxl.jar.
2. Copy jxl.jar to VerjIO\UfsServer\tomcat\webapps\ufs\WEB-INF\lib
3. Restart Verj IO
4. Create shared JavaScript script
5. Add the following imports at the top of the script:

        importPackage(java.io);
	
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

The first sheet will be named Sales Report and will have a table with the 6 columns headings specified in ColumnHeaders. The Columns property specify the database column names and their type. This sheet uses an SQL statement as the data source.

The second sheet will be named Shipment Report and have a table with the 3 columns specified. This sheet uses a table resource as the data source.

createExcelReport() returns an array with 2 indexes

result[0] is the status which can be "OK", "OK-NODATA" or "ERROR"

result[1] is the filename if result[0] is OK, the error message if result[0] returns ERROR or empty if result[0] returned OK-NODATA which means that there wasn't any data to write to the report.

General Tips
------------
1. The number of columns specified in ColumnHeaders must match the number of arrays provided in Columns or an error will be thrown.

2. When an SQL statement is provided as a data source, you must also provide the name of the database connection or an error will be thrown. Since database resources in Verj IO have a link to the database connection name and tables are linked to a database resource, this isn't needed for table data.

3. ColumnSize is optional and can be used to specify the column width. If you do not provide ColumnSize, all columns will be set to autosize the width automatically. If you provide an array with null for any values like in the example above, the columns that have null for the width will be set to autosize the width automatically.

4. Valid data types are "BOOLEAN", "CHAR", "CURRENCY", "DATE", "DATETIME", "INTEGER" (or "INT" as an alias for INTEGER") and "NUMERIC". If you specify that a column data type is CHAR but the value is an INTEGER, it will be written as an INTEGER. IF you want to force the column to be written as CHAR, add a true parameter. Ex. ["SOMECOLUMN","CHAR",true]. If the data type is specified as INT but the data value is a CHAR, it will be written as a CHAR.

5. Valid color values can be found in the object colorObject defined in createStyleFormat(). You can also supply an RGB string ike "83,162,240" as the color value to use a specific RGB value instead of a color name.

6. Valid border styles can be found in the object borderStylesObject defined in createStyleFormat(). 

7. Valid underline styles can be found in underlineStyleObject defined in createStyleFormat().

8. Valid alignment styles can be found in alignmentStylesObject defined in createStyleFormat().

9. When checking the return value of createExcelReport(), if the call to createExcelReport() is made from a client side function, the result array returned by createExcelReport() must be passed down to the client control that called the server side function. Otherwise, the user will not see the result message. In the server side function that gets called from the client, the call to createExcelReport() might look like this:
       
       // Part of the logic from the server side function generateKPIReport() that gets called from the client side
       var result=createExcelReport(excelReportObj);
       
       if (result[0]=="ERROR" || result[0]=="OK-NODATA") {
            return result;
        }
	
	   // Do something like email the report here if you want
	   
	   return ["OK","The report has been generated"];
	   } // End of generateKPIReport() server side shared function
	   
      In the HTML entities for the control that calls createExcelReport() directly or calls another server side function which calls createExcelReport() you would need to have something like this:

       var result=$eb.executeFunction("generateKPIReport",reportYear,false,true); // Call from client to server side function that calls createExcelReport()
      
       if (result[0]=="ERROR" || result[1]=="OK")
            alert(result[1]);
       else if (result[0]=="OK-NODATA")
            alert("There is no data available for the report");

10. If the call to createExcelReport() in a script that is run in a before form event, you should print the result if an error occurrs.
       
            var result=createExcelReport(excelReportObj);	
	    
	    if (result[0]=="ERROR")
                 print(result[1]);
            else if (result[0]=="OK-NODATA")
                 print("There is data available for the report");

11. If the call to createExcelReport() is made in a server side script, you can check the return value as follows:
     
            var result=createExcelReport(excelReportObj);	
	    
	    if (result[0]=="ERROR") {
                 alert(result[1]);
		 event.stopExecution();
            } else if (result[0]=="OK-NODATA")
                 alert("There is data available for the report");
		 event.stopExecution();
            else if (result[0]=="OK")
	         alert(result[1]);
12. If you have a table with a formula in one of the columns and want to have a table heder, you can add [null] as a place holder when specifying the column. If you do not do this, the number of column headers and columns will not match and the application will display an error message.

13. 

JSON Reference
--------------
The report object below contains all of the possible properties that you can pass to createExcelReport()
but not all of them are required.

       var reportObj = {
            FileName: "EOL Report",
            NamedStyles: [{ // Optional named style
                 Name: "Heading",
                 Color: "white",
                 BackgroundColor: "red"
            },
            Sheets: [{ 
                 StartRow: 3, // Start on row 3
                 FitWidth: true, // Optional. Will force sheet to fit all columns on the page width-wise
                 SheetName: "EOL Parts",
                 SheetIndex: 0,
                 SheetHeader: [{ // Add heading in Cell D1
                      Value: "EOL Parts Report as of " + getCurrentDate(),
                      Column: 3,
                      Row: 0,
		      MergeCells: "3,5", // Optional. Merges columns 3 thru 5 on the same row. You can also specify col,row,col,row to merge across more than 1 row
                  NamedStyle: "Heading", // Used named style definition. Don't use NamedStyle and Style together. Use only one
		      Style: [{ // Optional style sub object
                       Alignment: "left", // Optional. defaults to General if not specified 
                       Color: "WHITE", // Optional. Defaults to black if not specified
                       Size: "14", // Optional. Defaults to 12 if not specified
                       BackgroundColor: "green", // Optional. Defaults to white if not specified
                       Bold: true, // Optional. Defaults to false if not specified
                       Italic: true, // Optional. Defaults to false if not specified
                       Underline: true, // Optional. Defaults to false if not specified
                       UnderlineStyle: "single", // Optional/ Defauts to single if not specified
                       Borders: true, // Optional. Defaults to false if not specified
                       BorderStyle: "THICK", // Optional. Defaults to THIN if not specified but Borders: true is specified
               }],
               }], 
               ColumnSize: [20,10,null,15], // (Optional) Use null if you want to auto size         
               ColumnHeaders: "ID, Part Num,Part Description,Mfg Part Num,Ship Date,Status,Close Date",
               Columns: [["EOLID","INTEGER"],["PartNum","CHAR"],["PartDescription","CHAR"],["MfgPartNum","CHAR"],["LastShipDate","DATE"],["Status","CHAR"],["CloseDate","DATE"]],
	       MergeCells: ["1,2","3,1,4,1"], // Optional. Merges columns 1 & 2 on the last row after data has been writtem. The second merge merges columns 3 & 4 on row 1. 
               SQL: "SELECT * FROM EOL", // SQL based data
               DBConnection: "PRODUCTION", // (Mandatory if SQL statement is provided
          
               Formulas: [{
                    Column: 0, // Column A since columns start with 0
                    Row: 10, // Optional. Will be written after last row of data if not specified               
                    Formula: "SUM(A1:A<CURRENTROW>)", // Use <CURRENTROW> as a placeholder for the current row number
                    DataType: "INTEGER",
                    FormulaRowOffset: -1, // Optional) offset the value of <CURRENTROW>. So you could use -1 to to refer to CURRENTROW-1 
                    LineFormula: true, // (Optional flag to indicate that this formula needs to be written for each row)
               },],
               HyperLinks: [{
                    Column: 0, // Column A since columns start with 0
                    Row: 10, // Optional. Will be written after last row of data if not specified               
                    Value: "http://www.google.com",
                    DestinationSheet: "sheet2",
                    DestinationColumn: 2, // Column C since columns start with 0
                    DestinationRow: 2, // Row 3 since rows start with 0
               },],
            ],
            CustomCellText: [{
                 DestinationSheet: "sheet1", // (Optional)
                 Column: 0, // Column A since columns start with 0
                 Row: 10, // Optional. Will be written after last row of data if not specified,
                 Value: "Some Text",
                 DataType: "CHAR", // Optional. Defaults to CHAR if not specified
                 MergeCells: ["0,10,3,10"], // Optional. If specified, you must specify start_column,start_row ,end_column,end_row
		 Style: [{ // Optional style sub object
		          Alignment: "center", // Optional. defaults to General if not specified
                      Color: "white", // Optional. Defaults to black if not specified
                      Size: "14", // Optional. Defaults to 12 if not specified
                      BackgroundColor: "green", // Optional. Defaults to white if not specified
                      Bold: true, // Optional. Defaults to false if not specified
                      Italic: true, // Optional. Defaults to false if not specified
                      Underline: true, // Optional. Defaults to false if not specified
                      UnderlineStyle: "single", // Optional/ Defauts to single if not specified
                      Borders: true, // Optional. Defaults to false if not specified
                      BorderStyle: "THICK", // Optional. Defaults to THIN if not specified but Borders: true is specified
                 }],
                 },],
            };

Known Issues
------------
Using RGB color as background does not always work correctly.
