# VerjIO_Excel_Report_Generator

VerjIO_Excel_Report_Generator is a Verj IO function that will automatically create an Excel report based on an SQL statement or a Verj IO Table. To use it, you need to create a JSON object as a parameter which is a JSON "recipe" to build the report. A sample JSON object is provided below. The report is saved as an XLSX format Excel document.

## Features
1. Create multiple sheets with a table in each sheet using an SQL statement or a table resource as the data source for each sheet
2. Add formulas to each sheet including the option to specify that a formula should be written for each row of a table
3. Merge cells
4. Add hyperlinks in each sheet
5. Add custom text anywhere that you specify
6. Many custom style options
7. Re-useable named style definitions
8. Password protect a sheet
9. Specify conditional formatting rules.
10. Insert images into a sheet.
11. Freeze the first row or column of the sheet
12. Supports nested tables (currently supports SQL based data and not Table data)
13. Create pivot table

## Initial Installation
1. Download the latest version of [Apache POI](https://poi.apache.org/), extract the zip file and locate the following JAR files: poi-X.jar,poi-excelant-X.jar, poi-ooxml-X.jar and poi-ooxml-schemas-X.jar where X is the current version number.

2. Download the latest version of [commons-collections](https://commons.apache.org/proper/commons-collections/), extract the zip and locate commons-collectionsX.jar where X is the current version number.

3. Download the latest version of [commons-compress](https://commons.apache.org/proper/commons-compress/download_compress.cgi), extract the zip and locate commons-compress-X where X is the current version number.

4. Download the latest version of [XML Beans](https://xmlbeans.apache.org/download/index.html#XMLBeans-3.0.1), extract the zip and locate the jar xbean.jar.

5. Copy all the jar files that you extracted above to VerjIO's lib folder. On Verj IO 5.6+, the path is C:\VerjIOData\apps\ufs\webapp\WEB-INF\lib. On versions of Verj IO before 5.6, the path is  VerjIO\UfsServer\tomcat\webapps\ufs\WEB-INF\lib. You may have 2 different versions of commons-collections. This is ok.

6. Restart Verj IO

7. Create shared JavaScript script

8. Add the following imports at the top of the script:
   - importPackage(java.io);
   - importPackage(Packages.org.apache.poi);
   - importPackage(Packages.org.apache.poi.hssf.usermodel);
   - importPackage(Packages.org.apache.poi.xssf.usermodel);
   - importPackage(Packages.org.apache.poi.ss.usermodel);
   - importPackage(Packages.org.apache.poi.hssf.util);
	
9. Paste the contents of createExcelReport.js into the script.

## Upgrading VerjIO_Excel_Report_Generator

If you are using POI 3.X jar files, you need to update POI to 4.0 or higher using the instructions above. If you try to use the newest version of createExcelReport with POI 3.X jar files, you will most likely run into errors or issues because there are some breaking changes between POI 3 and 4.

If you are already using POI 4.X jar files, you only need to replace createExcelReport() with the newer version.

## Usage
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
     
     print ("The file name is " + result[1]);`

In the example above, the generated Excel file will be named "Sales Report as of 11-16-2020-09-20-10.xlsx". The date & time will automatically be added to the file name to make sure that each report always has a unique name.

The generated document will have 2 sheets. 

The first sheet will be named Sales Report and will have a table with the 6 columns headings specified in ColumnHeaders. ColumnSize is used to set the size of each column (specifying null for a column size will set that column to autosize). The Columns property specify the database column names and their type. This sheet uses an SQL statement as the data source with the specified database connection called PRODUCTION.

The second sheet will be named Shipment Report and have a table with the 3 columns specified. ColumnSize is not provided so all columns will auto size automatically. This sheet uses a table resource as the data source instead of an SQL statement.

createExcelReport() will return a 2 dimensional array with the results

The return value will be one of the following:

["OK","Filename.xlsx"] // Report was successfully created and the file was saved as Filename.xlsx
["OK-NODATA",""] // There was no data returned from the SQL query or table resource so the report was not generated
["ERROR","Error Message"] // An error occurred during the creation of the report with the exact message provided

Example 2: Nested tables
    
    var reportObj = {
          FileName: "TEST",
          Sheets: [{ 
               SheetName: "TEST",
               SheetIndex: 0,
               NestedTables:true,   
               ColumnHeadersParent: "User ID,Name,User name",
               ColumnsParent: [["UserID","CHAR"],["RealName","CHAR"],["Username","CHAR"]],
               SQLParent: "SELECT RealName,Username,UserID FROM Users",
               DBConnectionParent: "ParentDBCOnnection",
               ColumnHeadersChild: "Menu Name,Menu Description",
               ColumnsChild: [["MenuName","CHAR"],["MenuDescription","CHAR"]],
               SQLChild: "SELECT UserID,MenuName,MenuDescription FROM Menus_Auth",
               DBConnectionChild: "ChildDBCOnnection",
               JoinWhereClause: ["UserID",0], // Index 0 = column that links the 2 queries and 0 is the index of the column starting with 0 for the first column
               ChildIndent: 1 
          },
          ]
     };
     var result=createExcelReport(reportObj);

In the example above, the generated Excel file will be named "Test as of 11-16-2020-09-20-10.xlsx". 
When using nested tables, you have to specify the column headers, columns, SQL and DB Connection for the parent and child tables. In addition you need to specify JoinWhereClause as an array with 2 values; the name of the column that joins the parent and child tables and the index in the parent column headers (ColumnsParent). If you want to shift the child table to the right (which I recommend because it makes it easier to read the data) you can set ChildIndent/

## General Tips
1. The number of columns specified in ColumnHeaders must match the number of arrays provided in Columns or an error will be thrown.

2. When an SQL statement is provided as a data source, you must also provide the name of the database connection or an error will be thrown. Since Verj IO tables are already linked to a database connection, you don't need to specify the database connection name when providing a table source.

3. ColumnSize is optional and can be used to specify the column width. If you do not provide ColumnSize, all columns will be set to autosize the width automatically. If you provide an array with null for any values like in the example above, the columns that have null for the width will be set to autosize the width automatically.

4. Valid data types are "BOOLEAN", "CHAR", "CURRENCY", "DATE", "DATETIME", "TIME", "INTEGER" (or "INT" as an alias for INTEGER") and "NUMERIC". If you specify that a column data type is CHAR but the value that is going to be written to the Excel document is an INTEGER, it will be written as an INTEGER. IF you want to force the column to always be written as CHAR, add a true parameter after the CHAR data type. Ex. ["SOMECOLUMN","CHAR",true]. If the data type is specified as INT but the data value is not an INT, it will be written to the Excel document as a CHAR unless you add a true parameter. Ex ["SOMECOLUMN","INT",true].

5. Valid color values can be found in the object colorObject defined in getColor(). You can also supply an RGB string ike "83,162,240" as the color value to use a specific RGB value instead of a predefined color.

6. Valid border styles can be found in the object borderStylesObject defined in createStyleFormat(). 

7. Valid underline styles can be found in underlineStyleObject defined in createStyleFormat().

8. Valid alignment styles can be found in alignmentStylesObject defined in createStyleFormat().

9. If you are calling createExcelReport() in a client callable function, you cannot alert the result using event.owner.addWarningMessage on the server side. Instead you must return the warning message to the script that called the client side function that executes createExcelReport(). In the server side function that gets called from the client, the call to createExcelReport() might look like this:
       
       // Part of the logic from the server side function generateKPIReport() that gets called from the client side
       var result=createExcelReport(excelReportObj);
       
       if (result[0]=="ERROR" || result[0]=="OK-NODATA") {
            return result;
        }
	
	   // Do something like email the report here if you want
	   
	   return ["OK","The report has been generated"];
	   }
	   
      In the HTML entities of the control that called the client side function, you would need to have something like this:

       var result=$eb.executeFunction("generateKPIReport",reportYear,false,true); // Call from client to server side function that calls createExcelReport()
      
       if (result[0]=="ERROR"")
            alert(result[1]);
       else if (result[0]=="OK-NODATA")
            alert("There is no data available for the report");
       else if (result[0]=="OK")
            alert("The name of the file is " + result[1]);

10. If the call to createExcelReport() in a script that is run in a before form event, you should print the result if an error occurs using print instead of using event.owner.addErrorMessage or event.owner.addWarningMessage because you cannot initiate these types of alerts in a Before Form event.
       
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
12. If you have a column with a formula for each row, you  can use a line formula to specify that a formula is to be written for each row. When doing this, you will have a column header such as Total in ColumnHeader. In order to make sure that the number of columns and column headers match, put [null] as a place holder in the Columns property in place of an actual column name. Ex.

        Columns: [["OrderNum","INTEGER"],[null],["StatusDescription","CHAR"],["Customer","CHAR"],["PartNum","CHAR"],["PartDescription","CHAR"]], 
    ColumnHeaders: "Order Number,Order Count,Status,Customer,Part Num,Part Description",

13. When using a TableData source, the column names are case sensitive and must be provided in the same case that they appear in the table resource or this script will throw an error because the coulmn could not be located.

14. If you are using custom formatting, make sure to use a formula that evaluates to a boolean true or false value. Use $ in front of the column letter to allow the formula to be evaluated correctly for each row of a table.

15. If you open an Excel document that has a pivot table inside of it from an email client like Outlook, you have to manually click on the pivot table and choose Refresh.

JSON Reference
--------------
The report object below contains all of the possible properties that you can pass to createExcelReport()
but not all of them are required. All of the optional properties have a comment indicating that they are optional.

     `var reportObj = {
          FileName: "EOL Report",
          NamedStyles: [{ // Optional named style. All of the style options are listed below
          Name: "Heading",
               Alignment: "left", // Optional. defaults to General if not specified
               BackgroundColor: "green", // Optional. Defaults to white if not specified
               Bold: true, // Optional. Defaults to false if not specified
               Borders: true, // Optional. Defaults to false if not specified
               BorderStyle: "THICK", // Optional. Defaults to THIN if not specified but Borders: true is specified
               Color: "WHITE", // Optional. Defaults to black if not specified
               Italic: true, // Optional. Defaults to false if not specified
               Rotation: 90 ,// Optional. Specify in degrees. Valid values are from 0 to 180 degrees
               Size: "14", // Optional. Defaults to 12 if not specified
               Strikeout: true // puts horizontal line through the text
               Underline: true, // Optional. Defaults to false if not specified
               UnderlineStyle: "single", // Optional/ Defauts to single if not specified
               VerticalAlignment: "top" // Optional. Valid
               Wrap: true, // Optional. Forces text to wrap around to the next line
          },],
          Sheets: [{ 
               StartRow: 3, // Optional: Start on row 3
               SheetName: "EOL Parts",
               SheetIndex: 0,
               TopMargin: 1, // Optional top margin
               BottomMargin: 1, // Optional bottom margin
               LeftMargin: 1, // Optional left margin
               RightMargin: 1, // Optional right margin
               AllMargins: "1", // Optional. Providing 1 value applies the margin to all 4 sides
               AllMargins: "1,0.5", // Optional. Providing 2 values applies the first value to the top & bottom and the second value to the left and right margins
               AllMargins: "1,0.5,1.5,2", // Optional. Providing 4 values applies the 1st value to the top, 2nd value to the bottom , 3rd value to the left and 4th value to the right margin
               HeaderMargin: 1, // Optional header margin
               FooterMargin: 1, // Optional footer margin
               HeaderLeft: "Left header string goes here", // Optional left header
               HeaderCenter: "Center header string goes here", // Optional center header
               HeaderRight: "Right header string goes here", // Optional right header
               FooterLeft: "Left footer string goes here", // Optional left footer
               FooterCenter: "Center footer string goes here", // Optional center footer
               FooterRight: "Right footer string goes here", // Optional right footer
               FitWidth: true, // Optional. Will force sheet to fit all columns on the page width-wise
               FitHeight: true, // Optional. Will force sheet to fit all columns on the page height-wise
               FitToPages: true, // Optional. Will force sheet to fit all of the content into 1 page
               Orientation: "landscape", // Optional. Valid values are "landscape" or "portrait"
               Password: "somepassword": // Optional. Will protect the sheet from being edited unless the user enters this password
               PivotTable: true // Optional. Creates a pivot table based off of the data in the current sheet
               NestedTables: true, // Optional. This is only required if you are using nested tables
               ColumnHeadersParent: "UserID,RealName,Username", // Optional. This is only required if you are using nested tables. Parent table column headers when using nested tables
               ColumnsParent: [["UserID","CHAR"],["RealName","CHAR"],["Username","CHAR"]], // Optional. This is only required if you are using nested tables. Parent table columns when using nested tables
               SQLParent: "SELECT RealName,Username,UserID FROM Users", // Optional. This is only required if you are using nested tables. Parent SQL
               DBConnectionParent: "ParentDBCOnnection",  // Optional. This is only required if you are using nested tables. Parent DB connection for the SQL statement used above
               ColumnHeadersChild: "Menu Name,Menu Description", // Optional. This is only required if you are using nested tables. Child table column headers when using nested tables
               ColumnsChild: [["MenuName","CHAR"],["MenuDescription","CHAR"]], // Optional. This is only required if you are using nested tables. Child table columns when using nested tables
               SQLChild: "SELECT UserID,MenuName,MenuDescription FROM Menus", // Optional. This is only required if you are using nested tables. Child SQL
               DBConnectionChild: "GENESIS_INTRANET", // Optional. This is only required if you are using nested tables. Child DB connection for the SQL statement used above
               JoinWhereClause: ["UserID",0], // Optional. This is only required if you are using nested tables. Index 0 is the joining column name that links the parent and child
               ChildIndent: 1 //  Optional. This is only required if you are using nested tables. Shift the child table to the right by this many columns
               SheetHeader: [{ // Add heading in Cell D1
                    Value: "EOL Parts Report as of " + getCurrentDate(),
                    Column: 3,
                    Row: 0,
                    MergeCells: "3,5", // Optional. Merges columns 3 thru 5 on the same row. You can also specify start row,end row,start column,end column to merge across more than 1 row
                    NamedStyle: "Heading", // Used named style definition. Don't use NamedStyle and Style together. Use only one
                    Style: [{ // Optional style sub object
                         Alignment: "left", // Optional. defaults to General if not specified 
                         BackgroundColor: "green", // Optional. Defaults to white if not specified
                         Bold: true, // Optional. Defaults to false if not specified		       
                    }],
               }], 
               ColumnSize: [20,10,null,15], // (Optional) Use null if you want to auto size         
               ColumnHeaders: "ID, Part Num,Part Description,Mfg Part Num,Ship Date,Status,Close Date",
               Columns: [["EOLID","INTEGER"],["PartNum","CHAR"],["PartDescription","CHAR"],["MfgPartNum","CHAR"],["LastShipDate","DATE"],["Status","CHAR"],["CloseDate","DATE"]],
               ConditionalFormatting: [{ // Optional
                    Formula: "$H4=\"\"", // The formula must evaluate to true or false
                    StartRow: 3, // Starting row.
                    EndingRow: 20, // Optional. If not specified, will default to the last row of the table
                    StartColumn: 0, // Starting column to apply the formatting to
                    EndColumn: 8, // Ending column to apply the formatting to
                    Style:[{ // You must specify at least one of the style formats below when using ocnditional formatting
                         Bold: true,
                         Italic: false,
                         Color: "red",
                         Size: 12,
                         Underline: true,
                         UnderlineStyle: "thick"
                         BackgroundColor: "yellow",			 
                    }],
               },{
                    Formula: "$I4 < 0",
                    StartRow: 3,
                    StartColumn: 0,
                    EndColumn: 8,
                    Style:[{
                         BackgroundColor: "red",
                    }],
               },],
               MergeCells: ["1,2","3,1,4,1"], // Optional. Merges columns 1 & 2 on the last row after data has been writtem. The second merge merges columns 3 & 4 on row 1. 
               SQL: "SELECT * FROM EOL", // SQL based data
               DBConnection: "PRODUCTION", // (Mandatory if SQL statement is provided   
	       FreezePane: [0,1], // The first parameter should be 1 to freeze the first column, the second parameter should be 1 to freeze the first row
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
               Image: [{
                    FileName: "Logo.png", // Valid image formats are DIB,EMF,JPEG/JPG,PICT,PNG OR WMF
                    AnchorType: "DONT_MOVE_AND_RESIZE", // Optional anchor options. DONT_MOVE_AND_RESIZE is the default. Valid options are DONT_MOVE_AND_RESIZE, DONT_MOVE_DO_RESIZE, MOVE_AND_RESIZE or MOVE_DONT_RESIZE
                    StartRow: 0, // Starting row to write the image at
                    EndRow: 3, // Optional. Defaults to StartRow if not provided
                    StartColumn: 1, // Starting column to write the image at
                    EndColumn: 1, // Optional. Defaults to StartColumn if not provided
                    ScaleX: 2.92, // Optional. Scales the width of the image based on an integer. 1 is 100%, 2 is 200% and so on
                    ScaleY: 1.38, // Optional. Scales the height of the image based on an integer. 1 is 100%, 2 is 200% and so on
                    Scale: Scales both the width and height of the image based on an integer. 1 is 100%, 2 is 200% and so on
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
     };`

## CreateExcelReportObj

I have added createExcelReportObj.js which has a helper method. If you call this function with a Verj table name, it wll automatically create the createExcelReport object for you and output it. You can then re-order the columns or change anything if you wish.

Known Issues
------------
This report generator will not work for Verj IO tables that are part of a table repeater but you can use nested SQL statements.
