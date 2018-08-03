function createExcelReport(reportObj) {	
     var allRows,rows;
     var blacklistedColumn, blacklistedColumnCounter;
     var blacklistColumnIndexes=[];
     var columnCounter,columnSizeCounter;
     var columnNames,currColumnName,currColumnValue;
     var formulaCounter;
     var rowCounter;
     var today=new Date();
     var todayStr=((today.getMonth()+1) < 10 ? "0" + (today.getMonth() + 1) : (today.getMonth()+1)) + "-" + (today.getDate() < 10 ? "0" + today.getDate() : today.getDate()) + "-" + today.getFullYear() + " " + today.getHours() + "-" + (today.getMinutes() < 10 ? "0" : "") + today.getMinutes() + "-" + today.getSeconds();
     var header = new HeaderFooter();
     var footer = new HeaderFooter();
     var reportObjSheetCounter;
     var sheet;
     var stmt,rs,con;
     var rowWritten=false;
     
     // *** START OF STYLE DEFINITIONS ***
     
     var mainHeadingFont=new WritableFont(WritableFont.TIMES,24, WritableFont.BOLD,false);
     var mainHeadingStyle=new WritableCellFormat(mainHeadingFont);

     var BOMHeadingFont=new WritableFont(WritableFont.TIMES,20, WritableFont.BOLD,false);
     var BOMHeadingFormat=new WritableCellFormat(BOMHeadingFont);

     // Set the border and background color for the headerFormat style
     var headerFont=new WritableFont(WritableFont.TIMES,10, WritableFont.BOLD,false);
     headerFont.setColour(Colour.WHITE);

     var headerFormat=new WritableCellFormat(headerFont);
     headerFormat.setBorder(Border.ALL,BorderLineStyle.THIN);
     headerFormat.setBackground(Colour.LIGHT_BLUE);

     // Create the cell style: Single line borders on all sides
     var cellFormat=new WritableCellFormat();
     cellFormat.setBorder(Border.ALL,BorderLineStyle.THIN);

     var alignLeftFormat=new WritableCellFormat();
     alignLeftFormat.setAlignment(Alignment.LEFT);
     alignLeftFormat.setBorder(Border.ALL,BorderLineStyle.THIN); 

     var cellCurrency=new NumberFormat(NumberFormat.CURRENCY_DOLLAR + "##,###,##0.00", NumberFormat.COMPLEX_FORMAT);
     var cellCurrencyFormat=new WritableCellFormat(cellCurrency);
     cellCurrencyFormat.setBorder(Border.ALL,BorderLineStyle.THIN);

     var autosize=new CellView()
     autosize.setAutosize(true);
     
     // *** END OF STYLE DEFINITIONS ***

     // *** START OF VALIDATION ***

	   // Validate that filename was provided
	   if (typeof reportObj.FileName == 'undefined') {
	        return ["ERROR","The property FileName was not specified"]; 	
	   }

	   // Validate that reportObj has a Sheets property
	   if (typeof reportObj.Sheets == 'undefined') {
	        return ["ERROR","The property Sheets was not specified"]; 	
	   }

	   var sheetArray=[]; // Holds names of all sheets
	   
	   // Loop through each Sheet object
	   for (reportObjSheetCounter=0;reportObj.Sheets[reportObjSheetCounter] != null;reportObjSheetCounter++) {	        
	   	
	        // Validate that the current sheet has a SheetName property
	        if (typeof reportObj.Sheets[reportObjSheetCounter].SheetName == 'undefined') {
	             return ["ERROR","The property SheetName in Sheet " + reportObjSheetCounter + " was not specified"]; 	
	        }

	        // Make sure that the sheet doesnt exist already
	        for (var i=0;i < sheetArray.length;i++) {
	             if (sheetArray[i][0]==reportObj.Sheets[reportObjSheetCounter].SheetName) {
	                  return ["ERROR","The property SheetName in Sheet " + reportObjSheetCounter + " has the same name as the sheet of sheet " + sheetArray[i][1]]; 	
	             }
	        }

          // Add the sheet name and sheet index to an array
	        sheetArray.push(new Array(reportObj.Sheets[reportObjSheetCounter].SheetName,reportObjSheetCounter));

	        // Validate that the current sheet has a SheetIndex property
	        if (typeof reportObj.Sheets[reportObjSheetCounter].SheetIndex == 'undefined') {
	             return ["ERROR","The property SheetIndex in Sheet " + reportObjSheetCounter + " was not specified"]; 	
	        }

	        // Validate that TableData or SQL query was provided
	        if (typeof reportObj.Sheets[reportObjSheetCounter].TableData == 'undefined' && typeof reportObj.Sheets[reportObjSheetCounter].SQL == 'undefined') {
	             return ["ERROR","The property TableData or SQL in Sheet " + reportObjSheetCounter + " was not specified"]; 	
	        }

	        // Validate that only TableData or SQL query were provided but not both
	        if (typeof reportObj.Sheets[reportObjSheetCounter].TableData !== 'undefined' && typeof reportObj.Sheets[reportObjSheetCounter].SQL != 'undefined') {
	             return ["ERROR","The properties TableData and SQL in Sheet " + reportObjSheetCounter + " were both specified. Please specify only one."]; 	
	        }	        

	        // Validate that Columns was provided	                       
	        if (typeof reportObj.Sheets[reportObjSheetCounter].Columns == 'undefined') {
	             return ["ERROR","The property Columns in Sheet " + reportObjSheetCounter + " was not specified"];	
	        }

          // Validate that ColumnHeaders was provided	                       
	        if (typeof reportObj.Sheets[reportObjSheetCounter].ColumnHeaders == 'undefined') {
	             return ["ERROR","The property ColumnHeaders in Sheet " + reportObjSheetCounter + " was not specified"];	
	        }
	             
          // Validate that the size of Columns and ColumnHeaders match
          if (reportObj.Sheets[reportObjSheetCounter].Columns.length != reportObj.Sheets[reportObjSheetCounter].ColumnHeaders.split(",").length) {
               return ["ERROR","The properties Columns and ColumnHeaders in Sheet " + reportObjSheetCounter + " are of different lengths. Column length=" + reportObj.Sheets[reportObjSheetCounter].Columns.length + " and ColumnHeaders length=" + reportObj.Sheets[reportObjSheetCounter].ColumnHeaders.split(",").length];
          }
          
          // If SQL was provided, make sure that all of the necessary properties were provided
          if (typeof reportObj.Sheets[reportObjSheetCounter].SQL != 'undefined') {
               // Validate that DBConnection was provided	                       
	             if (typeof reportObj.Sheets[reportObjSheetCounter].DBConnection == 'undefined') {
	                  return ["ERROR","The property DBConnection in Sheet " + reportObjSheetCounter + " was not specified"];	
	             }              
          }

          // If 1 or more formulas were provided, validate the formula related properties
	        if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas != 'undefined') {
               for (formulaCounter=0;formulaCounter < reportObj.Sheets[reportObjSheetCounter].Formulas.length;formulaCounter++) {
                    // Validate that Column was provided	                       
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Column == 'undefined') {
	                       return ["ERROR","The property Column in Sheet " + reportObjSheetCounter + ", Formulas[" + formulaCounter + "] was not specified"];	
	                  }

	                  // Validate that Formula was provided	                       
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Formula == 'undefined') {
	                       return ["ERROR","The property Formula in Sheet " + reportObjSheetCounter + ", Formulas[" + formulaCounter + "] was not specified"];	
	                  }

	                  // Validate that DataType was provided	                       
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].DataType == 'undefined') {
	                       return ["ERROR","The property DataType in Sheet " + reportObjSheetCounter + ", Formulas[" + formulaCounter + "] was not specified"];
	                  }	                  
               } // end of for (formulaCounter=0;formulaCounter < re
	        } // end of  if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas != 'undefined') {

          // If 1 or more hyperlinks were specified, validate the hyperlink properties
	        if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks != 'undefined') {
               for (hyperlinkCounter=0;hyperlinkCounter < reportObj.Sheets[reportObjSheetCounter].Hyperlinks.length;hyperlinkCounter++) {
                    // Validate that Column was provided	                       
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].Column == 'undefined') {
	                       return ["ERROR","The property Column in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] was not specified"];
	                  }

	                  // Validate that Row was provided	                       
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].Row == 'undefined') {
	                       return ["ERROR","The property Row in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] was not specified"];
	                  }

	                  // Validate that Value was provided	                       
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].Value == 'undefined') {
	                       return ["ERROR","The property Value in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] was not specified"];
	                  }

	                  // Validate that DestinationSheet was provided	- The destination sheet may not exist yet so don't validate that the sheet exists here                       
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].DestinationSheet == 'undefined') {
	                       return ["ERROR","The property DestinationSheet in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] was not specified"];
	                  }

	                  // Validate that DestinationColumn was provided
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].DestinationColumn == 'undefined') {
	                       return ["ERROR","The property DestinationColumn in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] was not specified"];
	                  }

	                  // Validate that DestinationRow was provided
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].DestinationRow == 'undefined') {
	                       return ["ERROR","The property DestinationCRow in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] was not specified"];
	                  }
               } // end of for (hyperlinkCounter=0;
	        } // end of if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks != 'undefined') {        
	   }

	   // If 1 or more custom celll texts were specified, validate the custom cell text properties
     if (typeof reportObj.CustomCellText != 'undefined') {
          for (customCellTextCounter=0;customCellTextCounter < reportObj.CustomCellText.length;customCellTextCounter++) {
               // Validate that DestinationSheet was provided	                       
	             if (typeof reportObj.CustomCellText[customCellTextCounter].DestinationSheet == 'undefined') {
	                  return ["ERROR","The property DestinationSheet in CustomCellText[" + customCellTextCounter + "] was not specified"];
	             }
	                  
               // Validate that Column was provided	                       
	             if (typeof reportObj.CustomCellText[customCellTextCounter].Column == 'undefined') {
	                  return ["ERROR","The property Column in CustomCellText[" + customCellTextCounter + "] was not specified"];
	             }

	             // Validate that Value was provided	                       
	             if (typeof reportObj.CustomCellText[customCellTextCounter].Value == 'undefined') {
	                  return ["ERROR","The property Value in CustomCellText[" + customCellTextCounter + "] was not specified"];
	             }
          }
	   }
	        
     // *** END OF VALIDATION ***
	   
     // Build the file name
	   reportObj.FileName=reportObj.FileName + " as of " + todayStr + ".xls";

	   // Create formatted date string for the file name
     workbook=Workbook.createWorkbook(new File(reportObj.FileName)); 

     // Override the default light blue with our own RGB colors
     workbook.setColourRGB(Colour.LIGHT_BLUE,83,162,240);

     // *** Start creating the Excel document ***
     
     // *** Loop through the report object for each sheet object ***
     for (reportObjSheetCounter=0;reportObj.Sheets[reportObjSheetCounter] != null;reportObjSheetCounter++) {
          // Create the sheet based on the specified name and index
          sheet=workbook.createSheet(reportObj.Sheets[reportObjSheetCounter].SheetName, reportObj.Sheets[reportObjSheetCounter].SheetIndex);

          // *** SET THE COLUMN WIDTHS ***
          // Use the provided ColumnSize property if it was provided
          if (typeof reportObj.Sheets[reportObjSheetCounter].ColumnSize !== 'undefined') {
               // Loop through each item in ColumnSize array. I loop through ColumnHeaders because its length represents the total # of actual columns and this way you can only provide 1 column width and not specify the rest of the column widths
               for (columnSizeCounter=0;columnSizeCounter<reportObj.Sheets[reportObjSheetCounter].ColumnHeaders.length;columnSizeCounter++) {
                    // If a non-null value was passed use it. Otherwise default to autosize
                    if (reportObj.Sheets[reportObjSheetCounter].ColumnSize[columnSizeCounter] != null) {
                         sheet.setColumnView(columnSizeCounter,reportObj.Sheets[reportObjSheetCounter].ColumnSize[columnSizeCounter]);
                    } else {
                    	   sheet.setColumnView(columnSizeCounter,autosize);
                    }                    
               }
          } else { // When ColumnSize isn't provided, default first 100 columns to autosize
               // Set first 100 columns to autosize
               for (var i=0;i<100;i++) {
                    sheet.setColumnView(i,autosize);
               }
          }

          rowCounter=0;

          if (reportObj.Sheets[reportObjSheetCounter].StartRow != null) {
               rowCounter+=parseInt(reportObj.Sheets[reportObjSheetCounter].StartRow);	
          }
          
          // *** SHEET HEADING ***
          if (reportObj.Sheets[reportObjSheetCounter].SheetHeader != null) {
          	   try {
          	        var sheetHeader=reportObj.Sheets[reportObjSheetCounter].SheetHeader[0];
          	   } catch (e) {
                    alert("An error occurred when SheetHeader= " + reportObj.Sheets[reportObjSheetCounter].SheetHeader);
               	    event.stopExecution();
               }

               var styledFormat=null;
                    
               if (sheetHeader.Style != null) {
                    styledFormat=createStyleFormat(sheetHeader.Style[0]);
               }

               // Write the heading factoring in the DataType property
               switch (reportObj.Sheets[reportObjSheetCounter].SheetHeader.DataType) {
                    case "BOOLEAN":
                    case "INTEGER":
                    case "NUMERIC":
                         sheet.addCell(new Packages.jxl.write.Number(sheetHeader.Column,sheetHeader.Row,sheetHeader.Value,(styledFormat != null ? styledFormat : mainHeadingStyle)));
                         break
                    case "DATE":
              	    case "DATETIME":
                         dateVal=new Date(sheetHeader.Value).mmddyyyy();

                         sheet.addCell(new Label(sheetHeader.Column,sheetHeader.Row,dateVal,(styledFormat != null ? styledFormat : mainHeadingStyle)));
                         break;
                    case "CHAR":
                    default:
                         if (styledFormat != null)
                              sheet.addCell(new Label(sheetHeader.Column,sheetHeader.Row,sheetHeader.Value,styledFormat));
                         else 
                              sheet.addCell(new Label(sheetHeader.Column,sheetHeader.Row,sheetHeader.Value));
              }
                    
               //sheet.addCell(new Label(sheetHeader.Column,sheetHeader.Row,sheetHeader.Value,(styledFormat != null ? styledFormat : mainHeadingStyle)));
               rowCounter++;
          }

          // *** TABLE COLUMN HEADERS ***
          var columnHeaders=reportObj.Sheets[reportObjSheetCounter].ColumnHeaders.split(",")
          
          for (columnCounter=0;columnCounter<columnHeaders.length;columnCounter++) {
               try {
                    if (columnHeaders[columnCounter].toUpperCase() != "WHERECLAUSE") {
                         sheet.addCell(new Label(columnCounter,rowCounter,columnHeaders[columnCounter],headerFormat));
                    }
               } catch (e) {
                    	alert("An error occurred when columnHeaders= " + columnHeaders + ", length=" + columnHeaders.length + ",columnCounter="+columnCounter+ " and value=" + columnHeaders[columnCounter]);
               	     event.stopExecution();
               }
          }

          rowCounter++;
          
          // Get all columns
          try {
               var columns=eval(reportObj.Sheets[reportObjSheetCounter].Columns);
          } catch (e) {
               alert("An error occurred when Columns= " + reportObj.Sheets[reportObjSheetCounter].Columns);
               event.stopExecution();
          }

          // *** Save all of the data in an array ***
          var data = [];

          // *** SQL based data ***
          if (reportObj.Sheets[reportObjSheetCounter].SQL != null) { 
               // Read the data
               services.database.executeSelectStatement(reportObj.Sheets[reportObjSheetCounter].DBConnection,reportObj.Sheets[reportObjSheetCounter].SQL,     
               function (columnData) {
               	   lineArr = [];
               	   rowArr = [];

               	   // Loop through all columns in provided column list
                    for (var key in columns) {
                    	   // Loop through each column of the current row
                         for (var colName in columnData) { 
                         	    // If the current column is the one we are looking for
                              if (reportObj.Sheets[reportObjSheetCounter].Columns[key][0].toUpperCase() == colName.toUpperCase()) {
                                   // add column name, column value, column type and date format for date type
                              	   lineArr = new Array(columns[key][0],(columnData[columns[key][0]] != null ? columnData[columns[key][0]] : null),columns[key][1],columns[key][2]);

                                   rowArr.push(lineArr);
                                   break;
                              }
                         }
                    }

                    // push line array
                    data.push(rowArr);
               });
          } else if (reportObj.Sheets[reportObjSheetCounter].TableData != null) { // *** TABLE BASED DATA *** 
               allRows=tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableData);
               rows=allRows.getRows();               
               
               // Loop through each row
               while (rows.next()) {
               	    lineArr = [];
               	    rowArr = [];
               	    
                    // Loop through each column for the current row
                    for (columnCounter=0;columnCounter<columnHeaders.length;columnCounter++) {
                         try {
                              currColumnValue=tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableData).getColumn(columns[columnCounter][0]).displayValue;                         
                         } catch(e) {
                              alert("An error occurred when columns[columnCounter]=" + columns[columnCounter] + " for the index " + columnCounter + " when columns=" + reportObj.Sheets[reportObjSheetCounter].Columns);
                              event.stopExecution();
                         }

                         currColumnType=columns[columnCounter][1];

                         currColumnIndex=columns[columnCounter][2];

                         // add column name, column value, column type and column index
                         lineArr = new Array(columns[columnCounter][0],(currColumnValue != null ? currColumnValue : null), currColumnType);

                         rowArr.push(lineArr);
                    }                    

                    // push line array
                    data.push(rowArr);
               }
          }

          // Output the data
          /*for (var dataCounter=0;dataCounter<data.length;dataCounter++) {
                 for (var colCounter=0;colCounter<columnHeaders.length;colCounter++) {
                      alert("[" + colCounter + "]=" + data[dataCounter][colCounter]);
                 }
          }

          if (system.securityManager.getCredential("REALNAME")=="Segi Hovav")
               return;*/
               
          currColumnIndex=0;

          // Loop through each "line" in the array
          for (var dataCounter=0;dataCounter<data.length;dataCounter++) {
          	   if (currColumnIndex=columnHeaders.length)
          	        currColumnIndex=0;

               // Write the data
               for (var colCounter=0;colCounter<columnHeaders.length;colCounter++) {
                    // If the type is CHAR but the value is an INT, change the type to an INT so it will be written as an INT so
                    // that Excel doesn't complaign that the field is a number in a text cell
                    if (data[dataCounter][colCounter][2] == "CHAR" && data[dataCounter][colCounter][1] != null && isInt(data[dataCounter][colCounter][1]))
                         data[dataCounter][colCounter][2]="INTEGER";
                         
                    switch(data[dataCounter][colCounter][2]) { // Type
                         case "BOOLEAN":
                         case "INTEGER":
                         case "NUMERIC":
                              // If the data is null, don't attempt to write a null value as a number because it will throw an error message
                              // Instead, let the switch fall though so its written as a char type
                              if (data[dataCounter][colCounter][1] != null) {
                                   rowWritten=true;
                         	         sheet.addCell(new Packages.jxl.write.Number(currColumnIndex,rowCounter, data[dataCounter][colCounter][1] ,cellFormat));
                         	         break;
                              }
                         case "CHAR":
                              rowWritten=true;
                              sheet.addCell(new Label(currColumnIndex,rowCounter, data[dataCounter][colCounter][1] ,cellFormat));                         
                              break;
                         case "CURRENCY":
                    	         rowWritten=true;
                    	         sheet.addCell(new Packages.jxl.write.Number(currColumnIndex,rowCounter, data[dataCounter][colCounter][1],cellCurrencyFormat));
                    	         break;
                    	   case "DATE":
                    	   case "DATETIME":
                    	         rowWritten=true;

                    	         var dateVal=null;

                              // If index 1 (value) isn't null and [3] (date format) isn't null
                    	         if (data[dataCounter][colCounter][1] != null && data[dataCounter][colCounter][3] != null) {
                    	              switch (data[dataCounter][colCounter][3]) {
                    	                   case "yyyymmdd":
                    	                       dateVal=new Date(data[dataCounter][colCounter][1]).yyyymmdd();
                    	                  	   break;
                    	                   case "mmddyy":
                    	                  	   dateVal=new Date(data[dataCounter][colCounter][1]).mmddyy();
                    	                  	   break;
                    	                  	case "mm/dd/yyyy":
                    	                  	default:
                    	                  	    dateVal=new Date(data[dataCounter][colCounter][1]).mmddyyyy();
                    	              }  
                    	         } else if (data[dataCounter][colCounter][1] != null) {
                    	              dateVal=new Date(data[dataCounter][colCounter][1]).mmddyyyy();
                    	         }

                    	         // dateVal shouldn't ever be null
               	              sheet.addCell(new Label(currColumnIndex,rowCounter,(dateVal != null ? dateVal : ""),cellFormat));                    	         
                    }

                    if (rowWritten==true) {
                         currColumnIndex++;
                    }
               }

               // **** Formulas *** Write any line item formulas if specified
               if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas != 'undefined') {
                    for (formulaCounter=0;formulaCounter < reportObj.Sheets[reportObjSheetCounter].Formulas.length;formulaCounter++) {
                         try {
                              // Only process formulas where LineFormula=true
                             	if (reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].LineFormula != true)
                             	    continue;

                              var columnNum=reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Column;
                              var rowNum=(typeof reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Row != 'undefined' ? reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Row : rowCounter);

                              if (reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].FormulaRowOffset != null) {
                                   rowNum+=parseInt(reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].FormulaRowOffset);
                              }
                                   
                              var formula=reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Formula.replaceAll("<CURRENTROW>",(rowCounter+1));
                              var format=(reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].DataType == "CURRENCY" ? cellCurrencyFormat : cellFormat);
                         } catch(e) {
                              alert("An error occurred when Formulas=" + reportObj.Sheets[reportObjSheetCounter].Formulas);
                              event.stopExecution();
                         }

                         sheet.addCell(new Formula(columnNum,rowNum,formula,format));
                    }
               }
                                   
               rowCounter++;
          }
          
          // **** Formulas *** After writing the current sheet, write any non-line formulas if specified
          if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas != 'undefined') {
               for (formulaCounter=0;formulaCounter < reportObj.Sheets[reportObjSheetCounter].Formulas.length;formulaCounter++) {
                    try {
                    	   // Only process formulas where LineFormula != true
                    	   if (reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].LineFormula == true) {
                              continue;
                         }
                              
                         var columnNum=reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Column;
                         var rowNum=(typeof reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Row != 'undefined' ? reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Row : rowCounter);

                         if (reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].FormulaRowOffset != null) {
                              rowNum+=parseInt(reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].FormulaRowOffset);
                         }
                         
                         var formula=reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Formula.replaceAll("<CURRENTROW>",(rowNum));
                         var format=(reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].DataType == "CURRENCY" ? cellCurrencyFormat : cellFormat);
                    } catch(e) {
                         alert("An error occurred when Formulas=" + reportObj.Sheets[reportObjSheetCounter].Formulas);
                         event.stopExecution();
                    }

                    try {
                         sheet.addCell(new Formula(columnNum,(rowNum),formula,format));
                    } catch(e) {
                         alert("an error occurred writing the formula with the error " + e + " when columnNum="+columnNum+", rownum="+(rowNum-1)+", formula="+formula+", format=" + reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].DataType);
                    }
               }
          }

          // **** Hyperlinks *** After writing the current sheet, write hyperlinks if specified
          if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks != 'undefined') {
               for (hyperlinkCounter=0;hyperlinkCounter < reportObj.Sheets[reportObjSheetCounter].Hyperlinks.length;hyperlinkCounter++) {
                    try {
                         var columnNum=reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].Column;
                         var rowNum=(typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].Row != 'undefined' ? reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].Row : rowCounter);                         
                         var value=reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].Value;

                         var destinationSheet=workbook.getSheet(reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].DestinationSheet);

                         if (destinationSheet==null)
                              destinationSheet=sheet;
                         
                         var destinationColumn=reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].DestinationColumn;

                         var destinationRow=reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].DestinationRow;
                    } catch(e) {
                         alert("An error occurred when Hyperlinks=" + reportObj.Sheets[reportObjSheetCounter].Hyperlinks);
                         event.stopExecution();
                    }
                    
                    // Validate that DestinationSheet is a valid sheet
                    if (workbook.getSheet(reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].DestinationSheet) == null) {	                            	
	                       workbook.close();
	                       FileServices.deleteFile(reportObj.FileName);
	                       return ["ERROR","The property DestinationSheet in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + formulaCounter + "] refers to a sheet that does not exist"];
                    }
                         
                    sheet.addHyperlink(new WritableHyperlink(columnNum,rowNum,value,destinationSheet,destinationColumn,destinationRow));
               }
          }

          // If 1 or more custom celll texts were specified, validate the custom cell text properties
          if (typeof reportObj.CustomCellText != 'undefined') {         
               for (customCellTextCounter=0;customCellTextCounter < reportObj.CustomCellText.length;customCellTextCounter++) {
                    try {
                    	   // Skip if the destination sheet name doesn't match the sheet that we are on
                    	   if (reportObj.Sheets[reportObjSheetCounter].SheetName != reportObj.CustomCellText[customCellTextCounter].DestinationSheet)
                    	        continue;
                    	        
                         var columnNum=reportObj.CustomCellText[customCellTextCounter].Column;
                         var rowNum=(typeof reportObj.CustomCellText[customCellTextCounter].Row != 'undefined' ? reportObj.CustomCellText[customCellTextCounter].Row : rowCounter);
                         var value=reportObj.CustomCellText[customCellTextCounter].Value;
                         var destinationSheet=workbook.getSheet(reportObj.CustomCellText[customCellTextCounter].DestinationSheet);

                         if (destinationSheet==null) {
                              alert("DestinationSheet is null for CustomCellText["+customCellTextCounter+"] when DestinationSheetValue="+reportObj.CustomCellText[customCellTextCounter].DestinationSheet);
                              event.stopExecution();
                         }
                    } catch(e) {
                         alert("An error occurred when CustomCellText=" + reportObj.CustomCellText);
                         event.stopExecution();
                    }

                   
                    var styledFormat=null;
                    
                    if (reportObj.CustomCellText[customCellTextCounter].Style != null) {
                    	alert("custom cell style");
                         styledFormat=createStyleFormat(reportObj.CustomCellText[customCellTextCounter].Style[0]);
                    }

                    // Write the CustomCellText factoring in the DataType property
                    switch (reportObj.CustomCellText[customCellTextCounter].DataType) {
                         case "BOOLEAN":
                         case "INTEGER":
                         case "NUMERIC":
                              destinationSheet.addCell(new Packages.jxl.write.Number(columnNum,rowNum,value,(styledFormat != null ? styledFormat : null)));
                              break;
                         case "DATE":
                    	   case "DATETIME":
                              dateVal=new Date(value).mmddyyyy();

                    	        destinationSheet.addCell(new Label(columnNum,rowNum,dateVal,(styledFormat != null ? styledFormat : null)));
                    	        break;
                         case "CHAR":
                         default:
                              if (styledFormat != null)
                                   destinationSheet.addCell(new Label(columnNum,rowNum,value,styledFormat));
                              else 
                                   destinationSheet.addCell(new Label(columnNum,rowNum,value));
                    }
               }
          }

          // Delete any blacklisted rows
          for (var i=0;i<blacklistColumnIndexes.length;i++) {
               sheet.removeColumn(blacklistColumnIndexes[i]);
          }
     }

     try {
          workbook.write();
          workbook.close();
     } catch(e) {
          alert("An error saving the workbook with the error " + e);
          event.stopExecution();
     }
     
     if (rowWritten==true) {
          return ["OK",reportObj.FileName];
     } else {
     	    FileServices.deleteFile(reportObj.FileName);
          return ["OK-NODATA",""];	
     }
}

// Build Jexcel format style based on the style properties specified in the style object
function createStyleFormat(style) {
     var color=null,BGColor=null;

     if (style.Color != null)
          color=style.Color.toString().toUpperCase();
          
     // Since the JExcel API doesn't offer a way to translate a color string into a Colour property, I have to use a switch
     switch (color) {
          case "BLACK":
               color=Colour.BLACK
               break;
          case "BLUE":
               color=Colour.BLUE
               break;
          case "BROWN":
               color=Colour.BROWN
               break;
          case "GOLD":
               color=Colour.GOLD
               break;
          case "GREEN":
               color=Colour.GREEN 
               break;
          case "ORANGE":
               color=Colour.ORANGE
               break;
          case "PINK":
               color=Colour.PINK
               break;
          case "RED":
               color=Colour.RED
               break;
          case "WHITE":
               color=Colour.WHITE
               break;
          case "YELLOW":
               color=Colour.YELLOW
               break;
          default:
               color=Colour.BLACK               
     }

      if (style.BackgroundColor != null)
          BGColor=style.BackgroundColor.toString().toUpperCase();
          
      switch (BGColor) {
          case "BLACK":
               BGColor=Colour.BLACK
               break;
          case "BLUE":
               BGColor=Colour.BLUE
               break;
          case "BROWN":
               BGColor=Colour.BROWN
               break;
          case "GOLD":
               BGColor=Colour.GOLD
               break;
          case "GREEN":
               BGColor=Colour.GREEN 
               break;
          case "ORANGE":
               BGColor=Colour.ORANGE
               break;
          case "PINK":
               BGColor=Colour.PINK
               break;
          case "RED":
               BGColor=Colour.RED
               break;
          case "WHITE":
               BGColor=Colour.WHITE
               break;
          case "YELLOW":
               BGColor=Colour.YELLOW
               break;
          default:
               BGColor=Colour.BLACK    
     }

     var size=(style.Size != null ? style.Size : 12);     
     var bold=(style.Bold == true ? true : false);
     var italic=(style.Italic == true ? true : false);
     var underline=(style.Underline == true ? true : false);
     var borders=(style.Borders == true ? true : false);

     if (borders==true) {
               switch (style.BorderStyle.toUpperCase()) {
          case "DASH_DOT":
               borderStyle=BorderLineStyle.DASH_DOT;
               break;
          case "DASH_DOT_DOT":
               borderStyle=BorderLineStyle.DASH_DOT_DOT;
               break;
          case "DASHED":
               borderStyle=BorderLineStyle.DASHED;
               break;
          case "DOTTED":
               borderStyle=BorderLineStyle.DOTTED;
               break;
          case "DOUBLE":
               borderStyle=BorderLineStyle.DOUBLE;
               break;
          case "HAIR":
               borderStyle=BorderLineStyle.HAIR;
               break;
          case "MEDIUM":
               borderStyle=BorderLineStyle.MEDIUM;
               break;
          case "MEDIUM_DASH_DOT":
               borderStyle=BorderLineStyle.MEDIUM_DASH_DOT;
               break;
          case "MEDIUM_DASH_DOT_DOT":
               borderStyle=BorderLineStyle.MEDIUM_DASH_DOT_DOT;
               break;
          case "MEDIUM_DASHED":
               borderStyle=BorderLineStyle.MEDIUM_DASHED;
               break;
          case "NONE":
               borderStyle=BorderLineStyle.NONE;
               break;
          case "SLANTED_DASH_DOT":
               borderStyle=BorderLineStyle.SLANTED_DASH_DOT;
               break;
          case "THICK":
               borderStyle=BorderLineStyle.THICK;
               break;
          case "THIN":
               borderStyle=BorderLineStyle.THIN;
               break;
          default:
               borderStyle=BorderLineStyle.THIN;
          }
     }

     var formatFont=new WritableFont(WritableFont.TIMES,size,(bold==true ? WritableFont.BOLD : WritableFont.NO_BOLD),italic);
   
     /*
     if (underline == true) {
     	    var u=Packages.jxl.Format.UnderlineStyle;

     	    // throws the error TypeError: [JavaPackage jxl.Format.UnderlineStyle] is not a function, it is object
          formatFont.setUnderlineStyle(Packages.jxl.Format.UnderlineStyle.SINGLE);
     }*/
          
     formatFont.setColour(color);

     var format=new WritableCellFormat(formatFont);
          
     format.setBackground(BGColor);
     
     if (borders == true) {
          format.setBorder(Border.ALL,borderStyle);
     }

     return format;
}