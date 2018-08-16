function createExcelReport(reportObj) {	
     var allRows,rows;
     var blacklistedColumn, blacklistedColumnCounter;
     var blacklistColumnIndexes=[];
     var namedStyles=[];
     var columnCounter,columnSizeCounter;
     var columnFound=false,columnNames,currColumnName,currColumnValue;
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

     // Add all named styles to an array. Used with the validation
     if (typeof reportObj.NamedStyles != 'undefined') {
          for (namedStylesCounter=0;namedStylesCounter <  reportObj.NamedStyles.length;namedStylesCounter++) {
               styledObj=createStyleFormat(reportObj.NamedStyles[namedStylesCounter]);

               namedStyles.push([reportObj.NamedStyles[namedStylesCounter].Name,styledObj]);
          }
     }
     
     // *** START OF VALIDATION ***
	   if (typeof reportObj.FileName == 'undefined')
	        return ["ERROR","The property FileName was not specified"];

     // If named styles was provided
     if (typeof reportObj.NamedStyles != 'undefined') {
          // Store the names in an array to make sure that it is unique
          styleNames = [];
          
          for (namedStylesCounter=0;namedStylesCounter <  reportObj.NamedStyles.length;namedStylesCounter++) {
               // Make sure that the name
               if (reportObj.NamedStyles[namedStylesCounter].Name == null)
                    return ["ERROR","The named style at index " + namedStylesCounter + " does not have a name property"]; 	

               for (styleNameCounter=0;styleNameCounter<styleNames.length;styleNameCounter++) {
                    if (reportObj.NamedStyles[namedStylesCounter].Name.toString().toUpperCase()==styleNames[styleNameCounter].toString().toUpperCase())
                         return ["ERROR","The named style " + reportObj.NamedStyles[namedStylesCounter].Name + " is defined more than once. Please make sure that the named style is unique."]; 	
               }

               styleNames.push(reportObj.NamedStyles[namedStylesCounter].Name);
          }
     }
     
	   // Validate that reportObj has a Sheets property
	   if (typeof reportObj.Sheets == 'undefined')
	        return ["ERROR","The property Sheets was not specified"];

	   var sheetArray=[]; // Holds names of all sheets
	   
	   // Loop through each Sheet object
	   for (reportObjSheetCounter=0;reportObj.Sheets[reportObjSheetCounter] != null;reportObjSheetCounter++) {	   	
	        // Validate that the current sheet has a SheetName property
	        if (typeof reportObj.Sheets[reportObjSheetCounter].SheetName == 'undefined')
	             return ["ERROR","The property SheetName in Sheet " + reportObjSheetCounter + " was not specified"];

	        // Make sure that the sheet doesnt exist already
	        for (var i=0;i < sheetArray.length;i++) {
	             if (sheetArray[i][0]==reportObj.Sheets[reportObjSheetCounter].SheetName)
	                  return ["ERROR","The property SheetName in Sheet " + reportObjSheetCounter + " has the same name as the sheet of sheet " + sheetArray[i][1]];
	        }

          // Add the sheet name and sheet index to an array
	        sheetArray.push(new Array(reportObj.Sheets[reportObjSheetCounter].SheetName,reportObjSheetCounter));

	        // Validate that the current sheet has a SheetIndex property
	        if (typeof reportObj.Sheets[reportObjSheetCounter].SheetIndex == 'undefined')
	             return ["ERROR","The property SheetIndex in Sheet " + reportObjSheetCounter + " was not specified"]; 	

          // Validate the AllMargins property
          if (typeof reportObj.Sheets[reportObjSheetCounter].AllMargins !== 'undefined') {
               var len=reportObj.Sheets[reportObjSheetCounter].AllMargins.split(",").length;

               if (len != 1 && len != 2 && len != 4)
	                  return ["ERROR","The property AllMargins in Sheet " + reportObjSheetCounter + " has an invalid size. You can either specify AllMargins:\"1\" to set all 4 margins to 1, AllMargins:\"1,0.25\" to set the top and bottom margins to 1 and the left and right margins to 0.25 or AllMargins:\"1,1,0.25,0,25\" to set all 4 margins at once"];
          }

          // Validate the orientation if provided
          if (typeof reportObj.Sheets[reportObjSheetCounter].Orientation !== 'undefined') {
               if (reportObj.Sheets[reportObjSheetCounter].Orientation.toString().toUpperCase() != "PORTRAIT" && reportObj.Sheets[reportObjSheetCounter].Orientation.toString().toUpperCase() != "LANDSCAPE")
                    return ["ERROR","The property Orientation in Sheet " + reportObjSheetCounter + " has an invalid value. Valid values are landscape or portrait."];
          }
          
	        // If the sheet has a header, validate the MergeCells property if defined
	        if (typeof reportObj.Sheets[reportObjSheetCounter].SheetHeader != 'undefined') {
	             // Validate MergeCells length
	             if (typeof reportObj.Sheets[reportObjSheetCounter].SheetHeader[0].MergeCells != 'undefined') {
                    var len=reportObj.Sheets[reportObjSheetCounter].SheetHeader[0].MergeCells.split(",").length;
	                  
	                  if (len != 2 && len != 4)
	                       return ["ERROR","The property SheetHeader in Sheet " + reportObjSheetCounter + " has a MergeCells property with an invalid size. You can either specify MergeCells:\"2,4\" to merge rows 2-4 on the same row or MergeCells:\"2,2,4,4\" to merge from cell 2,2-4,4"];
	             }

	             // Make sure that the user only specifies Style or NamedStyle but not both
	             if (typeof reportObj.Sheets[reportObjSheetCounter].SheetHeader[0].NamedStyle != 'undefined' && typeof reportObj.Sheets[reportObjSheetCounter].SheetHeader[0].Style != 'undefined')
	                  return ["ERROR","The sheet header at index " + reportObjSheetCounter + " has a Style and NamedStyle property. Please specify only one"];

               // If a named style was provided, make sure that the name references a valid named style
               if (typeof reportObj.Sheets[reportObjSheetCounter].SheetHeader[0].NamedStyle != 'undefined') {
	                  namedStyleFound=false;

	                  for (namedStylesCounter=0;namedStylesCounter <  namedStyles.length;namedStylesCounter++) {
                         if (namedStyles[namedStylesCounter][0].toString().toUpperCase() === reportObj.Sheets[reportObjSheetCounter].SheetHeader[0].NamedStyle.toString().toUpperCase()) {
                              namedStyleFound=true;
                              break;
                         }
                    }

                    if (namedStyleFound==false)
                         return ["ERROR","The NamedStyle property " + reportObj.Sheets[reportObjSheetCounter].SheetHeader[0].NamedStyle + " in sheet " + reportObjSheetCounter + " does not appear to be a valid NamedStyle. Please refer to a valid named style"];
               }
	        }

	        // Validate that TableData or SQL query was provided
	        if (typeof reportObj.Sheets[reportObjSheetCounter].TableData == 'undefined' && typeof reportObj.Sheets[reportObjSheetCounter].SQL == 'undefined')
	             return ["ERROR","The property TableData or SQL in Sheet " + reportObjSheetCounter + " was not specified"];

	        // Validate that only TableData or SQL query were provided but not both
	        if (typeof reportObj.Sheets[reportObjSheetCounter].TableData !== 'undefined' && typeof reportObj.Sheets[reportObjSheetCounter].SQL != 'undefined')
	             return ["ERROR","The properties TableData and SQL in Sheet " + reportObjSheetCounter + " were both specified. Please specify only one."];

          if (typeof reportObj.Sheets[reportObjSheetCounter].TableData !== 'undefined' && tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableData)==null)
               return ["ERROR","The table " + reportObj.Sheets[reportObjSheetCounter].TableData + " referenced in sheet " + reportObjSheetCounter + " is not a valid table"]; 	
          
	        // Validate that Columns was provided	                       
	        if (typeof reportObj.Sheets[reportObjSheetCounter].Columns == 'undefined')
	             return ["ERROR","The property Columns in Sheet " + reportObjSheetCounter + " was not specified"];

          // Validate that ColumnHeaders was provided	                       
	        if (typeof reportObj.Sheets[reportObjSheetCounter].ColumnHeaders == 'undefined')
	             return ["ERROR","The property ColumnHeaders in Sheet " + reportObjSheetCounter + " was not specified"];
	             
          // Validate that the size of Columns and ColumnHeaders match
          if (reportObj.Sheets[reportObjSheetCounter].Columns.length != reportObj.Sheets[reportObjSheetCounter].ColumnHeaders.split(",").length)
               return ["ERROR","The properties Columns and ColumnHeaders in Sheet " + reportObjSheetCounter + " are of different lengths. Column length=" + reportObj.Sheets[reportObjSheetCounter].Columns.length + " and ColumnHeaders length=" + reportObj.Sheets[reportObjSheetCounter].ColumnHeaders.split(",").length];

          // Validate merge cells
          if (typeof reportObj.Sheets[reportObjSheetCounter].MergeCells != 'undefined') {
               for (var mergeCounter=0;mergeCounter < reportObj.Sheets[reportObjSheetCounter].MergeCells.length;mergeCounter++) {
                    var len=reportObj.Sheets[reportObjSheetCounter].MergeCells[mergeCounter].split(",").length;

                    if (len != 2 && len != 4)
	                       return ["ERROR","The property MergeCells in Sheet " + reportObjSheetCounter + " at index " + mergeCounter + " has an invalid size. You can either specify MergeCells:\"2,4\" to merge rows 2-4 on the same row or MergeCells:\"2,2,4,4\" to merge from cell 2,2-4,4"];
               }
          }
          
          // If SQL was provided, make sure that all of the necessary properties were provided
          if (typeof reportObj.Sheets[reportObjSheetCounter].SQL != 'undefined') {
               // Validate that DBConnection was provided	                       
	             if (typeof reportObj.Sheets[reportObjSheetCounter].DBConnection == 'undefined')
	                  return ["ERROR","The property DBConnection in Sheet " + reportObjSheetCounter + " was not specified"];
          }

          // If 1 or more formulas were provided, validate the formula related properties
	        if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas != 'undefined') {
               for (formulaCounter=0;formulaCounter < reportObj.Sheets[reportObjSheetCounter].Formulas.length;formulaCounter++) {
                    // Validate that Column was provided
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Column == 'undefined')
	                       return ["ERROR","The property Column in Sheet " + reportObjSheetCounter + ", Formulas[" + formulaCounter + "] was not specified"];

	                  // Validate that Formula was provided
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Formula == 'undefined')
	                       return ["ERROR","The property Formula in Sheet " + reportObjSheetCounter + ", Formulas[" + formulaCounter + "] was not specified"];

	                  // Validate that DataType was provided
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].DataType == 'undefined')
	                       return ["ERROR","The property DataType in Sheet " + reportObjSheetCounter + ", Formulas[" + formulaCounter + "] was not specified"];

	                  // Make sure that if LineFormula is specified, Row cannot be specified
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].LineFormula != 'undefined' && typeof reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Row != 'undefined')
	                       return ["ERROR","The properties Row and LineFormula in Sheet " + reportObjSheetCounter + ", Formulas[" + formulaCounter + "] were both specified. Please specify only one"];
               } // end of  for (formulaCounter=0;formulaCounter < reportObj.Sheets[reportObjSheetCounter].Formulas.length;formulaCounter++) {
	        } // end of  if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas != 'undefined') {

          // If 1 or more hyperlinks were specified, validate the hyperlink properties
	        if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks != 'undefined') {
               for (hyperlinkCounter=0;hyperlinkCounter < reportObj.Sheets[reportObjSheetCounter].Hyperlinks.length;hyperlinkCounter++) {
                    // Validate that Column was provided	                       
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].Column == 'undefined')
	                       return ["ERROR","The property Column in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] was not specified"];

	                  // Validate that Row was provided	                       
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].Row == 'undefined')
	                       return ["ERROR","The property Row in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] was not specified"];

	                  // Validate that Value was provided	                       
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].Value == 'undefined')
	                       return ["ERROR","The property Value in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] was not specified"];

	                  // Validate that DestinationSheet was provided	- The destination sheet may not exist yet so don't validate that the sheet exists here                       
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].DestinationSheet == 'undefined')
	                       return ["ERROR","The property DestinationSheet in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] was not specified"];

	                  // Validate that DestinationColumn was provided
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].DestinationColumn == 'undefined')
	                       return ["ERROR","The property DestinationColumn in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] was not specified"];

	                  // Validate that DestinationRow was provided
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].DestinationRow == 'undefined')
	                       return ["ERROR","The property DestinationRow in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] was not specified"];
               } // end of for (hyperlinkCounter=0;
	        } // end of if (typeof reportObj.Sheets[reportObjSheetCounter].Hyperlinks != 'undefined') {        

          // Make sure that the user only specifies Style or NamedStyle but not both
	        if (typeof reportObj.Sheets[reportObjSheetCounter].NamedStyle != 'undefined' && typeof reportObj.Sheets[reportObjSheetCounter].Style != 'undefined')
	             return ["ERROR","The sheet at index " + reportObjSheetCounter + " has a Style and NamedStyle property. Please specify only one"];

          // If a named style was provided, make sure that the name references a valid named style
          if (typeof reportObj.Sheets[reportObjSheetCounter].NamedStyle != 'undefined') {
	             namedStyleFound=false;

	             for (namedStylesCounter=0;namedStylesCounter <  namedStyles.length;namedStylesCounter++) {
                    if (namedStyles[namedStylesCounter][0].toString().toUpperCase() === reportObj.Sheets[reportObjSheetCounter].NamedStyle.toString().toUpperCase()) {
                         namedStyleFound=true;
                         break;
                    }
               }

               if (namedStyleFound==false)
                    return ["ERROR","The NamedStyle property " + reportObj.Sheets[reportObjSheetCounter].NamedStyle + " in sheet " + reportObjSheetCounter + " does not appear to be a valid NamedStyle. Please refer to a valid named style"];
          }
	   }

	   // If 1 or more custom celll texts were specified, validate the custom cell text properties
     if (typeof reportObj.CustomCellText != 'undefined') {
          for (customCellTextCounter=0;customCellTextCounter < reportObj.CustomCellText.length;customCellTextCounter++) {
               // Validate that DestinationSheet was provided	                       
	             if (typeof reportObj.CustomCellText[customCellTextCounter].DestinationSheet == 'undefined')
	                  return ["ERROR","The property DestinationSheet in CustomCellText[" + customCellTextCounter + "] was not specified"];
	                  
               // Validate that Column was provided	                       
	             if (typeof reportObj.CustomCellText[customCellTextCounter].Column == 'undefined')
	                  return ["ERROR","The property Column in CustomCellText[" + customCellTextCounter + "] was not specified"];

	             // Validate that Value was provided	                       
	             if (typeof reportObj.CustomCellText[customCellTextCounter].Value == 'undefined')
	                  return ["ERROR","The property Value in CustomCellText[" + customCellTextCounter + "] was not specified"];

	             // Validate merge cells if specified
               if (typeof reportObj.CustomCellText[customCellTextCounter].MergeCells != 'undefined') {
                    for (var mergeCounter=0;mergeCounter < reportObj.CustomCellText[customCellTextCounter].MergeCells.length;mergeCounter++) {
                         var len=reportObj.CustomCellText[customCellTextCounter].MergeCells[mergeCounter].split(",").length;

                         if (len != 4)
	                            return ["ERROR","The property MergeCells in CustomCellText at index " + customCellTextCounter + " with MergeCell index " + mergeCounter + " has an invalid size. You must specify start column,start row,end column,end row"];
                    }
               }

               // Make sure that the user only specifies Style or NamedStyle but not both
               if (typeof reportObj.CustomCellText[customCellTextCounter].NamedStyle != 'undefined' && typeof reportObj.CustomCellText[customCellTextCounter].Style != 'undefined')
	                  return ["ERROR","The CustomCellText at index " + customCellTextCounter + " has a Style and NamedStyle property. Please specify only one"];

               // If a named style was provided, make sure that the name references a valid named style
	             if (typeof reportObj.CustomCellText[customCellTextCounter].NamedStyle != 'undefined') {
	             	    namedStyleFound=false;

	             	    for (namedStylesCounter=0;namedStylesCounter <  namedStyles.length;namedStylesCounter++) {
                         if (namedStyles[namedStylesCounter][0].toString().toUpperCase() === reportObj.Sheets[reportObjSheetCounter].NamedStyle.toString().toUpperCase()) {
                              namedStyleFound=true;
                              break;
                         }
                    }

                    if (namedStyleFound==false)
                         return ["ERROR","The NamedStyle property " + reportObj.Sheets[reportObjSheetCounter].NamedStyle + " in sheet " + reportObjSheetCounter + " does not appear to be a valid NamedStyle. Please refer to a valid named style"];
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

     // *** START OF GENERATING THE EXCEL DOCUMENT ***
          
     // *** Loop through the report object for each sheet object ***
     for (reportObjSheetCounter=0;reportObj.Sheets[reportObjSheetCounter] != null;reportObjSheetCounter++) {
          // Create the sheet based on the specified name and index
          sheet=workbook.createSheet(reportObj.Sheets[reportObjSheetCounter].SheetName, reportObj.Sheets[reportObjSheetCounter].SheetIndex);

          // Set header / footer
          header.getLeft().appendWorkbookName(); // Add the workbook name to the top right of the header
          header.getRight().appendWorkSheetName(); // Add the sheet name to the top right of the header
          sheet.getSettings().setHeader(header);

          // Add the page numbers to the footer
          footer.getCentre().append("Page ");
          footer.getCentre().appendPageNumber();
          footer.getCentre().append("/");
          footer.getCentre().appendTotalPages();
          sheet.getSettings().setFooter(footer);
     
          // *** SET THE COLUMN WIDTHS ***
          // Use the provided ColumnSize property if it was provided
          if (typeof reportObj.Sheets[reportObjSheetCounter].ColumnSize !== 'undefined') {
               // Loop through each item in ColumnSize array. I loop through ColumnHeaders because its length represents the total # of actual columns and this way you can only provide 1 column width and not specify the rest of the column widths
               for (columnSizeCounter=0;columnSizeCounter<reportObj.Sheets[reportObjSheetCounter].ColumnHeaders.length;columnSizeCounter++) {
                    // If a non-null value was passed use it. Otherwise default to autosize
                    if (reportObj.Sheets[reportObjSheetCounter].ColumnSize[columnSizeCounter] != null)
                         sheet.setColumnView(columnSizeCounter,reportObj.Sheets[reportObjSheetCounter].ColumnSize[columnSizeCounter]);
                    else
                    	   sheet.setColumnView(columnSizeCounter,autosize);
               }
          } else { // When ColumnSize isn't provided, default first 100 columns to autosize
               // Set first 100 columns to autosize
               for (var i=0;i<100;i++) {
                    sheet.setColumnView(i,autosize);
               }
          }

          if (typeof reportObj.Sheets[reportObjSheetCounter].TopMargin !== 'undefined' && isInt(reportObj.Sheets[reportObjSheetCounter].TopMargin))
               sheet.getSettings().setTopMargin(parseInt(reportObj.Sheets[reportObjSheetCounter].TopMargin));

          if (typeof reportObj.Sheets[reportObjSheetCounter].BottomMargin !== 'undefined' && isInt(reportObj.Sheets[reportObjSheetCounter].BottomMargin))
               sheet.getSettings().setBottomMargin(parseInt(reportObj.Sheets[reportObjSheetCounter].BottomMargin));

          if (typeof reportObj.Sheets[reportObjSheetCounter].LeftMargin !== 'undefined' && isInt(reportObj.Sheets[reportObjSheetCounter].LeftMargin))
               sheet.getSettings().setLeftMargin(parseInt(reportObj.Sheets[reportObjSheetCounter].LeftMargin));

          if (typeof reportObj.Sheets[reportObjSheetCounter].RightMargin !== 'undefined' && isInt(reportObj.Sheets[reportObjSheetCounter].RightMargin))
               sheet.getSettings().setRightMargin(parseInt(reportObj.Sheets[reportObjSheetCounter].RightMargin));

          if (typeof reportObj.Sheets[reportObjSheetCounter].AllMargins !== 'undefined') {
               var len=reportObj.Sheets[reportObjSheetCounter].AllMargins.split(",").length;

               if (len==1) {
               	    var val=parseInt(reportObj.Sheets[reportObjSheetCounter].AllMargins);
               	    
                    sheet.getSettings().setTopMargin(parseInt(val));
                    sheet.getSettings().setBottomMargin(parseInt(val));
                    sheet.getSettings().setLeftMargin(parseInt(val));
                    sheet.getSettings().setRightMargin(parseInt(val));
               } else if (len==2) {
               	    var mergeCell=reportObj.Sheets[reportObjSheetCounter].AllMargins.split(",");

               	    sheet.getSettings().setTopMargin(parseInt(mergeCell[0]));
                    sheet.getSettings().setBottomMargin(parseInt(mergeCell[0]));
                    sheet.getSettings().setLeftMargin(parseInt(mergeCell[1]));
                    sheet.getSettings().setRightMargin(parseInt(mergeCell[1]));
               } else if (len==4) {
                    var mergeCell=reportObj.Sheets[reportObjSheetCounter].AllMargins.split(",");

                    sheet.getSettings().setTopMargin(parseInt(mergeCell[0]));
                    sheet.getSettings().setBottomMargin(parseInt(mergeCell[1]));
                    sheet.getSettings().setLeftMargin(parseInt(mergeCell[2]));
                    sheet.getSettings().setRightMargin(parseInt(mergeCell[3]));
               }
          }

          rowCounter=0;

          // HeaderMargin
          if (typeof reportObj.Sheets[reportObjSheetCounter].HeaderMargin !== 'undefined')
               sheet.getSettings().setHeaderMargin(parseInt(reportObj.Sheets[reportObjSheetCounter].HeaderMargin));

          // FooterMargin
          if (typeof reportObj.Sheets[reportObjSheetCounter].FooterMargin !== 'undefined')
               sheet.getSettings().setFooterMargin(parseInt(reportObj.Sheets[reportObjSheetCounter].FooterMargin));

          // FitWidth
          if (typeof reportObj.Sheets[reportObjSheetCounter].FitWidth !== 'undefined' && reportObj.Sheets[reportObjSheetCounter].FitWidth == true)
               sheet.getSettings().setFitWidth(1);

          // FitHeight
          if (typeof reportObj.Sheets[reportObjSheetCounter].FitHeight !== 'undefined' && reportObj.Sheets[reportObjSheetCounter].FitHeight == true)
               sheet.getSettings().setFitHeight(1);

          // FitToPages
          if (typeof reportObj.Sheets[reportObjSheetCounter].FitToPages !== 'undefined' && reportObj.Sheets[reportObjSheetCounter].FitToPages == true)
               sheet.getSettings().setFitToPages(true);

          // Orientation
          if (typeof reportObj.Sheets[reportObjSheetCounter].Orientation !== 'undefined') {
               var orientation;

               if (reportObj.Sheets[reportObjSheetCounter].Orientation.toString().toUpperCase() == "PORTRAIT")
                    orientation=Packages.jxl.format.PageOrientation.PORTRAIT;
               else if (reportObj.Sheets[reportObjSheetCounter].Orientation.toString().toUpperCase() == "LANDSCAPE")
                    orientation=Packages.jxl.format.PageOrientation.LANDSCAPE;

               sheet.getSettings().setOrientation(orientation);
          }

          // Password
          if (typeof reportObj.Sheets[reportObjSheetCounter].Password !== 'undefined') {
               sheet.getSettings().setPassword(reportObj.Sheets[reportObjSheetCounter].Password);
               sheet.getSettings().setProtected(true);
          }
          
          
          // Optional setting that If provided, indicates which row to start writing the data to
          if (reportObj.Sheets[reportObjSheetCounter].StartRow != null)
               rowCounter+=parseInt(reportObj.Sheets[reportObjSheetCounter].StartRow);
          
          // *** SHEET HEADING ***
          if (reportObj.Sheets[reportObjSheetCounter].SheetHeader != null) {
          	   try {
          	        var sheetHeader=reportObj.Sheets[reportObjSheetCounter].SheetHeader[0];
          	   } catch (e) {
                    // Close and delete the workbook
                    workbook.close();
                    
                    if (FileServices.existsFile(reportObj.FileName)) FileServices.deleteFile(reportObj.FileName);

                    return ["ERROR","An error occurred when SheetHeader= " + reportObj.Sheets[reportObjSheetCounter].SheetHeader];
               }

               var styledFormat=null;
                    
               if (sheetHeader.Style != null) {
                    styledFormat=createStyleFormat(sheetHeader.Style[0]);
               } else if (sheetHeader.NamedStyle != null) {
               	    for (namedStylesCounter=0;namedStylesCounter <  namedStyles.length;namedStylesCounter++) {
                         if (namedStyles[namedStylesCounter][0].toString().toUpperCase() === sheetHeader.NamedStyle.toString().toUpperCase()) {
                              styledFormat=namedStyles[namedStylesCounter][1];
                              break;
                         }
                    }
               }

               if (sheetHeader.DataType=="INT") // Alias for INTEGER data type
                    sheetHeader.DataType="INTEGER";
               
               // Write the heading factoring in the DataType property
               switch (sheetHeader.DataType) {
                    case "BOOLEAN":
                    case "INT":                       
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

              // Merge cells if specifed
              if (sheetHeader.MergeCells != null) {
                   var mergeCell=sheetHeader.MergeCells.split(",");

                   if (mergeCell.length == 2)
                        sheet.mergeCells(parseInt(mergeCell[0]),sheetHeader.Row,parseInt(mergeCell[1]),sheetHeader.Row);
                   else
                        sheet.mergeCells(parseInt(mergeCell[0]),parseInt(mergeCell[1]),parseInt(mergeCell[2]),parseInt(mergeCell[3]));
              }
              
              rowCounter++;
          }

          // *** TABLE COLUMN HEADERS ***
          var columnHeaders=reportObj.Sheets[reportObjSheetCounter].ColumnHeaders.split(",")
          
          for (columnCounter=0;columnCounter<columnHeaders.length;columnCounter++) {
               try {
                    if (columnHeaders[columnCounter].toUpperCase() != "WHERECLAUSE")
                         sheet.addCell(new Label(columnCounter,rowCounter,columnHeaders[columnCounter],headerFormat));
               } catch (e) {
                    workbook.close();
                    if (FileServices.existsFile(reportObj.FileName))
                         FileServices.deleteFile(reportObj.FileName);
                    
                    return ["ERROR","An error occurred when columnHeaders= " + columnHeaders + ", length=" + columnHeaders.length + ",columnCounter="+columnCounter+ " and value=" + columnHeaders[columnCounter]];
               }
          }

          rowCounter++;

          if (reportObj.Sheets[reportObjSheetCounter].Style != null) {
               styledFormat=createStyleFormat(reportObj.Sheets[reportObjSheetCounter].Style[0]);
          } else if (reportObj.Sheets[reportObjSheetCounter].NamedStyle != null) {
               for (namedStylesCounter=0;namedStylesCounter <  namedStyles.length;namedStylesCounter++) {
                    if (namedStyles[namedStylesCounter][0].toString().toUpperCase() === reportObj.Sheets[reportObjSheetCounter].NamedStyle.toString().toUpperCase()) {
                         styledFormat=namedStyles[namedStylesCounter][1];
                         break;
                    }
               }
          } else {
          	  styledFormat=null;
          }
          
          // Get all columns
          try {
               var columns=eval(reportObj.Sheets[reportObjSheetCounter].Columns);
          } catch (e) {
               return ["ERROR","An error occurred when Columns= " + reportObj.Sheets[reportObjSheetCounter].Columns];
          }

          // *** Save all of the data in an array ***
          var data = [];

          // *** SQL based data ***
          if (reportObj.Sheets[reportObjSheetCounter].SQL != null) { 
               columnNotFound=false;
               invalidColumn="";
               
               // Read the data
               services.database.executeSelectStatement(reportObj.Sheets[reportObjSheetCounter].DBConnection,reportObj.Sheets[reportObjSheetCounter].SQL,     
               function (columnData) {
               	   lineArr = [];
               	   rowArr = [];

               	   // Loop through all columns in provided column list
                    for (var key in columns) {
                    	   columnFound=false;
                    	   
                    	   // Loop through each column of the current row
                         for (var colName in columnData) { 
                         	    // If the current column is the one we are looking for
                              if (columns[key][0] != null && columns[key][0].toUpperCase() == colName.toUpperCase()) {

                                   // add column name, column value, column type and date format for date type
                              	   lineArr = new Array(columns[key][0],(columnData[columns[key][0]] != null ? columnData[columns[key][0]] : null),columns[key][1],columns[key][2]);

                                   rowArr.push(lineArr);
                                   
                                   columnFound=true;
                                   
                                   break;
                              }
                         }

                         if (columnFound == false) {
                         	    invalidColumn=columns[key][0];
                              break;
                         }
                    }

                    if (columnFound == false)
                         return;

                    // push line array
                    data.push(rowArr);
               });

               if (columnFound == false && invalidColumn != null && invalidColumn != "")
                    return ["ERROR","The sql column " + invalidColumn + " was not found in the database. Please check the spelling of the column name"];
          // *** TABLE BASED DATA *** 
          } else if (reportObj.Sheets[reportObjSheetCounter].TableData != null) {
               allRows=tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableData);
               rows=allRows.getRows();               
               
               // Loop through each row
               while (rows.next()) {
               	    lineArr = [];
               	    rowArr = [];
               	    
                    // Loop through each column for the current row
                    for (columnCounter=0;columnCounter<columnHeaders.length;columnCounter++) {
                         try {
                         	    if (columns[columnCounter][0] == null)
                         	         continue;
                         	         
                         	    // Make sure htat the column name is valid
                         	    if (tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableData).getColumn(columns[columnCounter][0]) == null)
                         	         return ["ERROR","The table column " + columns[columnCounter][0] + " was not found in the database. Please check the spelling of the column name"];
                              
                              currColumnValue=tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableData).getColumn(columns[columnCounter][0]).displayValue;

                              if (currColumnValue != null)
                                   currColumnValue=currColumnValue.replaceAll("<BR>","");
                         } catch(e) {
                         	    workbook.close();
                              if (FileServices.existsFile(reportObj.FileName)) FileServices.deleteFile(reportObj.FileName);
                              
                              return ["ERROR","An error occurred when columnCounter=" + columnCounter + ", columns[columnCounter][0]=" + columns[columnCounter][0] + ", columns[columnCounter]=" + columns[columnCounter] + " for the index " + columnCounter + " when columns=" + reportObj.Sheets[reportObjSheetCounter].Columns + " with the error message " + e];
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
          }*/

          currColumnIndex=0;

          // *** START OF LOOP THAT GOES THROUGH DATA ARRAY AND WRITES THE DATA ***
          for (var dataCounter=0;dataCounter<data.length;dataCounter++) {
          	   if (currColumnIndex=columnHeaders.length)
          	        currColumnIndex=0;

               // Write the data
               for (var colCounter=0;colCounter<columnHeaders.length;colCounter++) {
               	    if (data[dataCounter][colCounter]==null)
               	         continue;

                    // If the type is CHAR but the value is an INT, change the type to an INT so it will be written as an INT so
                    // that Excel doesn't complaign that the field is a number in a text cell
                    if (data[dataCounter][colCounter][2] == "CHAR" && data[dataCounter][colCounter][1] != null && isInt(data[dataCounter][colCounter][1]) && reportObj.Sheets[reportObjSheetCounter].Columns[colCounter][2] != true)
                         data[dataCounter][colCounter][2]="INTEGER";

                    // If the type is INT but the value is a CHAR, change the type to an CHAR so it will be written as a CHAR
                    if ((data[dataCounter][colCounter][2] == "INTEGER" || data[dataCounter][colCounter][2] == "INT") && data[dataCounter][colCounter][1] != null && !isInt(data[dataCounter][colCounter][1]))
                         data[dataCounter][colCounter][2]="CHAR";
                    
                    switch(data[dataCounter][colCounter][2]) { // Type
                         case "BOOLEAN":
                         case "INT":
                         case "INTEGER":
                         case "NUMERIC":
                              // If the data is null, don't attempt to write a null value as a number because it will throw an error message
                              // Instead, write empty string if its null
                              rowWritten=true;
                              
                              
                              if (data[dataCounter][colCounter][1] != null)                                   
                         	         sheet.addCell(new Packages.jxl.write.Number(currColumnIndex,rowCounter, data[dataCounter][colCounter][1] ,(styledFormat != null ? styledFormat : cellFormat)));
                              else
                                   sheet.addCell(new Label(currColumnIndex,rowCounter, "" ,(styledFormat != null ? styledFormat : cellFormat)));

                              break;
                         case "CHAR":
                              rowWritten=true;

                              sheet.addCell(new Label(currColumnIndex,rowCounter, data[dataCounter][colCounter][1] ,(styledFormat != null ? styledFormat : cellFormat)));                         

                              break;
                         case "CURRENCY":
                    	         rowWritten=true;

                    	         // If the data is null, don't attempt to write a null value as a number because it will throw an error message
                              // Instead, write empty string if its null
                    	         if (data[dataCounter][colCounter][1] != null)
                    	              sheet.addCell(new Packages.jxl.write.Number(currColumnIndex,rowCounter, data[dataCounter][colCounter][1],(styledFormat != null ? styledFormat : cellCurrencyFormat)));
                    	         else
                                   sheet.addCell(new Label(currColumnIndex,rowCounter, "" ,(styledFormat != null ? styledFormat : cellFormat)));
                                   
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
                    	         } else if (data[dataCounter][colCounter][1] != null)
                    	              dateVal=new Date(data[dataCounter][colCounter][1]).mmddyyyy();

                    	         // dateVal shouldn't ever be null
               	              sheet.addCell(new Label(currColumnIndex,rowCounter,(dateVal != null ? dateVal : ""),(styledFormat != null ? styledFormat : cellFormat)));                    	         
                    }

                    if (rowWritten==true)
                         currColumnIndex++;
               }

               // *** Merge cells if MergeCells was specified ***
               if (typeof reportObj.Sheets[reportObjSheetCounter].MergeCells != 'undefined') {
                    for (var mergeCounter=0;mergeCounter < reportObj.Sheets[reportObjSheetCounter].MergeCells.length;mergeCounter++) {
                        var mergeCell=reportObj.Sheets[reportObjSheetCounter].MergeCells[mergeCounter].split(",");

                        if (mergeCell.length == 2)
                             sheet.mergeCells(parseInt(mergeCell[0]),rowCounter,parseInt(mergeCell[1]),rowCounter);
                        else
                             sheet.mergeCells(parseInt(mergeCell[0]),parseInt(mergeCell[1]),parseInt(mergeCell[2]),parseInt(mergeCell[3]));  
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

                              if (reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].FormulaRowOffset != null)
                                   rowNum+=parseInt(reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].FormulaRowOffset);
                                   
                              var formula=reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Formula.replaceAll("<CURRENTROW>",(rowCounter+1));
                              var format=(reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].DataType == "CURRENCY" ? cellCurrencyFormat : cellFormat);
                         } catch(e) {
                         	    workbook.close();
                         	    
                              if (FileServices.existsFile(reportObj.FileName))
                                   FileServices.deleteFile(reportObj.FileName);
                              
                              return ["ERROR","An error occurred when Formulas=" + reportObj.Sheets[reportObjSheetCounter].Formulas];
                         }

                         sheet.addCell(new Formula(columnNum,rowNum,formula,format));
                    }
               }
                                   
               rowCounter++;
          }
          // *** END OF LOOP THAT GOES THROUGH DATA ARRAY AND WRITES THE DATA ***
          
          // **** Formulas *** After writing the current sheet, write any non-line formulas if specified
          if (typeof reportObj.Sheets[reportObjSheetCounter].Formulas != 'undefined') {
               for (formulaCounter=0;formulaCounter < reportObj.Sheets[reportObjSheetCounter].Formulas.length;formulaCounter++) {
                    try {
                    	   // Only process formulas where LineFormula != true
                    	   if (reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].LineFormula == true)
                              continue;
                              
                         var columnNum=reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Column;
                         var rowNum=(typeof reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Row != 'undefined' ? reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Row : rowCounter);

                         if (reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].FormulaRowOffset != null)
                              rowNum+=parseInt(reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].FormulaRowOffset);
                         
                         var formula=reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Formula.replaceAll("<CURRENTROW>",(rowNum));
                         var format=(reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].DataType == "CURRENCY" ? cellCurrencyFormat : cellFormat);
                    } catch(e) {
                    	   workbook.close();
                         if (FileServices.existsFile(reportObj.FileName))
                              FileServices.deleteFile(reportObj.FileName);
                              
                         return ["ERROR","An error occurred when Formulas=" + reportObj.Sheets[reportObjSheetCounter].Formulas];
                    }

                    try {
                         sheet.addCell(new Formula(columnNum,(rowNum),formula,format));
                    } catch(e) {
                    	   workbook.close();
                         if (FileServices.existsFile(reportObj.FileName))
                              FileServices.deleteFile(reportObj.FileName);
                         
                         return ["ERROR","An error occurred writing the formula with the error " + e + " when columnNum="+columnNum+", rownum="+(rowNum-1)+", formula="+formula+", format=" + reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].DataType];
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
                    	   workbook.close();
                         if (FileServices.existsFile(reportObj.FileName))
                              FileServices.deleteFile(reportObj.FileName);
                         
                         return ["ERROR","An error occurred when Hyperlinks=" + reportObj.Sheets[reportObjSheetCounter].Hyperlinks];
                    }
                    
                    // Validate that DestinationSheet is a valid sheet. We can't do this in the validation because the sheet won't exist use in the section that does the validation
                    if (workbook.getSheet(reportObj.Sheets[reportObjSheetCounter].Hyperlinks[hyperlinkCounter].DestinationSheet) == null) {	                            	
	                       workbook.close();

	                       if (FileServices.existsFile(reportObj.FileName))
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

                         if (destinationSheet==null)
                              return ["ERROR","DestinationSheet is null for CustomCellText["+customCellTextCounter+"] when DestinationSheetValue="+reportObj.CustomCellText[customCellTextCounter].DestinationSheet];
                    } catch(e) {
                    	   workbook.close();
                         if (FileServices.existsFile(reportObj.FileName))
                              FileServices.deleteFile(reportObj.FileName);
                              
                         return ["ERROR","An error occurred when CustomCellText=" + reportObj.CustomCellText];
                    }

                   
                    var styledFormat=null;
                    
                    if (reportObj.CustomCellText[customCellTextCounter].Style != null) {
                         styledFormat=createStyleFormat(reportObj.CustomCellText[customCellTextCounter].Style[0]);
                    } else if (reportObj.CustomCellText[customCellTextCounter].NamedStyle != null) {
                    	   for (namedStylesCounter=0;namedStylesCounter <  namedStyles.length;namedStylesCounter++) {
                              if (namedStyles[namedStylesCounter][0].toString().toUpperCase() === reportObj.CustomCellText[customCellTextCounter].NamedStyle.toString().toUpperCase()) {
                                   styledFormat=namedStyles[namedStylesCounter][1];
                                   break;
                              }
                         }
                    } else
          	             styledFormat=null;

                    // Write the CustomCellText factoring in the DataType property
                    switch (reportObj.CustomCellText[customCellTextCounter].DataType) {
                         case "BOOLEAN":
                         case "INT":
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

                    if (typeof reportObj.CustomCellText[customCellTextCounter].MergeCells != 'undefined') {
                         for (var mergeCounter=0;mergeCounter < reportObj.CustomCellText[customCellTextCounter].MergeCells.length;mergeCounter++) {
                              var mergeCell=reportObj.CustomCellText[customCellTextCounter].MergeCells[mergeCounter].split(",");

                              sheet.mergeCells(parseInt(mergeCell[0]),parseInt(mergeCell[1]),parseInt(mergeCell[2]),parseInt(mergeCell[3]));
                         }
                    }
               }
          }

          // Delete any blacklisted rows
          for (var i=0;i<blacklistColumnIndexes.length;i++) {
               sheet.removeColumn(blacklistColumnIndexes[i]);
          }
     }
     // *** END OF GENERATING THE EXCEL DOCUMENT ***

     try {
          workbook.write();
          workbook.close();
     } catch(e) {
          if (FileServices.existsFile(reportObj.FileName))
               FileServices.deleteFile(reportObj.FileName);

          return ["ERROR","An error saving the workbook with the error " + e];
     }
     
     if (rowWritten==true)
          return ["OK",reportObj.FileName];
     else {
     	    FileServices.deleteFile(reportObj.FileName);
          return ["OK-NODATA",""];	
     }
}

// Build Jexcel format style based on the style properties specified in the style object
function createStyleFormat(style) {
     var color=null,BGColor=null;

     // Since the JExcel API doesn't provide a way to translate a string color into a Colour object, I use an object instead
     var colorObject = {
          "AQUA" : Colour.AQUA,
          "BLACK": Colour.PALETTE_BLACK,
          "BLUE": Colour.BLUE,
          "BLUE_GREY" : Colour.BLUE_GREY,
          "BRIGHT_GREEN" : Colour.BRIGHT_GREEN,
          "BROWN" : Colour.BROWN,
          "CORAL" : Colour.CORAL,
          "DARK_BLUE" : Colour.DARK_BLUE,
          "DARK_GREEN" : Colour.DARK_GREEN,
          "DARK_PURPLE" : Colour.DARK_PURPLE,
          "DARK_RED" : Colour.DARK_RED,
          "DARK_TEAL" : Colour.DARK_TEAL,
          "DARK_YELLOW" : Colour.DARK_YELLOW,
          "GOLD" : Colour.GOLD,
          "GRAY_25" : Colour.GRAY_25,
          "GRAY_50" : Colour.GRAY_50,
          "GRAY_80" : Colour.GRAY_80,
          "GREEN" : Colour.GREEN,
          "GRAY_25_PERCENT" : Colour.GREY_25_PERCENT,
          "GRAY_40_PERCENT" : Colour.GREY_40_PERCENT,
          "GRAY_50_PERCENT" : Colour.GREY_50_PERCENT,
          "GRAY_80_PERCENT" : Colour.GREY_80_PERCENT,
          "ICE_BLUE" : Colour.ICE_BLUE,
          "INDIGO" : Colour.INDIGO,
          "IVORY" : Colour.IVORY,
          "LAVENDER" : Colour.LAVENDER,
          "LIGHT_BLUE" : Colour.LIGHT_BLUE,
          "LIGHT_GREEN" : Colour.LIGHT_GREEN,
          "LIGHT_ORANGE" : Colour.LIGHT_ORANGE,
          "LIGHT_TURQUOISE" : Colour.LIGHT_TURQUOISE,
          "LIME" : Colour.LIME,
          "OCEAN_BLUE" : Colour.OCEAN_BLUE,
          "OLIVE_GREEN" : Colour.OLIVE_GREEN,          
          "ORANGE" : Colour.ORANGE,
          "PALE_BLUE" : Colour.PALE_BLUE,
          "PALETTE_BLACK" : Colour.PALETTE_BLACK,
          "PERIWINKLE" : Colour.PERIWINKLE,
          "PINK" : Colour.PINK,
          "PLUM" : Colour.PLUM,
          "RED" : Colour.RED,
          "ROSE" : Colour.ROSE,
          "SEA_GREEN" : Colour.SEA_GREEN,
          "SKY_BLUE" : Colour.SKY_BLUE,
          "TAN" : Colour.TAN,
          "TEAL": Colour.TEAL,
          "TURQUOISE" : Colour.TURQUOISE,
          "UNKNOWN" : Colour.UNKNOWN, // This color is used when the user wants to use an RGB color
          "VERY_LIGHT_YELLOW" : Colour.VERY_LIGHT_YELLOW,
          "VIOLET" : Colour.VIOLET,
          "WHITE" : Colour.WHITE,
          "YELLOW" : Colour.YELLOW
     }

     var borderStylesObject = {
          "DASH_DOT" : BorderLineStyle.DASH_DOT,
          "DASH_DOT_DOT" : BorderLineStyle.DASH_DOT_DOT,
          "DASHED" : BorderLineStyle.DASHED,
          "DOTTED" : BorderLineStyle.DOTTED,
          "DOUBLE" : BorderLineStyle.DOUBLE,
          "HAIR" : BorderLineStyle.HAIR,
          "MEDIUM" : BorderLineStyle.MEDIUM,
          "MEDIUM_DASH_DOT" : BorderLineStyle.MEDIUM_DASH_DOT,
          "MEDIUM_DASH_DOT_DOT" : BorderLineStyle.MEDIUM_DASH_DOT_DOT,
          "MEDIUM_DASHED" : BorderLineStyle.MEDIUM_DASHED,
          "NONE" : BorderLineStyle.NONE,
          "SLANTED_DASH_DOT" : BorderLineStyle.SLANTED_DASH_DOT,
          "THICK" : BorderLineStyle.THICK,
          "THIN" : BorderLineStyle.THIN,
     }

     var underlineStylesObject = {
          "DOUBLE" : Packages.jxl.format.UnderlineStyle.DOUBLE,
          "DOUBLE_ACCOUNTING" : Packages.jxl.format.UnderlineStyle.DOUBLE_ACCOUNTING,
          "NO_UNDERLINE" : Packages.jxl.format.UnderlineStyle.NO_UNDERLINE,
          "SINGLE" : Packages.jxl.format.UnderlineStyle.SINGLE,
          "SINGLE_ACCOUNTING" : Packages.jxl.format.UnderlineStyle.SINGLE_ACCOUNTING,
     }

     var alignmentStylesObject = {
          "CENTER": Alignment.CENTRE,
          "FILL": Alignment.FILL,
          "GENERAL": Alignment.GENERAL,
          "JUSTIFY": Alignment.JUSTIFY,
          "LEFT": Alignment.LEFT,
          "RIGHT": Alignment.RIGHT,
     }

     if (style.Color != null) {
          if (style.Color.toString().indexOf(",") == -1 && colorObject[style.Color.toString().toUpperCase()] != null)
               color=colorObject[style.Color.toString().toUpperCase()];
          else {
          	   var rgb=style.Color.toString().split(",");
          	   
          	   workbook.setColourRGB(Colour.UNKNOWN,parseInt(rgb[0]),parseInt(rgb[1]),parseInt(rgb[2]));
          	   color=Colour.UNKNOWN;
          }
     } else
          color=Colour.BLACK;

     if (style.BackgroundColor != null)
          if (style.BackgroundColor.toString().indexOf(",") == -1 && colorObject[style.BackgroundColor.toString().toUpperCase()] != null)
               BGColor=colorObject[style.BackgroundColor.toString().toUpperCase()];
          else {
          	   var rgb=style.BackgroundColor.toString().split(",");
          	   
          	   workbook.setColourRGB(Colour.UNKNOWN,parseInt(rgb[0]),parseInt(rgb[1]),parseInt(rgb[2]));
          	   BGColor=Colour.UNKNOWN;
          }
     else
          BGColor=Colour.WHITE;

     var size=(style.Size != null ? style.Size : 12);     
     var bold=(style.Bold == true ? true : false);
     var italic=(style.Italic == true ? true : false);
     var underline=(style.Underline == true ? true : false);
     var borders=(style.Borders == true ? true : false);
     var alignment=(style.Alignment != null && alignmentStylesObject[style.Alignment.toString().toUpperCase()] != null ? alignmentStylesObject[style.Alignment.toString().toUpperCase()] : null);
     
     if (borders==true) {
          if (style.BorderStyle != null && borderStylesObject[style.BorderStyle.toString().toUpperCase()] != null)
               borderStyle=borderStylesObject[style.BorderStyle.toString().toUpperCase()];
          else
               borderStyle=BorderLineStyle.THIN;
     }

     var formatFont=new WritableFont(WritableFont.TIMES,size,(bold==true ? WritableFont.BOLD : WritableFont.NO_BOLD),italic);

     // Set the underline if specified
     if (underline == true) {
     	    if (style.UnderlineStyle != null && underlineStylesObject[style.UnderlineStyle] != null)
     	         formatFont.setUnderlineStyle(underlineStylesObject[style.UnderlineStyle]);
          else     	    
               formatFont.setUnderlineStyle(Packages.jxl.format.UnderlineStyle.SINGLE);
     }

     formatFont.setColour(color);

     var format=new WritableCellFormat(formatFont);
          
     format.setBackground(BGColor);
     
     if (borders == true)
          format.setBorder(Border.ALL,borderStyle);

     if (alignment != null)
          format.setAlignment(alignment);

     return format;
}
