function createExcelReport(reportObj) {	
     // Gets CellType based on name. Because of changes to referencing a cell type in POI 4, we have to get all Cell types and return the index for the cell type that we are looking for
     var getCellType=function(name) {
          var cellTypes=CellType.values();

          for (var i=0;i<cellTypes.length;i++)
               if (cellTypes[i]==name)
                    return cellTypes[i];          
     }
     
     // Returns color object
     var getColor=function(color) {
          // If color is not specified, it will default to black
          if (color==null) {
          	    print("Color not provided. Defaulting to black");
                color="black";
          }

          // RGB color
          if (color.indexOf(",") != -1) {
     	         var colorArr=color.split(",");
                
               return new XSSFColor(new java.awt.Color(parseFloat(colorArr[0]/255),parseFloat(colorArr[1]/255),parseFloat(colorArr[2]/255)), new org.apache.poi.xssf.usermodel.DefaultIndexedColorMap());
          }

          // Predefined color
          switch(color.toString().toUpperCase()) {
               case "AQUA":
    			          return HSSFColor.HSSFColorPredefined.AQUA.index;
    			          break;
 		           case "AUTOMATIC":
    			          return HSSFColor.HSSFColorPredefined.AUTOMATIC.index;
    			          break;
 		           case "BLACK":
    			          return HSSFColor.HSSFColorPredefined.BLACK.index;
    			          break;
 		           case "BLUE":
    			          return HSSFColor.HSSFColorPredefined.BLUE.index;
    			          break;
 		           case "BLUE_GREY":
    			          return HSSFColor.HSSFColorPredefined.BLUE_GREY.index;
    			          break;
 		           case "BRIGHT_GREEN":
    			          return HSSFColor.HSSFColorPredefined.BRIGHT_GREEN.index;
         			      break;
 		           case "BROWN":
    			          return HSSFColor.HSSFColorPredefined.BROWN.index;
    			          break;
 		           case "CORAL":
    			          return HSSFColor.HSSFColorPredefined.CORAL.index;
    			          break;
 		           case "CORNFLOWER_BLUE":
    			          return HSSFColor.HSSFColorPredefined.CORNFLOWER_BLUE.index;
    			          break;
     		       case "DARK_BLUE":
    			          return HSSFColor.HSSFColorPredefined.DARK_BLUE.index;
    			          break;
 		           case "DARK_GREEN":
    			          return HSSFColor.HSSFColorPredefined.DARK_GREEN.index;
    			          break;
 		           case "DARK_RED":
    			          return HSSFColor.HSSFColorPredefined.DARK_RED.index;
    			          break;
 		           case "DARK_TEAL":
    			          return HSSFColor.HSSFColorPredefined.DARK_TEAL.index;
    			          break;
 		           case "DARK_YELLOW":
    			          return HSSFColor.HSSFColorPredefined.DARK_YELLOW.index;
    			          break;
 		           case "GOLD":
    			          return HSSFColor.HSSFColorPredefined.GOLD.index;
    			          break;
 		           case "GREEN":
    			          return HSSFColor.HSSFColorPredefined.GREEN.index;
    			          break;
 		           case "GREY_25_PERCENT":
    			          return HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.index;
    			          break;
 		           case "GREY_40_PERCENT":
         			      return HSSFColor.HSSFColorPredefined.GREY_40_PERCENT.index;
    	     		      break;
 		           case "GREY_50_PERCENT":
    			          return HSSFColor.HSSFColorPredefined.GREY_50_PERCENT.index;
    			          break;
 		           case "GREY_80_PERCENT":
    			          return HSSFColor.HSSFColorPredefined.GREY_80_PERCENT.index;
           		      break;
 		           case "INDIGO":
    			          return HSSFColor.HSSFColorPredefined.INDIGO.index;
    			          break;
 		           case "LAVENDER":
    			          return HSSFColor.HSSFColorPredefined.LAVENDER.index;
    			          break;
 		           case "LEMON_CHIFFON":
    			          return HSSFColor.HSSFColorPredefined.LEMON_CHIFFON.index;
    			          break;
 		           case "LIGHT_BLUE":
    			          return HSSFColor.HSSFColorPredefined.LIGHT_BLUE.index;
    			          break;
 		           case "LIGHT_CORNFLOWER_BLUE":
    			          return HSSFColor.HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE.index;
    			          break;
 		           case "LIGHT_GREEN":
    			          return HSSFColor.HSSFColorPredefined.LIGHT_GREEN.index;
    			          break;
 		           case "LIGHT_ORANGE":
    			          return HSSFColor.HSSFColorPredefined.LIGHT_ORANGE.index;
    			          break;
 		           case "LIGHT_TURQUOISE":
    			          return HSSFColor.HSSFColorPredefined.LIGHT_TURQUOISE.index;
    			          break;
 		           case "LIGHT_YELLOW":
    			          return HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.index;
    			          break;
 		           case "LIME":
    			          return HSSFColor.HSSFColorPredefined.LIME.index;
    			          break;
 		           case "MAROON":
    			          return HSSFColor.HSSFColorPredefined.MAROON.index;
    			          break;
 		           case "OLIVE_GREEN":
         			      return HSSFColor.HSSFColorPredefined.OLIVE_GREEN.index;
    			          break;
 		           case "ORANGE":
    			          return HSSFColor.HSSFColorPredefined.ORANGE.index;
    			          break;
 		           case "ORCHID":
    			          return HSSFColor.HSSFColorPredefined.ORCHID.index;
    			          break;
 		           case "PALE_BLUE":
    			          return HSSFColor.HSSFColorPredefined.PALE_BLUE.index;
    			          break;
 		           case "PINK":
    			          return HSSFColor.HSSFColorPredefined.PINK.index;
    			          break;
 		           case "PLUM":
    			          return HSSFColor.HSSFColorPredefined.PLUM.index;
    			          break;
 		           case "RED":
    			          return HSSFColor.HSSFColorPredefined.RED.index;
    			          break;
 		           case "ROSE":
    			          return HSSFColor.HSSFColorPredefined.ROSE.index;
    			          break;
 		           case "ROYAL_BLUE":
    			          return HSSFColor.HSSFColorPredefined.ROYAL_BLUE.index;
    			          break;
 		           case "SEA_GREEN":
    			          return HSSFColor.HSSFColorPredefined.SEA_GREEN.index;
    			          break;
 		           case "SKY_BLUE":
    			          return HSSFColor.HSSFColorPredefined.SKY_BLUE.index;
    			          break;
 		           case "TAN":
    			          return HSSFColor.HSSFColorPredefined.TAN.index;
    			          break;
 		           case "TEAL":
    			          return HSSFColor.HSSFColorPredefined.TEAL.index;
    			          break;
 		           case "TURQUOISE":
    			          return HSSFColor.HSSFColorPredefined.TURQUOISE.index;
    			          break;
 		           case "VIOLET":
    			          return HSSFColor.HSSFColorPredefined.VIOLET.index;
    			          break;
 		           case "WHITE":
    			          return HSSFColor.HSSFColorPredefined.WHITE.index;
    			          break;
 		           case "YELLOW":
    			          return HSSFColor.HSSFColorPredefined.YELLOW.index;
 		           default:
 		                return HSSFColor.HSSFColorPredefined.BLACK.index;
          }
     }

     // Returns style object
     // Creates style from specified object
     var createStyleFormat=function (wb,styleDefinition) {	
          var alignmentStylesObject = {
               "CENTER": HorizontalAlignment.CENTER,
               "FILL": HorizontalAlignment.FILL,
               "GENERAL": HorizontalAlignment.GENERAL,
               "JUSTIFY": HorizontalAlignment.JUSTIFY,
               "LEFT": HorizontalAlignment.LEFT,
               "RIGHT": HorizontalAlignment.RIGHT,
          }
     
          var borderStylesObject = {
               "DASH_DOT" : BorderStyle.DASH_DOT,
               "DASH_DOT_DOT" : BorderStyle.DASH_DOT_DOT,
               "DASHED" : BorderStyle.DASHED,
               "DOTTED" : BorderStyle.DOTTED,
               "DOUBLE" : BorderStyle.DOUBLE,
               "HAIR" : BorderStyle.HAIR,
               "MEDIUM" : BorderStyle.MEDIUM,
               "MEDIUM_DASH_DOT" : BorderStyle.MEDIUM_DASH_DOT,
               "MEDIUM_DASH_DOT_DOT" : BorderStyle.MEDIUM_DASH_DOT_DOT,
               "MEDIUM_DASHED" : BorderStyle.MEDIUM_DASHED,
               "NONE" : BorderStyle.NONE,
               "SLANTED_DASH_DOT" : BorderStyle.SLANTED_DASH_DOT,
               "THICK" : BorderStyle.THICK,
               "THIN" : BorderStyle.THIN,
          }
     
          var underlineStylesObject = {
               "DOUBLE" : Packages.org.apache.poi.ss.usermodel.U_DOUBLE,
               "DOUBLE_ACCOUNTING" : Packages.org.apache.poi.ss.usermodel.U_DOUBLE_ACCOUNTING,
               "NO_UNDERLINE" : Packages.org.apache.poi.ss.usermodel.U_NONE,
               "SINGLE" : Packages.org.apache.poi.ss.usermodel.U_SINGLE,
               "SINGLE_ACCOUNTING" : Packages.org.apache.poi.ss.usermodel.U_SINGLE_ACCOUNTING,
          }

          var verticalAlignmentStylesObject = {
               "BOTTOM": Packages.org.apache.poi.ss.usermodel.BOTTOM,
               "CENTER": Packages.org.apache.poi.ss.usermodel.CENTER,
               "DISTRIBUTED": Packages.org.apache.poi.ss.usermodel.DISTRIBUTED,
               "JUSTIFY": Packages.org.apache.poi.ss.usermodel.JUSTIFY,
               "TOP": Packages.org.apache.poi.ss.usermodel.TOP,
          }
     
          var style=wb.createCellStyle();

          // Set the background color
          if (styleDefinition.BackgroundColor != null) {
               style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
               style.setFillForegroundColor(getColor(styleDefinition.BackgroundColor));
          }

          // A font object is needed for font specific styles
          var font=wb.createFont();

          // Set the foreground color
          if (styleDefinition.Color != null)
               font.setColor(getColor(styleDefinition.Color));
          else
     	         font.setColor(getColor("BLACK"));

          // Set the font size
          if (styleDefinition.Size != null)
               font.setFontHeightInPoints(styleDefinition.Size);
          else
               font.setFontHeightInPoints(12);

          // Set bold style
          if (styleDefinition.Bold == true)
               font.setBold(true);
          else
     	         font.setBold(false);

          // Set italic
          if (styleDefinition.Italic == true)
               font.setItalic(true);
          else
               font.setItalic(false);

          // Set underline
          if (styleDefinition.Underline == true)
               if (styleDefinition.UnderlineStyle != null && underlineStylesObject[styleDefinition.UnderlineStyle.toString().toUpperCase()] != null)
                    font.setUnderline(underlineStylesObject[styleDefinition.UnderlineStyle.toString().toUpperCase()]);
                else
                    font.setUnderline(underlineStylesObject["SINGLE"]);    

          // Set strikeout
          if (styleDefinition.Strikeout == true)
               font.setStrikeout(true);
           else
               font.setStrikeout(false);

          // Set borders
          if (styleDefinition.Borders == true) {
               if (styleDefinition.BorderStyle != null && borderStylesObject[styleDefinition.Borders.toString().toUpperCase()] != null) {
                    style.setBorderTop(borderStylesObject[styleDefinition.Borders.toString().toUpperCase()]);
                    style.setBorderBottom(borderStylesObject[styleDefinition.Borders.toString().toUpperCase()]);
                    style.setBorderLeft(borderStylesObject[styleDefinition.Borders.toString().toUpperCase()]);
                    style.setBorderRight(borderStylesObject[styleDefinition.Borders.toString().toUpperCase()]);
               } else {
                    style.setBorderTop(borderStylesObject["THIN"]);
                    style.setBorderBottom(borderStylesObject["THIN"]);
                    style.setBorderLeft(borderStylesObject["THIN"]);
                    style.setBorderRight(borderStylesObject["THIN"]);
               }
          }

          // Set alignment          
          if (styleDefinition.Alignment != null && alignmentStylesObject[styleDefinition.Alignment.toString().toUpperCase()] != null)
               style.setAlignment(alignmentStylesObject[styleDefinition.Alignment.toString().toUpperCase()]);
          else
               style.setAlignment(alignmentStylesObject["LEFT"]);


          // Set specific data format - Currently, only CURRENCY data type is supported
          if (styleDefinition.DataFormat != null) {
               df=wb.createDataFormat();
               
               switch (styleDefinition.DataFormat.toString().toUpperCase()) {
                    case "CURRENCY":
                         style.setDataFormat(df.getFormat("##,###,##0.00"));
                         break;
               }
          }

          // Set wrap style
          if (styleDefinition.Wrap == true) {
               style.setWrapText(true);
          }

          // Set vertical orientation
          if (styleDefinition.VerticalAlignment != null && verticalAlignmentStylesObject[styleDefinition.VerticalAlignment.toString().toUpperCase()] != null) {
               style.setVerticalAlignment(verticalAlignmentStylesObject[styleDefinition.VerticalAlignment.toString().toUpperCase()]);
          }

          // Set Rotation style 
          if (styleDefinition.Rotation != null && isInt(styleDefinition.Rotation) ) {
               style.setRotation(styleDefinition.Rotation);
          }
     
          style.setFont(font);
     
          return style;
     }
     
     var allRows,rows;
     var blacklistedColumn, blacklistedColumnCounter;
     var blacklistColumnIndexes=[];
     var namedStyles=[];
     var columnCounter,columnSizeCounter;
     var columnFound=false,columnNames,currColumnName,currColumnValue;
     var formulaCounter;
     var header, footer;
     var rowCounter=0;
     var today=new Date();
     //var todayStr=((today.getMonth()+1) < 10 ? "0" + (today.getMonth() + 1) : (today.getMonth()+1)) + "-" + (today.getDate() < 10 ? "0" + today.getDate() : today.getDate()) + "-" + today.getFullYear() + " " + today.getHours() + "-" + (today.getMinutes() < 10 ? "0" : "") + today.getMinutes() + "-" + today.getSeconds();     
     var todayStr=((today.getMonth()+1) < 10 ? "0" : "") + (today.getMonth() + 1) + "-" + (today.getDate() < 10 ? "0" : "") + today.getDate() + "-" + today.getFullYear() + " " + (today.getHours() < 10 ? "0" : "") + today.getHours() + "-" + (today.getMinutes() < 10 ? "0" : "") + today.getMinutes() + "-" + (today.getSeconds() < 10 ? "0" : "") + today.getSeconds();     
     
     var reportObjSheetCounter;
     var rowWritten=false;
     var sheet, workbook;
     var noData=false;
     var anyRowWritten=false;
     var currColumnIndex=0;
     
     // anchor types when placing an image in the sheet
     var anchorTypesObject = {
          "DONT_MOVE_AND_RESIZE" : ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE,
          "DONT_MOVE_DO_RESIZE" : ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE,
          "MOVE_AND_RESIZE" : ClientAnchor.AnchorType.MOVE_AND_RESIZE,
          "MOVE_DONT_RESIZE" : ClientAnchor.AnchorType.MOVE_DONT_RESIZE,          
     };
     
     // Needs to be declared here because it is used in createStyleFormat()
     workbook=XSSFWorkbook();
     
     // *** START OF STYLE DEFINITIONS ***
     var mainHeadingstyle=createStyleFormat(workbook,{Size:24,Bold: true,Color: "White"});

     // RGB shade doesnt work right
     var headerFormat=createStyleFormat(workbook,{Size:12, Bold: true,Color: "White", Borders: true, BackgroundColor: "83,162,240"},true);
     
     var cellFormat=createStyleFormat(workbook,{Size:12,Borders: true});

     var cellFormatNoBorder=createStyleFormat(workbook,{Size:16, Bold: true});

     var alignLeftFormat=createStyleFormat(workbook,{Size:12,Borders: true, Alignment: "left"});
     
     var cellCurrencyFormat=createStyleFormat(workbook,{DataFormat: "currency",Borders: true});

     var timeFont = workbook.createCellStyle();
     font = workbook.createFont();
     font.setFontHeightInPoints(12);
     timeFont.setFont(font);
		 timeFont.setDataFormat(workbook.createDataFormat().getFormat("hh:mm"));
		 timeFont.setBorderTop(BorderStyle.THIN);
     timeFont.setBorderBottom(BorderStyle.THIN);
     timeFont.setBorderLeft(BorderStyle.THIN);
     timeFont.setBorderRight(BorderStyle.THIN);
		 
     // *** END OF STYLE DEFINITIONS ***

     // Add all named styles to an array. Used with the validation
     if (typeof reportObj.NamedStyles != 'undefined') {
          for (namedStylesCounter=0;namedStylesCounter <  reportObj.NamedStyles.length;namedStylesCounter++) {
               styledObj=createStyleFormat(workbook,reportObj.NamedStyles[namedStylesCounter]);

               namedStyles.push([reportObj.NamedStyles[namedStylesCounter].Name,styledObj]);
          }
     }

     // *** START OF VALIDATION ***
	   if (typeof reportObj.FileName == 'undefined')
	        return ["ERROR","The property FileName was not specified"];

     // Remove these characters if they are going to be in the generated file name because they aren't isn't allowed in Windows
     reportObj.FileName=reportObj.FileName.replaceAll("\"","");
     reportObj.FileName=reportObj.FileName.replace("*",""); // This is here twice on purpose. Using replaceAll leads to an error Invalid quantifier * in the regex used by ReplaceAll
     reportObj.FileName=reportObj.FileName.replace("*","");
     reportObj.FileName=reportObj.FileName.replaceAll("/","");
     reportObj.FileName=reportObj.FileName.replaceAll("\\\\","");
     reportObj.FileName=reportObj.FileName.replaceAll(":","");
     reportObj.FileName=reportObj.FileName.replace("?","");
     reportObj.FileName=reportObj.FileName.replaceAll("<","");
     reportObj.FileName=reportObj.FileName.replaceAll(">","");
     reportObj.FileName=reportObj.FileName.replaceAll("|","");

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
          var nestedTables=false;
          
          if (typeof reportObj.Sheets[reportObjSheetCounter].NestedTables == 'undefined')
               nestedTables=false;
          else if (reportObj.Sheets[reportObjSheetCounter].NestedTables == true)
               nestedTables=true;
          
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

          if (typeof reportObj.Sheets[reportObjSheetCounter].FreezePane != 'undefined' && reportObj.Sheets[reportObjSheetCounter].FreezePane.length != 2) {
               return ["ERROR","The property FreezePane in Sheet " + reportObjSheetCounter + " need to specify a start and end row in the format FreezePane: [0,1]"];
          }

          // Validate columns for Non-nested tables
          if (nestedTables == false) {
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

               // If SQL was provided, make sure that all of the necessary properties were provided
               if (typeof reportObj.Sheets[reportObjSheetCounter].SQL != 'undefined') {
                    // Validate that DBConnection was provided	                       
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].DBConnection == 'undefined')
	                       return ["ERROR","The property DBConnection in Sheet " + reportObjSheetCounter + " was not specified"];
               }
          } else { // Nested table
          	   // Parent required properties
          	   
          	   // Validate that ColumnHeadersParent and ColumnHeadersChild were provided	                       
	             if (typeof reportObj.Sheets[reportObjSheetCounter].ColumnHeadersParent == 'undefined')
	                  return ["ERROR","The property ColumnHeadersParent in Sheet " + reportObjSheetCounter + " was not specified"];

               // Validate that Columns 1 & 2 were provided	                       
	             if (typeof reportObj.Sheets[reportObjSheetCounter].ColumnsParent == 'undefined')
	                  return ["ERROR","The property ColumnsParent in Sheet " + reportObjSheetCounter + " was not specified"];
	                  
          	   // Child required properties
          	   if (typeof reportObj.Sheets[reportObjSheetCounter].ColumnHeadersChild == 'undefined')
	                  return ["ERROR","The property ColumnHeadersChild in Sheet " + reportObjSheetCounter + " was not specified"];

	             if (typeof reportObj.Sheets[reportObjSheetCounter].ColumnsChild == 'undefined')
	                  return ["ERROR","The property ColumnsChild in Sheet " + reportObjSheetCounter + " was not specified"];

               if (typeof reportObj.Sheets[reportObjSheetCounter].TableDataParent != 'undefined' && typeof reportObj.Sheets[reportObjSheetCounter].SQLParent == 'undefined') {
                    return ["ERROR","The properties TableDataParent and SQLParent in Sheet " + reportObjSheetCounter + " cannot be specified together."];
               }
               
               // Nested tables - SQL based data
               if (typeof reportObj.Sheets[reportObjSheetCounter].TableDataParent == 'undefined') {
               	    if (typeof reportObj.Sheets[reportObjSheetCounter].SQLParent == 'undefined')
	                       return ["ERROR","The property SQLParent in Sheet " + reportObjSheetCounter + " was not specified"];

	                  if (typeof reportObj.Sheets[reportObjSheetCounter].SQLChild == 'undefined')
	                       return ["ERROR","The property SQLChild in Sheet " + reportObjSheetCounter + " was not specified"];

	                  if (typeof reportObj.Sheets[reportObjSheetCounter].DBConnectionParent == 'undefined')
	                       return ["ERROR","The property DBConnectionParent in Sheet " + reportObjSheetCounter + " was not specified"];

	                  if (typeof reportObj.Sheets[reportObjSheetCounter].DBConnectionChild == 'undefined')
	                       return ["ERROR","The property DBConnectionChild in Sheet " + reportObjSheetCounter + " was not specified"];
               } else { // Nested tables - Table based data
               	    if (tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableDataParent) == null)
               	          return ["ERROR","The property TableDataParent in Sheet " + reportObjSheetCounter + " refers to an invalid table"];
               	          
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].TableDataChild == 'undefined')
	                        return ["ERROR","The property TableDataChild in Sheet " + reportObjSheetCounter + " was not specified"];

	                  if (tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableDataChild) == null)
               	          return ["ERROR","The property TableDataChild in Sheet " + reportObjSheetCounter + " refers to an invalid table"];
               	          
	                  if (typeof reportObj.Sheets[reportObjSheetCounter].SQLChild != 'undefined')
	                       return ["ERROR","The property SQLChild in Sheet " + reportObjSheetCounter + " is not valid when using table data"];	                  

	                  if (typeof reportObj.Sheets[reportObjSheetCounter].DBConnectionChild != 'undefined')
	                       return ["ERROR","The property DBConnectionChild in Sheet " + reportObjSheetCounter + " is not valid when using table data"];
               }

               // Validate column sizes	             

               // Validate that the size of Columns and ColumnHeaders match
               if (reportObj.Sheets[reportObjSheetCounter].ColumnsParent.length != reportObj.Sheets[reportObjSheetCounter].ColumnHeadersParent.split(",").length)
                    return ["ERROR","The properties ColumnsParent and ColumnHeadersParent in Sheet " + reportObjSheetCounter + " are of different lengths. ColumnParent length=" + reportObj.Sheets[reportObjSheetCounter].ColumnsParent.length + " and ColumnHeadersParent length=" + reportObj.Sheets[reportObjSheetCounter].ColumnHeadersParent.split(",").length];

               if (reportObj.Sheets[reportObjSheetCounter].ColumnsChild.length != reportObj.Sheets[reportObjSheetCounter].ColumnHeadersChild.split(",").length)
                    return ["ERROR","The properties ColumnsChild and ColumnHeadersChild in Sheet " + reportObjSheetCounter + " are of different lengths. Column2 length=" + reportObj.Sheets[reportObjSheetCounter].ColumnsChild.length + " and ColumnHeadersChild length=" + reportObj.Sheets[reportObjSheetCounter].ColumnHeadersChild.split(",").length];

               // Join Where clause must always be specified and have 2 arguments
               if (typeof reportObj.Sheets[reportObjSheetCounter].JoinWhereClause == 'undefined')
	                  return ["ERROR","The property JoinWhereClause in Sheet " + reportObjSheetCounter + " was not specified"];

	             if (reportObj.Sheets[reportObjSheetCounter].JoinWhereClause.length != 2)
	                  return ["ERROR","The property JoinWhereClause in Sheet " + reportObjSheetCounter + " must have 2 properties ['MATCHINGCOLUM',parentIndex]"];
          }
          
          // Validate merge cells
          if (typeof reportObj.Sheets[reportObjSheetCounter].MergeCells != 'undefined') {
               for (var mergeCounter=0;mergeCounter < reportObj.Sheets[reportObjSheetCounter].MergeCells.length;mergeCounter++) {
                    var len=reportObj.Sheets[reportObjSheetCounter].MergeCells[mergeCounter].split(",").length;

                    if (len != 2 && len != 4)
	                       return ["ERROR","The property MergeCells in Sheet " + reportObjSheetCounter + " at index " + mergeCounter + " has an invalid size. You can either specify MergeCells:\"2,4\" to merge rows 2-4 on the same row or MergeCells:\"2,2,4,4\" to merge from cell 2,2-4,4"];
               }
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

          // If conditional formatting is provided, validate its properties
          if (typeof reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting != 'undefined') {
               for (cfCounter=0;cfCounter < reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting.length;cfCounter++) {
                    // Formula must be provided
                    if (typeof reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Formula == 'undefined')
                         return ["ERROR","The conditional formatting at index " + cfCounter + " in sheet " + reportObjSheetCounter + " does not have a Formula property"];

                    if (typeof reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style == 'undefined')
                         return ["ERROR","The conditional formatting at index " + cfCounter + " in sheet " + reportObjSheetCounter + " does not have a Style property"];

                    if (typeof reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].StartRow == 'undefined')
                         return ["ERROR","The conditional formatting at index " + cfCounter + " in sheet " + reportObjSheetCounter + " does not have a StartRow property"];

                    if (typeof reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].StartColumn == 'undefined')
                         return ["ERROR","The conditional formatting at index " + cfCounter + " in sheet " + reportObjSheetCounter + " does not have a StartColumn property"];

                    if (typeof reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].EndColumn == 'undefined')
                         return ["ERROR","The conditional formatting at index " + cfCounter + " in sheet " + reportObjSheetCounter + " does not have a EndColumn property"];
               }
          }

          // If image property is provided, validate its properties
          if (typeof reportObj.Sheets[reportObjSheetCounter].Image != 'undefined') {
               for (imgCounter=0;imgCounter < reportObj.Sheets[reportObjSheetCounter].Image.length;imgCounter++) {
                    // Validate that file name was provided
                    if (typeof reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].FileName == 'undefined')
                         return ["ERROR","The Image at index " + imgCounter + " in sheet " + reportObjSheetCounter + " does not have a FileName property"];

                    // Validate that the file is a valid file
                    if (FileServices.isFile(reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].FileName)==false)
                         return ["ERROR","The Image at index " + imgCounter + " in sheet " + reportObjSheetCounter + " refers to a file that does not exist"];
                    
                    var ext=reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].FileName.substring(reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].FileName.lastIndexOf(".")+1).toUpperCase();
                         
                    if (ext != "DIB" && ext != "EMF" && ext != "JPEG" && ext != "JPG" && ext != "PICT" && ext != "PNG" && ext != "WHF" )
                         return ["ERROR","The Image at index " + imgCounter + " in sheet " + reportObjSheetCounter + " refers to a file with an unknown format. Valid image formats are DIB,EMF,JPEG/JPG,PICT,PNG OR WMF."];
                         
                    //if (typeof reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].AnchorType == 'undefined')
                         //return ["ERROR","The Image at index " + imgCounter + " in sheet " + reportObjSheetCounter + " does not have an AnchorType property"];

                    
                    if (anchorTypesObject[reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].AnchorType.toString().toUpperCase()] == null)
                         return ["ERROR","The Image at index " + imgCounter + " in sheet " + reportObjSheetCounter + " does not have a valid value for the AnchorType property. Valid values are: DONT_MOVE_AND_RESIZE, DONT_MOVE_DO_RESIZE, MOVE_AND_RESIZE or MOVE_DONT_RESIZE"];

                    if (typeof reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].StartRow == 'undefined')
                         return ["ERROR","The Image at index " + imgCounter + " in sheet " + reportObjSheetCounter + " does not have a StartRow property"];

                    if (typeof reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].StartColumn == 'undefined')
                         return ["ERROR","The Image at index " + imgCounter + " in sheet " + reportObjSheetCounter + " does not have a StartColumn property"];
               }
          }
	   }

	   // If 1 or more custom cell texts were specified, validate the custom cell text properties
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
     reportObj.FileName=reportObj.FileName + " as of " + todayStr + ".xlsx";

     // *** START OF GENERATING THE EXCEL DOCUMENT ***

     // *** Start of Loop through the report object for each sheet object ***
     for (reportObjSheetCounter=0;reportObj.Sheets[reportObjSheetCounter] != null;reportObjSheetCounter++) {
     	    rowCounter=0;
     	    
          // Create the sheet based on the specified name and index
          sheet = workbook.createSheet(reportObj.Sheets[reportObjSheetCounter].SheetName);

          // *** START OF SETTING THE HEADER AND FOOTER ***
          header = sheet.getHeader();
          footer = sheet.getFooter();

          // *** END OF SETTING THE HEADER AND FOOTER ***

          // *** START OF SETTING THE MARGINS ***
          if (typeof reportObj.Sheets[reportObjSheetCounter].TopMargin !== 'undefined' && isInt(reportObj.Sheets[reportObjSheetCounter].TopMargin))
               sheet.setMargin(sheet.TopMargin,parseInt(reportObj.Sheets[reportObjSheetCounter].TopMargin));

          if (typeof reportObj.Sheets[reportObjSheetCounter].BottomMargin !== 'undefined' && isInt(reportObj.Sheets[reportObjSheetCounter].BottomMargin))
               sheet.setMargin(sheet.BottomMargin,parseInt(reportObj.Sheets[reportObjSheetCounter].BottomMargin));

          if (typeof reportObj.Sheets[reportObjSheetCounter].LeftMargin !== 'undefined' && isInt(reportObj.Sheets[reportObjSheetCounter].LeftMargin))
               sheet.setMargin(sheet.LeftMargin,parseInt(reportObj.Sheets[reportObjSheetCounter].LeftMargin));

          if (typeof reportObj.Sheets[reportObjSheetCounter].RightMargin !== 'undefined' && isInt(reportObj.Sheets[reportObjSheetCounter].RightMargin))
               sheet.setMargin(sheet.RightMargin,parseInt(reportObj.Sheets[reportObjSheetCounter].RightMargin));

          // All margins
          if (typeof reportObj.Sheets[reportObjSheetCounter].AllMargins !== 'undefined') {
               var len=reportObj.Sheets[reportObjSheetCounter].AllMargins.split(",").length;

               if (len==1) {
               	    var val=parseInt(reportObj.Sheets[reportObjSheetCounter].AllMargins);

               	    sheet.setMargin(sheet.TopMargin,parseInt(val));
               	    sheet.setMargin(sheet.BottomMargin,parseInt(val));
               	    sheet.setMargin(sheet.LeftMargin,parseInt(val));
               	    sheet.setMargin(sheet.RightMargin,parseInt(val));
               } else if (len==2) {
               	    var mergeCell=reportObj.Sheets[reportObjSheetCounter].AllMargins.split(",");

                    sheet.setMargin(sheet.TopMargin,parseInt(mergeCell[0]));
               	    sheet.setMargin(sheet.BottomMargin,parseInt(mergeCell[0]));
               	    sheet.setMargin(sheet.LeftMargin,parseInt(mergeCell[1]));
               	    sheet.setMargin(sheet.RightMargin,parseInt(mergeCell[1]));
               } else if (len==4) {
                    var mergeCell=reportObj.Sheets[reportObjSheetCounter].AllMargins.split(",");

                    sheet.setMargin(sheet.TopMargin,parseInt(mergeCell[0]));
               	    sheet.setMargin(sheet.BottomMargin,parseInt(mergeCell[1]));
               	    sheet.setMargin(sheet.LeftMargin,parseInt(mergeCell[2]));
               	    sheet.setMargin(sheet.RightMargin,parseInt(mergeCell[3]));
               }
          }

          // HeaderMargin
          if (typeof reportObj.Sheets[reportObjSheetCounter].HeaderMargin !== 'undefined')
               sheet.setMargin(sheet.HeaderMargin,parseInt(reportObj.Sheets[reportObjSheetCounter].HeaderMargin));

          // FooterMargin
          if (typeof reportObj.Sheets[reportObjSheetCounter].FooterMargin !== 'undefined')
               sheet.setMargin(sheet.FooterMargin,parseInt(reportObj.Sheets[reportObjSheetCounter].HeaderMargin));
          // *** END OF SETTING THE MARGINS ***         

          // *** START OF  SETTING THE HEADER ***
          if (typeof reportObj.Sheets[reportObjSheetCounter].HeaderLeft !== 'undefined')
               header.setLeft(reportObj.Sheets[reportObjSheetCounter].HeaderLeft);
          else
               header.setLeft("&F"); // Add the workbook name to the top right of the header

          if (typeof reportObj.Sheets[reportObjSheetCounter].HeaderCenter !== 'undefined')
               header.setCenter(reportObj.Sheets[reportObjSheetCounter].HeaderCenter);

          if (typeof reportObj.Sheets[reportObjSheetCounter].HeaderRight !== 'undefined')
               header.setRight(reportObj.Sheets[reportObjSheetCounter].HeaderRight);
          else
               header.setRight("&A"); // Add the sheet name to the top right of the header
          
          // *** END OF  SETTING THE HEADER ***

          // *** START OF  SETTING THE FOOTER ***
          if (typeof reportObj.Sheets[reportObjSheetCounter].FooterLeft !== 'undefined')
               footer.setLeft(reportObj.Sheets[reportObjSheetCounter].FooterLeft);

          if (typeof reportObj.Sheets[reportObjSheetCounter].FooterCenter !== 'undefined')
               footer.setCenter(reportObj.Sheets[reportObjSheetCounter].FooterCenter);
           else
               footer.setCenter("Page &P / &N"); // Add the page numbers to the footer to Page <current page #> / <Total # of pages>

          if (typeof reportObj.Sheets[reportObjSheetCounter].FooterRight !== 'undefined')
               footer.setRight(reportObj.Sheets[reportObjSheetCounter].FooterRight);
          // *** END OF  SETTING THE FOOTER ***
          
          // *** START OF MISC SHEET OPTIONS ***
          // FitWidth
          if (typeof reportObj.Sheets[reportObjSheetCounter].FitWidth !== 'undefined' && reportObj.Sheets[reportObjSheetCounter].FitWidth == true)
               sheet.getPrintSetup().setFitWidth(1);

          // FitHeight
          if (typeof reportObj.Sheets[reportObjSheetCounter].FitHeight !== 'undefined' && reportObj.Sheets[reportObjSheetCounter].FitHeight == true)
               sheet.getPrintSetup().setFitHeight(1);

          // FitToPages
          if (typeof reportObj.Sheets[reportObjSheetCounter].FitToPages !== 'undefined' && reportObj.Sheets[reportObjSheetCounter].FitToPages == true)
               sheet.setFitToPage(true);

          // Orientation
          if (typeof reportObj.Sheets[reportObjSheetCounter].Orientation !== 'undefined')
               if (reportObj.Sheets[reportObjSheetCounter].Orientation.toString().toUpperCase() == "PORTRAIT")
                    sheet.getPrintSetup().setLandscape(false);
               else if (reportObj.Sheets[reportObjSheetCounter].Orientation.toString().toUpperCase() == "LANDSCAPE")
                    sheet.getPrintSetup().setLandscape(true);

          // Password
          if (typeof reportObj.Sheets[reportObjSheetCounter].Password !== 'undefined')
               sheet.protectSheet(reportObj.Sheets[reportObjSheetCounter].Password);

          if (typeof reportObj.Sheets[reportObjSheetCounter].FreezePane != 'undefined')
               sheet.createFreezePane(reportObj.Sheets[reportObjSheetCounter].FreezePane[0],reportObj.Sheets[reportObjSheetCounter].FreezePane[1]);
          
          // *** END OF MISC SHEET OPTIONS ***

          // *** START OF SHEET HEADING ***
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
                    styledFormat=createStyleFormat(workbook,sheetHeader.Style[0]);
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

               var row = sheet.createRow(sheetHeader.Row);
               var cell = row.createCell(sheetHeader.Column);

               // In order to prevent errors, always default the type to CHAR if not specified.
                if (sheetHeader.DataType==null) sheetHeader.DataType="CHAR";
                    
               // Write the heading factoring in the DataType property
               switch (sheetHeader.DataType.toUpperCase()) {
                    case "BOOLEAN":
                    case "INT":
                    case "INTEGER":
                    case "NUMERIC":
                        cell.setCellValue(sheetHeader.Value);
                        cell.setCellStyle((styledFormat != null ? styledFormat : mainHeadingStyle));
                        cell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC);
                         break
                    case "DATE":
              	    case "DATETIME":
                         dateVal=new Date(sheetHeader.Value).mmddyyyy();
                         cell.setCellValue(dateVal);
                         cell.setCellStyle((styledFormat != null ? styledFormat : mainHeadingStyle));
                         break;
                    case "CHAR":
                    default:
                         cell.setCellValue(sheetHeader.Value);
                         
                         if (styledFormat != null)
                              cell.setCellStyle(styledFormat);
               }

              // Merge cells if specifed
              if (sheetHeader.MergeCells != null) {
                   var mergeCell=sheetHeader.MergeCells.split(",");

                   if (mergeCell.length == 2)
                        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(sheetHeader.Row,sheetHeader.Row,parseInt(mergeCell[0]),parseInt(mergeCell[1])));
                   else
                        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(parseInt(mergeCell[0]),parseInt(mergeCell[1]),parseInt(mergeCell[2]),parseInt(mergeCell[3])));
              }
              
              //rowCounter++;
          }
          // *** END OF SHEET HEADING ***

          // Optional setting that If provided, indicates which row to start writing the data to
          if (reportObj.Sheets[reportObjSheetCounter].StartRow != null)
               rowCounter=parseInt(reportObj.Sheets[reportObjSheetCounter].StartRow)-1;

          // *** START OF TABLE COLUMN HEADERS ***
          if (nestedTables == false) { // Non-nested table column headers
               var columnHeaders=reportObj.Sheets[reportObjSheetCounter].ColumnHeaders.split(",");

               row=sheet.createRow(rowCounter);
          
               for (columnCounter=0;columnCounter<columnHeaders.length;columnCounter++) {
                    try {                    
                         if (columnHeaders[columnCounter].toUpperCase() != "WHERECLAUSE") {
                    	        // When TableData is passed to createExcelReport() validate the supplied data type vs the actual data type
                    	        if (reportObj.Sheets[reportObjSheetCounter].TableData != null && tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableData) != null && tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableData).getColumn(columnHeaders[columnCounter]) != null) {
                    	             actualDataType=tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableData).getColumn(columnHeaders[columnCounter]).type;
                    	             reportedDataType=reportObj.Sheets[reportObjSheetCounter].Columns[columnCounter][1];

                    	             if (reportedDataType != reportedDataType) {
                    	                  if (actualDataType == "DATE" && reportedDataType != "DATE")
                    	                       print("Data type error found in createExcelReport for the report " + reportObj.FileName + ": The column " + reportObj.Sheets[reportObjSheetCounter].Columns[columnCounter][0] + " has a data type of " + tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableData).getColumn(reportObj.Sheets[reportObjSheetCounter].Columns[columnCounter][0]).type + " but the data type passed to createExcelReport() is " + reportObj.Sheets[reportObjSheetCounter].Columns[columnCounter][1]);
                    	             }
                    	        }

                              row.createCell(columnCounter).setCellValue(columnHeaders[columnCounter]);

                              row.getCell(columnCounter).setCellStyle(headerFormat);
                         }
                    } catch (e) {
                         workbook.close();
                    
                         if (FileServices.existsFile(reportObj.FileName))
                              FileServices.deleteFile(reportObj.FileName);
                    
                         return ["ERROR","The error " + e + " occurred when table is " + reportObj.Sheets[reportObjSheetCounter].TableData + ", columnHeaders= " + columnHeaders + ", length=" + columnHeaders.length + ",columnCounter="+columnCounter+ " and value=" + columnHeaders[columnCounter]];
                    }
               }

               rowCounter++;
          } else { // We will write the headers later when using nested tables
               var columnHeadersParent=reportObj.Sheets[reportObjSheetCounter].ColumnHeadersParent.split(",");               
          }
          
          // *** END OF TABLE COLUMN HEADERS ***

          // *** START OF GET STYLE FOR CURRENT SHEET ***
          // Get the styled format if provided
          if (reportObj.Sheets[reportObjSheetCounter].Style != null) {
               styledFormat=createStyleFormat(workbook,reportObj.Sheets[reportObjSheetCounter].Style[0]);
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

          // *** END OF GET STYLE FOR CURRENT SHEET ***

          // *** START OF GET ALL DATA ***

          // *** START OF GET ALL COLUMNS ***
          if (nestedTables == false) { // Non-nested table columns
               // Get all columns
               try {
                    var columns=eval(reportObj.Sheets[reportObjSheetCounter].Columns);
               } catch (e) {
                    return ["ERROR","An error occurred when Columns= " + reportObj.Sheets[reportObjSheetCounter].Columns];
               }
          } else {
               // Get all columns
               try {
                    var columnsParent=eval(reportObj.Sheets[reportObjSheetCounter].ColumnsParent);
               } catch (e) {
                    return ["ERROR","An error occurred when ColumnsParent= " + reportObj.Sheets[reportObjSheetCounter].Parent];
               }

               try {
                    var columnsChild=eval(reportObj.Sheets[reportObjSheetCounter].ColumnsChild);
               } catch (e) {
                    return ["ERROR","An error occurred when ColumnsChild= " + reportObj.Sheets[reportObjSheetCounter].ColumnsChild];
               }
          }
          // *** END OF GET ALL COLUMNS ***

          // *** Save all of the data in an array ***
          var data = [];

          if (nestedTables == false) {
               // *** SQL based data ***
               if (reportObj.Sheets[reportObjSheetCounter].SQL != null) {
                    columnNotFound=false;
                    invalidColumn="";

                    try {
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
                                             // Added try catch around this line because there seems to be a global recursion on it
                                             try {
                                                  // add column name, column value, column type and date format for date type
                              	                  lineArr = new Array(columns[key][0],(columnData[columns[key][0]] != null ? columnData[columns[key][0]] : null),columns[key][1],columns[key][2],(reportObj.Sheets[reportObjSheetCounter].Columns[key][2] != null ? reportObj.Sheets[reportObjSheetCounter].Columns[key][2] : null));
                                             } catch (e) {
                                                  return ["ERROR","An error occurred generating the report with the error " + e + " while adding the current row to lineArr"];
                                             }

                                             rowArr.push(lineArr);
                                   
                                             columnFound=true;
                                   
                                             break;
                                        }
                                   }

                                   if (columnFound == false && columns[key][0] != null) {
                         	              invalidColumn=columns[key][0];
                                        break;
                                   }
                              }

                              if (columnFound == false && columns[key][0] != null)
                                   return;

                              // push line array
                              data.push(rowArr);
                         });
                    } catch(e) {
               	         workbook.close();
                         
                         if (FileServices.existsFile(reportObj.FileName))
                              FileServices.deleteFile(reportObj.FileName);
                         
                          return ["ERROR","An error occurred generating the report with the error " + e + " while executing the SQL statement " + reportObj.Sheets[reportObjSheetCounter].SQL];
                    }
               
                    if (columnFound == false && invalidColumn != null && invalidColumn != "")
                         return ["ERROR","The sql column " + invalidColumn + " was not found in the database. Please check the spelling of the column name"];
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
                         	         if (columns[columnCounter] == null || typeof columns[columnCounter][0] == 'undefined' || columns[columnCounter][0] == null ) {
                         	              continue;
                         	         }                         	         
                         	         
                         	         // Make sure that the column name is valid
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

                              // add column name, column value, column type and force type if set
                              lineArr = new Array(columns[columnCounter][0],(currColumnValue != null ? currColumnValue : null), currColumnType,(columns[columnCounter][2] != null ? columns[columnCounter][2] : null));

                              rowArr.push(lineArr);
                         }                    

                         // push line array
                         data.push(rowArr);
                    }
               }
          } else { // Nested tables
          	   columnNotFound=false;
               invalidColumn="";

          	   if (typeof reportObj.Sheets[reportObjSheetCounter].SQLParent != 'undefined') {
                    try {
               	         var parentSQL=reportObj.Sheets[reportObjSheetCounter].SQLParent;
               	    
                         // Read the data
                         services.database.executeSelectStatement(reportObj.Sheets[reportObjSheetCounter].DBConnectionParent,parentSQL,
                         function (columnData) {
               	              lineArr = [];
               	              rowArr = [];

               	              // Loop through all columns in provided column list
                              for (var key in columnsParent) {
                    	             columnFound=false;
                    	   
                    	             // Loop through each column of the current row
                                   for (var colName in columnData) {
                         	              // If the current column is the one we are looking for
                                        if (columnsParent[key][0] != null && columnsParent[key][0].toUpperCase() == colName.toUpperCase()) {
                                             // add column name, column value, column type and date format for date type
                              	             lineArr = new Array(columnsParent[key][0],(columnData[columnsParent[key][0]] != null ? columnData[columnsParent[key][0]] : null),columnsParent[key][1],columnsParent[key][2],(reportObj.Sheets[reportObjSheetCounter].ColumnsParent[key][2] != null ? reportObj.Sheets[reportObjSheetCounter].ColumnsParent[key][2] : null),(reportObj.Sheets[reportObjSheetCounter].ColumnsParent[key][3] != null ? reportObj.Sheets[reportObjSheetCounter].ColumnsParent[key][3] : null));

                                             rowArr.push(lineArr);
                                   
                                             columnFound=true;
                                   
                                             break;
                                        }
                                   }

                                   // Exit for loop if invalid column was found
                                   if (columnFound == false && columnsParent[key][0] != null) {
                         	              invalidColumn=columnsParent[key][0];
                                        break;
                                   }
                              }

                              // Exit function loop if invalid column was found
                              if (columnFound == false && columnsParent[key][0] != null)
                                   return;

                              // push line array
                              data.push(rowArr);
                         });
                    } catch(e) {
               	         workbook.close();
                         
                         if (FileServices.existsFile(reportObj.FileName))
                              FileServices.deleteFile(reportObj.FileName);
                         
                         return ["ERROR","An error occurred generating the report with the error " + e + " while executing the SQL statement " + reportObj.Sheets[reportObjSheetCounter].SQLParent];
                    }
               
                    //if (columnFound == false && invalidColumn != null && invalidColumn != "")
                    //     return ["ERROR","The sql column " + invalidColumn + " was not found in the database. Please check the spelling of the column name"];
               } else if (typeof reportObj.Sheets[reportObjSheetCounter].TableDataParent != 'undefined') {
                    var allRows=tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableDataParent);
                    var rows=allRows.getRows();

                    while (rows.next()) {
                    	   lineArr = [];
               	         rowArr = [];
               	         
                         // Loop through all columns in provided column list
                         for (var key in columnsParent) {
                              if (allRows.getColumn(columnsParent[key][0]) != null) {
                                   // add column name, column value, column type and date format for date type
                              	   lineArr = new Array(columnsParent[key][0],(allRows.getColumn(columnsParent[key][0]).value != null ? allRows.getColumn(columnsParent[key][0]).value : null),columnsParent[key][1],columnsParent[key][2],(reportObj.Sheets[reportObjSheetCounter].ColumnsParent[key][2] != null ? reportObj.Sheets[reportObjSheetCounter].ColumnsParent[key][2] : null),(reportObj.Sheets[reportObjSheetCounter].ColumnsParent[key][3] != null ? reportObj.Sheets[reportObjSheetCounter].ColumnsParent[key][3] : null));

                                   rowArr.push(lineArr);
                              } else {
                                   invalidColumn=columnsParent[key][0];
                                   break;
                              }
                         }

                         // push line array
                         data.push(rowArr); 
                    }                                        
               }
          }
          // *** END OF GET ALL DATA ***
          
          // Output the data
          //for (var dataCounter=0;dataCounter<data.length;dataCounter++) {
          //       for (var colCounter=0;colCounter<columnHeaders.length;colCounter++) {
          //            alert("[" + colCounter + "]=" + data[dataCounter][colCounter]);
          //       }
          //}

          currColumnIndex=0;
          
          rowWritten=false;

          // *** START OF LOOP THAT GOES THROUGH DATA ARRAY AND WRITES THE DATA ***
          for (var dataCounter=0;dataCounter<data.length;dataCounter++) {
          	   if (nestedTables==true) { // Nested tables parent header
          	        if (dataCounter > 0)
          	             rowCounter++;
          	        
          	        row=sheet.getRow(rowCounter) != null ? sheet.createRow(rowCounter) : sheet.createRow(rowCounter);

                    // Write column headers
                    for (columnCounter=0;columnCounter<columnHeadersParent.length;columnCounter++) {
                         if (reportObj.Sheets[reportObjSheetCounter].ColumnsParent[columnCounter] != null && reportObj.Sheets[reportObjSheetCounter].ColumnsParent[columnCounter][3] == true)
                              continue;
                    	   //currColumnIndex=0;
                    	   
                         try {                    
                              row.createCell(columnCounter).setCellValue(columnHeadersParent[columnCounter]);

                              row.getCell(columnCounter).setCellStyle(headerFormat);
                         } catch (e) {
                              workbook.close();
                    
                              if (FileServices.existsFile(reportObj.FileName))
                                   FileServices.deleteFile(reportObj.FileName);
                    
                              return ["ERROR","The error " + e + " occurred when sql is " + reportObj.Sheets[reportObjSheetCounter].SQLParent + ", columnHeaders= " + columnHeadersParent + ", length=" + columnHeadersParent.length + ",columnCounter="+columnCounter+ " and value=" + columnHeadersParent[columnCounter]];
                         }
                    }

                    rowCounter++;
          	   }
          	   
               row = (sheet.getRow(rowCounter) != null ? sheet.getRow(rowCounter) : sheet.createRow(rowCounter));
               
          	   if (currColumnIndex == (nestedTables == false ? columns.length : columnsParent.length))
          	        currColumnIndex=0;  

               var parentMatchingValue=-1;

               // Write the data (or parent data if using nested tables)
               for (var colCounter=0;colCounter<(nestedTables == false ? columnHeaders.length : columnHeadersParent.length);colCounter++) {               	    
               	    if (data[dataCounter][colCounter]==null) {
               	         continue;
               	    }

                    // If the Column is equal to [null], increment currColumnIndex so the data is shifted to the right
               	    if (nestedTables == false) {
               	         while (currColumnIndex<columns.length && columns[currColumnIndex][0] == null) {
               	              currColumnIndex++;
               	         }
               	    } else {
               	    	  // Only assign to this value once per row
               	    	  if (colCounter == parseInt(reportObj.Sheets[reportObjSheetCounter].JoinWhereClause[1]) && parentMatchingValue==-1)
               	    	       parentMatchingValue=data[dataCounter][colCounter][1];
               	    	  
               	    	  while (reportObj.Sheets[reportObjSheetCounter].ColumnsParent[currColumnIndex][0] == null) {
               	              currColumnIndex++;
               	         }
               	    }

               	    if (data[dataCounter][colCounter][5]==true) {
               	         continue;
               	    }
                    
                    // If the type is CHAR but the value is an INT, change the type to an INT so it will be written as an INT so
                    // that Excel doesn't complaign that the field is a number in a text cell
                    if (data[dataCounter][colCounter][2] == "CHAR" && data[dataCounter][colCounter][1] != null && isInt(data[dataCounter][colCounter][1]) && data[dataCounter][colCounter][3] != true)
                         data[dataCounter][colCounter][2]="INTEGER";

                    // If the type is INT but the value is a CHAR, change the type to an CHAR so it will be written as a CHAR. Ignore percentage values because we want to still write them as a number
                    if ((data[dataCounter][colCounter][2] == "INTEGER" || data[dataCounter][colCounter][2] == "INT") && data[dataCounter][colCounter][1] != null && !isInt(data[dataCounter][colCounter][1]) && data[dataCounter][colCounter][1].indexOf("%") == -1)
                         data[dataCounter][colCounter][2]="CHAR";

                    if (currColumnIndex >= (!nestedTables ? columns.length : columnsParent))
                         currColumnIndex=0;
                         
                    cell = row.createCell(currColumnIndex);

                    // In order to prevent errors, always default the type to CHAR if not specified.
                    if (data[dataCounter][colCounter][2]==null) data[dataCounter][colCounter][2]="CHAR";

                    // Type
                    switch(data[dataCounter][colCounter][2].toUpperCase()) {
                         case "BOOLEAN":
                         case "INT":
                         case "INTEGER":
                         case "NUMERIC":
                              // If the data is null, don't attempt to write a null value as a number because it will throw an error message
                              // Instead, write empty string if its null
                              rowWritten=true;                              
                              
                              if (data[dataCounter][colCounter][1] != null) {                            
                                   cell.setCellValue(parseInt(data[dataCounter][colCounter][1]));
                                   cell.setCellStyle((styledFormat != null ? styledFormat : cellFormat));
                              } else {
                                   cell.setCellStyle((styledFormat != null ? styledFormat : cellFormat));
                              }
                              
                              cell.setCellType(getCellType("NUMERIC"));
                              
                              break;
                         case "CHAR":
                              rowWritten=true;

                              if (data[dataCounter][colCounter][1] != null) {                    
                                   cell.setCellStyle((styledFormat != null ? styledFormat : cellFormat));
                                   cell.setCellValue(data[dataCounter][colCounter][1]);
                              } else {
                              	   cell.setCellValue("");
                                   cell.setCellStyle((styledFormat != null ? styledFormat : cellFormat));
                              }

                              break;
                         case "CURRENCY":
                    	         rowWritten=true;

                    	         // If the data is null, don't attempt to write a null value as a number because it will throw an error message
                              // Instead, write empty string if its null
                    	         if (data[dataCounter][colCounter][1] != null) {
                    	              cell.setCellValue(parseFloat(data[dataCounter][colCounter][1]));
                                    cell.setCellStyle((styledFormat != null ? styledFormat : cellCurrencyFormat));
                    	         } else
                                   cell.setCellStyle((styledFormat != null ? styledFormat : cellFormat));

                               cell.setCellType(getCellType("NUMERIC"));
                               
                    	         break;
                    	   case "TIME":
                    	   			 if (data[dataCounter][colCounter][1] != null) {
                    	   			      cell.setCellValue(data[dataCounter][colCounter][1]);
                    	   			      cell.setCellStyle(timeFont);
                    	   			 }
                    	   		   break;												                     	         
                    	   case "DATE":
                    	   case "DATETIME":
                    	         rowWritten=true;

                    	         var dateVal="";

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
                    	         cell.setCellValue(dateVal);
                    	         cell.setCellStyle((styledFormat != null ? styledFormat : cellFormat));
                    } // end of switch

                    if (rowWritten==true)
                         currColumnIndex++;
               } // end of for (var colCounter=0;colCounter<columnHeaders.length;colCounter++) {
               
               rowCounter++;

               // write child data if we are using nested tables
               if (nestedTables == true) {
                    // Write child table headers
                    var columnHeadersChild=reportObj.Sheets[reportObjSheetCounter].ColumnHeadersChild.split(",");

                    row=(sheet.getRow(rowCounter) != null ? sheet.getRow(rowCounter) : sheet.createRow(rowCounter));

                    var childIndent=0;

                    try {
                         childIndent=(typeof reportObj.Sheets[reportObjSheetCounter].ChildIndent !== 'undefined' ? parseInt(reportObj.Sheets[reportObjSheetCounter].ChildIndent) : 0);
                         childIntent=parseInt(childIndent);
                    } catch (e) {
                    }
                         
                    for (columnCounter=0;columnCounter<columnHeadersChild.length;columnCounter++) {
                         try {
                              if (columnHeadersChild[columnCounter].toUpperCase() != "WHERECLAUSE") {
                                   row.createCell(columnCounter+childIndent).setCellValue(columnHeadersChild[columnCounter]);

                                   row.getCell(columnCounter+childIndent).setCellStyle(headerFormat);
                              }
                         } catch (e) {
                              workbook.close();
                    
                              if (FileServices.existsFile(reportObj.FileName))
                                   FileServices.deleteFile(reportObj.FileName);
                              
                              return ["ERROR","The error " + e + " occurred when sql is " + reportObj.Sheets[reportObjSheetCounter].SQLChild + ", columnHeaders= " + columnHeadersChild + ", length=" + columnHeadersChild.length + ",columnCounter="+columnCounter+ " and value=" + columnHeadersChild[columnCounter]];
                         }
                    }

                    rowCounter++;                    

                    // Write child table data
                    var rowArr=[];

                    // SQL Based nested table
                    if (reportObj.Sheets[reportObjSheetCounter].SQLChild != null) {
                         try {
                              var childSQL=reportObj.Sheets[reportObjSheetCounter].SQLChild.replace("<WHERECLAUSE>",(reportObj.Sheets[reportObjSheetCounter].SQLChild.indexOf("WHERE ") == -1 ?  " WHERE " : " AND ") + reportObj.Sheets[reportObjSheetCounter].JoinWhereClause[0] + "=" + (typeof parentMatchingValue == 'string' ? "'" : "") + parentMatchingValue + (typeof parentMatchingValue == 'string' ? "'" : ""));

                              row=(sheet.getRow(rowCounter) != null ? sheet.getRow(rowCounter) : sheet.createRow(rowCounter));
                              
                              // Read the data and save it into childData array
                              var childData=[];
                              
                              services.database.executeSelectStatement(reportObj.Sheets[reportObjSheetCounter].DBConnectionChild,childSQL,
                              function (columnData) {                              	   
                              	   for (var colCounter=0;colCounter<columnsChild.length;colCounter++) {
                              	        for (var colName in columnData) {                       	        
                              	        	   if (columnsChild[colCounter][0] === colName) {
                              	        	   	    if (columnsChild[colCounter] == null) {
                              	        	   	         currColumnIndex++;
                              	        	   	         continue
                              	        	   	    }

                              	        	        // add column name, column value, column type and date format for date type
                                   	              var lineArr = new Array(columnsChild[colCounter][0],(columnData[colName] != null ? columnData[colName] : null),columnsChild[colCounter][1],(columnsChild[colCounter][2] != null ? columnsChild[colCounter][2] : null));
                                   	              rowArr.push(lineArr);
                              	        	   }
                              	        }
                              	   }

                              	   childData.push(rowArr);
                              	   rowArr=[];
                             });
                        } catch(e) {
               	             workbook.close();
                              
                             if (FileServices.existsFile(reportObj.FileName))
                                  FileServices.deleteFile(reportObj.FileName);
                         
                             return ["ERROR","An error occurred generating the report with the error " + e + " while executing the SQL statement " + reportObj.Sheets[reportObjSheetCounter].SQLChild + " WHERE " + reportObj.Sheets[reportObjSheetCounter].JoinWhereClause[0] + "=" + parentMatchingValue];
                        }
                   } else if (reportObj.Sheets[reportObjSheetCounter].TableDataChild != null) {
                        row=(sheet.getRow(rowCounter) != null ? sheet.getRow(rowCounter) : sheet.createRow(rowCounter));

                        // Read the data and save it into childData array
                        var childData=[];

                        var allRows=tables.getTable(reportObj.Sheets[reportObjSheetCounter].TableDataChild);
                        var rows=allRows.getRows();

                        while (rows.next()) {
                    	       lineArr = [];
               	             rowArr = [];
               	         
                             // Loop through all columns in provided column list  reportObj.Sheets[reportObjSheetCounter].JoinWhereClause[0]
                             for (var key in columnsChild) {
                                  if (allRows.getColumn(columnsChild[key][0]) != null && allRows.getColumn(reportObj.Sheets[reportObjSheetCounter].JoinWhereClause[0]) != null && allRows.getColumn(reportObj.Sheets[reportObjSheetCounter].JoinWhereClause[0]).value == parentMatchingValue) {
                                       // add column name, column value, column type and date format for date type
                              	       lineArr = new Array(columnsChild[key][0],(allRows.getColumn(columnsChild[key][0]).value != null ? allRows.getColumn(columnsChild[key][0]).value : null),columnsChild[key][1],columnsChild[key][2],(reportObj.Sheets[reportObjSheetCounter].ColumnsChild[key][2] != null ? reportObj.Sheets[reportObjSheetCounter].ColumnsChild[key][2] : null),(reportObj.Sheets[reportObjSheetCounter].ColumnsChild[key][3] != null ? reportObj.Sheets[reportObjSheetCounter].ColumnsChild[key][3] : null));

                                       rowArr.push(lineArr);
                                   } else {
                                       invalidColumn=columnsParent[key][0];
                                       break;
                                   }
                             }

                             // push into line array
                             if (rowArr.length > 0)
                                  childData.push(rowArr); 
                        }
                   }

                   // Write the child data
                   for (var childDataCounter=0;childDataCounter<childData.length;childDataCounter++) {                                  
                        for (var columnItem=0;columnItem<childData[childDataCounter].length;columnItem++) {
                             var item=childData[childDataCounter][columnItem];                                      
                                       
                             cell = row.createCell(childIndent+columnItem);
                                                                              
                             // If the type is INT but the value is a CHAR, change the type to an CHAR so it will be written as a CHAR. Ignore percentage values because we want to still write them as a number
                             if ((item[2] == "INTEGER" || item[2] == "INT") && item[1] != null && !isInt(item[1]) && item[1].indexOf("%") == -1)
                                  item[2]="CHAR";                                        

                             // In order to prevent errors, always default the type to CHAR if not specified.
                             if (item[2]==null) item[2]="CHAR";

                             // Type
                             switch(item[2].toUpperCase()) {
                                  case "BOOLEAN":
                                  case "INT":
                                  case "INTEGER":
                                  case "NUMERIC":
                                       // If the data is null, don't attempt to write a null value as a number because it will throw an error message
                                       // Instead, write empty string if its null
                                       rowWritten=true;                              
                              
                                       if (item[1] != null) {                            
                                            cell.setCellValue(parseInt(item[1]));
                                            cell.setCellStyle((styledFormat != null ? styledFormat : cellFormat));
                                       } else {
                                            cell.setCellStyle((styledFormat != null ? styledFormat : cellFormat));
                                       }
                              
                                       cell.setCellType(getCellType("NUMERIC"));
                              
                                       break;
                                  case "CHAR":
                                       rowWritten=true;

                                       if (item[1] != null) {                    
                                            cell.setCellStyle((styledFormat != null ? styledFormat : cellFormat));
                                            cell.setCellValue(item[1]);
                                       } else {
                         	                  cell.setCellValue("");
                                            cell.setCellStyle((styledFormat != null ? styledFormat : cellFormat));
                                       }
                                             
                                       break;
                                 case "CURRENCY":
                    	                 rowWritten=true;

                    	                 // If the data is null, don't attempt to write a null value as a number because it will throw an error message
                                       // Instead, write empty string if its null
                    	                 if (item[1] != null) {
                    	                      cell.setCellValue(parseFloat(item[1]));
                                            cell.setCellStyle((styledFormat != null ? styledFormat : cellCurrencyFormat));
                    	                 } else
                                            cell.setCellStyle((styledFormat != null ? styledFormat : cellFormat));
                                                  
                                       cell.setCellType(getCellType("NUMERIC"));
                               
                    	                 break;
                    	            case "TIME":
                    	                 rowWritten=true;
                    	   			          
                    	   			         if (item[1] != null) {
                    	   			              cell.setCellValue(item[1]);
                    	   			              cell.setCellStyle(timeFont);
                    	   			         }
                    	   		                  
                    	   		           break;												                     	         
                    	            case "DATE":
                    	            case "DATETIME":
                    	                 rowWritten=true;
    
                        	             var dateVal="";

                                       // If index 1 (value) isn't null and [3] (date format) isn't null
                        	             if (item[1] != null && item[3] != null) {
                        	                  switch (item[3]) {
                        	                       case "yyyymmdd":
                         	                            dateVal=new Date(item[1]).yyyymmdd();
                         	                  	        break;
                         	                       case "mmddyy":
                         	                  	        dateVal=new Date(item[1]).mmddyy();
                         	                  	        break;
                    	                       	   case "mm/dd/yyyy":
                    	                       	   default:
                    	                       	        dateVal=new Date(item[1]).mmddyyyy();
                    	                           }
                    	                 } else if (item[1] != null)
                    	                      dateVal=new Date(item[1]).mmddyyyy();

                    	                 // dateVal shouldn't ever be null
                    	                 cell.setCellValue(dateVal);
                    	                 cell.setCellStyle((styledFormat != null ? styledFormat : cellFormat));
                                  } // end of switch                                   
                             } 

                             rowCounter++;

                             row=(sheet.getRow(rowCounter) != null ? sheet.getRow(rowCounter) : sheet.createRow(rowCounter));
                        }

                        if (columnFound == false && invalidColumn != null && invalidColumn != "")
                             return ["ERROR","The column " + invalidColumn + " was not found in the database. Please check the spelling of the column name"];
               }

               // *** Merge cells if MergeCells was specified ***
               if (typeof reportObj.Sheets[reportObjSheetCounter].MergeCells != 'undefined') {
                    for (var mergeCounter=0;mergeCounter < reportObj.Sheets[reportObjSheetCounter].MergeCells.length;mergeCounter++) {
                        var mergeCell=reportObj.Sheets[reportObjSheetCounter].MergeCells[mergeCounter].split(",");

                        if (mergeCell.length == 2)
                             sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(rowCounter,rowCounter,parseInt(mergeCell[0]),parseInt(mergeCell[1])));
                        else
                             sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(parseInt(mergeCell[0]),parseInt(mergeCell[1]),parseInt(mergeCell[2]),parseInt(mergeCell[3])));
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
                                   
                              var formula=reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].Formula.replaceAll("<CURRENTROW>",(rowCounter));
                              var format=(reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].DataType == "CURRENCY" ? cellCurrencyFormat : cellFormat);
                         } catch(e) {
                         	    workbook.close();
                         	    
                              if (FileServices.existsFile(reportObj.FileName))
                                   FileServices.deleteFile(reportObj.FileName);
                              
                              return ["ERROR","An error occurred when Formulas=" + reportObj.Sheets[reportObjSheetCounter].Formulas];
                         }

                         row=(sheet.getRow(rowNum) != null ? sheet.getRow(rowNum) : sheet.createRow(rowNum));

                         if (row !=null) {
                              cell = (row.getCell(columnNum) != null ? row.getCell(columnNum) : row.createCell(columnNum));

                              cell.setCellFormula(formula);
                         
                    	        cell.setCellStyle(format);
                         }
                    }
               }
          } // end of for (var dataCounter=0;dataCounter<data.length;dataCounter++) {
          
          // *** END OF LOOP THAT GOES THROUGH DATA ARRAY AND WRITES THE DATA ***          

          // Disable Number stored as text warning
          sheet.addIgnoredErrors(new org.apache.poi.ss.util.CellRangeAddress(0, rowCounter, 0,100), IgnoredErrorType.NUMBER_STORED_AS_TEXT);
          
          // Check to see if any data was written to the current sheet
          if (rowWritten==false) {
               row = sheet.createRow(1);
          	   cell = row.createCell(0);

          	   cell.setCellValue("No data found");
          	   cell.setCellStyle(cellFormatNoBorder);
          } else {
          	anyRowWritten=true; // rowWritten is referenced for the current sheet. anyRowWritten is used to determine if any data was written on any sheet
          } 
          
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
                         row=(sheet.getRow(rowNum) != null ? sheet.getRow(rowNum) : sheet.createRow(rowNum));
                         
                         cell = (row.getCell(columnNum) != null ? row.getCell(columnNum) : row.createCell(columnNum));

                         cell.setCellFormula(formula);
                         
                    	   cell.setCellStyle(format);
                    } catch(e) {
                    	   workbook.close();
                         if (FileServices.existsFile(reportObj.FileName))
                              FileServices.deleteFile(reportObj.FileName);
                         
                         return ["ERROR","An error occurred writing the formula with the error " + e + " when columnNum="+columnNum+", rownum="+(rowNum-1)+", formula="+formula+", format=" + reportObj.Sheets[reportObjSheetCounter].Formulas[formulaCounter].DataType];
                    }
               }
          }

          // **** Hyperlinks *** After writing the current sheet, write hyperlinks if specified
          if (typeof reportObj.Sheets[reportObjSheetCounter].HyperLinks != 'undefined') {
               for (hyperlinkCounter=0;hyperlinkCounter < reportObj.Sheets[reportObjSheetCounter].HyperLinks.length;hyperlinkCounter++) {
                    try {
                         var columnNum=reportObj.Sheets[reportObjSheetCounter].HyperLinks[hyperlinkCounter].Column;
                         var rowNum=(typeof reportObj.Sheets[reportObjSheetCounter].HyperLinks[hyperlinkCounter].Row != 'undefined' ? reportObj.Sheets[reportObjSheetCounter].HyperLinks[hyperlinkCounter].Row : rowCounter);                         
                         var value=reportObj.Sheets[reportObjSheetCounter].HyperLinks[hyperlinkCounter].Value;

                         var destinationSheet=workbook.getSheet(reportObj.Sheets[reportObjSheetCounter].HyperLinks[hyperlinkCounter].DestinationSheet);

                         if (destinationSheet==null)
                              destinationSheet=sheet;
                         
                         var destinationColumn=reportObj.Sheets[reportObjSheetCounter].HyperLinks[hyperlinkCounter].DestinationColumn;

                         var destinationRow=reportObj.Sheets[reportObjSheetCounter].HyperLinks[hyperlinkCounter].DestinationRow;
                    } catch(e) {
                    	   workbook.close();
                         if (FileServices.existsFile(reportObj.FileName))
                              FileServices.deleteFile(reportObj.FileName);
                         
                         return ["ERROR","An error occurred when Hyperlinks=" + reportObj.Sheets[reportObjSheetCounter].HyperLinks];
                    }
                    
                    // Validate that DestinationSheet is a valid sheet. We can't do this in the validation because the sheet won't exist use in the section that does the validation
                    if (workbook.getSheet(reportObj.Sheets[reportObjSheetCounter].HyperLinks[hyperlinkCounter].DestinationSheet) == null) {	                            	
	                       workbook.close();

	                       if (FileServices.existsFile(reportObj.FileName))
	                            FileServices.deleteFile(reportObj.FileName);
	                            
	                       return ["ERROR","The property DestinationSheet in Sheet " + reportObjSheetCounter + ", Hyperlinks[" + hyperlinkCounter + "] refers to a sheet that does not exist"];
                    }

                    if (rowNum==null) row=createRow(rowNum);
                         
                    cell = row.getCell(columnNum);

                    if (cell==null) cell=row.createCell(columnNum);

                    cell.setCellValue(value);

                    createHelper = workbook.getCreationHelper();
                    link = createHelper.createHyperlink(org.apache.poi.common.usermodel.HyperlinkType.URL);
                    link.setAddress(value);
                    cell.setHyperlink(link);

                    hlinkfont = workbook.createFont();
                    hlinkfont.setUnderline(XSSFFont.U_SINGLE);
                    hlinkfont.setColor(HSSFColor.BLUE.index);
                    hlinkstyle = workbook.createCellStyle();
                    hlinkstyle.setFont(hlinkfont);
                    cell.setCellStyle(hlinkstyle);
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
                         styledFormat=createStyleFormat(workbook,reportObj.CustomCellText[customCellTextCounter].Style[0]);
                    } else if (reportObj.CustomCellText[customCellTextCounter].NamedStyle != null) {
                    	   for (namedStylesCounter=0;namedStylesCounter <  namedStyles.length;namedStylesCounter++) {
                              if (namedStyles[namedStylesCounter][0].toString().toUpperCase() === reportObj.CustomCellText[customCellTextCounter].NamedStyle.toString().toUpperCase()) {
                                   styledFormat=namedStyles[namedStylesCounter][1];
                                   break;
                              }
                         }
                    } else
          	             styledFormat=null;

                    row = sheet.getRow(rowNum);

                    if (row==null) row = sheet.createRow(rowNum);
                    cell = row.createCell(columnNum);

                    // In order to prevent errors, always default the type to CHAR if not specified.
                    if (reportObj.CustomCellText[customCellTextCounter].DataType==null) reportObj.CustomCellText[customCellTextCounter].DataType="CHAR";
                    
                    // Write the CustomCellText factoring in the DataType property
                    switch (reportObj.CustomCellText[customCellTextCounter].DataType.toUpperCase()) {
                         case "BOOLEAN":
                         case "CURRENCY":
                         case "INT":
                         case "INTEGER":
                         case "NUMERIC":
                              if (value.indexOf("$") != -1)                              	   
                                   value=value.toString().replace("$","");

                              cell.setCellValue(parseFloat(value));

                              cell.setCellType(getCellType("NUMERIC"));

                              if (styledFormat != null) cell.setCellStyle(styledFormat);
                              break;
                         case "DATE":
                    	   case "DATETIME":
                              dateVal=new Date(value).mmddyyyy();

                              cell.setCellValue(dateVal);
                    	        if (styledFormat != null) cell.setCellStyle(styledFormat);
                    	        
                    	        break;
                         case "CHAR":
                         default: 
                              if (value==null) value="";
                              
                              cell.setCellValue(value);
                              if (styledFormat != null) cell.setCellStyle(styledFormat);
                    }

                    if (typeof reportObj.CustomCellText[customCellTextCounter].MergeCells != 'undefined') {
                         for (var mergeCounter=0;mergeCounter < reportObj.CustomCellText[customCellTextCounter].MergeCells.length;mergeCounter++) {
                              var mergeCell=reportObj.CustomCellText[customCellTextCounter].MergeCells[mergeCounter].split(",");

                              sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(parseInt(mergeCell[0]),parseInt(mergeCell[1]),parseInt(mergeCell[2]),parseInt(mergeCell[3])));
                         }
                    } // end of if
               } // end of for (customCellTextCounter=0;customCellTextCounter < reportObj.CustomCellText.length;customCellTextCounter++) { 
          }   // end of if (typeof reportObj.CustomCellText != 'undefined') {

          // *** START OF SETTING THE COLUMN WIDTHS ***
          
          // Use the provided ColumnSize property if it was provided to autosize all of the columns in the current sheet
          if (typeof reportObj.Sheets[reportObjSheetCounter].ColumnSize !== 'undefined') {
               if (nestedTables == false) {
                    // Loop through each item in ColumnSize array. I loop through ColumnHeaders because its length represents the total # of actual columns and this way you can only provide 1 column width and not specify the rest of the column widths
                    for (columnSizeCounter=0;columnSizeCounter<reportObj.Sheets[reportObjSheetCounter].ColumnHeaders.length;columnSizeCounter++) {
                         // If a non-null value was passed use it. Otherwise default to autosize
                         if (reportObj.Sheets[reportObjSheetCounter].ColumnSize[columnSizeCounter] != null)
                              sheet.setColumnWidth(columnSizeCounter,(reportObj.Sheets[reportObjSheetCounter].ColumnSize[columnSizeCounter]*256) );
                         else
                              sheet.autoSizeColumn(columnSizeCounter);
                    }
               } else {
               	    // Loop through each item in ColumnSize array. I loop through ColumnHeaders because its length represents the total # of actual columns and this way you can only provide 1 column width and not specify the rest of the column widths
                    for (columnSizeCounter=0;columnSizeCounter<reportObj.Sheets[reportObjSheetCounter].ColumnHeadersParent.length;columnSizeCounter++) {
                         // If a non-null value was passed use it. Otherwise default to autosize
                         if (reportObj.Sheets[reportObjSheetCounter].ColumnSizeParent[columnSizeCounter] != null)
                              sheet.setColumnWidth(columnSizeCounter,(reportObj.Sheets[reportObjSheetCounter].ColumnSizeParent[columnSizeCounter]*256) );
                         else
                              sheet.autoSizeColumn(columnSizeCounter+childIndent);
                    }
               }
          } else { // When ColumnSize isn't provided, default first 100 columns to autosize
               if (nestedTables == false) {
                    // Autosize all of the columns               
                    for (var columnSizeCounter=0;columnSizeCounter<reportObj.Sheets[reportObjSheetCounter].Columns.length;columnSizeCounter++) {
               	         try {
               	    	        //print("columnSizeCounter="+ columnSizeCounter + " and the var" + (isInt(columnSizeCounter) ? " is an int" : " is not an int"));
               	    	   
                              sheet.autoSizeColumn(columnSizeCounter);
               	         } catch(e) {
               	              //print("An error occurred sizing column " + columnSizeCounter + " when " + (reportObj.Sheets[reportObjSheetCounter].SQL != null ? "the SQL=" + reportObj.Sheets[reportObjSheetCounter].SQL : " the table is " + reportObj.Sheets[reportObjSheetCounter].TableData) + " with the error" + e);
               	         }
                    }
               } else {
                    // Autosize all of the columns for columns2             
                    for (var columnSizeCounter=0;columnSizeCounter<reportObj.Sheets[reportObjSheetCounter].ColumnsParent.length;columnSizeCounter++) {
               	         try {
               	    	        //print("columnSizeCounter="+ columnSizeCounter + " and the var" + (isInt(columnSizeCounter) ? " is an int" : " is not an int"));
               	    	   
                              sheet.autoSizeColumn(columnSizeCounter);
               	         } catch(e) {
               	              //print("An error occurred sizing column " + columnSizeCounter + " when " + (reportObj.Sheets[reportObjSheetCounter].SQL != null ? "the SQL=" + reportObj.Sheets[reportObjSheetCounter].SQL : " the table is " + reportObj.Sheets[reportObjSheetCounter].TableData) + " with the error" + e);
               	         }
                    }

                    // Autosize all of the columns for columns2            
                    for (var columnSizeCounter=0;columnSizeCounter<reportObj.Sheets[reportObjSheetCounter].ColumnsChild.length;columnSizeCounter++) {
               	         try {
               	    	        //print("columnSizeCounter="+ columnSizeCounter + " and the var" + (isInt(columnSizeCounter) ? " is an int" : " is not an int"));
               	    	   
                              sheet.autoSizeColumn(columnSizeCounter+childIndent);
               	         } catch(e) {
               	              //print("An error occurred sizing column " + columnSizeCounter + " when " + (reportObj.Sheets[reportObjSheetCounter].SQL != null ? "the SQL=" + reportObj.Sheets[reportObjSheetCounter].SQL : " the table is " + reportObj.Sheets[reportObjSheetCounter].TableData) + " with the error" + e);
               	         }
                    }
               }
          }
          // *** END OF OF SETTING THE COLUMN WIDTHS ***

          // *** START OF CONDITIONAL FORMATTING ***
          if (typeof reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting != 'undefined') {
               var underlineStylesObject = {
                    "DOUBLE" : HSSFFont.U_DOUBLE,
                    "DOUBLE_ACCOUNTING" : HSSFFont.U_DOUBLE_ACCOUNTING,
                    "NO_UNDERLINE" : HSSFFont.U_NONE,
                    "SINGLE" : HSSFFont.U_SINGLE,
                    "SINGLE_ACCOUNTING" : HSSFFont.U_SINGLE_ACCOUNTING,
               }

               // Loop through all conditional formatting rules
               for (cfCounter=0;cfCounter < reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting.length;cfCounter++) {                    
                    sheetCF = sheet.getSheetConditionalFormatting();

                    var formula=reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Formula.replaceAll("<CURRENTROW>",(rowCounter+1));

                    rule = sheetCF.createConditionalFormattingRule(formula);

                    fill = rule.createPatternFormatting();

                    fontFmt = rule.createFontFormatting();

                    // Conditional formatting format style
                    if (reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].Bold != null && reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].Bold == true)
                         cfBold=true
                    else
                    	   cfBold=false;

                    if (reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].Italic != null && reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].Italic == true)
                         cfItalic=true
                    else
                    	   cfItalic=false;

                    fontFmt.setFontStyle(cfItalic,cfBold);

                    if (reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].Color != null)
                         fontFmt.setFontColorIndex(getColor(reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].Color));
                    else
                         fontFmt.setFontColorIndex(getColor("BLACK"));

                    if (reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].Size != null)
                         fontFmt.setFontHeight(reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].Size*20); // hight is in 1/20th of a point
                    else
                         fontFmt.setFontHeight(12*20); // height is in 1/20th of a point

                    if (reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].Underline == true) {
                         if (reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].UnderlineStyle != null && underlineStylesObject[reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].UnderlineStyle] != null)
                              fontFmt.setUnderlineType(underlineStylesObject[reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].UnderlineStyle]);
                         else
                    	        fontFmt.setUnderlineType(underlineStylesObject["SINGLE"]);
                    }

                    if (reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].BackgroundColor != null) {
                         fill.setFillBackgroundColor(getColor(reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].Style[0].BackgroundColor));
                         fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
                    }

                    var endRow=(reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].EndRow != null ? reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].EndRow : rowCounter-1);

                    try {
                         var region = new org.apache.poi.ss.util.CellRangeAddress(reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].StartRow,endRow,reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].StartColumn,reportObj.Sheets[reportObjSheetCounter].ConditionalFormatting[cfCounter].EndColumn);

                         regions = new Array();

                         regions.push(region);
                     
                         sheetCF.addConditionalFormatting(regions, rule);
                    } catch(e) {
                         //sendAdminEmail("An error occurred with the report " + reportObj.FileName,"The error " + e + " occurred with the report " + reportObj.FileName);
                    }
               } // end of for (cfCounter=0;cfCounter
          } // *** END OF CONDITIONAL FORMATTING ***

          // *** START OF WRITING AN IMAGE ***
          if (typeof reportObj.Sheets[reportObjSheetCounter].Image != 'undefined') {
               for (imgCounter=0;imgCounter < reportObj.Sheets[reportObjSheetCounter].Image.length;imgCounter++) {
                    stream = new FileInputStream(reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].FileName);

                    helper = workbook.getCreationHelper();

                    drawing = sheet.createDrawingPatriarch();

                    anchor = helper.createClientAnchor();

                    if (reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].AnchorType != null)
                         anchor.setAnchorType(anchorTypesObject[reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].AnchorType.toString().toUpperCase()]);
                    else
                         anchor.setAnchorType(anchorTypesObject["DONT_MOVE_AND_RESIZE"]);
                         
                    var ext=reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].FileName.substring(reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].FileName.lastIndexOf(".")+1).toUpperCase();

                    var type;
                    
                    switch (ext) {
                         case "DIB":
                              type=Workbook.PICTURE_TYPE_DIB;
                              break;
                         case "EMF":
                              type=Workbook.PICTURE_TYPE_EMF;
                              break;
                         case "JPEG":
                         case "JPG":
                               type=Workbook.PICTURE_TYPE_JPEG;
                               break;
                         case "PICT":
                               type=Workbook.PICTURE_TYPE_PICT;
                               break;
                         case "PNG":
                               type=Workbook.PICTURE_TYPE_PNG;
                               break;
                         case "WMF":
                               type=Workbook.PICTURE_TYPE_WMF;
                               break
                    }
                    
                    pictureIndex = workbook.addPicture(Packages.org.apache.commons.io.IOUtils.toByteArray(stream), type);

                    endCol=(reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].EndColumn != null ? reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].EndColumn : reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].StartColumn);
                    endRow=(reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].EndRow != null ? reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].EndRow : reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].StartRow);

                    anchor.setCol1(reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].StartColumn);
                    anchor.setCol2(endCol);
                    anchor.setRow1(reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].StartRow);
                    anchor.setRow2(endRow);

                    pict = drawing.createPicture(anchor, pictureIndex);

                    if (reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].ScaleX != null && reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].ScaleY != null)
                         pict.resize(reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].ScaleX,reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].ScaleY);
                    else if (reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].Scale != null)
                         pict.resize(reportObj.Sheets[reportObjSheetCounter].Image[imgCounter].Scale);
                    else
                         pict.resize();
               }
          }        
          // *** END OF WRITING AN IMAGE ***

          // *** START OF CREATING PIVOT TABLE ***
          if (typeof reportObj.Sheets[reportObjSheetCounter].PivotTable != 'undefined' && reportObj.Sheets[reportObjSheetCounter].PivotTable == true) {
               var firstRow = sheet.getFirstRowNum();
               var lastRow = sheet.getLastRowNum();
               var firstCol = sheet.getRow(0).getFirstCellNum();
               var lastCol = sheet.getRow(0).getLastCellNum();

               var topLeft = new CellReference(firstRow, firstCol);
               var botRight = new CellReference(lastRow, lastCol - 1);

               var aref = new org.apache.poi.ss.util.AreaReference(topLeft, botRight,null);
               var pos = new CellReference(firstRow, 0);
               var pivotSheet = workbook.createSheet("Pivot");
               
               var pivotTable = pivotSheet.createPivotTable(aref, pos,sheet);
                             
               pivotTable.addRowLabel(0);
               pivotTable.addRowLabel(1);
               
               pivotTable.addColumnLabel(DataConsolidateFunction.SUM,6, "Pass Percent");
               pivotTable.addColumnLabel(DataConsolidateFunction.SUM,7, "Fail Percent");
          }
          // *** END OF CREATING PIVOT TABLE ***
     } // *** Start of Loop through the report object for each sheet object ***
	   
     var fos = new FileOutputStream(reportObj.FileName);
     
     workbook.write(fos);
     
     fos.close();
     workbook.close();

     if (anyRowWritten == true) {
          return ["OK",reportObj.FileName];
     } else {
          return ["OK-NODATA",reportObj.FileName];
     }
}