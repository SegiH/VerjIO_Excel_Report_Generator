function createExcelReportObj(tableName) {
    var columnHeaders="";
    var columns="["
    if (tableName==null || tableName == "") {
         return "Error: Table name not provided";
    }

    if (tables.getTable(tableName) == null) {
         return "Error: Table name is not valid";
    }

    var cols = tables.getTable(tableName).getColumns();
    
    for each (var col in cols) {
            if (col.labelText.text.toLowerCase().indexOf("whereclause") != -1)
                 continue;
                 
         columnHeaders+=col.labelText.text + ","

         columns+="[\"" + col.elementName.replace(tableName + "-","") + "\",\"" + col.type + "\"],";
    }

    // Remove trailing comma from columnHeaders and column
    columnHeaders=columnHeaders.substring(0,columnHeaders.length-1);
    columns=columns.substring(0,columns.length-1);
    columns+="]";

    // Create formatted filename based on the table name
    // so "ComponentQualificationFollowingSubsystem" becomes Component Qualification Following Subsystem
    var fileName="";
    var currcharCode;

    for (var i=0;i<tableName.length;i++) {
         if (i>0 && i<tableName.length-1 && tableName[i].charCodeAt(0) >= 97 && tableName[i].charCodeAt(0) < 122 && tableName[i+1].charCodeAt(0) >= 65 && tableName[i+1].charCodeAt(0) <=90)
              fileName+=tableName[i] + " ";
         else      
              fileName+=tableName[i];
    }
    
    var reportObjStr="var excelReportObj = {<BR>" + "&nbsp;".repeat(8) + "FileName: \"" + fileName + "\",<BR>" + "&nbsp;".repeat(8) + "Sheets:&nbsp;[{<BR>" + "&nbsp;".repeat(16) + "SheetName:&nbsp;\"" + fileName + "\",<BR>" + "&nbsp;".repeat(16) + "SheetIndex:&nbsp;0,<BR>" + "&nbsp;".repeat(16) + "ColumnHeaders:&nbsp;\"" + columnHeaders + "\",<BR>" + "&nbsp;".repeat(16) + "Columns:&nbsp;" + columns + ",<BR>" + "&nbsp;".repeat(16) + "TableData:&nbsp;\"" + tableName + "\",<BR>" + "&nbsp;".repeat(8) + "}<BR>" + "&nbsp;".repeat(8) + "]" + "<BR>}";
    
    //event.owner.addWarningMessage("Please use the following create Excel Report object:<BR><BR>" + reportObjStr);     
    return reportObjStr;
}