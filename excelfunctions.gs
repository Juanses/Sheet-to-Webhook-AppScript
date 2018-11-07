var ExcelClass = function(){
  
  //SET SHEET FROM SPREEDSHEET ID
  this.setsheetbyId = function(id,sheetnumber){
    this.ss = SpreadsheetApp.openById(id);
    this.sheet = this.ss.getSheets()[sheetnumber];
  }
  
  //SET COLS AUTOMATICALLY IN EXCEL OBJECT
  this.setColsAuto = function (){
    var columns = this.sheet.getRange(1, 1, 1, this.sheet.getLastColumn()).getValues()[0]; //read headers
    this.cols = columns;
  }
  
  //GET ALL THE VALUES IN A RANGE
  this.getallvalues = function(startline,startcolumn,endcolumn){
    var lastrow = this.sheet.getLastRow();
    return this.sheet.getRange(startline,startcolumn,lastrow-(startline-1),endcolumn).getValues(); 
  }
  
  //GET ALL THE VALUES AND CREATE AN OBJECT FROM A ROW BASED ON THE COLUMNS 
  this.getvaluesfromrow = function(values,startrow,rownumber){
    var row;
    var rowdata = {};
    for (var i = 0; i < values.length; i++) {
      if (i == (rownumber-startrow)){
        for(key in this.cols){
          rowdata[this.cols[key]] = values[i][key];
        }
        break;
      }
    }   
    return rowdata;
  } 
  
}
