// sets up our custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Sentence Menu' )
      .addItem('Sentence', 'SENTENCE' )  
      .addItem("Clean",'CLEAN')    
      .addItem("Tally", "TALLY")
      .addItem("Working", 'cleanWorking')
      .addToUi();
}

function cleanWorking() {
  var ss = SpreadsheetApp.getActive();
  var src_sheet = ss.getSheetByName("Data");
  var wrk_sheet = ss.getSheetByName("Working")

  if (src_sheet.getRange("A1:A1").getValue() == "Effective Date") {
     src_sheet.deleteColumn(1) // get rid of empty A column
  }
  // copy headings
  var sr = src_sheet.getRange("A1:C1");
  var dr = wrk_sheet.getRange("A1:C1");
  sr.copyTo(dr)
  dr = wrk_sheet.getRange("A1:D1")
  dr.setBackground("#274e13")
  dr.setFontColor("white")

  // find the transacations in the correct date range - ie last row to copy
  var timeZone = ss.getSpreadsheetTimeZone();
  var date = ss.getSheetByName("Results").getRange("B2").getDisplayValue();
  var haystack = src_sheet.getRange("A1:A");
  var finder =  haystack.createTextFinder(date)
  var next = finder.findNext()
  while (next == null) {
    var now = new Date(toStdDate(date));
    const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    var yesterday = new Date(now.getTime() - MILLIS_PER_DAY);   
    date = Utilities.formatDate(yesterday, timeZone, 'dd/MM/yyyy');
    finder =  haystack.createTextFinder(date)
    next = finder.findNext()
  }
  var last_row = next.getRowIndex();
  last_row -= 2;

  
  var data_range = wrk_sheet.getRange("dataRange" );
  var src_range = src_sheet.getRange("A2:C" + last_row );
  data_range.clear();

  // need to offset so we only copy transactions in the date range
  src_range = src_range.offset(0, 0, last_row, 3);
  data_range = data_range.offset(0, 0,last_row, 3);
  src_range.copyTo(data_range);
  CLEAN()
}

function toStdDate(str) {
  var pcs = str.split("/");
  return pcs[2] + "-" + pcs[1] + "-" + pcs[0];
}

function toBankDate(str) {
   const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    yesterday = new Date(now.getTime() - MILLIS_PER_DAY);   
    date = Utilities.formatDate(yesterday, timeZone, 'dd/MM/yyyy');
}
function indexOfStr(str, strArray) {
    for (var j=0; j<strArray.length; j++) {
        if (strArray[j].match(str)) return j;
    }
    return -1;
}

function transposeArr( arr ) {
  var rowlen = 0;
  var col = [];
  arr.forEach(function(row,index) {
    if (row.length > rowlen) {
      rowlen = row.length;
    }
  });
  for ( var i=0; i< rowlen; i++) {
    var row = [];
    col.push(row);
    for (var j=0; j < arr.length; j++ ) {
        row.push(arr[j][i]);
      
    }
  }
  Logger.log(col);
  return col;
}

// Tally the sentenced charges to their categories
function TALLY() {
  var data_range = SpreadsheetApp.getActive().getSheetByName("Working").getRange("sentenceR" );
  var data_vals = data_range.getValues();
  var index = [];
  var totals = [];
  var vals = []
  
  var data_index = 0;
  while(data_vals[data_index][0] != "") {
    var i = index.indexOf(data_vals[data_index][3].toLowerCase());
    if ( i > -1 ) {
      // found it so we have a tally.  
      totals[i] += data_vals[data_index][2];
      vals[i].push(data_vals[data_index][2]);
      data_index++;
      continue;
    } 
    index.push(data_vals[data_index][3].toLowerCase());
    totals.push(data_vals[data_index][2]);
    vals.push([]);
    vals[vals.length - 1][0] = data_vals[data_index][2];
    data_index++;
  }   
  
  var tmp = transposeArr(vals);
  var empty = new Array(index.length);
  tmp.unshift(index,totals,empty);
  var output = SpreadsheetApp.getActive().getSheetByName("Results").getRange(8,1,tmp.length,index.length);
  output.setValues(tmp);
  var hdr = SpreadsheetApp.getActive().getSheetByName("Results").getRange(1,1);
  bg = hdr.getBackground();
  fg = hdr.getFontColor();
  hdr = SpreadsheetApp.getActive().getSheetByName("Results").getRange(8,1,1,index.length);
  hdr.setBackground(bg);
  hdr.setFontColor(fg)
  hdr = SpreadsheetApp.getActive().getSheetByName("Results").getRange(9,1,1,index.length);
  hdr.setBorder(null, null, true, null, false, false)
}

// clean up the categorization column
function CLEAN() {
  // clear the sentencing
  var data_range = SpreadsheetApp.getActive().getSheetByName("Working").getRange("dataRange" );
  var last_col = data_range.getNumColumns(); // actually want the column after the range
  var last_row = data_range.getNumRows();
  var clr_range = data_range.offset(0,last_col,last_row,1);
  clr_range.clear();
    
  // clear the tally results
  var col = SpreadsheetApp.getActive().getSheetByName("Results").getLastColumn();
  var row = SpreadsheetApp.getActive().getSheetByName("Results").getLastRow();
  if (row <= 5) {
    return;
  }
  clr_range = SpreadsheetApp.getActive().getSheetByName("Results").getRange(6,1,row - 5, col);
  clr_range.clear();
}

// sentence menu handler
function SENTENCE() {
  var map_range = SpreadsheetApp.getActive().getSheetByName("Working").getRange("defnMap" );
  var data_range = SpreadsheetApp.getActive().getSheetByName("Working").getRange("sentenceR" );

  var data_vals = data_range.getValues();
  var map_vals = map_range.getValues();

  var data_index = 0;
  while(true) {
    if (data_vals[data_index][1] == "") {
      break;
    }
    var desc = data_vals[data_index][1].toLowerCase();
    var found = false;
    var map_index = 0;
    while(map_vals[map_index][0] != "") {  
      if (desc.indexOf(map_vals[map_index][0].toLowerCase()) > -1) {
        data_vals[data_index][3] = map_vals[map_index][1];
        found = true;
      }
      map_index++;
    }
    if (!found) {
        data_vals[data_index][3] = "??";
    }
    data_index++;
  }
  data_vals.splice(data_index);  
  write_range = data_range.offset(0, 0, data_vals.length);
  write_range.setValues(data_vals)
}

function my_notify(val) {
  var range = SpreadsheetApp.getActive().getSheetByName("Working").getRange("E1:E1" );
  range.setValue(val);
}
 