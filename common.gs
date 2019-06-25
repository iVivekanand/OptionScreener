function getLastRow(sheet, column) {
  var lastRow = sheet.getMaxRows();
  var values = sheet.getRange(column + "1:" + column + lastRow).getValues();

  for (; values[lastRow - 1] == "" && lastRow > 0; lastRow--) {}
  return lastRow;
}

function getLotSize() {
  lotSize.getRange("A1").setValue("=IMPORTDATA(CONCATENATE(URL!B3,URL!B1))");
}

function createSheet(sheetName) {
  var spreadSheet1 = SpreadsheetApp.openById("myspreadsheet");
  try {
    var sheetNameObject = spreadSheet1.getSheetByName(sheetName);
    spreadSheet1.setActiveSheet(sheetNameObject);
    return spreadSheet1.getSheetByName(sheetName);
  }
  catch (e) {
    Logger.log("Creating sheet: " + sheetName);
    spreadSheet1.insertSheet(sheetName);
    currentMonth = spreadSheet1.getSheetByName(sheetName);
    currentMonthRange = currentMonth.getRange("A1:U3");
    insertItmProbabilityScreenHeader();
    writeCalculations();
    formatItmProbabilitySheet();
    return spreadSheet.getSheetByName(sheetName);
  }
}

function printScrips() {
  var lotSizeRangeValues = lotSizeRange.getValues();
  for(var a=0;a<=scripCount;a++) {
    Logger.log(encodeURIComponent(lotSizeRangeValues[a]));
  }
}

function extractOptionChains() {
  date = new Date();
  var start = date.getTime();
  
  var Time = spreadSheet.getSheetByName("Time");
  Time.getRange("D1").setValue(date.toLocaleTimeString());
  currentMonthRange.getCell(1, 22).setValue(date.toLocaleTimeString());
  currentMonthRange.getCell(1, 22).setNote("Extract Options Chain Start Time");
  
  var dataRange = Time.getRange("J:J");
  dataRange.clear();
  var timeRangeValues = Time.getRange(1, 1, scripCount, 10).getValues(); //scripCount
  var data = "";
  var text = 0;
  var scrips = Time.getRange("A1:" + "A" + getLastRow(Time, "A")).getValues();
  //var scrips = Time.getRange("A1:A100").getValues();
  for(var counter = 1; counter <= scripCount; counter++) { // scripCount
    scrip = scrips[counter-1];
    try {
      data = UrlFetchApp.fetch("http://main.xfiddle.com/c82d7951/" + optionChainPage + ".php?url=https%3A%2F%2Fwww.nseindia.com%2Flive_market%2FdynaContent%2Flive_watch%2Foption_chain%2FoptionKeys.jsp%3F%26symbol%3D" + encodeURIComponent(encodeURIComponent(scrip)) + "%26date%3D" + expiry);
      if ( data.getContentText().length == 2 )
      {
        Logger.log("Empty data. Retrying for " + scrip); 
        Utilities.sleep(500);
        data = UrlFetchApp.fetch("http://main.xfiddle.com/c82d7951/" + optionChainPage + ".php?url=https%3A%2F%2Fwww.nseindia.com%2Flive_market%2FdynaContent%2Flive_watch%2Foption_chain%2FoptionKeys.jsp%3F%26symbol%3DNIFTY%26date%3D" + expiry);
        counter--;
        continue;
      }
    } catch (e) {
      Logger.log("Exception for " + scrip + ". Retrying after 1s delay");
      Utilities.sleep(500);
      data = UrlFetchApp.fetch("http://main.xfiddle.com/c82d7951/" + optionChainPage + ".php?url=https%3A%2F%2Fwww.nseindia.com%2Flive_market%2FdynaContent%2Flive_watch%2Foption_chain%2FoptionKeys.jsp%3F%26symbol%3D" + encodeURIComponent(encodeURIComponent(scrip)) + "%26date%3D" + expiry);
    }
    timeRangeValues[counter-1][9] = data;
    timeRangeValues[counter-1][2] = data.getContentText().length;
  }
  Time.getRange(1, 1, scripCount, 10).setValues(timeRangeValues); //scripCount
  date = new Date();
  var end = date.getTime();
  Time.getRange("G1").setValue((end-start)/1000);
  Time.getRange("E1").setValue(date.toLocaleTimeString());
  Time.getRange("B1").setValue("=COUNTIF(C:C,\"=2\")");
  currentMonthRange.getCell(1, 23).setValue(date.toLocaleTimeString());
  currentMonthRange.getCell(1, 23).setNote("Extract Options Chain End Time");
  currentMonthRange.getCell(1, 24).setValue((end - start)/(1000*60)).setNumberFormat("#.##");
  Logger.log("Extraction complete!");
}

function writeCalculations() {
  currentMonthRange.getCell(1, 13).setValue("=IF(TODAY()<CONCAT(\"$B\",\"$3\"), SUBTOTAL(9, Q:Q), SUBTOTAL(9, R:R))");
  currentMonthRange.getCell(1, 13).setNote("Current P/L");
  currentMonthRange.getCell(1, 14).setValue("=SUBTOTAL(9, G:G,I:I)-(SUBTOTAL(3, G:G)-1)*75");
  currentMonthRange.getCell(1, 14).setNote("Maximum profit potential");
  currentMonthRange.getCell(1, 15).setValue("=SUBTOTAL(9, P:P)");
  currentMonthRange.getCell(1, 15).setNote("Total investment");
  currentMonthRange.getCell(1, 19).setValue("=IFERROR(N1/O1,\"\")*0.7").setNumberFormat("0.##%");;
  currentMonthRange.getCell(1, 19).setNote("Maximum Profit %");
  currentMonthRange.getCell(1, 20).setValue("=IFERROR(IF(M1/O1<0, M1/O1,M1/O1*0.7),\"\")").setNumberFormat("0.##%");;
  currentMonthRange.getCell(1, 20).setNote("Realized P/L %");
}
