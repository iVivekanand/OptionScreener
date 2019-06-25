var spreadSheet = SpreadsheetApp.openById("myspreadsheet");
var lotSize = spreadSheet.getSheetByName("LotSize");
var time = spreadSheet.getSheetByName("Time");
var url = spreadSheet.getSheetByName("URL");
var backend = spreadSheet.getSheetByName("Backend");
var analysis = spreadSheet.getSheetByName("Analysis");
var phpFiddleUrl = "http://main.xfiddle.com/c82d7951/PassUrl.php?url=";
var expiry = analysis.getRange("G2").getValue();
var currentMonth = createSheet(expiry);
var date = new Date();
var analysisRange = analysis.getRange("A1:W42");
var lotSizeRange = backend.getRange("A1:A" + getLastRow(backend, "A"));
var currentMonthRange = currentMonth.getRange("A1:AA3");
var scripCount = getLastRow(time, "A");
var roiExpected = analysisRange.getCell(2, 9).getValue();
var expectedItmProbability = analysisRange.getCell(2, 16).getValue();

var scrips = lotSizeRange.getValues();
var lotSizes = backend.getRange("Z1:Z" + getLastRow(backend, "Z"));
var lotSizeRanges = lotSizes.getValues();
var analysisRangeValues = analysisRange.getValues();
var insertRange = currentMonth.getRange("A3:AA3");

var optionChainPage = "OptionChain";
