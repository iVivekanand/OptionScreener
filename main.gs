var expiry = analysis.getRange("G2").getValue();
var timeToExpiry = analysis.getRange("O2").getValue()/365;
var ceValue = 0;
var peValue = 0;
var spot = 0;
var strike = 0;
var lotSize = 0;
var investment = 0;
var iv = 0.0;
var currentItmProbability = 0.0;
var start = 1;
var scrip = "";
var scripData = time.getRange(1, 13, 1, 1).getValue();
//var scripData = "/";
var rows;
var columns;

try {
  rows = Utilities.parseCsv(scripData, "/");
  columns = Utilities.parseCsv(rows[0][0], " ");
}
catch (e) {
    Logger.log(e.message);
/*    scripData = "Underlying Stock: CEATLTD 1389.00As on Sep 14 2018 15:30:30 IST/
?	-?	-?	-?	-?	-?	-?	1400?	171.50?	-?	-?	1080.00?	700?	0.15?	-?	-?	-?	-?	-?	-?	-?	-/
?	-?	-?	-?	-?	-?	-?	1400?	151.50?	-?	-?	1100.00?	700?	0.15?	-?	-?	-?	-?	-?	-?	-?	-/
?	-?	-?	-?	-?	-?	-?	1400?	135.00?	-?	-?	1120.00?	700?	0.15?	-?	-?	-?	-?	-?	-?	-?	-/
?	-?	-?	-?	-?	-?	-?	1400?	112.00?	-?	-?	1140.00?	700?	0.15?	-?	-?	-?	-?	-?	-?	-?	-/
?	-?	-?	-?	-?	-?	-?	1400?	92.00?	-?	-?	1160.00?	700?	0.15?	-?	-?	-?	-?	-?	-?	-?	-/
?	-?	-?	-?	-?	-?	-?	1400?	74.00?	-?	-?	1180.00?	700?	0.15?	-?	-?	-?	-?	-?	-?	-?	-/
?	-?	-?	-?	-?	-?	-?	1750?	54.50?	-?	-?	1200.00?	700?	0.15?	-?	-?	-?	-?	-?	-?	-?	-/
?	-?	-?	-?	-?	-?	-?	1750?	104.60?	-?	-?	1220.00?	700?	0.15?	-?	-?	-?	-?	-?	-?	-?	-/
?	-?	-?	-?	-?	-?	-?	1750?	92.95?	-?	-?	1240.00?	1050?	1.00?	-?	-?	-?	11.00?	-?	-?	-?	1750/
?	-?	-?	-?	-?	-?	-?	3850?	111.65?	145.95?	2100?	1260.00?	350?	4.00?	5.90?	700?	-7.10?	4.00?	39.56?	2?	-?	4900/
?	-?	-?	-?	-?	-?	-?	4550?	95.60?	130.25?	2800?	1280.00?	1050?	1.25?	-?	-?	-?	10.25?	-?	-?	-?	700/
?	700?	-?	-?	-?	74.65?	-?	7700?	84.90?	112.95?	7700?	1300.00?	350?	7.45?	8.45?	700?	-7.75?	7.40?	37.63?	39?	700?	21700/
?	-?	-?	-?	-?	-?	-?	10850?	65.90?	94.35?	5600?	1320.00?	700?	8.25?	20.30?	4900?	-10.00?	15.30?	41.60?	2?	-700?	350/
?	5250?	-350?	3?	33.84?	68.00?	12.95?	1750?	66.05?	76.20?	1050?	1340.00?	350?	15.00?	18.05?	700?	-14.70?	15.15?	34.91?	30?	-5250?	23800/
?	11550?	-2100?	15?	31.28?	52.00?	13.80?	3850?	51.20?	58.95?	350?	1360.00?	700?	18.55?	22.45?	700?	-13.50?	23.50?	36.84?	13?	1750?	11900/
?	8400?	700?	85?	32.92?	41.65?	10.05?	350?	34.45?	43.00?	350?	1380.00?	700?	27.30?	37.55?	6300?	-19.05?	27.30?	-?	1?	350?	3500/
?	42000?	-1750?	376?	34.29?	32.50?	8.55?	350?	33.05?	34.00?	350?	1400.00?	3500?	29.00?	46.20?	8750?	-9.55?	46.20?	41.16?	20?	-4200?	26600/
?	30450?	5250?	99?	34.85?	25.25?	5.95?	350?	24.90?	26.30?	350?	1420.00?	5950?	41.95?	74.45?	5600?	-?	99.85?	-?	-?	-?	3850/
?	51800?	700?	34?	36.14?	19.75?	6.05?	350?	18.65?	19.50?	350?	1440.00?	6300?	52.10?	75.15?	9100?	-?	86.50?	-?	-?	-?	350/
?	22050?	1050?	14?	35.25?	13.60?	2.60?	700?	12.75?	13.95?	350?	1460.00?	6300?	68.30?	102.00?	7700?	-?	94.00?	-?	-?	-?	350/
?	8050?	-700?	10?	36.82?	10.70?	4.70?	700?	8.05?	10.95?	350?	1480.00?	700?	80.45?	108.65?	10150?	-3.75?	107.40?	50.18?	1?	-?	5250/
?	21350?	1750?	14?	37.62?	8.00?	1.60?	350?	6.50?	8.00?	350?	1500.00?	350?	107.50?	128.40?	7000?	-?	90.00?	-?	-?	-?	350/
?	6650?	-?	-?	-?	5.40?	-?	700?	1.30?	6.40?	700?	1520.00?	3150?	117.95?	147.75?	3150?	-?	-?	-?	-?	-?	-/
?	1750?	-?	1?	36.44?	3.30?	-0.65?	700?	0.70?	-?	-?	1540.00?	2450?	134.45?	157.10?	700?	-?	-?	-?	-?	-?	-/
?	2450?	-?	-?	-?	5.15?	-?	4900?	0.30?	-?	-?	1560.00?	2450?	144.55?	187.45?	700?	-?	-?	-?	-?	-?	-/
?	1400?	-?	-?	-?	3.50?	-?	1050?	0.75?	6.45?	1400?	1580.00?	700?	163.70?	207.10?	700?	-?	-?	-?	-?	-?	-/
?	5250?	-?	-?	-?	2.25?	-?	1750?	0.50?	-?	-?	1600.00?	1400?	183.20?	222.90?	1400?	-?	178.00?	-?	-?	-?	350/
?	-?	-?	-?	-?	-?	-?	700?	0.25?	-?	-?	1620.00?	700?	203.60?	246.50?	700?	-?	-?	-?	-?	-?	-/
?	-?	-?	-?	-?	-?	-?	700?	0.25?	-?	-?	1640.00?	1400?	223.30?	260.55?	1400?	-?	239.00?	-?	-?	-?	1400/"
*/
}

function getLotSize() {
  var lotSize = spreadSheet.getSheetByName("LotSize");
  var csvContent = UrlFetchApp.fetch("http://main.xfiddle.com/c82d7951/passurl.php?url=https%3A%2F%2Fwww.nseindia.com%2Fcontent%2Ffo%2Ffo_mktlots.csv").getContentText();
  var csvData = Utilities.parseCsv(csvContent);
  lotSize.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
}

function screenCe() {  
  for(var ceEntry = Math.round(rows.length/2); ceEntry <= rows.length; ceEntry++) {
    var strikeToCheck = Utilities.parseCsv(rows[ceEntry-1][0], "?");
    if(isNaN(Number(strikeToCheck[0][4]))) {
      // Proceed to next one if there isn't a buyer (No bids)
       continue;
    }
    strike = parseFloat(strikeToCheck[0][11]);
    iv = parseFloat(strikeToCheck[0][4])/100;
    ceValue = parseFloat(strikeToCheck[0][8]) * lotSize - 60;    
    currentItmProbability = NORMDIST(Math.log(spot/strike)/(iv*Math.sqrt(timeToExpiry)));    
    investment = ((spot + Math.abs(strike - spot)) * lotSize * 0.12);
    
    if(scrip == "NIFTY") {
      investment = 50000;
    }
    else if(scrip == "BANKNIFTY") {
      investment = 60000;
    }
    
    roiCalculated = ceValue/investment;
//    if( roiCalculated > roiExpected && roiCalculated < 0.05) {
    if( roiCalculated > roiExpected ) {
      if(currentItmProbability <= expectedItmProbability || currentItmProbability == "-") {
        copyCe();
      }
    }
  }
}

function screenPe() {
  for(var peEntry = Math.round(rows.length/2); peEntry > 0; peEntry--) {
    var strikeToCheck = Utilities.parseCsv(rows[peEntry-1][0], "?");
    if(isNaN(Number(strikeToCheck[0][18]))) {
      // Proceed to next one if there isn't a buyer (No bids)
       continue;
    }
    strike = parseFloat(strikeToCheck[0][11]);
    iv = parseFloat(strikeToCheck[0][18])/100;
    peValue = parseFloat(strikeToCheck[0][13]) * lotSize - 60;
    currentItmProbability = 1 - NORMDIST(Math.log(spot/strike)/(iv*Math.sqrt(timeToExpiry)));    
    investment = ((spot + Math.abs(strike - spot)) * lotSize * 0.12);
    
    if(scrip == "NIFTY") {
      investment = 50000;
    }
    else if(scrip == "BANKNIFTY") {
      investment = 60000;
    }
    
    roiCalculated = peValue/investment;
//    if( roiCalculated > roiExpected && roiCalculated < 0.05) {
    if( roiCalculated > roiExpected ) {
      if(currentItmProbability <= expectedItmProbability || currentItmProbability == "-") {
        copyPe();
      }
    }
  }
}

function getAtmIv(optionType) {
  var ceDivisor = 0;
  var peDivisor = 0;
  var callIv = 0.0;
  var putIv = 0.0;
  var otm1Index, otm2Index, itm1Index, itm2Index, atmIndex;
  var otm1CeIv, otm1PeIv, otm2CeIv, otm2PeIv, itm1CeIv, itm1PeIv, itm2CeIv, itm2PeIv, atmCeIv, atmPeIv;
  otm1Index = otm2Index = itm1Index = itm2Index = atmIndex = otm1CeIv = otm1PeIv = otm2CeIv = otm2PeIv = itm1CeIv = itm1PeIv = itm2CeIv = itm2PeIv = atmCeIv = atmPeIv = 0;
  
  try {
    for(var middle = 4; middle <= rows.length; middle++) {
      var prevCol = Utilities.parseCsv(rows[middle-1][0], "?");
      var middleCol = Utilities.parseCsv(rows[middle][0], "?");
      var nextCol = Utilities.parseCsv(rows[middle+1][0], "?");
      
      // Spot price between middle and next strike
      if( (spot > parseInt(middleCol[0][11]),10) && (spot < parseInt(nextCol[0][11],10)) ) {
        // Spot price closer to the middle strike
        if( (spot - parseInt(middleCol[0][11]),10) < (parseInt(nextCol[0][11],10) - spot) ) {
          atmIndex = middle;
          itm1Index = middle - 1;
          itm2Index = middle - 2;
          otm1Index = middle + 1;
          otm2Index = middle + 2;
          break;
        }
        // Spot price closer to the next strike
        else {
          atmIndex = middle + 1;
          itm1Index = middle;
          itm2Index = middle - 1;
          otm1Index = middle + 2;
          otm2Index = middle + 3;
          break;
        }          
      }
      // Spot price between previous and middle strike
      else if( (spot < parseInt(middleCol[0][11],10)) && (spot > parseInt(prevCol[0][11],10)) ) {
        // Spot price closer to the middle strike
        if( (parseInt(middleCol[0][11],10) - spot) < (spot - parseInt(prevCol[0][11],10)) ) {
          atmIndex = middle;
          itm1Index = middle - 1;
          itm2Index = middle - 2;
          otm1Index = middle + 1;
          otm2Index = middle + 2;
          break;
        }
        // Spot price closer to the previous strike
        else {
          atmIndex = middle - 1;
          itm1Index = middle - 2;
          itm2Index = middle - 3;
          otm1Index = middle;
          otm2Index = middle + 1;
          break;
        }       
      }
      // Spot price equals the middle strike
      else if(spot == parseInt(middleCol[0][11],10)) {
        atmIndex = middle;
        itm1Index = middle - 1;
        itm2Index = middle - 2;
        otm1Index = middle + 1;
        otm2Index = middle + 2;
        break;
      }
    } // End for each strike
    
    // Split IV row into individual colums using delimiter "?"
    var atmCol = Utilities.parseCsv(rows[atmIndex][0], "?");
    var ceItm1Col = Utilities.parseCsv(rows[itm1Index][0], "?");
    var ceItm2Col = Utilities.parseCsv(rows[itm2Index][0], "?");
    var ceOtm1Col = Utilities.parseCsv(rows[otm1Index][0], "?");
    var ceOtm2Col = Utilities.parseCsv(rows[otm2Index][0], "?");
    
    // Extract CE and PE IVs
    atmCeIv = atmCol[0][4];
    atmPeIv = atmCol[0][18];
    itm1CeIv = ceItm1Col[0][4];
    itm1PeIv = ceOtm1Col[0][18];
    itm2CeIv = ceItm2Col[0][4];
    itm2PeIv = ceOtm2Col[0][18];
    otm1CeIv = ceOtm1Col[0][4];
    otm1PeIv = ceItm1Col[0][18];
    otm2CeIv = ceOtm2Col[0][4];
    otm2PeIv = ceItm2Col[0][18];
    
    if(atmCeIv.indexOf("-") == -1) { ceDivisor++; atmCeIv = parseFloat(atmCeIv); }
    else atmCeIv = 0;
    if(itm1CeIv.indexOf("-") == -1) { ceDivisor++; itm1CeIv = parseFloat(itm1CeIv); }
    else itm1CeIv = 0;
    if(itm2CeIv.indexOf("-") == -1) { ceDivisor++; itm2CeIv = parseFloat(itm2CeIv); }
    else itm2CeIv = 0;
    if(otm1CeIv.indexOf("-") == -1) { ceDivisor++; otm1CeIv = parseFloat(otm1CeIv); }
    else otm1CeIv = 0;
    if(otm2CeIv.indexOf("-") == -1) { ceDivisor++; otm2CeIv = parseFloat(otm2CeIv); }
    else otm2CeIv = 0;
    
    if(atmPeIv.indexOf("-") == -1) { peDivisor++; atmPeIv = parseFloat(atmPeIv); }
    else atmPeIv = 0;
    if(itm1PeIv.indexOf("-") == -1) { peDivisor++; itm1PeIv = parseFloat(itm1PeIv); }
    else itm1PeIv = 0;
    if(itm2PeIv.indexOf("-") == -1) { peDivisor++; itm2PeIv = parseFloat(itm2PeIv); }
    else itm2PeIv = 0;
    if(otm1PeIv.indexOf("-") == -1) { peDivisor++; otm1PeIv = parseFloat(otm1PeIv); }
    else otm1PeIv = 0;
    if(otm2PeIv.indexOf("-") == -1) { peDivisor++; otm2PeIv = parseFloat(otm2PeIv); }
    else otm2PeIv = 0;
    
    callIv = (atmCeIv + itm1CeIv + itm2CeIv + otm1CeIv + otm2CeIv)/ceDivisor;
    putIv = (atmPeIv + itm1PeIv + itm2PeIv + otm1PeIv + otm2PeIv)/peDivisor;
  }
  catch (e) {
    Logger.log(e.message);
    return 0.0;
  }
  if(optionType == "call")
    return callIv;
  else return putIv;
}

function screen() {
  date = new Date();
  var startTime = date.getTime();
  currentMonthRange.getCell(1, 1).setValue(date.toLocaleTimeString());
  currentMonthRange.getCell(1, 1).setNote("Screen Start Time");
  
  var scripCount = getLastRow(backend, "A");
  
  for(var row = 1; row <= scripCount;  row++) {
    
    scripData = time.getRange(row, 10, 1, 1).getValue();
    rows = Utilities.parseCsv(scripData, "/");
    columns = Utilities.parseCsv(rows[0][0], " ");
    start = rows.length;
    spot = parseFloat(columns[0][3].substring(0, columns[0][3].indexOf("As")));
    scrip = columns[0][2];
    lotSize = lotSizeRanges[row-1];
    
    if(spot > 75) {
      // Screen only scrips that have an underlying value of above 75
      screenCe();
      screenPe();
    }
  }
  date = new Date();
  var endTime = date.getTime();
  currentMonthRange.getCell(1, 2).setValue(date.toLocaleTimeString());
  currentMonthRange.getCell(1, 2).setNote("Screen End Time");
  currentMonthRange.getCell(1, 3).setValue((endTime - startTime)/(1000*60)).setNumberFormat("#.##");
}

function getLastPrice() {
  date = new Date();
  var startTime = date.getTime();
  var entries = getLastRow(currentMonth, "A");
    
  currentMonthRange.getCell(1, 4).setValue(date.toLocaleTimeString());
  currentMonthRange.getCell(1, 4).setNote("Get Last Price Start Time");
  currentMonthRange = currentMonth.getRange(3, 1, entries, 20);
  currentMonthRange.sort({column: 3, ascending: true});
  currentMonthRange = currentMonth.getRange(1, 1, entries, 20);
  currentMonthRangeValues = currentMonthRange.getValues();
  
  currentMonthRangeValues[0][12] = "=IF(TODAY()<CONCAT(\"$B\", \"$3\"), SUBTOTAL(9, Q:Q), SUBTOTAL(9, R:R))";
  currentMonthRangeValues[0][13] = "=SUBTOTAL(9, G:G,I:I)-(SUBTOTAL(3, G:G)-1)*75";
  currentMonthRangeValues[0][14] = "=SUBTOTAL(9, P:P)";
  currentMonthRangeValues[0][16] = "=P1/0.7";
  currentMonthRangeValues[0][18] = "=IFERROR(N1/O1,\"\")*0.7";
  currentMonthRangeValues[0][19] = "=IFERROR(IF(M1/O1<0, M1/O1,M1/O1*0.7),\"\")";
  
  for(var row = 3; row <= entries;  row++) {
    var rangeToAssess = currentMonth.getRange(row, 1, 1, 20).getValues();
    var scrip = rangeToAssess[0][2];
    var strike = parseFloat(rangeToAssess[0][3]);
    var lotSize = rangeToAssess[0][5];
    var isCe = 0;
    var isPe = 0;
    currentMonthRangeValues[row-1][12] = 0;
    currentMonthRangeValues[row-1][16] = -75;
    if(rangeToAssess[0][6] == "-") {
      isPe = 1;
    } else if(rangeToAssess[0][8] == "-") {
      isCe = 1;
    }
    
    if(scrip != columns[0][2]) {
      for(var scripIterator = 1; scripIterator <= scripCount; scripIterator++) {
        scripData = time.getRange(scripIterator, 10, 1, 1).getValue();
        rows = Utilities.parseCsv(scripData, "/");
        columns = Utilities.parseCsv(rows[0][0], " ");
        start = rows.length;
        spot = parseFloat(columns[0][3].substring(0, columns[0][3].indexOf("As")));
        lotSize = lotSizeRanges[scripIterator-1];
        if(scrip == columns[0][2])
          break;
      }
    }
    currentMonthRangeValues[row-1][13] = spot;
    if(isCe == 1) {
      for(var ceEntry = Math.round(rows.length/2); ceEntry <= rows.length; ceEntry++) {
        var strikeToCheck = Utilities.parseCsv(rows[ceEntry-1][0], "?");
        if(strike == parseFloat(strikeToCheck[0][11])){
          iv = parseFloat(strikeToCheck[0][4])/100;   
          currentMonthRangeValues[row-1][12] = parseFloat(strikeToCheck[0][8]) * lotSize - 75;
          if(isNaN(Number(currentMonthRangeValues[row-1][12]))) {
            currentMonthRangeValues[row-1][12] = "-";
          }
          currentMonthRangeValues[row-1][14] = NORMDIST(Math.log(spot/strike)/(iv*Math.sqrt(timeToExpiry)));
          if(isNaN(Number(currentMonthRangeValues[row-1][14]))) {
            currentMonthRangeValues[row-1][14] = "-";
          }
          currentMonthRangeValues[row-1][16] = "=IFERROR(G"+ row + "-M" + row + "-75,-75)";
          currentMonthRangeValues[row-1][17] = "=IF(G"+ row + "<>\"-\",IF(D"+ row + ">N"+ row + ",G"+ row + ",ABS(D"+ row + "-N"+ row + ")*F"+ row + "*-1),IF(I"+ row + "<>\"-\",IF(D"+ row + "<N"+ row + ",I"+ row + ",ABS(D"+ row + "-N"+ row + ")*F"+ row + "*-1),ABS(D"+ row + "-N"+ row + ")*F"+ row + "*-1))-75";
          currentMonthRangeValues[row-1][18] = (currentMonthRangeValues[row-1][6] - 75)/currentMonthRangeValues[row-1][15];
          currentMonthRangeValues[row-1][19] = "=IF(IF(TODAY()<CONCAT(\"$B\",\"$3\"), Q"+ row + "/P"+ row + ", R"+ row + "/P"+ row + ") < 0, IF(TODAY()<CONCAT(\"$B\",\"$3\"), Q"+ row + "/P"+ row + ", R"+ row + "/P"+ row + "), IF(TODAY()<CONCAT(\"$B\",\"$3\"), Q"+ row + "/P"+ row + ", R"+ row + "/P"+ row + ")*0.7)";
        }
      }
    } else if(isPe == 1) {
      for(var peEntry = Math.round(rows.length/2); peEntry > 0; peEntry--) {
        var strikeToCheck = Utilities.parseCsv(rows[peEntry-1][0], "?");      
        if(strike == parseFloat(strikeToCheck[0][11])){
          iv = parseFloat(strikeToCheck[0][18])/100;
          currentMonthRangeValues[row-1][12] = parseFloat(strikeToCheck[0][13]) * lotSize - 75;
          if(isNaN(Number(currentMonthRangeValues[row-1][12]))) {
            currentMonthRangeValues[row-1][12] = "-";
          }
          currentMonthRangeValues[row-1][14] = 1 - NORMDIST(Math.log(spot/strike)/(iv*Math.sqrt(timeToExpiry)));
          if(isNaN(Number(currentMonthRangeValues[row-1][14]))) {
            currentMonthRangeValues[row-1][14] = "-";
          }
          currentMonthRangeValues[row-1][16] = "=IFERROR(I"+ row + "-M" + row + "-75,-75)";
          currentMonthRangeValues[row-1][17] = "=IF(G"+ row + "<>\"-\",IF(D"+ row + ">N"+ row + ",G"+ row + ",ABS(D"+ row + "-N"+ row + ")*F"+ row + "*-1),IF(I"+ row + "<>\"-\",IF(D"+ row + "<N"+ row + ",I"+ row + ",ABS(D"+ row + "-N"+ row + ")*F"+ row + "*-1),ABS(D"+ row + "-N"+ row + ")*F"+ row + "*-1))-75";
          currentMonthRangeValues[row-1][18] = (currentMonthRangeValues[row-1][8] - 75)/currentMonthRangeValues[row-1][15];
          currentMonthRangeValues[row-1][19] = "=IF(IF(TODAY()<CONCAT(\"$B\",\"$3\"), Q"+ row + "/P"+ row + ", R"+ row + "/P"+ row + ") < 0, IF(TODAY()<CONCAT(\"$B\",\"$3\"), Q"+ row + "/P"+ row + ", R"+ row + "/P"+ row + "), IF(TODAY()<CONCAT(\"$B\",\"$3\"), Q"+ row + "/P"+ row + ", R"+ row + "/P"+ row + ")*0.7)";
        }
      }
    }
  }
  
  date = new Date();
  var endTime = date.getTime();
  currentMonthRangeValues[0][4] = date.toLocaleTimeString();
  currentMonthRange.getCell(1, 5).setNote("Get Last Price End Time");
  currentMonthRangeValues[0][5] = ((endTime - startTime)/(1000*60));
  currentMonthRange.setValues(currentMonthRangeValues);
}
