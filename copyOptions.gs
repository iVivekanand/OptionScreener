var insertRangeValues = [];
insertRangeValues[0] = [];

function copyCe() {
  currentMonth.insertRows(3);
  
  insertRangeValues[0][0] = new Date();
  insertRangeValues[0][1] = expiry;
  insertRangeValues[0][2] = scrip;
  insertRangeValues[0][3] = strike;
  insertRangeValues[0][4] = spot;
  insertRangeValues[0][5] = lotSize;
  insertRangeValues[0][6] = ceValue;
  insertRangeValues[0][7] = currentItmProbability;
  insertRangeValues[0][8] = "-";
  insertRangeValues[0][9] = "-";
  insertRangeValues[0][10] = timeToExpiry * 365;
  insertRangeValues[0][11] = (strike-spot)/spot;
  insertRangeValues[0][12] = ceValue;
  insertRangeValues[0][13] = "";
  insertRangeValues[0][14] = "";
  insertRangeValues[0][15] = investment;
  insertRangeValues[0][16] = -75;
  insertRangeValues[0][17] = "=IF(G3<>\"-\",IF(D3>N3,G3,ABS(D3-N3)*F3*-1+G3),IF(I3<>\"-\",IF(D3<N3,I3,ABS(D3-N3)*F3*-1+I3),ABS(D3-N3)*F3*-1))+AA3";
  insertRangeValues[0][18] = (ceValue - 75)/investment*0.7;
  insertRangeValues[0][19] = "=IF(IF(TODAY()<B3, Q3/P3, R3/P3) < 0, IF(TODAY()<B3, Q3/P3, R3/P3), IF(TODAY()<B3, Q3/P3, R3/P3)*0.7)";
  insertRangeValues[0][20] = iv;
  insertRangeValues[0][21] = getAtmIv("call")/100;
  insertRangeValues[0][22] = "=IFERROR(1-COUNTIF(INDIRECT(CONCATENATE(\"Call_IV!\",LEFT(Y3,LEN(Y3)-2), \":\", LEFT(Y3,LEN(Y3)-2))), CONCAT(\">\",U3*100))/COUNT(INDIRECT(CONCATENATE(\"Call_IV!\",LEFT(Y3,LEN(Y3)-2), \":\", LEFT(Y3,LEN(Y3)-2)))), \"\")";
  insertRangeValues[0][23] = "=IFERROR(1-COUNTIF(INDIRECT(CONCATENATE(\"Call_IV!\",LEFT(Y3,LEN(Y3)-2), \":\", LEFT(Y3,LEN(Y3)-2))), CONCAT(\">\",V3*100))/COUNT(INDIRECT(CONCATENATE(\"Call_IV!\",LEFT(Y3,LEN(Y3)-2), \":\", LEFT(Y3,LEN(Y3)-2)))), \"\")";
  insertRangeValues[0][24] = "=ADDRESS(1,  MATCH(C3, Call_IV!$A$1:$IO$1, 0))";
  insertRangeValues[0][25] = "=AVERAGE(INDIRECT(CONCATENATE(\"Call_IV!\",LEFT(Y3,LEN(Y3)-2), \":\", LEFT(Y3,LEN(Y3)-2))))/100";
  insertRangeValues[0][26] = "=INDIRECT(CONCATENATE(\"IV_%ile_6m!\",Y3))";
  
  insertRange.setValues(insertRangeValues);
}

function copyPe() {
  currentMonth.insertRows(3);
  
  insertRangeValues[0][0] = new Date();
  insertRangeValues[0][1] = expiry;
  insertRangeValues[0][2] = scrip;
  insertRangeValues[0][3] = strike;
  insertRangeValues[0][4] = spot;
  insertRangeValues[0][5] = lotSize;
  insertRangeValues[0][6] = "-";
  insertRangeValues[0][7] = "-";
  insertRangeValues[0][8] = peValue;
  insertRangeValues[0][9] = currentItmProbability;
  insertRangeValues[0][10] = timeToExpiry * 365;
  insertRangeValues[0][11] = (strike-spot)/spot;
  insertRangeValues[0][12] = peValue;
  insertRangeValues[0][13] = "";
  insertRangeValues[0][14] = "";
  insertRangeValues[0][15] = investment;
  insertRangeValues[0][16] = -75;
  insertRangeValues[0][17] = "=IF(G3<>\"-\",IF(D3>N3,G3,ABS(D3-N3)*F3*-1+G3),IF(I3<>\"-\",IF(D3<N3,I3,ABS(D3-N3)*F3*-1+I3),ABS(D3-N3)*F3*-1))+AA3";
  insertRangeValues[0][18] = (peValue - 75)/investment*0.7;
  insertRangeValues[0][19] = "=IF(IF(TODAY()<B3, Q3/P3, R3/P3) < 0, IF(TODAY()<B3, Q3/P3, R3/P3), IF(TODAY()<B3, Q3/P3, R3/P3)*0.7)";
  insertRangeValues[0][20] = iv;
  insertRangeValues[0][21] = getAtmIv("put")/100;
  insertRangeValues[0][22] = "=IFERROR(1-COUNTIF(INDIRECT(CONCATENATE(\"Put_IV!\",LEFT(Y3,LEN(Y3)-2), \":\", LEFT(Y3,LEN(Y3)-2))), CONCAT(\">\",U3*100))/COUNT(INDIRECT(CONCATENATE(\"Put_IV!\",LEFT(Y3,LEN(Y3)-2), \":\", LEFT(Y3,LEN(Y3)-2)))), \"\")";
  insertRangeValues[0][23] = "=IFERROR(1-COUNTIF(INDIRECT(CONCATENATE(\"Put_IV!\",LEFT(Y3,LEN(Y3)-2), \":\", LEFT(Y3,LEN(Y3)-2))), CONCAT(\">\",V3*100))/COUNT(INDIRECT(CONCATENATE(\"Put_IV!\",LEFT(Y3,LEN(Y3)-2), \":\", LEFT(Y3,LEN(Y3)-2)))), \"\")";
  insertRangeValues[0][24] = "=ADDRESS(1,  MATCH(C3, Put_IV!$A$1:$IOs$1, 0))";
  insertRangeValues[0][25] = "=AVERAGE(INDIRECT(CONCATENATE(\"Put_IV!\",LEFT(Y3,LEN(Y3)-2), \":\", LEFT(Y3,LEN(Y3)-2))))/100";
  insertRangeValues[0][26] = "=INDIRECT(CONCATENATE(\"IV_%ile_6m!\",Y3))";
 
  insertRange.setValues(insertRangeValues);
}
