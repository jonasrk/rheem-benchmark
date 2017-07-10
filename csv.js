function importCSVFromGoogleDrive(mode, filename) {

    var file = DriveApp.getFilesByName("est_cards_" + mode + "-" + filename + ".csv").next();
    var csvData = Utilities.parseCsv(file.getBlob().getDataAsString(), ';');
    var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(mode + "-" + filename);
    sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
    sheet.insertColumnAfter(2);
    var cell = sheet.getRange(2,3)

    var formulas = [
        ["=CONCATENATE(B2,A2)"]
    ];

    cell.setFormulas(formulas);

    cell.copyTo(sheet.getRange(3, 3, sheet.getLastRow()));

    cell = sheet.getRange(1,4)
    cell.setValue(mode + " in lower")
    cell = sheet.getRange(1,5)
    cell.setValue(mode + " in upper")
    cell = sheet.getRange(1,6)
    cell.setValue(mode + " in conf")
    cell = sheet.getRange(1,7)
    cell.setValue(mode + " out lower")
    cell = sheet.getRange(1,8)
    cell.setValue(mode + " out upper")
    cell = sheet.getRange(1,9)
    cell.setValue(mode + " out conf")

}

function importCSVFromGoogleDriveThreeTimes(filename) {

    importCSVFromGoogleDrive("baseline", filename);
    importCSVFromGoogleDrive("training", filename);
    importCSVFromGoogleDrive("validation", filename);


    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("baseline" + "-" + filename);
    sheet.insertColumnAfter(5);
    sheet.insertColumnAfter(5);
    sheet.insertColumnAfter(10);
    sheet.insertColumnAfter(10);

    var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("validation" + "-" + filename);

    var cell = sheet2.getRange(1, 4, sheet.getLastRow() - 1)
    cell.copyTo(sheet.getRange(1, 6, sheet.getLastRow() - 1));
    cell = sheet2.getRange(1, 5, sheet.getLastRow())
    cell.copyTo(sheet.getRange(1, 7, sheet.getLastRow() - 1));
    cell = sheet2.getRange(1, 7, sheet.getLastRow())
    cell.copyTo(sheet.getRange(1, 11, sheet.getLastRow() - 1));
    cell = sheet2.getRange(1, 8, sheet.getLastRow())
    cell.copyTo(sheet.getRange(1, 12, sheet.getLastRow() - 1));

    sheet.insertColumnAfter(5);
    cell = sheet.getRange(1,6)
    cell.setValue("act in")
    cell = sheet.getRange(2,6)
    var formulas = [
        ["=VLOOKUP(C2,'act_cards_July06-00uhr23'!C2:I1120, 5, FALSE)"]
    ];
    cell.setFormulas(formulas);
    cell.copyTo(sheet.getRange(2, 6, sheet.getLastRow() - 2));

    sheet.insertColumnAfter(11);
    cell = sheet.getRange(1,12)
    cell.setValue("act out")
    cell = sheet.getRange(2,12)
    var formulas = [
        ["=VLOOKUP(C2,'act_cards_July06-00uhr23'!C2:I1120, 2, FALSE)"]
    ];
    cell.setFormulas(formulas);
    cell.copyTo(sheet.getRange(2, 12, sheet.getLastRow() - 2));

    cell = sheet.getRange(1,16)
    cell.setValue("baseline out error")
    cell = sheet.getRange(2,16)
    var formulas = [
        ["=( IF(L2<1,0,LOG10(L2))-IF(J2<1,0,LOG10(J2)))^2+( IF(L2<1,0,LOG10(L2))-IF(K2<1,0,LOG10(K2)))^2"]
    ];
    cell.setFormulas(formulas);
    cell.copyTo(sheet.getRange(2, 16, sheet.getLastRow() - 2));

    cell = sheet.getRange(1,17)
    cell.setValue("validation out error")
    cell = sheet.getRange(2,17)
    var formulas = [
        ["=( IF(L2<1,0,LOG10(L2))-IF(M2<1,0,LOG10(M2)))^2+( IF(L2<1,0,LOG10(L2))-IF(N2<1,0,LOG10(N2)))^2"]
    ];
    cell.setFormulas(formulas);
    cell.copyTo(sheet.getRange(2, 17, sheet.getLastRow() - 2));

    cell = sheet.getRange(1,18)
    cell.setValue("error diff")
    cell = sheet.getRange(2,18)
    var formulas = [
        ["=P2-Q2"]
    ];
    cell.setFormulas(formulas);
    cell.copyTo(sheet.getRange(2, 18, sheet.getLastRow() - 2));

    var sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("act_cards_July06-00uhr23");
    var range = sheet3.getRange(4, 10)
    range.copyFormatToRange(sheet, 18, 18, 2, sheet.getLastRow() - 2)

    // Fill Dashboard
    var sheet_dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");

    cell = sheet.getRange(135,18)
    var formulas = [
        ["=COUNTIF(R2:R131,\">0\")"]
    ];
    cell.setFormulas(formulas);
    cell.copyTo(sheet_dashboard.getRange(sheet.getLastRow(), 2));

    cell = sheet.getRange(136,18)
    var formulas = [
        ["=COUNTIF(R2:R131,\"<0\")"]
    ];
    cell.setFormulas(formulas);



}

importCSVFromGoogleDriveThreeTimes("July07-18uhr37");