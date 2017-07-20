function importCSVFromGoogleDrive(mode, filename) {

    if (DriveApp.getFilesByName("est_cards_" + mode + "-" + filename + ".csv").hasNext()){
        var file = DriveApp.getFilesByName("est_cards_" + mode + "-" + filename + ".csv").next();
        var csvData = Utilities.parseCsv(file.getBlob().getDataAsString(), ';');
        var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(mode + "-" + filename);
        sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);

        replaceInSheet(sheet, "0", "1");

        sheet.insertColumnAfter(2);
        var cell_concat = sheet.getRange(2,3);
        var formulas = [
            ["=CONCATENATE(INDEX(SPLIT(B2, \"-\"),1,1),A2)"]
        ];
        cell_concat.setFormulas(formulas);
        cell_concat.copyTo(sheet.getRange(3, 3, sheet.getLastRow() - 2));

        var cell = sheet.getRange(1,4);
        cell.setValue(mode + " in lower");
        cell = sheet.getRange(1,5);
        cell.setValue(mode + " in upper");
        cell = sheet.getRange(1,6);
        cell.setValue(mode + " in conf");
        cell = sheet.getRange(1,7);
        cell.setValue(mode + " out lower");
        cell = sheet.getRange(1,8);
        cell.setValue(mode + " out upper");
        cell = sheet.getRange(1,9);
        cell.setValue(mode + " out conf")

    }

}

function importCSVFromGoogleDriveThreeTimes(filename, delete_files, load_images, image_names) {

    importCSVFromGoogleDrive("baseline", filename);
    importCSVFromGoogleDrive("training", filename);
    importCSVFromGoogleDrive("validation", filename);


    var baseline_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("baseline" + "-" + filename);
    baseline_sheet.insertColumnAfter(5);
    baseline_sheet.insertColumnAfter(5);
    baseline_sheet.insertColumnAfter(10);
    baseline_sheet.insertColumnAfter(10);

    var validation_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("validation" + "-" + filename);

    var cell = validation_sheet.getRange(1, 4, baseline_sheet.getLastRow());
    cell.copyTo(baseline_sheet.getRange(1, 6, baseline_sheet.getLastRow()));
    cell = validation_sheet.getRange(1, 5, baseline_sheet.getLastRow());
    cell.copyTo(baseline_sheet.getRange(1, 7, baseline_sheet.getLastRow()));
    cell = validation_sheet.getRange(1, 7, baseline_sheet.getLastRow());
    cell.copyTo(baseline_sheet.getRange(1, 11, baseline_sheet.getLastRow()));
    cell = validation_sheet.getRange(1, 8, baseline_sheet.getLastRow());
    cell.copyTo(baseline_sheet.getRange(1, 12, baseline_sheet.getLastRow()));

    function fill_column(sheet, header, formula, column, after){
        if (after != null){
            baseline_sheet.insertColumnAfter(after);
        }
        var cell = sheet.getRange(1,column);
        cell.setValue(header);
        cell = sheet.getRange(2,column);
        var formulas = [
            [formula]
        ];
        cell.setFormulas(formulas);
        cell.copyTo(sheet.getRange(2, column, sheet.getLastRow() - 1));
    }

    fill_column(baseline_sheet,
        "act in",
        "=FILTER(UNIQUE(FILTER('act_cards_July13-23uhr10'!D$2:D$1120, 'act_cards_July13-23uhr10'!C$2:C$1120 = C2)), ABS(UNIQUE(FILTER('act_cards_July13-23uhr10'!D$2:D$1120, 'act_cards_July13-23uhr10'!C$2:C$1120 = C2)) - E2) =MIN(ABS(UNIQUE(FILTER('act_cards_July13-23uhr10'!D$2:D$1120, 'act_cards_July13-23uhr10'!C$2:C$1120 = C2)) - E2)))",
        6,
        after=5);
    fill_column(baseline_sheet,
        "act out",
        "=FILTER(UNIQUE(FILTER('act_cards_July13-23uhr10'!F$2:F$1120, 'act_cards_July13-23uhr10'!C$2:C$1120 = C2)), ABS(UNIQUE(FILTER('act_cards_July13-23uhr10'!F$2:F$1120, 'act_cards_July13-23uhr10'!C$2:C$1120 = C2)) - K2) =MIN(ABS(UNIQUE(FILTER('act_cards_July13-23uhr10'!F$2:F$1120, 'act_cards_July13-23uhr10'!C$2:C$1120 = C2)) - K2)))",
        12,
        after=11);
    fill_column(baseline_sheet, "baseline select lower", "=J2/D2", 16);
    fill_column(baseline_sheet, "baseline select upper", "=K2/E2", 17);
    fill_column(baseline_sheet, "est select lower", "=M2/G2", 18);
    fill_column(baseline_sheet, "est select upper", "=N2/H2", 19);
    fill_column(baseline_sheet, "act select", "=L2/F2", 20);
    fill_column(baseline_sheet, "error diff", "=(ABS(T2-P2)+ABS(T2-Q2))-(ABS(T2-R2)+ABS(T2-S2))", 21);
    fill_column(baseline_sheet, "if error", "=IFERROR(U2, \"\")", 22);


    // copy formatting
    var act_cards_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("act_cards_July13-23uhr10");
    var range = act_cards_sheet.getRange(4, 10);
    range.copyFormatToRange(baseline_sheet, 22, 22, 2, baseline_sheet.getLastRow() - 1);


    if (load_images){
        fill_column(baseline_sheet, "split", "=INDEX(SPLIT(B2, \"-\"),1,1)", 23);

        cell = baseline_sheet.getRange(2, 1, baseline_sheet.getLastRow() - 1, baseline_sheet.getLastColumn());
        cell.sort(23);

        cell = baseline_sheet.getRange(2, 23, baseline_sheet.getLastRow() - 1);

        // get unique operator keys
        var valuesp = cell.getValues();
        var values = [];
        for(var i=0 ; i < valuesp.length ; i++) {
            values.push(valuesp[i][0]);
        }
        function onlyUnique(value, index, self) {
            return self.indexOf(value) === index;
        }
        var unique = values.filter( onlyUnique );

        // find first occurence of each operator and insert image
        for(var i=0 ; i < unique.length ; i++) {
            var rng = baseline_sheet.getRange(2, 23, baseline_sheet.getLastRow() - 1, 2);
            var data = rng.getValues();
            var search = unique[i];

            var found_yet = false;
            for (var j=0; j < data.length; j++) {
                if (found_yet == false){
                    if (data[j][0] == search) {
                        found_yet = true;
                        getImage(search, filename, j+1, 25, image_names);
                    }
                }
            }
        }
    }


    // aggregate values and populate dashboard
    var sheet_dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
    cell = sheet_dashboard.getRange(sheet_dashboard.getLastRow() + 1,1);
    cell.setValue(filename);

    var last_row = baseline_sheet.getLastRow();

    function aggregate(last_row, offset, formula, target){
        var cell = baseline_sheet.getRange(last_row + offset,22);
        var formulas = [
            [formula]
        ];
        cell.setFormulas(formulas);
        cell.copyValuesToRange(sheet_dashboard, target, target, sheet_dashboard.getLastRow(), sheet_dashboard.getLastRow());
    }

    aggregate(last_row, 2, "=COUNTIF(V2:V" + last_row + ",\">0\")", 2);
    aggregate(last_row, 3, "=COUNTIF(V2:V" + last_row + ",\"<0\")", 3);
    aggregate(last_row, 4, "=MEDIAN(V2:V" + last_row + ")", 7);
    aggregate(last_row, 5, "=AVERAGE(V2:V" + last_row + ")", 8);
    aggregate(last_row, 6, "=COUNTIF(V2:V" + last_row + ",\"=0\")", 4);


    if (delete_files){
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("baseline" + "-" + filename);
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("training" + "-" + filename);
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("validation" + "-" + filename);
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
}

function getImage(op_name, data_id, row, column, image_names) {
    if (DriveApp.getFilesByName(image_names + "-" + op_name + ".png").hasNext()){
        var file = DriveApp.getFilesByName(image_names + "-" + op_name + ".png").next();
        var id = file.getId();
        var gfile = DriveApp.getFileById(id);
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("baseline-" + data_id);
        sheet.insertImage(gfile.getThumbnail(), column, row);
    }
}

function replaceInSheet(sheet, to_replace, replace_with) {
    //get the current data range values as an array
    var values = sheet.getDataRange().getValues();
    //loop over the rows in the array
    for(var row in values){
        //use Array.map to execute a replace call on each of the cells in the row.
        var replaced_values = values[row].map(function(original_value){
            if (original_value.toString() == to_replace){
                return original_value.toString().replace(to_replace,replace_with);
            } else {
                return original_value.toString();
            }
        });
        //replace the original row values with the replaced values
        values[row] = replaced_values;
    }
    //write the updated values to the sheet
    sheet.getDataRange().setValues(values);
}


function runthis(){
    importCSVFromGoogleDriveThreeTimes("July14-12uhr00", delete_files=false, load_images=true, image_names="linear_training_validation-July13-23uhr10");
}
