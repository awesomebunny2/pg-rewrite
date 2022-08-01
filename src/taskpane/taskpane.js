//validation sheet password: fissh
//just a comment
$(() => {
    // DOCUMENT LOADED
    //console.log("DOCUMENT LOADED");

    /**
     * Clicking this prints a mouse =============================
     */
    $("#meece").on("click", () => {
        //location.reload();
        console.log("CLICKED🐭");
        showElement("#fissh", "show");
    })

    $(".ok").on("click", function() {
        showElement("#fissh", "hide");
        showFisshGif();
        //location.reload();
    });

    $(".dont").on("click", function() {
        showElement("#fissh", "hide");
    });


    $("#reload").on("click", function() {

        // Hide the message
        // alert("YES YOU ARE!");
        showMessage(undefined, "hide");
        location.reload();

    });

    //console.log("Hi");

    // function showMessage(msg, showHide) {
    //     if (showHide === "hide") {
    //         $("#message-text").empty();
    //         $("#message").css("display", "none");
    //     } else if (showHide === "show") {
    //         $("#message-text").text(msg);
    //         $("#message").css("display", "flex");
    //     }
    // }
});

//#region GLOBAL -------------------------------------------------------------------------------------------------------------------------------

    //#region TEST SUBJECTS -----------------------------------------------------------------------------------------------------------------------
        //CREATIVE REQUEST -Alfredo's Pizza - West Babylon - MENU - ~/*1338,52130,1*/~
        //CREATIVE REQUEST -Bella Napoli - Canfield - Env #10 8.5x11 S2 - ~/*1837,65845,1*/~
        //Re: Artist Request - Brickhouse Pizzeria - Richfield Springs - MENU - ~/*30601,72301,1*/~
    //#endregion ------------------------------------------------------------------------------------------------------------------------

    //#region I MIGHT NEED THIS BEGINNING MATERIAL SOME DAY ---------------------------------------------------------------------------------------

        /*
        * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
        * See LICENSE in the project root for license information.
        */

        /* global console, document, Excel, Office */

        // Office.onReady((info) => {
        //     if (info.host === Office.HostType.Excel) {
        //         document.getElementById("sideload-msg").style.display = "none";
        //         document.getElementById("app-body").style.display = "flex";
        //         document.getElementById("run").onclick = run;
        //     }
        // });

        // export async function run() {
        //     try {
        //         await Excel.run(async (context) => {
        //             /**
        //                * Insert your Excel code here
        //                */
        //             const range = context.workbook.getSelectedRange();

        //             // Read the range address
        //             range.load("address");

        //             // Update the fill color
        //             range.format.fill.color = "yellow";

        //             await context.sync();
        //             console.log(`The range address was ${range.address}.`);
        //         });
        //     } catch (error) {
        //         console.error(error);
        //     }
        // }

    //#endregion -----------------------------------------------------------------------------------------------------------------------


    // var lookup = new Object();
    var productIDData = {};
    var projectTypeIDData = {};
    var pickupData = {};
    var proofToClientData = {};
    var creativeProofData = {};
    var officeHoursData = {};
    var tierLevelData = {};
    var changesData = {};
    var changesIDData = {};
    var printDateRefData = {};
    var groupRefData = {};
    var loop = true;
    var changeEvent;
    var selectionEvent;
    var snailPoop = {};
    var designManagersData = {};

    var rowIndexPostSort;

    var changedTable;

    var destinationTable;
    var destinationTableName;
    var destinationRows;
    var destinationTableRange;
    var destinationHeader;

    var completedTable;

    var activationEvent;

    var currentWorksheet;

    var previousSelection;

    var previousSelectionObj = {
        tableId: "",
        address: "",
        rowIndex: ""
    };

    var didTableChangeFire = false;
    //console.log(didTableChangeFire);

    var deactivationEvent;

    var deactivatedWorksheetId;

    var activatedWorksheet;

    var valPassword = "fissh";

    //var activatedTables;

    // var previousSelectionFill;

    // var previousSelectionFontColor;

    // var previousSelectionFontWeight;

    //var previousTableId;

    //var previousTableName;

    // var previousItems = new Object();

    // var cheeseFarts;

    // var test;

    //var activeProjectTable;


    /* Check if aevents are turned on
    Excel.run(async function(context) {
        context.runtime.load("enableEvents");
        await context.sync();
        console.log(context.runtime.enableEvents)
    });


    Excel.run(async function(context) {
        context.runtime.load("enableEvents");
        await context.sync();
        context.runtime.enableEvents = true;
        console.log(context.runtime.enableEvents);
    });
    */

//#endregion ----------------------------------------------------------------------------------------------------------------------------------

// Office.onReady((info) => {

//     // eventsOn();

//     if (info.host === Office.HostType.Excel) {


//         Excel.run(async (context) => {

//             registerEventHandlers();
//         });
//     };
// });





//#region ON READY ---------------------------------------------------------------------------------------------------------------------------

    //#region LOADS VALIDATION VALUES AND UPDATES DROPDOWN VALUES IN TASKPANE ------------------------------------------------

        Office.onReady((info) => {

            // eventsOn();
            // Excel.run(async (context) => {
            //     var activeSheet = context.workbook.worksheets.getActiveWorksheet();

            //     activeSheet.onActivated.add(function (event) {
            //         return Excel.run(async (context) => {
            //             console.log("The activated worksheet ID is: " + event.worksheetId);
            //             activeProjectTable = activeSheet.tables.getItemAt(0);

            //             await context.sync();
            //         });
            //     });
            // });

            if (info.host === Office.PlatformType.OfficeOnline) {
                console.log("You're currently using the online version of Excel!")
            };

            if (info.host === Office.HostType.Excel) {

                Excel.run(async (context) => {

                    activationEvent = registerOnActivateHandler();
                    //deactivationEvent = registerOnDeactivationHandler();

                    //#region LOADING VALUES ---------------------------------------------------------------------------------

                        //load up the validation tables being referenced
                        var sheet = context.workbook.worksheets.getItem("Validation");
                        var productIDValTable = sheet.tables.getItem("ProductIDTable");
                        var projectTypeIDTable = sheet.tables.getItem("ProjectTypeIDTable")
                        var pickedUpValTable = sheet.tables.getItem("PickupTurnaroundTime");
                        var proofToClientValTable = sheet.tables.getItem("ArtTurnaroundTime");
                        var tierLevelValTable = sheet.tables.getItem("TierLevelsTable");
                        var creativeProofTable = sheet.tables.getItem("CreativeProofAdjust");
                        var officeHoursTable = sheet.tables.getItem("OfficeHours");
                        var changesDataTable = sheet.tables.getItem("ChangesData");
                        var changesIDTable = sheet.tables.getItem("ChangesIDTable");
                        var groupPrintDateRefTable = sheet.tables.getItem("dateTable");
                        var designManagersTable = sheet.tables.getItem("DesignManagersTable");
                        var activeSheet = context.workbook.worksheets.getActiveWorksheet().load("worksheetId");
                        //activeSheet.onChanged.add(handleChange);
                        // context.runtime.load("enableEvents");

                        //var leAllTables = context.workbook.tables.load("items/name");

                        var activeProjectTable = activeSheet.tables.getItemAt(0);

                        // doingPasswords(sheet);

                        // async function doingPasswords(sheet) {
                        //     let password = await passwordHandler();
                        //     passwordHelper(password);
                        //     await Excel.run(async (context) => {
                        //         //let activeSheet = context.workbook.worksheets.getActiveWorksheet();
                        //         //var sheetProtection = sheet.protection.load("isPasswordProtected");
                        //         //sheet.load(["format/*", "format/protection", "format/protection/protected"]);
                        //         sheet.load("protection/protected");

                        //         await context.sync();

                        //         //var sheetPasswordProtection = sheetProtection.isPasswordProtected;

                        //         var isProtected = sheet.protection.protected;

                        //         await context.sync();

                        //         if (!sheet.protection.protected) {
                        //             sheet.protection.protect(null, password);
                        //         };
                        //     });

                        //     // await Excel.run(async (context) => {
                        //     //     let activeSheet = context.workbook.worksheets.getActiveWorksheet();
                        //     //     activeSheet.load("protection/protected");
                        //     //     await context.sync();
                            
                        //     //     if (!activeSheet.protection.protected) {
                        //     //         activeSheet.protection.protect();
                        //     //     }
                        //     // });

                        // };
                    

                       


                        var workbookName = context.workbook.load("name");

                        // var activeCompletedTable = activeSheet.tables.getItemAt(1);

                        //var activeProjectTable;


                        // activeSheet.onActivated.add(function (event) {
                        //     return Excel.run(async (context) => {
                        //         console.log("The activated worksheet ID is: " + event.worksheetId);
                        //         activeProjectTable = activeSheet.tables.getItemAt(0);

                        //         await context.sync();
                        //     });
                        // });




                        //get data from the tables
                        var productIDBodyRange = productIDValTable.getDataBodyRange().load("values");
                        var projectTypeIDBodyRange = projectTypeIDTable.getDataBodyRange().load("values");
                        var pickedUpBodyRange = pickedUpValTable.getDataBodyRange().load("values");
                        var proofToClientBodyRange = proofToClientValTable.getDataBodyRange().load("values");
                        var tierLevelBodyRange = tierLevelValTable.getDataBodyRange().load("values");
                        var creativeProofBodyRange = creativeProofTable.getDataBodyRange().load("values");
                        var officeHoursBodyRange = officeHoursTable.getDataBodyRange().load("values");
                        var changesDataBodyRange = changesDataTable.getDataBodyRange().load("values");
                        var changesIDBodyRange = changesIDTable.getDataBodyRange().load("values");
                        var designManagersBodyRange = designManagersTable.getDataBodyRange().load("values");
                        var groupPrintDateRefRange = groupPrintDateRefTable.getDataBodyRange().load("values");


                    //#endregion ----------------------------------------------------------------------------------------------

                    //eventsFunction();
                    //changeEvent = context.workbook.tables.onChanged.add(onTableChangedEvents);
                    //tryCatch(changeEvent);

                    await context.sync()




                        //console.log(workbookName.name);

                        //console.log("I sharkded");

                        // if (currentWorksheet == undefined) {
                        //     currentWorksheet = activeSheet.id;
                        // };

                        //var leCurrentWorksheet = context.workbook.worksheets.getItem(currentWorksheet);

                        //var leCurrentProjectTable = leCurrentWorksheet.tables.getItemAt(0);


                        // activeSheet.onActivated.add(function (event) {
                        //     return Excel.run(async (context) => {
                        //         console.log("The activated worksheet ID is: " + event.worksheetId);
                        //         activeProjectTable = activeSheet.tables.getItemAt(0);

                        //         await context.sync();
                        //     });
                        // });

                        // var listOfCompletedTables = [];

                        // leAllTables.items.forEach(function (table) { //for each table in the workbook...
                        //     if (table.name.includes("Completed")) { //if the table name includes the word "Completed" in it...
                        //         listOfCompletedTables.push(table.name); //push the name of that table into an array
                        //     };
                        // });

                        // //returns true if the changedTable is a completed table from the array previously made, false if it is anything else
                        // var completedTableChanged = listOfCompletedTables.includes(changedTable.name);

                        //#region GRABBING DATA FROM VALIDATION AND WRITING TO CODE ----------------------------------------------


                            //#region PRODUCT ID DATA ----------------------------------------------------------------

                                var productIDArr = productIDBodyRange.values;

                                for (var row of productIDArr) {
                                    productIDData[row[0].trim()] = {
                                        "productID":row[0].trim(),
                                        "relativeProduct":row[1].trim(),
                                        "productCode":row[2].trim()
                                    };
                                };

                                // console.log("The productIDData is:");
                                // console.log(productIDData);

                            //#endregion -----------------------------------------------------------------------------


                            //#region PROJECT TYPE ID DATA ----------------------------------------------------------------

                                var projectTypeIDArr = projectTypeIDBodyRange.values;

                                for (var row of projectTypeIDArr) {
                                    projectTypeIDData[row[0].trim()] = {
                                        "projectType":row[0].trim(),
                                        "projectTypeCode":row[1].trim(),
                                    };
                                };

                                // console.log(projectTypeIDData);

                            //#endregion -----------------------------------------------------------------------------

                            //#region DESIGN MANAGERS DATA -----------------------------------------------------------

                                var designManagersArr = designManagersBodyRange.values;

                                for (var row of designManagersArr) {
                                    designManagersData[row[0].trim()] = {
                                        "designManager":row[0],
                                        "worksheetTabColor":row[1]
                                    };
                                };

                            //#endregion -----------------------------------------------------------------------------

                            //#region PICKED UP TURN AROUND TIME DATA ------------------------------------------------

                                var pickupArr = pickedUpBodyRange.values;
                                // console.log(pickupArr);

                                for (var row of pickupArr) {
                                    pickupData[row[0].trim()] = {
                                    "brandNewBuild":row[1],
                                    "brandNewBuildFromNatives":row[2],
                                    "brandNewBuildFromTemplate":row[3],
                                    "changesToExistingNatives":row[4],
                                    "specCheck":row[5],
                                    "weTransferUpload":row[6],
                                    "specialRequest":row[7],
                                    "other":row[8]
                                    };
                                };

                                // console.log(pickupData);

                            //#endregion ------------------------------------------------------------------------------


                            //#region PROOF TO CLIENT TIME DATA ------------------------------------------------

                                var proofToClientArr = proofToClientBodyRange.values;
                                // console.log(proofToClientArr);

                                for (var row of proofToClientArr) {
                                    proofToClientData[row[0].trim()] = {
                                    "brandNewBuild":row[1],
                                    "brandNewBuildFromNatives":row[2],
                                    "brandNewBuildFromTemplate":row[3],
                                    "changesToExistingNatives":row[4],
                                    "specCheck":row[5],
                                    "weTransferUpload":row[6],
                                    "specialRequest":row[7],
                                    "other":row[8]
                                    };
                                };

                                //console.log(proofToClientData);

                            //#endregion ------------------------------------------------------------------------------


                            //#region TIER LEVEL DATA ------------------------------------------------

                                var tierLevelArr = tierLevelBodyRange.values;

                                for (var row of tierLevelArr) {
                                    tierLevelData[row[0].trim()] = {
                                    "brandNewBuild":row[1],
                                    "brandNewBuildFromNatives":row[2],
                                    "brandNewBuildFromTemplate":row[3],
                                    "changesToExistingNatives":row[4],
                                    "specCheck":row[5],
                                    "weTransferUpload":row[6],
                                    "specialRequest":row[7],
                                    "other":row[8]
                                    };
                                };

                            //#endregion ------------------------------------------------------------------------------


                            //#region CREATIVE PROOF DATA ----------------------------------------------------------------

                                var creativeProofArr = creativeProofBodyRange.values;

                                for (var row of creativeProofArr) {
                                    creativeProofData[row[0].trim()] = {
                                        "creativeReviewProcess":row[1]
                                    };
                                };

                                // console.log(creativeProofData);

                            //#endregion -----------------------------------------------------------------------------


                            //#region OFFICE HOURS DATA ----------------------------------------------------------------

                                var officeHoursArr = officeHoursBodyRange.values;

                                for (var row of officeHoursArr) {
                                    officeHoursData[row[0].trim()] = {
                                        "weekday":row[0],
                                        "startTime":row[1],
                                        "endTime":row[2],
                                        "workDay":row[3],
                                        "workDayWithBreak":row[4]
                                    };
                                };

                                // console.log(officeHoursData);

                            //#endregion -----------------------------------------------------------------------------


                            //#region CHANGES DATA ------------------------------------------------

                                var changesDataArr = changesDataBodyRange.values;
                                // console.log(changesDataArr);

                                for (var row of changesDataArr) {
                                    changesData[row[0].trim()] = {
                                    "lightChanges":row[1],
                                    "moderateChanges":row[2],
                                    "heavyChanges":row[3],
                                    };
                                };

                                //console.log(changesData);

                            //#endregion ------------------------------------------------------------------------------


                            //#region CHANGES ID DATA ----------------------------------------------------------------

                                var changesIDArr = changesIDBodyRange.values;

                                for (var row of changesIDArr) {
                                    changesIDData[row[0].trim()] = {
                                        "changes":row[0].trim(),
                                        "changesCode":row[1].trim(),
                                    };
                                };

                                // console.log(changesIDData);

                            //#endregion -----------------------------------------------------------------------------


                            //#region PRINT DATE DATA ------------------------------------------------

                                var printDateRefArr = groupPrintDateRefRange.values;
                                // console.log(proofToClientArr);


                                for (var row of printDateRefArr) {
                                    var serialPrint = row[3];
                                    var formattedPrintDate = convertToDate(serialPrint);
                                    var aNewDate = new Date(formattedPrintDate);
                                    //converts the date into a simplifed format for dropdown: mm/dd/yy
                                    formattedPrintDate = [('' + (aNewDate.getMonth() + 1)).slice(-2),
                                      ('' + aNewDate.getDate()).slice(-2),
                                        (aNewDate.getFullYear() % 100)].join('/');

                                    printDateRefData[formattedPrintDate] = {
                                    "basedOnNow":row[0],
                                    "yearBasedOnNow":row[1],
                                    "weekBasedOnNow":row[2],
                                    "printDate":formattedPrintDate,
                                    "weekday":row[4],
                                    "adjust":row[5],
                                    "group":row[6]
                                    };
                                };

                                //console.log(proofToClientData);

                            //#endregion ------------------------------------------------------------------------------


                            //#region GROUP DATA ------------------------------------------------

                                var groupRefArr = groupPrintDateRefRange.values;
                                // console.log(proofToClientArr);

                                var gArr = [];


                                for (var row of groupRefArr) { //for each row in the dateTable...

                                    var x = row[6]; //the group letter of the current row

                                    var isGroupAlreadyPresent = false;

                                    for (var y of gArr) { //for each element in gArr...
                                        if (y == x) { //if an element from gArr = the current row group letter, then isGroupAlreadyPresent is true
                                            isGroupAlreadyPresent = true;
                                        };
                                    };

                                    if (isGroupAlreadyPresent == false) { //if group letter is not already in the data, create the object and properties for the row

                                        groupRefData[row[6].trim()] = {
                                            "basedOnNow":row[0],
                                            "yearBasedOnNow":row[1],
                                            "weekBasedOnNow":row[2],
                                            "printDate":row[3],
                                            "weekday":row[4],
                                            "adjust":row[5],
                                            "group":row[6]
                                        };

                                        gArr.push(x); //pushes the group letter of the current row into the gArr for further calculations

                                    };

                                };

                                //console.log(proofToClientData);

                            //#endregion ------------------------------------------------------------------------------

                        //#endregion --------------------------------------------------------------------------------------------------

                        //changeEvent = context.workbook.tables.onChanged.add(onTableChangedEvents);

                        //selectionEvent = activeProjectTable.onSelectionChanged.add(onTableSelectionChangedEvents);

                        //selectionEvent = leCurrentProjectTable.onSelectionChanged.add(onTableSelectionChangedEvents);


                        //selectionEvent = context.workbook.onSelectionChanged.add(onTableSelectionChangedEvents);





                        // if (completedTableChanged == true) {
                        //     selectionEvent = activeCompletedTable.onSelectionChanged.add(onTableSelectionChangedEvents);
                        // } else {
                        //     selectionEvent = activeProjectTable.onSelectionChanged.add(onTableSelectionChangedEvents);
                        // };

                        //selectionEvent = activeSheet.onSelectionChanged.add(onTableSelectionChangedEvents);

                    });

                    // changeEvent = context.workbook.tables.onChanged.add(onTableChangedEvents);


                // console.log(info);
                tryCatch(updateDropDowns);

                eventsOn();
                console.log("Events: ON  →  turned on in onReady function!");
                //updateDropDowns();
            };
        });

    //#endregion ---------------------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------------------


function passwordHelper(password) {
    if (null == password || password.trim() == "") {
      let errorMessage = "Password is expected but not provided";
      console.log(errorMessage);
    };
  };
  
  async function passwordHandler() {
    let settingName = "TheTestPasswordUsedByThisSnippet";
    let savedPassword = Office.context.document.settings.get(settingName);
    var testPassword;
    if (null == savedPassword || savedPassword.trim() == "") {
      //let item = document.getElementById("test-password");
        
      testPassword = valPassword;

      //let testPassword = item.hasAttribute("value") ? item.getAttribute("value") : null;
      if (null != testPassword && testPassword.trim() != "") {
        // store test password for retrieval upon re-opening this workbook
        Office.context.document.settings.set(settingName, testPassword);
        await Office.context.document.settings.saveAsync();
  
        savedPassword = testPassword;
      }
    } else {
      //document.getElementById("test-password").setAttribute("value", savedPassword);
      testPassword = valPassword;
      savedPassword = testPassword;
    }
  
    console.log("Test password is " + savedPassword);
  
    return savedPassword;
  }





//enables onSelectionChanged event upon inital load and any reloads of the taskpane
async function registerOnActivateHandler() {
    await Excel.run(async (context) => {

        // console.log("Inital load Selection Event: ");
        // console.log(selectionEvent);
        // removeSelectionEvent();

        console.log("Reload activation function fired!");
        let sheets = context.workbook.worksheets;
        var activeSheet = context.workbook.worksheets.getActiveWorksheet().load("worksheetId");
        activeSheet.load("name");

        context.runtime.load("enableEvents");

        var theAllTable = context.workbook.tables.load("count"); //all of the tables in the workbook
        theAllTable.load("items");

        var worksheetTables = activeSheet.tables.load("items/count");


        await context.sync();

        context.runtime.enableEvents = false;
        console.log("Events: OFF - Occured in registerOnActivateHandler");

        activeSheetName = activeSheet.name;

        console.log(activeSheetName);

        if (activeSheet.name == "Validation") {
            console.log("Active sheet is the Validation sheet, so onSelection & onChange events have been bound");
            sheets.onActivated.add(onActivate);
            sheets.onDeactivated.add(onDeactivate);
            eventsOn();
            console.log("Events: ON  →  turned on in the registerOnActivateHandler function, but for the Validation sheet");
            return;
        };

        var worksheetTablesCount = worksheetTables.count; //the number of tables in the workbook

        //cycles through each table in the workbook
        for (var p = 0; p < worksheetTablesCount; p++) { // <-- looping your tables
            var cycleTables = worksheetTables.getItemAt(p).load("name/worksheet");

            var cycleTableRows = cycleTables.rows.load("items");

            var tablesWorksheet = cycleTables.worksheet.load("name");

            var cycleBodyRange = cycleTables.getDataBodyRange().load("values"); //gets range of table
            cycleBodyRange.load("columnIndex");

            cycleBodyRange.load(["rowCount", "columnCount", "cellCount"]);

            var headerRange = cycleTables.getHeaderRowRange().load("values");

            const usedDataRange = cycleBodyRange.getUsedRangeOrNullObject(
                true /* valuesOnly */
            );

            var propertiesToGet = cycleBodyRange.getRowProperties({ //gets format properties of the rows in the table
                format: {
                    fill: {
                        color: true
                    },
                    font: {
                        bold: true,
                        color: true
                    }
                },
                rowIndex: true
            });

            await context.sync();

            var head = headerRange.values;

            var leTable = cycleBodyRange.values

            var listOfCompletedTables = [];
            if (cycleTables.name.includes("Completed")) { //if the table name includes the word "Completed" in it...
                listOfCompletedTables.push(cycleTables.name); //push the name of that table into an array
            };

            //returns true if the changedTable is a completed table from the array previously made, false if it is anything else
            var completedTableChanged = listOfCompletedTables.includes(cycleTables.name);

            if (tablesWorksheet.name !== "Validation" && usedDataRange.isNullObject !== true) { //ignore all tables in Validation sheet
                //cycles through each row in the table
                for (var iRow = 0; iRow < cycleBodyRange.rowCount; iRow++) {
                    var theRows = cycleTableRows.items
                    var rowValues = theRows[iRow].values
                    var rowRange = cycleTables.rows.getItemAt(iRow).getRange(); //range of row we are currently on
                    var rowProperties = propertiesToGet.value[iRow]; //loads in those row properites from eariler
                    var tableStart = cycleBodyRange.columnIndex;

                    var theWorksheet = context.workbook.worksheets.getItem(tablesWorksheet.name).load("name");

                    var rowInfoSorted = new Object();

                    for (var name of head[0]) {
                        theGreatestFunctionEverWritten(head, name, rowValues, leTable, rowInfoSorted, iRow);
                    }

                    if (rowProperties.format.fill.color == "#F5D9FF") { //if the row is purple, do the following...  #F5D9FF
                        console.log("Found a purple row!");
                        console.log(`Table: ${cycleTables.name}\nRow Index: ${iRow}`);
                        rowRange.format.fill.clear();
                        //will need to run conditional formatting function next
                        await context.sync();

                        conditionalFormatting(rowInfoSorted, tableStart, theWorksheet, iRow, completedTableChanged, rowRange, null)
                    };
                };
            };
        };

        for (var y = 0; y < worksheetTablesCount; y++) {
            var bonTable = worksheetTables.getItemAt(y);

            // bonTable.onChanged.add(theFunction);
            // await context.sync();
            //bonTable.onSelectionChanged.add(theFunction);
            changeEvent = bonTable.onChanged.add(onTableChanged);

            selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);

            // await context.sync().then(function() {
            //     selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);
            // });

            //console.log("bonTable fired!");
        };

        sheets.onActivated.add(onActivate);
        sheets.onDeactivated.add(onDeactivate);

        //deactivationEvent = sheets.onDeactivated.add(onDeactivate);


        //removeSelectionEvent();

        //changeEvent = context.workbook.tables.onChanged.add(onTableChangedEvents);


        console.log("A handler has been registered for the OnActivate event.");

        eventsOn();
        console.log("Events: ON  →  turned on in the registerOnActivateHandler function, typically triggered by a reload");

    }).catch (err => {
        console.log(err) // <--- does this log?
        showMessage(err, "show");
        context.runtime.enableEvents = true;
    });
};

// async function registerOnDeactivationHandler() {
//     await Excel.run(async(context) => {
//         var sheets = context.workbook.worksheets;
//         sheets.onDeactivated.add(onDeactivate);

//         var activeSheet = context.workbook.worksheets.getActiveWorksheet().load("worksheetId");

//         //context.runtime.load("enableEvents");

//         var theAllTable = context.workbook.tables.load("count"); //all of the tables in the workbook
//         theAllTable.load("items");

//         var worksheetTables = activeSheet.tables.load("items/count");

//         await context.sync();
//         console.log("A handler has been registered for the OnDeactivate event.");

//         var worksheetTablesCount = worksheetTables.count; //the number of tables in the workbook

//         //cycles through each table in the workbook
        
//         for (var y = 0; y < worksheetTablesCount; y++) {
//             var bonTable = worksheetTables.getItemAt(y);

//             // bonTable.onChanged.add(theFunction);
//             // await context.sync();
//             //bonTable.onSelectionChanged.add(theFunction);
//             changeEvent = bonTable.onChanged.add(onTableChanged);

//             removeSelectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);

//             // await context.sync().then(function() {
//             //     selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);
//             // });

//             //console.log("bonTable fired!");
//         };
// };

async function onDeactivate(eventArgs) {
    await Excel.run(async(context) => {
        console.log("Source of the onDeactivate event: " + eventArgs.source);

        console.log("The worksheet Id that was deactivated was: " + eventArgs.worksheetId);
        //removeSelectionEvent();

        deactivatedWorksheetId = eventArgs.worksheetId
        // selectionEvent.remove();

        // await context.sync();

        // selectionEvent = null;
        // console.log("Selection Event was removed");

        // var activatedTables = activatedWorksheet.tables.load("items/count");

        // await context.sync();


        // var worksheetTablesCount = activatedTables.count; //the number of tables in the workbook

        // for (var y = 0; y < worksheetTablesCount; y++) {
        //     var bonTable = activatedTables.getItemAt(y);
        //     changeEvent = bonTable.onChanged.add(onTableChanged);
        //     //removeChangeEvent();
        //     selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);
        //     console.log("Added selectionEvent to each table in the current sheet!");
        //     //removeSelectionEvent();


        //     // await context.sync().then(function() {
        //     //     selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);
        //     // });
        // };

        // var currentWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId);
        // console.log("The worksheet Id that was deactivated was: " + eventArgs.worksheetId);

        // var worksheetTables = currentWorksheet.tables.load("items/count");

        // await context.sync();

        // var worksheetTablesCount = worksheetTables.count; //the number of tables in the workbook

        // for (var y = 0; y < worksheetTablesCount; y++) {
        //     var bonTable = worksheetTables.getItemAt(y);
        //     //changeEvent = bonTable.onChanged.add(onTableChanged);
        //     //removeChangeEvent();
        //     selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);
        //     //removeSelectionEvent();


        //     // await context.sync().then(function() {
        //     //     selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);
        //     // });
        // };
    });
};

// async function theFunction(eventArgs) {
//     await Excel.run(async (context) => {
//         console.log(eventArgs);
//         context.runtime.load("enableEvents");

//         var currentWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId);
//         var worksheetTables = currentWorksheet.tables.load("items/count");



//         await context.sync();

//         context.runtime.enableEvents = false;
//         console.log("Events: OFF - Occured in theFunction!");
//         console.log("The Table Changed Fire Variable is: " + didTableChangeFire);

//         var worksheetTablesCount = worksheetTables.count; //the number of tables in the workbook

//         for (var y = 0; y < worksheetTablesCount; y++) {
//             var bonTable = worksheetTables.getItemAt(y);

//             if (eventArgs.type == "TableChanged") {
//                 console.log("Only Changed Event Fired!");
//                 changeEvent = bonTable.onChanged.add(onTableChanged);
//                 return;
//             } else if (eventArgs.type == "TableSelectionChanged" && didTableChangeFire == false) {
//                 console.log("Only Selection Event Fired!");
//                 selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);
//                 return;
//             } else {
//                 eventsOn();
//                 console.log("Events: ON  →  turned on at the end of theFunction!");
//             };
//             //return;
//         };

//         // eventsOn();
//         // console.log("Events: ON  →  turned on at the end of theFunction!");

//     // }).catch (err => {
//     //     console.log(err) // <--- does this log?
//     //     showMessage(err, "show");
//     //     context.runtime.enableEvents = true;
//     });
//     eventsOn();
//     console.log("Events: ON  →  turned on at the end of theFunction!");
// };


//when the worksheet changes, this fires and binds the events to the first table that is selected in the sheet
async function onActivate(eventArgs) {
    await Excel.run(async (context) => {

        // console.log("Worksheet change Selection Event: ");
        // console.log(selectionEvent);
        // removeSelectionEvent();
        console.log("Source of the onActivate event: " + eventArgs.source);

        console.log("Worksheet Switched (onActivate) function fired");
        // console.log(args);
        // console.log(args.type);

        //changeEvent = context.workbook.tables.onChanged.add(onTableChangedEvents);



        activatedWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId);

        console.log("The activated worksheet Id : " + eventArgs.worksheetId);

        //activatedTables = currentWorksheet.tables.load("items/count");

        // context.runtime.load("enableEvents");

        var theAllTable = context.workbook.tables.load("count"); //all of the tables in the workbook
        theAllTable.load("items");

        var worksheetTables = activatedWorksheet.tables.load("items/count");





        await context.sync();

        // context.runtime.enableEvents = false;
        // console.log("Events: OFF - Occured in registerOnActivateHandler");

        console.log("The deactivation worksheet ID is: " + deactivatedWorksheetId)

        var oldWorksheet = context.workbook.worksheets.getItem(deactivatedWorksheetId);

        var oldWorksheetTables = oldWorksheet.tables.load("items/count");

        await context.sync();




        // console.log("Reload activation function fired!");
        // let sheets = context.workbook.worksheets;
        // var activeSheet = context.workbook.worksheets.getActiveWorksheet().load("worksheetId");





        var oldWorksheetTablesCount = oldWorksheetTables.count; //the number of tables in the workbook

        //cycles through each table in the workbook
        for (var p = 0; p < oldWorksheetTablesCount; p++) { // <-- looping your tables
            var oldCycleTables = oldWorksheetTables.getItemAt(p).load("name/worksheet");

            var oldCycleTableRows = oldCycleTables.rows.load("items");

            var oldTablesWorksheet = oldCycleTables.worksheet.load("name");

            var oldCycleBodyRange = oldCycleTables.getDataBodyRange().load("values"); //gets range of table
            oldCycleBodyRange.load("columnIndex");

            oldCycleBodyRange.load(["rowCount", "columnCount", "cellCount"]);

            var oldHeaderRange = oldCycleTables.getHeaderRowRange().load("values");

            const oldUsedDataRange = oldCycleBodyRange.getUsedRangeOrNullObject(
                true /* valuesOnly */
            );

            var oldPropertiesToGet = oldCycleBodyRange.getRowProperties({ //gets format properties of the rows in the table
                format: {
                    fill: {
                        color: true
                    },
                    font: {
                        bold: true,
                        color: true
                    }
                },
                rowIndex: true
            });

            await context.sync();

            var oldHead = oldHeaderRange.values;

            var oldLeTable = oldCycleBodyRange.values

            var oldListOfCompletedTables = [];
            if (oldCycleTables.name.includes("Completed")) { //if the table name includes the word "Completed" in it...
                oldListOfCompletedTables.push(oldCycleTables.name); //push the name of that table into an array
            };

            //returns true if the changedTable is a completed table from the array previously made, false if it is anything else
            var oldCompletedTableChanged = oldListOfCompletedTables.includes(oldCycleTables.name);

            if (oldTablesWorksheet.name !== "Validation" && oldUsedDataRange.isNullObject !== true) { //ignore all tables in Validation sheet
                //cycles through each row in the table
                for (var aRow = 0; aRow < oldCycleBodyRange.rowCount; aRow++) {
                    var theOldRows = oldCycleTableRows.items
                    var oldRowValues = theOldRows[aRow].values
                    var oldRowRange = oldCycleTables.rows.getItemAt(aRow).getRange(); //range of row we are currently on
                    var oldRowProperties = oldPropertiesToGet.value[aRow]; //loads in those row properites from eariler
                    var oldTableStart = oldCycleBodyRange.columnIndex;

                    var theOldWorksheet = context.workbook.worksheets.getItem(oldTablesWorksheet.name).load("name");

                    var oldRowInfoSorted = new Object();

                    for (var oldName of oldHead[0]) {
                        theGreatestFunctionEverWritten(oldHead, oldName, oldRowValues, oldLeTable, oldRowInfoSorted, aRow);
                    }

                    if (oldRowProperties.format.fill.color == "#F5D9FF") { //if the row is purple, do the following...  #F5D9FF
                        console.log("Found a purple row!");
                        console.log(`Table: ${oldCycleTables.name}\nRow Index: ${aRow}`);
                        oldRowRange.format.fill.clear();
                        //will need to run conditional formatting function next
                        await context.sync();

                        conditionalFormatting(oldRowInfoSorted, oldTableStart, theOldWorksheet, aRow, oldCompletedTableChanged, oldRowRange, null)
                    };
                };
            };
        };

        location.reload();

        // var worksheetTablesCount = worksheetTables.count;

        // for (var y = 0; y < worksheetTablesCount; y++) {
        //     var bonTable = worksheetTables.getItemAt(y);

        //     // bonTable.onChanged.add(theFunction);
        //     // await context.sync();
        //     //bonTable.onSelectionChanged.add(theFunction);
        //     changeEvent = bonTable.onChanged.add(onTableChanged);

        //     selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);

        //     // await context.sync().then(function() {
        //     //     selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);
        //     // });

        //     //console.log("bonTable fired!");
        // };

        // sheets.onActivated.add(onActivate);
        // sheets.onDeactivated.add(onDeactivate);

        //deactivationEvent = sheets.onDeactivated.add(onDeactivate);


        //removeSelectionEvent();

        //changeEvent = context.workbook.tables.onChanged.add(onTableChangedEvents);


        // console.log("A handler has been registered for the OnActivate event.");

        // eventsOn();
        // console.log("Events: ON  →  turned on in the registerOnActivateHandler function, typically triggered by a reload");

        // var worksheetTablesCount = worksheetTables.count; //the number of tables in the workbook

        // for (var y = 0; y < worksheetTablesCount; y++) {
        //     var bonTable = worksheetTables.getItemAt(y);
        //     changeEvent = bonTable.onChanged.add(onTableChanged);
        //     //removeChangeEvent();
        //     selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);
        //     console.log("Added selectionEvent to each table in the current sheet!");
        //     //removeSelectionEvent();


        //     // await context.sync().then(function() {
        //     //     selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);
        //     // });
        // };

        //removeEvent();
        //removeChangeEvent();
        //removeSelectionEvent();

        // registerOnDeactivationHandler();





            //return;
       // });

       //eventsOn();

    }).catch (err => {
        console.log(err) // <--- does this log?
        showMessage(err, "show");
        context.runtime.enableEvents = true;
    });
};

async function removeSelectionEvent() {
    await Excel.run(selectionEvent.context, async(context) => {

        //var worksheetOld = selectionEvent.context.workbook.worksheets.getItem()
        selectionEvent.remove();

        await context.sync();

        selectionEvent = null;
        console.log("Selection Event was removed");
    });
};

async function removeChangeEvent() {
    await Excel.run(changeEvent.context, async(context) => {
        changeEvent.remove();

        await context.sync();

        changeEvent = null;
        console.log("Change Event was removed");
    });
};

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function onTableSelectionChangedEvents(eventArgs) {
    await Excel.run(/*previousSelection,*/ async (context) => {

      //console.log("This is the onTableSelectionChangedEvents eventArgs");
      //console.log(eventArgs);

        // if (selectionEvent !== null) {
        //     removeSelectionEvent();
        //     console.log("onSelectionHandler for the previous sheet was removed.")
        // };

        //console.log("Source of the onTableSelectionChanged event: " + eventArgs.source);


        //console.log("Running onTableSelectionChangedEvents!");

        // if (didTableChangeFire == true) {
        //     return;
        // };

        var theActiveWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId);
        var activeWorksheetTables = theActiveWorksheet.tables.load("items/count");

        var currentRange = theActiveWorksheet.getRange(eventArgs.address);
        var currentRow = currentRange.getRow();




        //context.runtime.load("enableEvents");

        await context.sync();

        var activeWorksheetTablesCount = activeWorksheetTables.count;

        // context.runtime.enableEvents = false;
        // console.log("Events: OFF - Occured in onTableSelectionChangedEvents");

        if (previousSelectionObj.tableId !== "") { //if user has made a selection prior to the current selection without triggering a reload, the previousSelectionObj should have arguments that will bring the user into this function to load in variables to handle the previous row highlighting
            //var previousTableId = eventArgs.tableId; // Table we came from

            var previousTable = context.workbook.tables.getItem(previousSelectionObj.tableId);
            previousTable.load("name/id");

            var previousTableRows = previousTable.rows.load("items");

            var previousWorksheet = previousTable.worksheet.load("name");

            var previousRowIndex = previousSelectionObj.rowIndex - 1;

            var previousSelectionRange = previousTable.rows.getItemAt(previousRowIndex).getRange();

            var previousTableRange = previousTable.getDataBodyRange().load("values");
            previousTableRange.load("columnIndex");

            var previousSelectionAddress = previousSelectionObj.address;

            var tablesAll = context.workbook.tables.load("items");

        };


        await context.sync();

        // for (var b = 0; b < activeWorksheetTablesCount; b++) {
        //     var oneOfTheTables = activeWorksheetTables.getItemAt(b).load("name/worksheet");
        //     var tableHeader = oneOfTheTables.getHeaderRowRange().load("address/values");

        //     await context.sync();

        //     var headerAddress = tableHeader.address;
        //     var selectionAddress = eventArgs.address;
        // }

        if (previousTable !== undefined) { //if previousTable is undefined, then the last function never fired, meaning that thsi is the first time the user is selecting anything this run. Since there are no previous selection variables stored, we skip this function. 

            var previousTableName = previousTable.name

            var newLeTable = previousTableRange.values;

            //console.log("Previous Table Event: The previous table name is: " + previousTableName + " with the selection address of: " + previousSelectionAddress);

            var listOfCompletedTables = [];

            tablesAll.items.forEach(function (table) { //for each table in the workbook...
                if (table.name.includes("Completed")) { //if the table name includes the word "Completed" in it...
                    listOfCompletedTables.push(table.name); //push the name of that table into an array
                };
            });

            //returns true if the changedTable is a completed table from the array previously made, false if it is anything else
            var completedTableChanged = listOfCompletedTables.includes(previousTable.name);

            var headerRangeToo = previousTable.getHeaderRowRange().load("values");

        };

        var worksheetName = context.workbook.worksheets.getActiveWorksheet().load("name/id");

        var selectedTable = context.workbook.tables.getItem(eventArgs.tableId).load("name");

        var range = context.workbook.getSelectedRange();
        range.load(['address', 'values', 'rowIndex']);

        var selectedTableRows = selectedTable.rows.load("items");
        var selectedTableRowsCount = selectedTable.rows.load("count");


        await context.sync();

        //console.log(`Selected Table Event: The current table name is: ${selectedTable.name} with the selection address of: ${eventArgs.address}`);

        var isTableEmpty = selectedTableRowsCount.count;

        if (isTableEmpty == 0) {
            console.log("Table is empty, so no highlighting was applied");
            // eventsOn();
            // console.log("Events: ON  →  turned on in the onTableSelectionChangedEvents function when the selected range was a part of an empty table");
            return;
        };

        //adds formatting to current row
        if (eventArgs.address !== "") { //if the selection address is not a part of a table, this function is skipped

            //applies border to selected row
            var rI = range.rowIndex;

            if (rI == 0) {
                //previousTable = undefined;
                console.log("Selection is in the header row, so no formatting was applied")
            } else {
                var bees = selectedTableRows.getItemAt(rI - 1).getRange();
                bees.load(["format/*", "format/fill", "format/borders", "format/font"]);
                bees.load("address");
            };

            await context.sync();

            if (eventArgs.address == previousSelectionAddress) {
                previousTable = undefined;
                console.log("Current selection address was the same as the previous selection address, so previous row formatting was prevented");
            };

            //removes formatting from previous row in same table
            if (previousTable !== undefined) {

                var headTwo = headerRangeToo.values;

                var zeRows = previousTableRows.items;
                var zeRowValues = zeRows[previousRowIndex].values

                var tableStart = previousTableRange.columnIndex;

                var rowInfoSorted = new Object();

                for (var name of headTwo[0]) {
                    theGreatestFunctionEverWritten(headTwo, name, zeRowValues, newLeTable, rowInfoSorted, previousRowIndex);
                }


                conditionalFormatting(rowInfoSorted, tableStart, previousWorksheet, previousRowIndex, completedTableChanged, previousSelectionRange, null);

            };

            if (rI !== 0) {
                bees.format.fill.color = "#F5D9FF";
                bees.format.font.color = "black";
    
                previousSelectionObj.tableId = eventArgs.tableId;
    
                previousSelectionObj.address = eventArgs.address;
    
                previousSelectionObj.rowIndex = rI;
            };


        } else { //if the selection address is not a part of a table AND there is a previous selection still highlighted...

            if (previousTable !== undefined) {

                var headTwo = headerRangeToo.values;

                var zeRows = previousTableRows.items;
                var zeRowValues = zeRows[previousRowIndex].values

                var tableStart = previousTableRange.columnIndex;

                var rowInfoSorted = new Object();

                for (var name of headTwo[0]) {
                    theGreatestFunctionEverWritten(headTwo, name, zeRowValues, newLeTable, rowInfoSorted, previousRowIndex);
                }


                conditionalFormatting(rowInfoSorted, tableStart, previousWorksheet, previousRowIndex, completedTableChanged, previousSelectionRange, null);

            };

        }

        //removeSelectionEvent();


        // eventsOn();
        // console.log("Events: ON  →  turned on in the onTableSelectionChangeEvents function once it had successfully finished running!");

    }).catch (err => {
        console.log(err) // <--- does this log?
        showMessage(err, "show");
        context.runtime.enableEvents = true;
    });
};

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



//#region TASKPANE ---------------------------------------------------------------------------------------------------------------------------


    //#region STYLIZING TASKPANE ELEMENTS ---------------------------------------------------------------------------------------------------


        //#region STYLIZE SPECIFIC CHARACTERS --------------------------------------------------------------------------------

            //this stylizes all * characeters in the container element to use different CSS than other character elements
            $("#container").each(function () {
                $(this).html($(this).html().replace(/(\*)/g, '<span style="color: rgba(220, 20, 60, 0.50); font-size: 9pt; padding-left: 1px; padding-bottom: 1px;">$1</span>'));
            });

        //#endregion ---------------------------------------------------------------------------------------------------------


        //#region WHEN THESE TASKPANE ITEMS ARE NO LONGER FOCUSED ON, DO SOMETHING --------------------------------------------

            $("#client").on("focusout", function() {

                removeWarningClass("#client", "#warning2");

            });

            $("#product").on("focusout", function() {

                removeWarningClass("#product", "#warning3");

            });

            $("#project-type").on("focusout", function() {

                removeWarningClass("#project-type", "#warning4");

            });

        //#endregion ----------------------------------------------------------------------------------------------------------


        //#region ADD WARNING CLASS ----------------------------------------------------------------------------------------------------

            /**
             * Shows the input warning for the input object and adds warning CSS formatting
             * @param {Object*} object Taskpane UI element to add warning class to
             * @param {Object} warning The warning that we are adding
             */
            function addWarningClass(object, warning) {

                var cheese = $(object).val();

                if ((($(object).val()) == "") || (($(object).val()) == null)) {

                    $(warning).show().text(`Required`); //show error
                    $(object).addClass("warning-box")
                    $(object).addClass("warning-box + .label")

                };

            };

        //#endregion -------------------------------------------------------------------------------------------------------------------


        //#region REMOVE WARNING CLASS --------------------------------------------------------------------------------------------------

            /**
             * Hides the input warning from the input object and remove warning CSS formatting
             * @param {Object} object Taskpane UI element to remove warning class from
             * @param {Object} warning The warning that we are removing
             */
            function removeWarningClass(object, warning) {

                if ((($(object).val()) !== "") || (($(object).val()) !== null)) {

                    $(warning).hide(); // Don't show the error
                    $(object).removeClass("warning-box")
                    $(object).removeClass("warning-box + .label")

                };

            };

        //#endregion -------------------------------------------------------------------------------------------------------------------


    //#endregion -------------------------------------------------------------------------------------------------------------


    //#region UPDATE DROPDOWNS --------------------------------------------------------------------------------------------------------------

        /**
         * Populates taskpane dropdowns with items from cooresponding validation sheet tables
         */
        async function updateDropDowns() {

            await Excel.run(async (context) => {

                //#region LOAD VALUES ----------------------------------------------------------------------------------------

                    var sheet = context.workbook.worksheets.getItem("Validation");
                    var productIDValTable = sheet.tables.getItem("ProductIDTable");
                    var projectTypeValTable = sheet.tables.getItem("ProjectTypeIDTable");
                    var groupPrintValTable = sheet.tables.getItem("GroupPrintTable");
                    var designManagersValTable = sheet.tables.getItem("DesignManagersTable");
                    var queueValTable = sheet.tables.getItem("QueueTable");
                    var tierValTable = sheet.tables.getItem("TierTable");
                    var tagsValTable = sheet.tables.getItem("TagsTable");
                    var groupDateRefTable = sheet.tables.getItem("dateTable");

                    // Get data from the table.
                    var productIDBodyRange = productIDValTable.getDataBodyRange().load("values");
                    var projectTypeBodyRange = projectTypeValTable.getDataBodyRange().load("values");
                    var groupPrintBodyRange = groupPrintValTable.getDataBodyRange().load("values");
                    var designManagersBodyRange = designManagersValTable.getDataBodyRange().load("values");
                    var queueBodyRange = queueValTable.getDataBodyRange().load("values");
                    var tierBodyRange = tierValTable.getDataBodyRange().load("values");
                    var tagsBodyRange = tagsValTable.getDataBodyRange().load("values");
                    var groupDateRefRange = groupDateRefTable.getDataBodyRange().load("values");

                //#endregion --------------------------------------------------------------------------------------------------

                await context.sync();

                //#region PRODUCT ID VALUES -----------------------------------------------------------------------------------

                    var productIDBodyValues = productIDBodyRange.values;

                    $("#product").empty();
                    $("#product").append($("<option disabled selected hidden></option>").val("").text(""));

                    productIDBodyValues.forEach(function(row) {

                        // Add an option to the select box
                        var option = `<option product-id="${row[0]}" relative-product="${row[1]}" product-code="${row[2]}">${row[1]}</option>`;

                        var x = $(`#product > option[relative-product="${row[1]}"]`).length; //finds current relative-product in current option in the product dropdown and returns how many are currently in the dropdown

                        if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
                            if (row[1] !== "") { //if the relative-product in option is empty, do not add to list
                                $("#product").append(option);
                            };
                        };
                    });

                //#endregion ---------------------------------------------------------------------------------------------------

                //#region PROJECT TYPE VALUES -----------------------------------------------------------------------------------

                    var projectTypeBodyValues = projectTypeBodyRange.values;

                    $("#project-type").empty();
                    $("#project-type").append($("<option disabled selected hidden></option>").val("").text(""));

                    projectTypeBodyValues.forEach(function(row) {

                        // Add an option to the select box
                        var option = `<option project-type-id="${row[0]}">${row[0]}</option>`;

                        var x = $(`#project-type > option[project-type-id="${row[0]}"]`).length;

                        if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
                            $("#project-type").append(option);
                        };
                    });

                //#endregion ---------------------------------------------------------------------------------------------------

                // //#region PRINT DATE & GROUP VALUES -----------------------------------------------------------------------------------

                //     var groupPrintBodyValues = groupPrintBodyRange.values;

                //     $("#print-date").empty();
                //     $("#print-date").append($("<option disabled selected hidden></option>").val("").text(""));
                //     $("#group").empty();
                //     $("#group").append($("<option disabled selected hidden></option>").val("").text(""));

                //     groupPrintBodyValues.forEach(function(row) {

                //         // Add an option to the select box
                //         var option = `<option group-id="${row[0]}" print-date-id="${row[1]}">${row[0]}</option>`;

                //         var x = $(`#print-date > option[print-date-id="${row[1]}"]`).length;
                //         var y = $(`#group > option[group-id="${row[0]}"]`).length;


                //         if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
                //             var leDate = convertToDate(`${row[1]}`);

                //             var d = new Date(leDate);

                //             //converts the date into a simplifed format for dropdown: mm/dd/yy
                //             leDate = [('' + (d.getMonth() + 1)).slice(-2), ('' + d.getDate()).slice(-2), (d.getFullYear() % 100)].join('/');

                //             //create proper html formatting for option to be added to select box
                //             var printDateOption = `<option print-date-convert="${leDate}">${leDate}</option>`;

                //             $("#print-date").append(printDateOption);
                //         };

                //         if (y == 0) { // Meaning, it's not there yet, because it's length count is 0
                //             $("#group").append(option);
                //         };
                //     });

                // //#endregion ---------------------------------------------------------------------------------------------------

                //#region PRINT DATE & GROUP VALUES -----------------------------------------------------------------------------------

                    var groupDateRefValues = groupDateRefRange.values;

                    $("#print-date").empty();
                    $("#print-date").append($("<option disabled selected hidden></option>").val("").text(""));
                    $("#group").empty();
                    $("#group").append($("<option disabled selected hidden></option>").val("").text(""));

                    groupDateRefValues.forEach(function(row) {

                        // Add an option to the select box
                        var option = `<option based-on-now="${row[0]}" year-based-on-now="${row[1]}" week-based-on-now="${row[2]}" print-date="${row[3]}" weekday="${row[4]}" adjust="${row[5]}" group="${row[6]}">${row[6]}</option>`;

                        var x = $(`#print-date > option[print-date="${row[3]}"]`).length;
                        var y = $(`#group > option[group="${row[6]}"]`).length;


                        if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
                            var leDate = convertToDate(`${row[3]}`);

                            var d = new Date(leDate);

                            //converts the date into a simplifed format for dropdown: mm/dd/yy
                            leDate = [('' + (d.getMonth() + 1)).slice(-2), ('' + d.getDate()).slice(-2), (d.getFullYear() % 100)].join('/');

                            //create proper html formatting for option to be added to select box
                            var printDateOption = `<option print-date-convert="${leDate}">${leDate}</option>`;

                            $("#print-date").append(printDateOption);
                        };

                        if (y == 0) { // Meaning, it's not there yet, because it's length count is 0
                            $("#group").append(option);
                        };
                    });

                //#endregion ---------------------------------------------------------------------------------------------------

                //#region ARTIST LEAD VALUES -----------------------------------------------------------------------------------

                    var designManagersBodyValues = designManagersBodyRange.values;

                    $("#design-managers").empty();
                    $("#design-managers").append($("<option disabled selected hidden></option>").val("").text(""));

                    designManagersBodyValues.forEach(function(row) {

                        // Add an option to the select box
                        var option = `<option design-managers-id="${row[0]}">${row[0]}</option>`;

                        var x = $(`#design-managers > option[design-managers-id="${row[0]}"]`).length;

                        if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
                            $("#design-managers").append(option);
                        };
                    });

                //#endregion ---------------------------------------------------------------------------------------------------

                //#region QUEUE VALUES -----------------------------------------------------------------------------------

                    var queueBodyValues = queueBodyRange.values;

                    $("#queue").empty();
                    $("#queue").append($("<option disabled selected hidden></option>").val("").text(""));

                    queueBodyValues.forEach(function(row) {

                        // Add an option to the select box
                        var option = `<option queue-id="${row[0]}">${row[0]}</option>`;

                        var x = $(`#queue > option[queue-id="${row[0]}"]`).length;

                        if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
                            $("#queue").append(option);
                        };
                    });

                //#endregion ---------------------------------------------------------------------------------------------------

                //#region TIER VALUES -----------------------------------------------------------------------------------

                    var tierBodyValues = tierBodyRange.values;

                    $("#tier").empty();
                    $("#tier").append($("<option disabled selected hidden></option>").val("").text(""));

                    tierBodyValues.forEach(function(row) {

                        // Add an option to the select box
                        var option = `<option tier-id="${row[0]}">${row[0]}</option>`;

                        var x = $(`#tier > option[tier-id="${row[0]}"]`).length;

                        if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
                            $("#tier").append(option);
                        };
                    });

                //#endregion ---------------------------------------------------------------------------------------------------

                //#region TAGS VALUES -----------------------------------------------------------------------------------

                    var tagsBodyValues = tagsBodyRange.values;

                    $("#tags").empty();
                    $("#tags").append($("<option disabled selected hidden></option>").val("").text(""));

                    tagsBodyValues.forEach(function(row) {

                        // Add an option to the select box
                        var option = `<option tags-id="${row[0]}">${row[0]}</option>`;

                        var x = $(`#tags > option[tags-id="${row[0]}"]`).length;

                        if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
                            $("#tags").append(option);
                        };
                    });

                //#endregion ---------------------------------------------------------------------------------------------------

            });
        };

    //#endregion ------------------------------------------------------------------------------------------------------------------


    //#region AUTO POPULATE TASKPANE FIELDS -------------------------------------------------------------------------------------------------


        //#region AUTO POPULATE TASKPANE BASED ON SUBJECT -----------------------------------------------------------------------------------


            $("#subject").keyup(() => tryCatch(subjectPasted));
            //$("#subject").keyup(() => subjectPasted());



            //#region SUBJECT PASTED FUNCTION ---------------------------------------------------------------------------------------------

                /**
                 * Auto-fills certain taskpane inputs based on the value pasted into the subject line input
                 */
                async function subjectPasted() {
                    var paste = $("#subject").val();
                    if (paste.length == 0) { // If what's pasted is empty

                        $("#warning1").hide(); // Don't show the error
                        $(this).removeClass("warning-box")
                        $(this).removeClass("warning-box + .label")
                        $("#client, #location, #product, #code").val(""); // Empty all inputs

                    } else if (!paste.includes("~/*")) { // If what's pasted does not contain "~/*"

                            $("#warning1").show().text(`This subject does not contain "~/*"`);

                        //    var warningCSS = {
                        //        "border": "2px",
                        //        "border-color": "red"
                        //    }
                        //    $(this).css("border", "2px solid red");

                            $(this).addClass("warning-box")
                            $(this).addClass("warning-box + .label")


                        //    $(this).css("pointer-events", "none");
                            $("#client, #location, #product, #code").val(""); // Empty all inputs

                    } else { // Probably a valid subject (contains ~/*)

                        $("#warning1").hide() // Hide error
                        $(this).removeClass("warning-box")
                        $(this).removeClass("warning-box + .label")


                        /** ------------------------------------------------------------
                         Parse the subject, fill the other inputs
                        ------------------------------------------------------------ */

                        // Split at "-"s
                        var splitPaste = paste.split("-");

                        var blanks = splitPaste.includes("");

                        if (blanks == true) {

                            var noBlanksArr = splitPaste.filter(function(x) {
                                return x !== "";
                            });

                        } else {

                            var noBlanksArr = splitPaste;

                        };


                        if (noBlanksArr[0].includes(":")) {

                            var str = noBlanksArr[0];

                            str = str.substring(str.indexOf(":") + 1);

                            noBlanksArr.splice(0, 1, str);

                        };

                        var hasRequest = noBlanksArr[0].includes("CREATIVE REQUEST") || noBlanksArr[0].includes("Creative Request") || noBlanksArr[0].includes("ARTIST REQUEST") || noBlanksArr[0].includes("Artist Request");

                        if (hasRequest == true) {

                            noBlanksArr.shift();

                        };

                        var plasticS = (noBlanksArr[noBlanksArr.length - 2]).trim();

                        if (plasticS == "S" || plasticS == "Flat") {

                            var plasticSIndex = noBlanksArr.indexOf(noBlanksArr[noBlanksArr.length - 2]);

                            noBlanksArr.splice(plasticSIndex, 1);

                            if (plasticS == "Flat") {

                                var productPostFlatIndex = noBlanksArr.indexOf(noBlanksArr[noBlanksArr.length - 2]);

                                noBlanksArr[productPostFlatIndex] = noBlanksArr[noBlanksArr.length - 2] + "Flat";

                            };

                        };

                        // .NET stuff at end (~/*20104,51824,2*/~)
                        // Remove spaces (just in case), "~/*", "*/~", then split at ","
                        var splitCodes = noBlanksArr[noBlanksArr.length - 1].replace(' ','').replace('~/*','').replace('*/~','').split(",");

                        var theClient = noBlanksArr[0].trim();

                        var theLocation = noBlanksArr[1].trim();

                        var theProduct = noBlanksArr[noBlanksArr.length - 2].trim();
                        // updatedProduct = productID(updatedProduct, 1);
                        // productID(updatedProduct, 1).then((updatedProduct) => {
                        //     console.log("A snail was exceuted");
                        //     return updatedProduct;
                        // });

                        var theCode = splitCodes[0].trim();

                        try {
                            console.log("The productIDData is:");
                            console.log(productIDData);
                            // var snailFace = productIDData["MENU"].productID;
                            var match = productIDData[theProduct].productID;
                            var updatedProduct = productIDData[theProduct].relativeProduct;

                            $("#client").val(theClient);

                            if (noBlanksArr.length > 3) {
                                $("#location").val(theLocation);
                            };

                            $("#code").val(theCode);

                            if (match == undefined) {
                                console.log("Product is undefined...");
                            } else {
                                $("#product").val(updatedProduct).removeClass("grey-sel");
                                console.log(`You matched ${updatedProduct}!`)
                            }

                        } catch (e) {
                            // Something was wrong with the subject
                            $("#warning1").show().text(`Something's wrong with this subject. Error: ` + e);
                        };

                    };
                };

            //#endregion ---------------------------------------------------------------------------------------------------------------------


        //#endregion ----------------------------------------------------------------------------------------------------------------------


        //#region AUTO POPULATED PRINT DATE BASED ON GROUP ---------------------------------------------------------------------------------

            $("#group").change(() => tryCatch(groupToPrintLink));

            function groupToPrintLink() {
                var group = $("#group").val();

                if (group.length == 0) {

                    $("#print-date").val("");

                } else {

                    var theGroup = group.trim();

                    try {

                        var printDateMatch = groupRefData[theGroup].printDate;

                        var formattedPrintDateMatch = convertToDate(printDateMatch);
                        var leNewDate = new Date(formattedPrintDateMatch);
                        //converts the date into a simplifed format for dropdown: mm/dd/yy
                        formattedPrintDateMatch = [('' + (leNewDate.getMonth() + 1)).slice(-2), ('' + leNewDate.getDate()).slice(-2), (leNewDate.getFullYear() % 100)].join('/');


                        $("#print-date").val(formattedPrintDateMatch);

                    } catch (e) {
                        console.log("Error with print date autofill based on group letter input. Please debug to resolve.")
                    };

                };
            };

        //#endregion -----------------------------------------------------------------------------------------------------------------------


        //#region AUTO POPULATE GROUP BASED ON PRINT DATE ----------------------------------------------------------------------------------

            $("#print-date").change(() => tryCatch(printToGroupLink));

            function printToGroupLink() {
                var lePrintDate = $("#print-date").val();

                if (lePrintDate.length == 0) {

                    $("#group").val("");

                } else {

                    var thePrintDate = lePrintDate.trim();

                    try {

                        var groupMatch = printDateRefData[thePrintDate].group;


                        $("#group").val(groupMatch);

                    } catch (e) {
                        console.log("Error with print date autofill based on group letter input. Please debug to resolve.")
                    };

                };
            };

        //#endregion ------------------------------------------------------------------------------------------------------------------------



    //#endregion ---------------------------------------------------------------------------------------------------------------------------


    //#region TASKPANE BUTTONS --------------------------------------------------------------------------------------------------------------

        // $("#meece").on("click", () => {
        //     console.log("CLICKED🐭");
        // })

        //#region ON SUBMIT CLICK ------------------------------------------------------------------------------------------------

            $("#submit").on("click", async function() {

                //adds warnings and doesn't write values to sheet
                if ((($("#client").val()) == "") || (($("#product").val()) == null) || (($("#project-type").val()) == null)) {
                    addWarningClass("#client", "#warning2");
                    addWarningClass("#product", "#warning3");
                    addWarningClass("#project-type", "#warning4");
                    return;
                };

                //removes warnings and writes values to sheet (through addAProject function)
                if ((($("#client").val()) !== "") || (($("#product").val()) !== null) || (($("#project-type").val()) !== null)) {
                    removeWarningClass("#client", "#warning2");
                    removeWarningClass("#product", "#warning3");
                    removeWarningClass("#project-type", "#warning4");
                    await addAProjectEvents().catch(err => {
                        console.log(err);
                        showMessage(err, "show");
                    });
                };
            });

        //#endregion -------------------------------------------------------------------------------------------------------------

        //#region ON CLEAR CLICK ----------------------------------------------------------------------------------------------------

        $("#clear").on("click", function() {

            $("#subject, #client, #location, #product, #code, #project-type, #csm, #print-date, #group, #design-managers, #queue, #tier, #tags, #start-override, #work-override, #notes").val(""); // Empty all inputs
            removeWarningClass("#subject", "#warning1");
            removeWarningClass("#client", "#warning2");
            removeWarningClass("#product", "#warning3");
            removeWarningClass("#project-type", "#warning4");

        });

        //#endregion ---------------------------------------------------------------------------------------------------------------

    //#endregion ------------------------------------------------------------------------------------------------------------------


//#endregion --------------------------------------------------------------------------------------------------------------------------------


//#region EDITING THE TABLE ------------------------------------------------------------------------------------------------------------------


    //#region ADDING A PROJECT FROM TASKPANE -------------------------------------------------------------------------------------------------


        //#region TURN OFF EVENTS BEFORE EXECUTING ADD A PROJECT -------------------------------------------------------------------------------

            /**
             * Turns events off, then executes the addAProject function
             */
            async function addAProjectEvents() {
                // throw("YOU MESSED UP");
                await Excel.run(async (context) => {
                    context.runtime.load("enableEvents");
                    await context.sync().then(function () {

                        //turns events off
                        context.runtime.enableEvents = false;
                        console.log("Events: OFF - Occured in addAProjectEvents");

                    });
                }).then(function() {
                    addAProject();
                });
            };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------


        //#region ADD A PROJECT ----------------------------------------------------------------------------------------------------------------

            /**
             * Generates Added date/time, turn around times for both the Picked Up / Started By and Proof To Client columns adjusted for office hours, adds these values to the table, then generates a priority number for each row based on the value in the Picked Up / Started By column, then sorts the data by priority
             */
            async function addAProject() {

                console.log("Add A Project was fired!");

                await Excel.run(async (context) => {

                    //#region LOAD VALUES ------------------------------------------------------------------------------------------------------

                        var sheet = context.workbook.worksheets.getActiveWorksheet().load("name");
                        sheet.load("tabColor");
                        //updating this variable to work for the changedTable will not work since the taskpane doesn't trigger an onchanged event until afterward
                        var sheetTable = sheet.tables.getItemAt(0).load("name"); //this is fine since the user will only ever be adding new projects to the unassigned table or the artist tables, which are all the first tables in their documents.
                        sheetTable.rows.add(null);

                        var sheetTableRows = sheetTable.rows.load("items");
                        var sheetTableRange = sheetTable.getDataBodyRange().load("values");
                        var sheetTableHeader = sheetTable.getHeaderRowRange().load("values");
                        context.runtime.load("enableEvents");


                    //#endregion ---------------------------------------------------------------------------------------------------------------


                    //#region GET INPUT FROM TASKPANE ------------------------------------------------------------------------------------------

                        // Data from DOM
                        var designManagersVal = $("#design-managers").val();
                        var queueVal = $("#queue").val();
                        var tierVal = $("#tier").val();
                        var subjectVal = $("#subject").val();
                        var clientVal = $("#client").val();
                        var locationVal = $("#location").val();
                        var productVal = $("#product").val();
                        var projectTypeVal = $("#project-type").val();
                        var csmVal = $("#csm").val();
                        var printDateVal = $("#print-date").val();
                        var groupVal = $("#group").val();
                        var tagsVal = $("#tags").val();
                        var codeVal = $("#code").val();
                        var startOverrideVal = $("#start-override").val();
                        var workOverrideVal = $("#work-override").val();
                        var notes = $("#notes").val();

                    //#endregion --------------------------------------------------------------------------------------------------------------


                    //#region WRITE ARRAY -----------------------------------------------------------------------------------------------------

                        // Data to send to Table
                        var write = [[
                            "", // 0 - Priority
                            designManagersVal, // 1 - Design Manager
                            queueVal, // 2 - Queue
                            tierVal, // 3 - Tier
                            subjectVal, // 4 - Subject
                            clientVal, // 5 - Client
                            locationVal, // 6 - Location
                            productVal, // 7 - Product
                            projectTypeVal, // 8 - Project Type
                            csmVal, // 9 - CSM
                            "", // 10 - Added
                            printDateVal, // 11 - Print Data
                            groupVal, // 12 - Group
                            "", // 13 - Picked Up / Started By
                            "", // 14 - Proof to Client
                            "", // 15 - Date of Last Edit
                            tagsVal, // 16 - Tags
                            "", // 17 - Status
                            codeVal, // 18 - Code
                            "", // 19 - Artist
                            notes, // 20 - Notes
                            startOverrideVal, // 21 - Start Override
                            workOverrideVal // 22 - Work Override
                        ]];

                    //#endregion -----------------------------------------------------------------------------------------------------------


                    await context.sync(); // BOOM!

                    var tableRowIndex = sheetTable.rows.count - 1;

                    var tableRowItems = sheetTableRows.items

                    var tableName = sheetTable.name;

                    var leSheetName = sheet.name;




                    var rangeOfTable = sheetTableRange.values;
                    var rowValuesOfTable = tableRowItems[tableRowIndex].values;
                    var headerOfTable = sheetTableHeader.values;

                    var tableRowInfo = new Object();

                    for (var name of headerOfTable[0]) {
                        theGreatestFunctionEverWritten(headerOfTable, name, rowValuesOfTable, rangeOfTable, tableRowInfo, tableRowIndex);
                    };

                    //#region GENERATE ADDED DATE -----------------------------------------------------------------------------------------

                        var now = new Date();
                        var toSerial = JSDateToExcelDate(now);

                        write[0][tableRowInfo.added.columnIndex] = toSerial;

                    //#endregion ---------------------------------------------------------------------------------------------------------


                    if (startOverrideVal == "") {
                        write[0][tableRowInfo.startOverride.columnIndex] = 0;
                        //startOverrideVal = 0;
                    };

                    if (workOverrideVal == "") {
                        write[0][tableRowInfo.workOverride.columnIndex] = 0;
                    };

                    var leStatus = statusAutofill(tableName);

                    write[0][tableRowInfo.status.columnIndex] = leStatus;

                    var leArtist = artistAutofill(tableName, leSheetName);

                    write[0][tableRowInfo.artist.columnIndex] = leArtist;

                    if ((designManagersVal == "" || designManagersVal == null) && (sheet.name !== "Unassigned Projects") && (sheet.name !== "Validation")) {

                        var sheetTabColor = sheet.tabColor;

                        var theDesignManager = "";

                        if (sheetTabColor == designManagersData.Emily.worksheetTabColor) {
                            theDesignManager = "Emily";
                        } else if (sheetTabColor == designManagersData.Peter.worksheetTabColor) {
                            theDesignManager = "Peter";
                        } else if (sheetTabColor == designManagersData.Luke.worksheetTabColor) {
                            theDesignManager = "Luke";
                        };

                        write[0][tableRowInfo.designManager.columnIndex] = theDesignManager;

                    };

                    //get the Project Type Coded variable from the Project Type ID Data based on the returned Project Type from the taskpane
                    var theProjectTypeCode = projectTypeIDData[projectTypeVal].projectTypeCode;

                    if((tierVal == "" || tierVal == null) && (sheet.name !== "Validation")) {
                        var defaultTier = tierLevelData[productVal][theProjectTypeCode];
                        write[0][tableRowInfo.tier.columnIndex] = defaultTier;
                    };


                    //#region GENERATE PICKED UP / TURN AROUND TIME VALUE -----------------------------------------------------------------

                        //returns turn around time value from the PickedUp Turn Around Time table based on the product and project type values
                        var pickedUpTurnAroundTime = pickupData[productVal][theProjectTypeCode];

                        //add start override time to # of hours
                        var pickedUpHours = pickedUpTurnAroundTime + Number(startOverrideVal);

                        //add new time to date added, then adjust for office hours
                        var addedDate = new Date(now);
                        var pickupOfficeHours = officeHours(addedDate, pickedUpHours);

                        //converts to excel readable format
                        var excelPickupOfficeHours = Number(JSDateToExcelDate(pickupOfficeHours));



                        //#region OFFICE HOURS TESTING VARIABLES ----------------------------------------------------------------------------

                            //BEFORE DAY: 44670.31389 (4/19/22 7:32 AM)
                            //AFTER DAY: 44670.78264 (4/19/22 6:47 PM)
                            //DURING DAY: 44670.59444 (4/19/22 2:16 PM)
                            //ON WEEKEND: 44667.29167 (4/16/22 7:00 AM)
                            //JUST BEFORE WEEKEND: 44666.45833 (4/15/22 11:00 AM)

                            // testingDate = 44667.29167;
                            // testingDate = convertToDate(testingDate);

                            // testingHours = 24;

                            // var PickupOHAdjust = officeHours(testingDate, testingHours, officeHoursData);

                        //#endregion --------------------------------------------------------------------------------------------------------

                        write[0][tableRowInfo.pickedUpStartedBy.columnIndex] = excelPickupOfficeHours;

                    //#endregion --------------------------------------------------------------------------------------------------------


                //#region GENERATE ART TURN AROUND TIME VALUE --------------------------------------------------------------------------

                        //returns turn around time value from the Proof To Client Turn Around Time table based on the product and project type values
                        var proofToClient = proofToClientData[productVal][theProjectTypeCode];

                        //returns the Creative Review Process value from said table based on the product
                        var creativeReview = creativeProofData[productVal].creativeReviewProcess;

                        //adds proof to client value to the creative review turn around time
                        var proofWithReview = proofToClient + creativeReview;

                        //add work override time to # of hours
                        var artTurnAround = proofWithReview + Number(workOverrideVal);

                        //add new time to the value previouskly found in the pickUpOfficeHours variable, then adjust for office hours
                        var proofToClientOfficeHours = officeHours(pickupOfficeHours, artTurnAround);

                        //converts to excel readable format
                        var excelProofToClientOfficeHours = Number(JSDateToExcelDate(proofToClientOfficeHours));

                        //#region OFFICE HOURS TESTING VARIABLES ----------------------------------------------------------------------------

                            //BEFORE DAY: 44670.31389 (4/19/22 7:32 AM)
                            //AFTER DAY: 44670.78264 (4/19/22 6:47 PM)
                            //DURING DAY: 44670.59444 (4/19/22 2:16 PM)
                            //ON WEEKEND: 44667.29167 (4/16/22 7:00 AM)
                            //JUST BEFORE WEEKEND: 44666.45833 (4/15/22 11:00 AM)

                            // testingDate = 44667.29167;
                            // testingDate = convertToDate(testingDate);

                            // testingHours = 24;

                            // var PickupOHAdjust = officeHours(testingDate, testingHours, officeHoursData);

                        //#endregion --------------------------------------------------------------------------------------------------------

                        write[0][tableRowInfo.proofToClient.columnIndex] = excelProofToClientOfficeHours;

                    //#endregion --------------------------------------------------------------------------------------------------------

                    var tablePickedUpColumnIndex = tableRowInfo.pickedUpStartedBy.columnIndex;
                    var tableProofToClientColumnIndex = tableRowInfo.proofToClient.columnIndex;

                    rangeOfTable[tableRowIndex] = write[0];

                    if (leSheetName == "Unassigned Projects") {

                        var gee = leSorting(tableRowInfo, rangeOfTable, tablePickedUpColumnIndex, write[0]);

                    } else {

                        var gee = leSorting(tableRowInfo, rangeOfTable, tableProofToClientColumnIndex, write[0]);

                    }


                    var kale = rowIndexPostSort;

                    sheetTableRange.values = gee;

                    console.log("Content has been added to the table through the taskpane successfully!")





                    // if (changedColumnIndex == rowInfo.artist.columnIndex || changedColumnIndex == rowInfo.status.columnIndex) {

                    //     var newChangedTableRows = destinationTable.rows.load("items");

                    //     var newBodyValues = destinationTable.getDataBodyRange().load("values");

                    // } else {

                    //     var newChangedTableRows = changedTable.rows.load("items");

                    //     var newBodyValues = changedTable.getDataBodyRange().load("values");

                    // };

                    var newSheetTableRows = sheetTable.rows.load("items");
                    var newSheetTableRange = sheetTable.getDataBodyRange().load("values");



                    await context.sync();

                    // var newTableRowItems = newSheetTableRows.items;

                    // var newRangeOfTable = newSheetTableRange.values;

                    // var newRowValuesOfTable = newTableRowItems[rowIndexPostSort].values;

                    // var newRowRange = newSheetTableRows.getItemAt(rowIndexPostSort).getRange();

                    // var newTableRowInfo = new Object();

                    // for (var name of headerOfTable[0]) {
                    //     theGreatestFunctionEverWritten(headerOfTable, name, newRowValuesOfTable, newRangeOfTable, newTableRowInfo, rowIndexPostSort);
                    // };

                    // conditionalFormatting(newTableRowInfo, 0, sheet, rowIndexPostSort, false, newRowRange, completedTable);





                    // var leTableSorted = newBodyValues.values

                    // var rowRangeSorted = newChangedTableRows.getItemAt(rowIndexPostSort).getRange();

                    // var tableRowsSorted = newChangedTableRows.items;

                    // var rowValuesSorted = tableRowsSorted[rowIndexPostSort].values;

                    // var rowInfoSorted = new Object();

                    // for (var name of head[0]) {
                    //     theGreatestFunctionEverWritten(head, name, rowValuesSorted, leTableSorted, rowInfoSorted, rowIndexPostSort);
                    // };




                    for (var m = 0; m < rangeOfTable.length; m++) {


                        var newTableRowItems = newSheetTableRows.items;

                        var newRangeOfTable = newSheetTableRange.values;

                        var newRowValuesOfTable = newTableRowItems[m].values;

                        var newRowRange = newSheetTableRows.getItemAt(m).getRange();

                        var newTableRowInfo = new Object();

                        for (var name of headerOfTable[0]) {
                            theGreatestFunctionEverWritten(headerOfTable, name, newRowValuesOfTable, newRangeOfTable, newTableRowInfo, m);
                        };

                        conditionalFormatting(newTableRowInfo, 0, sheet, m, false, newRowRange, null);

                    };








                    //writes the write array to the table
                    //sheetTable.rows.add(null /*add rows to the end of the table*/, write);

                   // await context.sync(); // BOOM!


                    //#region PRIORITY NUMBER GENERATION AND SORTING -------------------------------------------------------------------------

                        //assign priority numbers and sorts table
                        //priorityGenerationAndSortation();

                    //#endregion -------------------------------------------------------------------------------------------------------------


                    //console.log("Add a project function has completed, but the sub-function for priority & sorting is still running asyncronously");

                    //await context.sync();





                    // context.runtime.enableEvents = true;
                    // console.log("Events are turned on");

                });

                eventsOn();
                console.log("Events: ON  →  turned on in the addAProject function after a project was added to the sheet through the taskpane!");

            };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------




        //#region ADD A PROJECT SUB-FUNCTIONS -------------------------------------------------------------------------------------------------


            //#region AUTOFILL STATUS COLUMN --------------------------------------------------------------------

                function statusAutofill(tableName) {

                    if (tableName == "UnassignedProjects") { //if the table the row was inserted into is "UnassignedProjects", set status column to "Awaiting Artist"
                        var status = "Awaiting Artist";
                    };

                    if (tableName !== "UnassignedProjects") { //if the table the row was inserted into is not "UnassaignedProjects"...
                        var status = "Not Working";
                    };

                    return status;

                };

            //#endregion ----------------------------------------------------------------------------------------


            //#region AUTOFILL ARTIST COLUMN --------------------------------------------------------------------


                function artistAutofill(tableName, leSheetName) {

                    if (tableName == "UnassignedProjects") {
                        var artist = "Unassigned";
                    } else if (tableName !== "UnassignedProjects") {
                        var artist = leSheetName;
                    };
                    return artist;
                };


            //#endregion ----------------------------------------------------------------------------------------


            //#region OFFICE HOURS FUNCTION ---------------------------------------------------------------------------------------------------

                /**
                 * Adds adjustment hours to the date and adjusts to fit within office hours
                 * @param {Date} date the added date
                 * @param {Number} hoursToAdd The humber of adjustment hours to add to the added date
                 * @returns Date
                 */
                function officeHours(date, hoursToAdd) {

                    //#region FUNCTION VARIABLES ---------------------------------------------------------------------------------------------------

                        //date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
                        //gets the day of the week
                        var theDay = date.getDay();
                        if (theDay == 0) {theDay = "Sunday"} else if (theDay == 1) {theDay = "Monday"} else if (theDay == 2) {theDay = "Tuesday"} else if (theDay == 3) {theDay = "Wednesday"} else if (theDay == 4) {theDay = "Thursday"} else if (theDay == 5) {theDay = "Friday"} else if (theDay == 6) {theDay = "Saturday"};

                        var adjustmentMinutes = hoursToAdd * 60; // 12.5 hours = 750 minutes
                        var includesWeekends = false;

                        var current = new Date(date); //clone of the date variable that calculations will be made to
                        //current.setMinutes(current.getMinutes() - current.getTimezoneOffset()); //removes time zone offset to bring all dates to the same level

                        //#region SET DATES WITH 0 TIME -----------------------------------------------------------------------------------------------

                            //set workDayStart date to = date, but have the time be 0 (will assign times to later)
                            var workDayStart = new Date(date);
                            workDayStart.setHours(0);
                            workDayStart.setMinutes(0);
                            workDayStart.setSeconds(0);
                            workDayStart.setMilliseconds(0);

                            //set workDayEnd date to = date, but have the time be 0 (will assign times to later)
                            var workDayEnd = new Date(date);
                            workDayEnd.setHours(0);
                            workDayEnd.setMinutes(0);
                            workDayEnd.setSeconds(0);
                            workDayEnd.setMilliseconds(0);

                        //#endregion --------------------------------------------------------------------------------------------------------------

                        //#region CREATE START TIME AND END TIME VARIABLES FOR THE DATE ---------------------------------------------------------------

                            var weekdayVars = officeHoursData[theDay] //returns all the info for the weekday that date lands on

                            //#region CREATE START TIME DATE --------------------------------------------------------------------------------------

                                //this varibale will have the correct start time but will still be using the ground 0 date for serial numbers. Will be adjusted up next
                                var theStart = convertToDate(weekdayVars.startTime); //converts serial number to JSDate for start of work day
                                //sets that date of theStart to be at ground 0 for JSDates
                                theStart.setFullYear(1970);
                                theStart.setMonth(0);
                                theStart.setDate(1);
                                theStart.setMinutes(theStart.getMinutes() - theStart.getTimezoneOffset()); //removes time zone offset to bring all dates to the same level
                                //gives us the milliseconds between 0 and this time
                                var fartTime = theStart.getTime();
                                workDayStart.setMilliseconds(fartTime); //adds the startTime to the correct date variable from eariler

                            //#endregion ---------------------------------------------------------------------------------------------------------

                            //#region CREATE END TIME DATE ---------------------------------------------------------------------------------------

                                //this varibale will have the correct end time but will still be using the ground 0 date for serial numbers. Will be adjusted up next
                                var theEnd = convertToDate(weekdayVars.endTime); //converts serial number to JSDate for end of work day
                                //sets that date of theEnd to be at ground 0 for JSDates
                                theEnd.setFullYear(1970);
                                theEnd.setMonth(0);
                                theEnd.setDate(1);
                                theEnd.setMinutes(theEnd.getMinutes() - theEnd.getTimezoneOffset()); //removes time zone offset to bring all dates to the same level
                                //gives us the milliseconds between 0 and this time
                                var shartTime = theEnd.getTime();
                                workDayEnd.setMilliseconds(shartTime); //adds the endTime to the correct date variable from eariler

                            //#endregion -----------------------------------------------------------------------------------------------------------

                        //#endregion ---------------------------------------------------------------------------------------------------------------

                    //#endregion -------------------------------------------------------------------------------------------------------------------


                    //#region WHILE ADJUSTMENT NUMBER REMAINS POSITIVE ----------------------------------------------------------------------------

                        // if (current.toLocaleDateString('en-US') == date.toLocaleDateString('en-US')) {
                        //     current.setDate(current.getDate() + 1);
                        // }

                        while(adjustmentMinutes > 0) {

                            //#region RECALCULATE START AND END TIMES IF DATE ADVANCES -----------------------------------------------------------

                                var currentInfo = shortDate(current);
                                var dateInfo = shortDate(date);

                                // if (current.toLocaleDateString('en-US') !== date.toLocaleDateString('en-US')) { //if we go on into another day, recalculate start and end time dates
                                //if (current.getDay() !== date.getDay()) { //if we go on into another day, recalculate start and end time dates
                                if (currentInfo !== dateInfo) {


                                    //gets the day of the week
                                    var theDay = current.getDay();

                                    if (theDay == 0) {theDay = "Sunday"} else if (theDay == 1) {theDay = "Monday"} else if (theDay == 2) {theDay = "Tuesday"} else if (theDay == 3) {theDay = "Wednesday"} else if (theDay == 4) {theDay = "Thursday"} else if (theDay == 5) {theDay = "Friday"} else if (theDay == 6) {theDay = "Saturday"};

                                    //#region SET DATES WITH 0 TIME (CURRENT) ----------------------------------------------------------------------

                                        //set workDayStart date to = current date, but have the time be 0 (will assign times to later)
                                        var workDayStart = new Date(current);
                                        workDayStart.setHours(0);
                                        workDayStart.setMinutes(0);
                                        workDayStart.setSeconds(0);
                                        workDayStart.setMilliseconds(0);

                                        //set workDayEnd date to = current date, but have the time be 0 (will assign times to later)
                                        var workDayEnd = new Date(current);
                                        workDayEnd.setHours(0);
                                        workDayEnd.setMinutes(0);
                                        workDayEnd.setSeconds(0);
                                        workDayEnd.setMilliseconds(0);

                                    //#endregion ---------------------------------------------------------------------------------------------------

                                    //#region CREATE START TIME AND END TIME VARIABLES FOR THE DATE (CURRENT) --------------------------------------

                                        //gets start and end times of date's work day
                                        weekdayVars = officeHoursData[theDay] //returns all the info for the weekday that date lands on

                                        //#region CREATE START TIME DATE (CURRENT) -----------------------------------------------------------------

                                            theStart = convertToDate(weekdayVars.startTime); //converts serial number to JSDate for start of work day
                                            theStart.setFullYear(1970);
                                            theStart.setMonth(0);
                                            theStart.setDate(1);
                                            theStart.setMinutes(theStart.getMinutes() - theStart.getTimezoneOffset()); //removes time zone offset to bring all dates to the same level
                                            fartTime = theStart.getTime();
                                            workDayStart.setMilliseconds(fartTime);

                                        //#endregion ------------------------------------------------------------------------------------------------

                                        //#region CREATE END TIME DATE (CURRENT) --------------------------------------------------------------------

                                            theEnd = convertToDate(weekdayVars.endTime); //converts serial number to JSDate for end of work day
                                            theEnd.setFullYear(1970);
                                            theEnd.setMonth(0);
                                            theEnd.setDate(1);
                                            theEnd.setMinutes(theEnd.getMinutes() - theEnd.getTimezoneOffset()); //removes time zone offset to bring all dates to the same level

                                            shartTime = theEnd.getTime();
                                            workDayEnd.setMilliseconds(shartTime);

                                        //#endregion ------------------------------------------------------------------------------------------------

                                        //date = new Date(current);

                                    //#endregion ----------------------------------------------------------------------------------------------------

                                };

                            //#endregion -----------------------------------------------------------------------------------------------------------

                            //#region INCREMENT --------------------------------------------------------------------------------------------------

                                //if current is still within the workday and not on a weekend, subtract 1 minute from the adjustment number
                                if(current > workDayStart && current < workDayEnd && (includesWeekends ? current.getDay() !== 0 && current.getDay() !== 6 : true)) {
                                    adjustmentMinutes--;
                                };
                                current.setTime(current.getTime() + 1000 * 60); //adds 1 minute to current time

                            //#endregion ----------------------------------------------------------------------------------------------------------

                        };

                    //#endregion -------------------------------------------------------------------------------------------------------------------

                    return current;

                };

            //#endregion ----------------------------------------------------------------------------------------------------------------------


            function shortDate(aDate) {
                var day = aDate.getDate();
                var month = aDate.getMonth();
                var year = aDate.getFullYear();
                var output = `${month} ${day} ${year}`;
                return output;
            };




            //#region OLD OFFICE HOURS CODE ---------------------------------------------------------------------------------


                // //#region OFFICE HOURS ---------------------------------------------------------------------------------------

                //     /**
                //      * Sets weekday variables and loops through the withinOfficeHours function, which adjusts the date to be within office hours
                //      * @param {Date} date Date to be adjusted to be within office hours
                //      * @param {Number} number Number of adjustment hours to add to date
                //      * @returns Date
                //      */
                //     function oldOfficeHours(day, number) {

                //         while (loop == true) { //loops through the office hours function until the value returns within office hours
                //             var officeHours = withinOfficeHours(day, number);
                //             day = officeHours.date;
                //             number = officeHours.adjustmentNumber;
                //             loop = officeHours.loop;
                //         };
                //         //console.log("The correct date & time is: " + day);
                //         loop = true;
                //         // console.log(day);
                //         return day;

                //     };


                //     //#region OFFICE HOURS FUNCTIONS -----------------------------------------------------------------------------------------------------------


                //         //#region WITHIN OFFICE HOURS FUNCTION -------------------------------------------------------------------------------------------------

                //             /**
                //              * Adjusts date to be within office hours while maintaining an accurate turn around time variable for the adjustment number
                //              * @param {Date} date Date to be adjusted to be within office hours
                //              * @param {Number} adjustmentNumber Number of adjustment hours to add to date
                //              * @returns An object with properties (date, adjustment number, and loop)
                //              */
                //             function withinOfficeHours(date, adjustmentNumber) {

                //                 //#region VARIABLES ------------------------------------------------------------------------------------------------------------

                //                     //#region SETS DATE VARIABLES ----------------------------------------------------------------------------------------------

                //                         var dateSerial = Number(JSDateToExcelDate(date)); //converts date to excel serial for calculations
                //                         var adjusted = parseFloat(adjustmentNumber); //converts adjustment number from String to Number for calculations
                //                         var numberMinutes = adjusted * 60;
                //                         var adjustmentNumberSerial = minutesToSerial(numberMinutes);

                //                         //gets day of the week attributes for the date variable
                //                         var dateDayOfWeek = date.getDay(); //returns a dayID (0-6) for the day of the week of the date object
                //                         var dayTitle = titleDOW(dateDayOfWeek); //returns a day title based on the dayID of the dateDayOfWeek variable
                //                         var theWeekdayVar = officeHoursData[dayTitle];

                //                         var startOfWorkDay = setToStartOfDay(date, theWeekdayVar);

                //                         var endOfWorkDay = setToEndOfDay(date, theWeekdayVar);

                //                     //#endregion -------------------------------------------------------------------------------------------------------------

                //                     //#region ADJUSTS DATES IN CASE REQUEST WAS SUBMITTED OUTSIDE OF OFFICE HOURS ---------------------------------------

                //                         if (dateSerial < startOfWorkDay) { //if date is between 12AM and start time, adjust hours to be the start time
                //                             dateSerial = startOfWorkDay;
                //                             date = convertToDate(dateSerial);
                //                             //   dateSerial = Number(JSDateToExcelDate(date)); //converts date to excel serial for calculations

                //                             //   dateMilli = date.getTime();
                //                             //   bookendVars = startEndMidnight(date, theWeekdayVar);
                //                         };

                //                         if (dateSerial > endOfWorkDay) { //if date is after end time and before 12AM, go to next day and adjust hours to be the start time of that next day
                //                             date.setDate(date.getDate() + 1);
                //                             dateDayOfWeek = date.getDay();
                //                             dayTitle = titleDOW(dateDayOfWeek);
                //                             theWeekdayVar = officeHoursData[dayTitle];
                //                             startOfWorkDay = setToStartOfDay(date, theWeekdayVar);
                //                             endOfWorkDay = setToEndOfDay(date, theWeekdayVar);
                //                             dateSerial = startOfWorkDay;
                //                             date = convertToDate(dateSerial); //converts date to excel serial for calculations

                //                             //   dateMilli = date.getTime();
                //                             //   bookendVars = startEndMidnight(date, theWeekdayVar);
                //                         };

                //                     //#endregion ------------------------------------------------------------------------------------------------------------

                //                     //#region ADJUSTS DATES IN CASE REQUEST WAS SUBMITTED ON WEEKEND ----------------------------------------------------

                //                     if ((dateDayOfWeek == 6) || (dateDayOfWeek == 0)) { //if date was submitted on a weekend...
                //                         date = weekendAdjust(date, dateDayOfWeek);
                //                         dateDayOfWeek = date.getDay();
                //                         dayTitle = titleDOW(dateDayOfWeek);
                //                         theWeekdayVar = officeHoursData[dayTitle];
                //                         startOfWorkDay = setToStartOfDay(date, theWeekdayVar);
                //                         endOfWorkDay = setToEndOfDay(date, theWeekdayVar);
                //                         dateSerial = startOfWorkDay;
                //                         date = convertToDate(dateSerial); //converts date to excel serial for calculations

                //                         // dateMilli = date.getTime();
                //                         // bookendVars = startEndMidnight(date, theWeekdayVar);
                //                     };

                //                 //#endregion ------------------------------------------------------------------------------------------------------------

                //                     //#region SETS ADJUSTMENT DATE VARIABLES -----------------------------------------------------------------------------------

                //                         //adds adjustmentNumber to date to get an adjustedDate value that will be used in later checks and calculations
                //                         var adjustedDate = new Date(date);
                //                         var adjustedDateSerial = Number(JSDateToExcelDate(adjustedDate));
                //                         adjustedDateSerial = adjustedDateSerial + adjustmentNumberSerial;
                //                         adjustedDate = convertToDate(adjustedDateSerial);

                //                     //#endregion ---------------------------------------------------------------------------------------------------------------

                //                     //#region SETS ADD A DAY VARIABLES -----------------------------------------------------------------------------------------

                //                         //gets day of the week attributes for the day after the date variable
                //                         var nextDay = new Date(date);

                //                         var newNextDay = getNextDay(nextDay); //also sets this variable to the start time of the next day
                //                         var addADaySerial = newNextDay.nextDay;
                //                         var addADay = convertToDate(addADaySerial);
                //                         var addADayTitle = newNextDay.nextDayTitle;
                //                         var addADayWeekdayVar = officeHoursData[addADayTitle];
                //                         var addADayEnd = setToEndOfDay(addADay, addADayWeekdayVar);

                //                     //#endregion ----------------------------------------------------------------------------------------------------------------

                //                 //#endregion ----------------------------------------------------------------------------------------------------------------

                //                 //#region ACTION: SETS ADJUSTED DATE TO BE WITHIN OFFICE HOURS ------------------------------------------------------------------

                //                     //if adjustedDate falls outside of office hours, do this...
                //                     if (adjustedDateSerial < startOfWorkDay || adjustedDateSerial > endOfWorkDay) { //since the bookendVars is in reference to the date variable, this function will still trigger if adjustedDate is technically within office hours, but on a different day

                //                         //#region SETS ADJUSTMENT NUMBER VALUES ---------------------------------------------------------------------------------

                //                             var dayRemainder = (endOfWorkDay - dateSerial) // / 1000) / 60) / 60; //time between end of work day and the original date time
                //                             var remainingAdjust = adjustmentNumberSerial - dayRemainder; //gives us the remaining adjustment hours based off of what was already used to get to the end of the work day
                //                             // var remainingAdjustMilli = remainingAdjust * 3600000;

                //                         //#endregion ------------------------------------------------------------------------------------------------------------

                //                         //#region NEW DAY CALCULATIONS ------------------------------------------------------------------------------------------

                //                             var newDay = new Date(addADay);
                //                             var newDaySerial = Number(JSDateToExcelDate(newDay));

                //                             //adds remaining adjustment hours to the beginning of the work day the next day after date (addADay)
                //                             var dateTimeAdjusted = newDaySerial + remainingAdjust;

                //                             var dateTimeAdjustedConvert = convertToDate(dateTimeAdjusted); //convert serial number to date object

                //                             date = dateTimeAdjustedConvert; //not sure if it should be date or something else yet. Need to make sure that the function works with this

                //                         //#endregion ------------------------------------------------------------------------------------------------------------

                //                         //#region SET LOOP VARIABLES IF STILL NOT WITHIN OFFICE HOURS OR EXCEEDS OFFICE HOURS OF NEXT DAY -----------------------

                //                             //if the new date exceeds the office hours of addADay, then do this...
                //                             if (dateTimeAdjusted > addADayEnd) {
                //                             var addADayWorkDayLength = parseFloat(addADayWeekdayVar.workDay); //converts adjustment number from String to Number for calculations
                //                             var addADayLengthMinutes = addADayWorkDayLength * 60;
                //                             var addADayLengthSerial = minutesToSerial(addADayLengthMinutes);
                //                             adjustmentNumber = (remainingAdjust - addADayLengthSerial) //subtracts remainingAdjust hours from the total workDay hours in the addADay variable
                //                             var dayAfterTomorrow = new Date(addADay);
                //                             var newDayAfterTomorrow = getNextDay(dayAfterTomorrow);
                //                             dateSerial = newDayAfterTomorrow.nextDay;
                //                             date = convertToDate(dateSerial);
                //                             loop = true;
                //                             var newAdjustmentNumber = convertToDate(adjustmentNumber);

                //                             if (adjustmentNumber > 1) {
                //                                 var wheezy = newAdjustmentNumber.getDate();
                //                                 var zeDays = wheezy*24 //converts days into hours
                //                             } else {
                //                                 var zeDays = 0;
                //                             };
                //                             // var wheezy = newAdjustmentNumber.getTime();

                //                             // var sleezy = ((wheezy/1000)/60)/60;



                //                             //Months....probably don't even go this far
                //                             // var meezy = newAdjustmentNumber.getMonth();
                //                             // var monthArr = [];

                //                             // if (wheezy !== 0) {

                //                             //     var jan = 31*24; //hours in janurary 1900
                //                             //     monthArr.push(jan);
                //                             //     var feb = (28*24) + jan;
                //                             //     monthArr.push(feb);
                //                             //     var mar = (31*24) + feb;
                //                             //     monthArr.push(mar);
                //                             //     var apr = (30*24) + mar;
                //                             //     monthArr.push(apr);
                //                             //     var may = (31*24) + apr;
                //                             //     monthArr.push(may);
                //                             //     var jun = (30*24) + may;
                //                             //     monthArr.push(jun);
                //                             //     var jul = (31*24) + jun;
                //                             //     monthArr.push(jul);
                //                             //     var aug = (31*24) + jul;
                //                             //     monthArr.push(aug);
                //                             //     var sep = (30*24) + aug;
                //                             //     monthArr.push(sep);
                //                             //     var oct = (31*24) + sep;
                //                             //     monthArr.push(oct);
                //                             //     var nov = (30*24) + oct;
                //                             //     monthArr.push(nov);
                //                             //     var dec = (31*24) + nov;
                //                             //     monthArr.push(dec);

                //                             //     if (meezy !== 0) {
                //                             //         meezy - 1; //adjusts so that the days are calculated from the month prior (and all other prior months)
                //                             //         var monthHours = monthArr[meezy];
                //                             //     } else {
                //                             //         var monthHours = 0;
                //                             //     };

                //                             // } else {
                //                             //     var monthHours = 0;
                //                             // };

                //                             var cheesey = newAdjustmentNumber.getHours();
                //                             var squeezy = newAdjustmentNumber.getMinutes();
                //                             var truMinutes = squeezy/60;
                //                             adjustmentNumber = cheesey + truMinutes + zeDays;
                //                             return {
                //                                 date,
                //                                 adjustmentNumber,
                //                                 loop
                //                             };
                //                             } else {
                //                             loop = false;
                //                             return {
                //                                 date,
                //                                 adjustmentNumber,
                //                                 loop
                //                             };
                //                             };

                //                         //#endregion -------------------------------------------------------------------------------------------------------------

                //                     } else {
                //                         date = adjustedDate;
                //                         loop = false;
                //                         return {
                //                         date,
                //                         adjustmentNumber,
                //                         loop
                //                         };
                //                     };

                //                 //#endregion --------------------------------------------------------------------------------------------------------------------

                //             };

                //         //#endregion ---------------------------------------------------------------------------------------------------------------------------


                //         //#region TITLE DAY OF WEEK FUNCTION ---------------------------------------------------------------------------------------------------

                //             /**
                //              * Returns the weekday variable, with all it's associated properties, from the weekday index input value
                //              * @param {Number} d The indexed number (0-6) of the weekday
                //              * @returns An object with properties
                //              */
                //             function titleDOW(d) { //returns the day of the week (refered to directly in another variable) based on the dayID index number
                //                 if (d == 0) {
                //                 return "Sunday";
                //                 } else if (d == 1) {
                //                 return "Monday";
                //                 } else if (d == 2) {
                //                 return "Tuesday";
                //                 } else if (d == 3) {
                //                 return "Wednesday";
                //                 } else if (d == 4) {
                //                 return "Thursday";
                //                 } else if (d == 5) {
                //                 return "Friday";
                //                 } else if (d == 6) {
                //                 return "Saturday";
                //                 };
                //             };

                //         //#endregion ----------------------------------------------------------------------------------------------------------------------------------


                //         //#region START/END/MIDNIGHT FUNCTIONS --------------------------------------------------------------------------------------------------


                //             //I used to use Milliseconds to do my calculations, but since I am loading in date serial #'s from the excel sheet that could chnage at any time,
                //             //it makes more since it instead work within the Excel Serial Number and do all my calculations as serial instead of milliseconds that I then later convert to serial

                //             //I also decided to break these apart into separate functions so I can reference them one at a time later on in the code


                //             //#region SET TO START OF THE WORK DAY --------------------------------------------------------------------------------------------------

                //                 /**
                //                  * Set date to the start of the workday based on the weekday
                //                  * @param {Date} date The date variable
                //                  * @param {Object} theWeekdayVar The object associated with the specific weekday including all of its properties
                //                  * @returns Date
                //                  */
                //                 function setToStartOfDay(date, theWeekdayVar) {

                //                     var theDateBlank = new Date(date);
                //                     theDateBlank.setHours(0);
                //                     theDateBlank.setMinutes(0);
                //                     theDateBlank.setSeconds(0);
                //                     var theDateBlankSerial = Number(JSDateToExcelDate(theDateBlank));
                //                     //   var theDateBlankMilli = theDateBlank.getTime();

                //                     if (theWeekdayVar.startTime == "--") {
                //                         var startOfWorkDay = theDateBlankSerial //+ 8:30 as a serial number
                //                     } else {
                //                         var startOfWorkDay = theDateBlankSerial + theWeekdayVar.startTime;
                //                     };


                //                     var startWorkDayReadable = convertToDate(startOfWorkDay);

                //                     return startOfWorkDay;

                //                 };

                //             //#endregion ----------------------------------------------------------------------------------------------------------------------------


                //             //#region SET TO END OF THE WORK DAY ----------------------------------------------------------------------------------------------------

                //                 /**
                //                  * Set the date to the end of the workday based on the weekday
                //                  * @param {Date} date The date variable
                //                  * @param {Object} theWeekdayVar The object associated with the specific weekday including all of its properties
                //                  * @returns Date
                //                  */
                //                 function setToEndOfDay(date, theWeekdayVar) {

                //                     var theDateBlank = new Date(date);
                //                     theDateBlank.setHours(0);
                //                     theDateBlank.setMinutes(0);
                //                     theDateBlank.setSeconds(0);
                //                     var theDateBlankSerial = Number(JSDateToExcelDate(theDateBlank));
                //                     //   var theDateBlankMilli = theDateBlank.getTime();

                //                     var endOfWorkDay = theDateBlankSerial + theWeekdayVar.endTime;

                //                     var endWorkDayReadable = convertToDate(endOfWorkDay);

                //                     return endOfWorkDay;

                //                 };

                //             //#endregion ----------------------------------------------------------------------------------------------------------------------------


                //             //#region SET TO MIDNIGHT ---------------------------------------------------------------------------------------------------------------

                //                 /**
                //                  * Sets date to serial number of the next day at midnight (very beginning of the day)
                //                  * @param {Date} date The date variable
                //                  * @returns Number
                //                  */
                //                 function setToMidnight(date) {

                //                     var midnight = new Date(date);
                //                     midnight.setDate(midnight.getDate() + 1);
                //                     midnight.setHours(0);
                //                     midnight.setMinutes(0);
                //                     midnight.setSeconds(0);
                //                     var midnightSerial = Number(JSDateToExcelDate(midnight));

                //                     return midnightSerial;

                //                 };

                //             //#endregion -----------------------------------------------------------------------------------------------------------------------------


                //         //#endregion ----------------------------------------------------------------------------------------------------------------------------------


                //         //#region GET NEXT DAY FUNCTION --------------------------------------------------------------------------------------------------------

                //             /**
                //              * Adds a day to the date variable and sets it to the start time of that new day's day of the week. Also adjusts for weekends if needed.
                //              * @param {Date} date A date object
                //              * @returns An object with properties
                //              */
                //             function getNextDay(date) {

                //                 var nextDay = new Date(date);
                //                 var newNextDay = nextDay.setDate(nextDay.getDate() + 1); //returns the day after the original date
                //                 nextDay = new Date(newNextDay);
                //                 var nextDayDayOfWeek = nextDay.getDay();
                //                 var nextDayTitle = titleDOW(nextDayDayOfWeek); //returns a day title based on the dayID of the addADay variable
                //                 var theWeekdayVar = officeHoursData[nextDayTitle];

                //                 if ((nextDayDayOfWeek == 6) || (nextDayDayOfWeek == 0)) { //checks if nextDay falls on a weekend
                //                     nextDay = weekendAdjust(nextDay, nextDayDayOfWeek); //adjusts nextDay output to not fall on a weekend
                //                     nextDayDayOfWeek = nextDay.getDay();
                //                     nextDayTitle = titleDOW(nextDayDayOfWeek);
                //                     theWeekdayVar = officeHoursData[nextDayTitle];
                //                 };

                //                 nextDay = setToStartOfDay(nextDay, theWeekdayVar);

                //                 return {
                //                     nextDay,
                //                     nextDayTitle
                //                 };
                //             };

                //         //#endregion ----------------------------------------------------------------------------------------------------------------------------------


                //         //#region MINUTES TO SERIAL ------------------------------------------------------------------------------------------------------------

                //             /**
                //              * Converts from a time in minutes to an Excel serial number, starting from the beginning of time (otherwise known as Dec 31, 1899)
                //              * @param {Number} minutes A time in minutes
                //              * @returns Number
                //              */
                //             function minutesToSerial(minutes) {
                //             //   var date = new Date();
                //             //   date.setDate(0);
                //                 var date = 0;
                //                 date = convertToDate(date);
                //                 date.setMinutes(minutes);
                //                 var numberSerial = Number(JSDateToExcelDate(date));
                //                 return numberSerial;
                //             };

                //         //#endregion ----------------------------------------------------------------------------------------------------------------------------


                //         //#region WEEKEND ADJUST FUNCTION ------------------------------------------------------------------------------------------------------

                //             /**
                //              * If input date falls on a weekend, returns a new date adjusted to start on the next upcoming Monday
                //              * @param {Date} date A date variable
                //              * @param {Number} dateWeekday A number indexed 0-6 representing the weekday of the date variable
                //              * @returns Date
                //              */
                //             function weekendAdjust(date, dateWeekday) {
                //                 if (dateWeekday == 6) {
                //                     var weekend = new Date(date);
                //                     weekend.setDate(weekend.getDate() + 2);
                //                     return weekend;
                //                 } else if (dateWeekday == 0) {
                //                     var weekend = new Date(date);
                //                     weekend.setDate(weekend.getDate() + 1);
                //                     return weekend;
                //                 };
                //             };

                //         //#endregion ------------------------------------------------------------------------------------------------------------------------------

            //#endregion ---------------------------------------------------------------------------------------------


            //#region CONVERT DATE TO SERIAL ----------------------------------------------------------------------------------------------------------

                /**
                 * Converts input date into serial number that excel can apply conditional formatting to
                 * @param {Date} inDate A date variable
                 * @returns String
                 */
                function JSDateToExcelDate(inDate) {

                    var returnDateTime = 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
                    //var returnDateTime = 25569.0 + ((inDate.getTime()) / (1000 * 60 * 60 * 24));
                    return returnDateTime.toString().substr(0,20);

                };

            //#endregion --------------------------------------------------------------------------------------------------------------------------------


            //#region CONVERT SERIAL TO DATE --------------------------------------------------------------------------------------------------------

                /**
                 * Finds the value of Date Added in the changed row and converts it to be a date object in EST.
                 * @param {Number} serial The serial number to be converted
                 * @returns Date
                 */
                function convertToDate(serial) {

                    var date = new Date(Math.round((serial - 25569)*86400*1000)); //convert serial number to date object
                    date.setMinutes(date.getMinutes() + date.getTimezoneOffset()); //adjusting from GMT to EST (adds 4 hours)
                    return date;

                };

            //#endregion ---------------------------------------------------------------------------------------------------


            //#region PRIORITY GENERATION AND SORTATION ---------------------------------------------------------------------------------------

                /**
                 * Generates a priority number for each row in the table based on the values in the Picked Up / Started By column. Also sorts the data by priority
                 */
                async function priorityGenerationAndSortation() {

                    console.log("Priority and sorting function has fired!");

                    await Excel.run(async (context) => {

                        var sheet = context.workbook.worksheets.getActiveWorksheet();
                        var sheetTable = sheet.tables.getItemAt(0);
                        var priorityColumnData = sheetTable.columns.getItem("Priority").getDataBodyRange().load("values");
                        var bodyRange = sheetTable.getDataBodyRange().load("values");
                        var headerRange = sheetTable.getHeaderRowRange().load("values");
                        context.runtime.load("enableEvents");


                        await context.sync().then(function () {

                            // priorityColumnData.values.push([]);

                            var head = headerRange.values;

                            var pickedUpColumnIndex = findColumnIndex(head, "Picked Up / Started By"); //returns the index number of the "Picked Up / Started By" column based on it's position in the table header row

                            //need a function that will pull values from "pickedUpColumnIndex" position of the bodyRange.values for each row in sheet and put them in a new array

                            var activeTableValues = bodyRange.values; //loads all values of the active table

                            var pickedUpAllValuesArr = allColumnValues(activeTableValues, pickedUpColumnIndex); //makes an array of just the values from the Picked Up / Started By column
                            // pickedUpAllValuesArr.push(excelPickupOfficeHours);

                            //if the pickedUp array has duplicate values, this nested for statement will add 1 second to the times of each duplicate value to allow the priority number generation to work properly
                            for (var i = 0; i < pickedUpAllValuesArr.length; i++) {
                                for (var j = 0; j < pickedUpAllValuesArr.length; j++) {
                                    if (i !== j) { //makes sure that the values do not equal (so the first pass will fail, naturally)
                                        if (pickedUpAllValuesArr[i] == pickedUpAllValuesArr[j]) {
                                            console.log("A duplicate is present at index " + j + " of the array");
                                            pickedUpAllValuesArr[j] = pickedUpAllValuesArr[j] + 0.0000115740; //adds one second to the duplicate entry
                                        }
                                    }
                                }
                            }

                            var priorityNumbers = JSON.parse(JSON.stringify(pickedUpAllValuesArr)); //creates a duplicate of original array to be used for assigning the priority numbers, without having anything done to it affect oriignal array

                            var pickedUpAllValuesSorted = JSON.parse(JSON.stringify(pickedUpAllValuesArr)); //creates a duplicate of original array to be used to sort the original arrays values, without having anything done to it affect oriignal array
                            pickedUpAllValuesSorted.sort(); //sorts the array





                            for (var n = 0; n < pickedUpAllValuesSorted.length; n++) {
                                var index = pickedUpAllValuesArr.indexOf(pickedUpAllValuesSorted[n]); //finds the value at n in the sorted array, then finds that index of that value in the unsorted array
                                priorityNumbers[index] = [(n + 1)]; //in the new priority numbers array, inserts the n value (+1 to account for 0 index) at the index spot
                            };

                            priorityColumnData.values = priorityNumbers; //writes values to the priority column
                            // priorityColumnValues = priorityNumbers;

                            var priorityColumnIndex = findColumnIndex(head, "Priority"); //returns index number of the priority column

                            bodyRange.sort.apply([ //sorts entire table based on the priority column
                                {
                                    key: priorityColumnIndex,
                                    ascending: true
                                }
                            ])

                            console.log("Priority & Sorting function is now finished!");

                            // context.runtime.enableEvents = true;
                            // console.log("Events are turned on");

                        }).then(() => {
                            eventsOn();
                            console.log("Events: ON  →  turned on in the priorityGenerationAndSortation function, which I believe is an antiquated function that should not be used any longer. If you are seeing this, something has gone terribly wrong and you have somehow ended up in an alternate universe where this function was used. My condolances.");
                            return;
                        });

                    });
                };

            //#endregion ------------------------------------------------------------------------------------------------------------------


            //#region ALL COLUMN VALUES ---------------------------------------------------------------------------------------------------

                /**
                 * Returns an array of the all the values from a specific column in a table
                 * @param {Array} tableValues An array of arrays containing all the data from the table
                 * @param {Number} columnIndex Index number of the column we are trying to make an array of from its data
                 * @returns Array
                 */
                function allColumnValues(tableValues, columnIndex) {

                    var PUTimeArr = [];

                    for (var row of tableValues) { //for each row in the table
                        var PUTurnAroundTime = row[columnIndex]; //get the item where the row and columnIndex values meet
                        PUTimeArr.push(PUTurnAroundTime); //push this value to a new array
                    };

                    return PUTimeArr;

                };

            //#endregion ------------------------------------------------------------------------------------------------------------------


            //#region FIND COLUMN INDEX ---------------------------------------------------------------------------------------------------

            /**
             * Returns index of a column name based on it's position in the header row
             * @param {Array} header An array of arrays containing all the headers in the table
             * @param {String} columnName The name of the column that we are trying to find an index number for
             * @returns Number
             */
            function findColumnIndex(header, columnName) {
                var i = 0;
                var jelly;

                for (var column of header[0]) { //for each item in the header array
                    if (column == columnName) { //if the item matches the columnName input, return the value of i, otherwise increment i and continue through rest of array
                        jelly = i;
                        return jelly;
                    }
                    i++;
                };
            };

        //#endregion -------------------------------------------------------------------------------------------------------------------


        //#endregion ----------------------------------------------------------------------------------------------------------------------


    //#endregion -------------------------------------------------------------------------------------------------------------------------------


    //#region TURN OFF EVENTS BEFORE EXECUTING ON TABLE CHANGED -----------------------------------------------------------------------------

        /**
         * Turns events off, then executes the onTableChanged function
         */
        // async function onTableChangedEvents(eventArgs) {

        //     console.log("Running onTableChangedEvents!");

        //     await Excel.run(async (context) => {

        //         context.runtime.load("enableEvents");

        //         await context.sync();

        //         console.log("I awaited the context.sync().")
        //         context.runtime.enableEvents = false;
        //         console.log("Events: OFF - Occured in onTableChangedEvents");

        //         // var result = await onTableChanged(eventArgs).then(tableChangedPriorityAndSort(poop.rowInfo, poop.bodyRange, poop.priorityColumnData));
        //     });

        //     console.log("Excel.run() is done. Can we catch the error from the async onTableChanged()?? 🐭")

        //     await onTableChanged(eventArgs).catch(err => {
        //         console.log(err) // <--- does this log?
        //         showMessage(err, "show");
        //     })

        // };

    //#endregion ---------------------------------------------------------------------------------------------------------------------------


    //#region ON TABLE CHANGED --------------------------------------------------------------------------------------------------------------

        async function onTableChanged(eventArgs) {

            // console.log("Source of the onTableChanged event: " + eventArgs.source);

            // if (eventArgs.source == "Remote") {
            //     console.log("Content was changed by a remote user, exiting onTableChanged Event");
            //     return;
            // };

            // if (eventArgs.changeType == "RowInserted") {
            //     handleIllegalInsert(eventArgs);
            //     return;
            // }

            // async function handleIllegalInsert(eventArgs) {

            //     await Excel.run(async (context) => {

            //         var changedWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId).load("name");

            //         var changedTable = context.workbook.tables.getItem(eventArgs.tableId).load("name"); //Returns tableId of the table where the event occured

            //         var changedTableRows = changedTable.rows;

            //         var changedAddress = changedWorksheet.getRange(eventArgs.address);
            //         changedAddress.load("columnIndex");
            //         changedAddress.load("rowIndex");

            //         await context.sync();

            //         var changedRowIndex = changedAddress.rowIndex; //index of the row where the change was made (on a worksheet level)

            //         if (eventArgs.changeType == "RowDeleted") {
            //             var changedRowTableIndex = 0;
            //         } else {
            //             var changedRowTableIndex = changedRowIndex - 1; //adjusts index number for table level (-1 to skip header row)
            //         };

            //         var rowRange = changedTableRows.getItemAt(changedRowTableIndex).getRange();

            //         if (eventArgs.changeType == "RowInserted") {
            //             console.log("tsk tsk tsk...Don't forget the 7th commandment of the Art Queue Add-In:");
            //             console.log('"Thou shalt submit all requests to thy own sheet by means of the Add A Project taskpane. Manually adding rows of info to thyn sheet beith forbidden."');
            //             console.log("It's a simple mistake, but make sure not to do it again.");
            //             rowRange.delete("Up");
            //             // eventsOn();
            //             // console.log("Events: ON  →  triggered after a row was manually inserted into the sheet by the user, followed by the swift removal of said row and a slap on the wrist.");
            //             return;
            //         };

            //     });

            // };


            await Excel.run(async (context) => {

                console.log("Source of the onTableChanged event: " + eventArgs.source);

                if (eventArgs.source == "Remote") {
                    console.log("Content was changed by a remote user, exiting onTableChanged Event");
                    return;
                }


                //console.log("Running onTableChanged!");

                context.runtime.load("enableEvents");

                await context.sync();

                context.runtime.enableEvents = false;
                console.log("Events: OFF - Occured in onTableChanged!");

                if (eventArgs.changeType == "RowInserted") {
                    handleIllegalInsert(eventArgs);
                    showDennis();
                    return;
                }
    
                async function handleIllegalInsert(eventArgs) {
    
                    await Excel.run(async (context) => {
    
                        var changedWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId).load("name");
    
                        var changedTable = context.workbook.tables.getItem(eventArgs.tableId).load("name"); //Returns tableId of the table where the event occured
    
                        var changedTableRows = changedTable.rows;
    
                        var changedAddress = changedWorksheet.getRange(eventArgs.address);
                        changedAddress.load("columnIndex");
                        changedAddress.load("rowIndex");
    
                        await context.sync();
    
                        var changedRowIndex = changedAddress.rowIndex; //index of the row where the change was made (on a worksheet level)
    
                        // if (eventArgs.changeType == "RowDeleted") {
                        //     var changedRowTableIndex = 0;
                        // } else {
                        //     var changedRowTableIndex = changedRowIndex - 1; //adjusts index number for table level (-1 to skip header row)
                        // };
                      
                        var changedRowTableIndex = changedRowIndex - 1; //adjusts index number for table level (-1 to skip header row)
                    
                        var rowRange = changedTableRows.getItemAt(changedRowTableIndex).getRange();
    
                        console.log("tsk tsk tsk...Don't forget the 7th commandment of the Art Queue Add-In:");
                        console.log('"Thou shalt submit all requests to thy own sheet by means of the Add A Project taskpane. Manually adding rows of info to thyn sheet beith forbidden."');
                        console.log("It's a simple mistake, but make sure not to do it again.");

                        rowRange.delete("Up");

                        eventsOn();
                        console.log("Events: ON  →  triggered after a row was manually inserted into the sheet by the user, followed by the swift removal of said row and a slap on the wrist.");
                        
                        return;
    
                    });
    
                };

                //#region LOAD VARIABLES FROM WORKBOOK -----------------------------------------------------------------------------------------

                    var details = eventArgs.details;
                    var address = eventArgs.address;
                    var changeType = eventArgs.changeType;
                    //console.log(changeType);


                    var allWorksheets = context.workbook.worksheets;
                    allWorksheets.load("items/name/tables/id");
                    // console.log()
                    var changedWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId).load("name");
                    var worksheetTables = changedWorksheet.tables;

                    // var queueSheet = worksheetTables.getItemAt(0);
                    // var completedTable = worksheetTables.getItemAt(1);

                    // .load("items/name");
                    var valSheet = context.workbook.worksheets.getItem("Validation").load("name");

                    //Used to find the column and row index on a worksheet level
                    var changedAddress = changedWorksheet.getRange(address);
                    changedAddress.load("columnIndex");
                    changedAddress.load("rowIndex");

                    //var sheet = context.workbook.worksheets.getActiveWorksheet();
                    //var sheetTable = sheet.tables.getItemAt(0);
                    // var priorityColumnData = sheetTable.columns.getItem("Priority").getDataBodyRange().load("values");

                    //Used to load values of the changed row/table to be used in functions & to return updated values to the table
                    var allTables = context.workbook.tables;
                    allTables.load("items/name");
                    var tablesInWorksheet = changedWorksheet.tables.load("items/count");
                    changedTable = context.workbook.tables.getItem(eventArgs.tableId).load("name"); //Returns tableId of the table where the event occured
                    var changedTableColumns = changedTable.columns
                    changedTableColumns.load("items/name");
                    var changedTableRows = changedTable.rows;
                    changedTableRows.load("items");
                    var startOfTable = changedTable.getRange().load("columnIndex");

                    //var priorityColumnData = changedTable.columns.getItem("Priority").getDataBodyRange().load("values");


                    // LEGACY CODE THAT NEEDS TO BE UPDATED TO BE MORE FLEXIBLE ==========================================================
                    // ===================================================================================================================

                        //#region SPECIFIC TABLE VARIABLES --------------------------------------------------------------------------

                            var unassignedTable = context.workbook.tables.getItem("UnassignedProjects").load("worksheet");
                            var unassignedTableName = context.workbook.tables.getItem("UnassignedProjects").load("name");
                            var unassignedTableRows = unassignedTable.rows;
                            unassignedTableRows.load("items");
                            var unassignedRange = unassignedTable.getDataBodyRange().load("values");
                            var unassignedHeader = unassignedTable.getHeaderRowRange().load("values");


                            var peterTable = context.workbook.tables.getItem("PeterProjects").load("worksheet");
                            var peterTableName = context.workbook.tables.getItem("PeterProjects").load("name");
                            var peterTableRows = peterTable.rows;
                            peterTableRows.load("items");
                            var peterRange = peterTable.getDataBodyRange().load("values");
                            var peterHeader = peterTable.getHeaderRowRange().load("values");


                            var mattTable = context.workbook.tables.getItem("MattProjects").load("worksheet");
                            var mattTableName = context.workbook.tables.getItem("MattProjects").load("name");
                            var mattTableRows = mattTable.rows;
                            mattTableRows.load("items");
                            var mattRange = mattTable.getDataBodyRange().load("values");
                            var mattHeader = mattTable.getHeaderRowRange().load("values");


                            var alainaTable = context.workbook.tables.getItem("AlainaProjects").load("worksheet");
                            var alainaTableName = context.workbook.tables.getItem("AlainaProjects").load("name");
                            var alainaTableRows = alainaTable.rows;
                            alainaTableRows.load("items");
                            var alainaRange = alainaTable.getDataBodyRange().load("values");
                            var alainaHeader = alainaTable.getHeaderRowRange().load("values");


                            var joeTable = context.workbook.tables.getItem("JoeProjects").load("worksheet");
                            var joeTableName = context.workbook.tables.getItem("JoeProjects").load("name");
                            var joeTableRows = joeTable.rows;
                            joeTableRows.load("items");
                            var joeRange = joeTable.getDataBodyRange().load("values");
                            var joeHeader = joeTable.getHeaderRowRange().load("values");


                            var sarahTable = context.workbook.tables.getItem("SarahProjects").load("worksheet");
                            var sarahTableName = context.workbook.tables.getItem("SarahProjects").load("name");
                            var sarahTableRows = sarahTable.rows;
                            sarahTableRows.load("items");
                            var sarahRange = sarahTable.getDataBodyRange().load("values");
                            var sarahHeader = sarahTable.getHeaderRowRange().load("values");


                            var michaelTable = context.workbook.tables.getItem("MichaelProjects").load("worksheet");
                            var michaelTableName = context.workbook.tables.getItem("MichaelProjects").load("name");
                            var michaelTableRows = michaelTable.rows;
                            michaelTableRows.load("items");
                            var michaelRange = michaelTable.getDataBodyRange().load("values");
                            var michaelHeader = michaelTable.getHeaderRowRange().load("values");


                            var dannyTable = context.workbook.tables.getItem("DannyProjects").load("worksheet");
                            var dannyTableName = context.workbook.tables.getItem("DannyProjects").load("name");
                            var dannyTableRows = dannyTable.rows;
                            dannyTableRows.load("items");
                            var dannyRange = dannyTable.getDataBodyRange().load("values");
                            var dannyHeader = dannyTable.getHeaderRowRange().load("values");


                            var joshTable = context.workbook.tables.getItem("JoshProjects").load("worksheet");
                            var joshTableName = context.workbook.tables.getItem("JoshProjects").load("name");
                            var joshTableRows = joshTable.rows;
                            joshTableRows.load("items");
                            var joshRange = joshTable.getDataBodyRange().load("values");
                            var joshHeader = joshTable.getHeaderRowRange().load("values");


                            var lukeTable = context.workbook.tables.getItem("LukeProjects").load("worksheet");
                            var lukeTableName = context.workbook.tables.getItem("LukeProjects").load("name");
                            var lukeTableRows = lukeTable.rows;
                            lukeTableRows.load("items");
                            var lukeRange = lukeTable.getDataBodyRange().load("values");
                            var lukeHeader = lukeTable.getHeaderRowRange().load("values");


                            var breBTable = context.workbook.tables.getItem("BreBProjects").load("worksheet");
                            var breBTableName = context.workbook.tables.getItem("BreBProjects").load("name");
                            var breBTableRows = breBTable.rows;
                            breBTableRows.load("items");
                            var breBRange = breBTable.getDataBodyRange().load("values");
                            var breBHeader = breBTable.getHeaderRowRange().load("values");


                            var kristenTable = context.workbook.tables.getItem("KristenProjects").load("worksheet");
                            var kristenTableName = context.workbook.tables.getItem("KristenProjects").load("name");
                            var kristenTableRows = kristenTable.rows;
                            kristenTableRows.load("items");
                            var kristenRange = kristenTable.getDataBodyRange().load("values");
                            var kristenHeader = kristenTable.getHeaderRowRange().load("values");


                            var ethanTable = context.workbook.tables.getItem("EthanProjects").load("worksheet");
                            var ethanTableName = context.workbook.tables.getItem("EthanProjects").load("name");
                            var ethanTableRows = ethanTable.rows;
                            ethanTableRows.load("items");
                            var ethanRange = ethanTable.getDataBodyRange().load("values");
                            var ethanHeader = ethanTable.getHeaderRowRange().load("values");


                            var christianTable = context.workbook.tables.getItem("ChristianProjects").load("worksheet");
                            var christianTableName = context.workbook.tables.getItem("ChristianProjects").load("name");
                            var christianTableRows = christianTable.rows;
                            christianTableRows.load("items");
                            var christianRange = christianTable.getDataBodyRange().load("values");
                            var christianHeader = christianTable.getHeaderRowRange().load("values");


                            var jessicaTable = context.workbook.tables.getItem("JessicaProjects").load("worksheet");
                            var jessicaTableName = context.workbook.tables.getItem("JessicaProjects").load("name");
                            var jessicaTableRows = jessicaTable.rows;
                            jessicaTableRows.load("items");
                            var jessicaRange = jessicaTable.getDataBodyRange().load("values");
                            var jessicaHeader = jessicaTable.getHeaderRowRange().load("values");


                            var luisTable = context.workbook.tables.getItem("LuisProjects").load("worksheet");
                            var luisTableName = context.workbook.tables.getItem("LuisProjects").load("name");
                            var luisTableRows = luisTable.rows;
                            luisTableRows.load("items");
                            var luisRange = luisTable.getDataBodyRange().load("values");
                            var luisHeader = luisTable.getHeaderRowRange().load("values");


                            var emilyTable = context.workbook.tables.getItem("EmilyProjects").load("worksheet");
                            var emilyTableName = context.workbook.tables.getItem("EmilyProjects").load("name");
                            var emilyTableRows = emilyTable.rows;
                            emilyTableRows.load("items");
                            var emilyRange = emilyTable.getDataBodyRange().load("values");
                            var emilyHeader = emilyTable.getHeaderRowRange().load("values");


                            var lisaTable = context.workbook.tables.getItem("LisaProjects").load("worksheet");
                            var lisaTableName = context.workbook.tables.getItem("LisaProjects").load("name");
                            var lisaTableRows = lisaTable.rows;
                            lisaTableRows.load("items");
                            var lisaRange = lisaTable.getDataBodyRange().load("values");
                            var lisaHeader = lisaTable.getHeaderRowRange().load("values");


                            var ritaTable = context.workbook.tables.getItem("RitaProjects").load("worksheet");
                            var ritaTableName = context.workbook.tables.getItem("RitaProjects").load("name");
                            var ritaTableRows = ritaTable.rows;
                            ritaTableRows.load("items");
                            var ritaRange = ritaTable.getDataBodyRange().load("values");
                            var ritaHeader = ritaTable.getHeaderRowRange().load("values");


                            var robinTable = context.workbook.tables.getItem("RobinProjects").load("worksheet");
                            var robinTableName = context.workbook.tables.getItem("RobinProjects").load("name");
                            var robinTableRows = robinTable.rows;
                            robinTableRows.load("items");
                            var robinRange = robinTable.getDataBodyRange().load("values");
                            var robinHeader = robinTable.getHeaderRowRange().load("values");


                            var jordanTable = context.workbook.tables.getItem("JordanProjects").load("worksheet");
                            var jordanTableName = context.workbook.tables.getItem("JordanProjects").load("name");
                            var jordanTableRows = jordanTable.rows;
                            jordanTableRows.load("items");
                            var jordanRange = jordanTable.getDataBodyRange().load("values");
                            var jordanHeader = jordanTable.getHeaderRowRange().load("values");


                            var toddTable = context.workbook.tables.getItem("ToddProjects").load("worksheet");
                            var toddTableName = context.workbook.tables.getItem("ToddProjects").load("name");
                            var toddTableRows = toddTable.rows;
                            toddTableRows.load("items");
                            var toddRange = toddTable.getDataBodyRange().load("values");
                            var toddHeader = toddTable.getHeaderRowRange().load("values");

                        //#endregion ----------------------------------------------------------------------------------------------------

                    // ===================================================================================================================
                    // ===================================================================================================================



                    //all of the data from the changed table
                    var bodyRange = changedTable.getDataBodyRange().load("values");

                    //the header data of the changed table
                    var headerRange = changedTable.getHeaderRowRange().load("values");


                    //context.runtime.load("enableEvents"); //turn on and off events

                //#endregion --------------------------------------------------------------------------------------------------------------


                await context.sync();

                // var changedColumnIndexOG = changedAddress.columnIndex; //index of the column where the change was made (on a worksheet level)
                // var changedRowIndex = changedAddress.rowIndex; //index of the row where the change was made (on a worksheet level)

                // if (changeType == "RowDeleted") {
                //     var changedRowTableIndex = 0;
                // } else {
                //     var changedRowTableIndex = changedRowIndex - 1; //adjusts index number for table level (-1 to skip header row)
                // };

                // var myRow = changedTableRows.getItemAt(changedRowTableIndex); //loads the changed row in the changed table as an object
                // var rowRange = changedTableRows.getItemAt(changedRowTableIndex).getRange();

                // if (changeType == "RowInserted") {
                //     console.log("tsk tsk tsk...Don't forget the 7th commandment of the Art Queue Add-In:");
                //     console.log('"Thou shalt submit all requests to thy own sheet by means of the Add A Project taskpane. Manually adding rows of info to thyn sheet is forbidden."');
                //     console.log("It's a simple mistake, but make sure not to do it again.");
                //     rowRange.delete("Up");
                //     // eventsOn();
                //     // console.log("Events: ON  →  triggered after a row was manually inserted into the sheet by the user, followed by the swift removal of said row and a slap on the wrist.");
                //     return;
                // };

                // context.runtime.enableEvents = false;
                // console.log("Events: OFF - Occured in onTableChanged!");


                    //#region ASSIGNING VARIABLES -----------------------------------------------------------------------------------------------


                        //#region VALIDATION EXODUS AND TURNING OFF EVENTS -------------------------------------------------------------------------


                            //#region DON'T DO ANYTHING IF CHANGE WAS MADE TO VALIDATION SHEET -----------------------------------------------------

                                if (changedWorksheet.name == valSheet.name) { //if the change was made to the Validation sheet, exit the function
                                    console.log("Validation Sheet was changed, exiting the table changed event...")
                                    return;
                                };

                            //#endregion -----------------------------------------------------------------------------------------------------------


                            //#region TURN EVENTS OFF ----------------------------------------------------------------------------------------------

                                // context.runtime.enableEvents = false; //turns events off
                                // console.log("Events are turned off!!");

                            //#endregion -----------------------------------------------------------------------------------------------------------


                        //#endregion ---------------------------------------------------------------------------------------------------------------


                        //#region CREATING AND ASSIGNING WORKBOOK VARIABLES ------------------------------------------------------------------------


                            //#region CALL LOADED VARIABLES ---------------------------------------------------------------------------------------

                                if (changedWorksheet.name == "Unassigned Projects") {
                                    var completedTable = null;
                                } else {
                                    var completedTable = worksheetTables.getItemAt(1);
                                };

                                var changedColumnIndexOG = changedAddress.columnIndex; //index of the column where the change was made (on a worksheet level)
                                var changedRowIndex = changedAddress.rowIndex; //index of the row where the change was made (on a worksheet level)

                                var tableColumns = changedTableColumns.items; //loads all the changed table's columns
                                var tableRows = changedTableRows.items; //loads all the changed table's rows
                                if (changeType == "RowDeleted") {
                                    var changedRowTableIndex = 0;
                                } else {
                                    var changedRowTableIndex = changedRowIndex - 1; //adjusts index number for table level (-1 to skip header row)
                                };
                                var rowValues = tableRows[changedRowTableIndex].values; //loads the values of the changed row in the changed table
                                var myRow = changedTableRows.getItemAt(changedRowTableIndex); //loads the changed row in the changed table as an object
                                var rowRange = changedTableRows.getItemAt(changedRowTableIndex).getRange();
                                var justToCheck = rowIndexPostSort;
                                //var sortedRowInfo = new Object();
                                var tablesInWorksheetCount = tablesInWorksheet.count;


                                var tableContent = bodyRange.values; //all of the changed table's content
                                var head = headerRange.values; //all of the changed table's headers


                                var tableStart = startOfTable.columnIndex; //column index of the start of the table
                                var changedColumnIndex = changedColumnIndexOG - tableStart; //adjusts columnIndex to reflect the actual position in the table, no matter where the table is on the sheet


                            //#endregion ----------------------------------------------------------------------------------------------------------------


                            //#region RECREATES CHANGED TABLE IN CODE AND ASSIGNS COLUMN INDEX AND VALUE PROPERTIES OF THE CHANGED ROW TO AN OBJECT

                                var leTable = JSON.parse(JSON.stringify(tableContent)); //creates a duplicate array of the entire changed tables content to be used for making adjustments to the sheet, without having anything done to it affect oriignal array

                                var rowInfo = new Object(); //object that will contain the values and column indexs of every item in the changed row

                                for (var name of head[0]) { //for each header item in the head array...
                                    //creates keys with the header names of each column in the changed table and assigns them to the rowInfo object. For each key, the column index and cell values are added for the cell in that column in the changed row
                                    theGreatestFunctionEverWritten(head, name, rowValues, leTable, rowInfo, changedRowTableIndex);
                                };


                                var pickedUpColumnIndex = rowInfo.pickedUpStartedBy.columnIndex; //index of picked up column
                                var proofToClientColumnIndex = rowInfo.proofToClient.columnIndex; //index of proof to client cloumn
                                var addedColumnIndex = rowInfo.added.columnIndex;

                                var statusValue = rowInfo.status.value;

                                //console.log("I am a fart"); //hehe

                            //#endregion ------------------------------------------------------------------------------------------------------------

                            if (changeType == "RowInserted") {
                                console.log("tsk tsk tsk...Don't forget the 7th commandment of the Art Queue Add-In:");
                                console.log('"Thou shalt submit all requests to thy own sheet by means of the Add A Project taskpane. Manually adding rows of info to thyn sheet is forbidden."');
                                console.log("It's a simple mistake, but make sure not to do it again.");
                                rowRange.delete("Up");
                                // eventsOn();
                                // console.log("Events: ON  →  triggered after a row was manually inserted into the sheet by the user, followed by the swift removal of said row and a slap on the wrist.");
                                return;
                            };

                            //#region FINDS IF CHANGE WAS MADE TO THE UNASSIGNED PROJECTS TABLE OR NOT ----------------------------------------

                                var isUnassigned;

                                if (changedWorksheet.name == "Unassigned Projects") {
                                    isUnassigned = true;
                                } else {
                                    isUnassigned = false;
                                };

                            //#endregion ------------------------------------------------------------------------------------------------------


                            //#region FINDS IF CHANGED TABLE IS A COMPLETED TABLE OR NOT ------------------------------------------------------

                                //var listOfCompletedTables = [];

                                var changedTableName = changedTable.name;

                                var completedTableChanged = changedTableName.includes("Completed");

                                // allTables.items.forEach(function (table) { //for each table in the workbook...
                                //     if (table.name.includes("Completed")) { //if the table name includes the word "Completed" in it...
                                //         listOfCompletedTables.push(table.name); //push the name of that table into an array
                                //     };
                                // });
                                //
                                // //returns true if the changedTable is a completed table from the array previously made, false if it is anything else
                                // var completedTableChanged = listOfCompletedTables.includes(changedTable.name);

                            //#endregion ------------------------------------------------------------------------------------------------------

                        //#endregion ----------------------------------------------------------------------------------------------------------------


                    //#endregion ---------------------------------------------------------------------------------------------------------------

                    if (changeType == "RowDeleted") {

                        console.log("RowPeepee");

                        var leCheese = bodyRange.values;


                        if (changedTable.id == unassignedTable.id) {
                            leTable = leSorting(rowInfo, leTable, pickedUpColumnIndex, rowValues[0]);
                        };
                        if (changedTable.id !== unassignedTable.id && completedTableChanged == false) {
                            leTable = leSorting(rowInfo, leTable, proofToClientColumnIndex, rowValues[0]);
                        };

                        bodyRange.values = leTable;

                        await context.sync();

                        var newChangedTableRows = changedTable.rows;
                        newChangedTableRows.load("items");

                        await context.sync();

                        var tableRows = changedTableRows.items; //loads all the changed table's rows

                        for (var m = 0; m < leTable.length; m++) {

                            var rowRangeSorted = newChangedTableRows.getItemAt(m).getRange();

                            var rowValuesSorted = tableRows[m].values;

                            var rowInfoSorted = new Object();

                            for (var name of head[0]) {
                                theGreatestFunctionEverWritten(head, name, rowValuesSorted, leTable, rowInfoSorted, m);
                            };

                            conditionalFormatting(rowInfoSorted, tableStart, changedWorksheet, m, completedTableChanged, rowRangeSorted, null);

                        };

                        //conditionalFormatting(rowInfo, tableStart, changedWorksheet, changedRowTableIndex, completedTableChanged, rowRange, completedTable);
                        eventsOn();
                        console.log("Events: ON  →  turned on after a row was deleted within the onTableChanged function!");

                        return;

                    };


                    if ((changedColumnIndex == rowInfo.printDate.columnIndex) || (changedColumnIndex == rowInfo.group.columnIndex)) {

                        if (changedColumnIndex == rowInfo.printDate.columnIndex) {

                            var formattedDate = convertToDate(rowInfo.printDate.value);
                            var newerDate = new Date(formattedDate);
                            formattedDate = [('' + (newerDate.getMonth() + 1)).slice(-2), ('' + newerDate.getDate()).slice(-2), (newerDate.getFullYear() % 100)].join('/');

                            try {
                                var matchGroup = printDateRefData[formattedDate].group;
                            }
                            catch (e) {
                                if (matchGroup == undefined) {
                                    matchGroup = "N/A";
                                };
                            };

                            leTable[changedRowTableIndex][rowInfo.group.columnIndex] = matchGroup;
                            bodyRange.values = leTable;
                        }

                        if (changedColumnIndex == rowInfo.group.columnIndex) {
                            var groupUppercase = rowInfo.group.value.toUpperCase();
                            var matchPrintDate = groupRefData[groupUppercase].printDate;
                            if (matchPrintDate == undefined) {
                                matchPrintDate = "N/A";
                            };
                            leTable[changedRowTableIndex][rowInfo.printDate.columnIndex] = matchPrintDate;
                            leTable[changedRowTableIndex][rowInfo.group.columnIndex] = groupUppercase;
                            bodyRange.values = leTable;
                        };

                        conditionalFormatting(rowInfo, tableStart, changedWorksheet, changedRowTableIndex, completedTableChanged, rowRange, completedTable);
                        // eventsOn();
                        // console.log("Events: ON  →  turned on within the onTableChanged function after the print date or group columns were updated!");
                    };

                    var statusMove = false;

                    if (rowInfo.status.value == "Completed" || rowInfo.status.value == "Cancelled") {
                        statusMove = true;
                    };

                    //#region ADJUST TURN AROUND TIMES, SORTING, & PRIORITY NUMBERS ------------------------------------------------------------

                        //if any of these columns are changed, turn around times will be adjusted and the table will be sorted
                        if (changedColumnIndex == rowInfo.pickedUpStartedBy.columnIndex || changedColumnIndex == rowInfo.proofToClient.columnIndex || changedColumnIndex == rowInfo.priority.columnIndex || changedColumnIndex == rowInfo.product.columnIndex || changedColumnIndex == rowInfo.projectType.columnIndex || changedColumnIndex == rowInfo.added.columnIndex || changedColumnIndex == rowInfo.startOverride.columnIndex || changedColumnIndex == rowInfo.workOverride.columnIndex || (changedColumnIndex == rowInfo.status.columnIndex && completedTableChanged == false && statusMove == false)) {

                            console.log("I will update the turn around times, priority numbers, and sort the sheet before turning events back on!")

                            //adjusts picked up / started by turn around time
                            var lePickUpTime = getPickUpTime(rowInfo, leTable, changedRowTableIndex);

                            //adjusts proof to client turn around time
                            var leProofToClientTime = getProofToClientTime(rowInfo, leTable, lePickUpTime, changedRowTableIndex);

                            var changedRowValues = leTable[changedRowTableIndex];

                            if (changedTable.id == unassignedTable.id) {

                                //sorts based on pickedUp column values and assigns priority numbers
                                var sortAndPrioritize = leSorting(rowInfo, leTable, pickedUpColumnIndex, changedRowValues);

                            } else {

                                //sorts based on proof to client column values and assigns priority numbers
                                var sortAndPrioritize = leSorting(rowInfo, leTable, proofToClientColumnIndex, changedRowValues);

                            };

                            var check = rowIndexPostSort;

                            //writes updated values to the table
                            bodyRange.values = sortAndPrioritize; //overwrite changed table data with the new data from the sorted array

                        };

                    //#endregion ----------------------------------------------------------------------------------------------------------------------


                    //#region MOVE DATA BETWEEN TABLES -----------------------------------------------------------------------------------------

                        if (changedColumnIndex == rowInfo.artist.columnIndex || changedColumnIndex == rowInfo.status.columnIndex) {
                            console.log("Here is where all the complex move functions will take place!")



                            //LEGACY CODE THAT WORKS BUT NEEDS TO BE UPDATED TO BE MORE FLEXIBLE =================================================
                            // ===================================================================================================================

                                //#region ASSIGNS THE DESTINATION TABLE VALUE ---------------------------------------------------------

                                    if (rowInfo.artist.value == "Unassigned" && isUnassigned == false) {
                                        destinationTable = unassignedTable;
                                        destinationTableName = unassignedTableName.name;
                                        destinationRows = unassignedTableRows.items;
                                        destinationTableRange = unassignedRange;
                                        destinationHeader = unassignedHeader;
                                    } else if (rowInfo.artist.value == "Peter") {
                                        destinationTable = peterTable;
                                        destinationTableName = peterTableName.name;
                                        destinationRows = peterTableRows.items;
                                        destinationTableRange = peterRange;
                                        destinationHeader = peterHeader;
                                    } else if (rowInfo.artist.value == "Matt") {
                                        destinationTable = mattTable;
                                        destinationTableName = mattTableName.name;
                                        destinationRows = mattTableRows.items;
                                        destinationTableRange = mattRange;
                                        destinationHeader = mattHeader;
                                    } else if (rowInfo.artist.value == "Alaina") {
                                        destinationTable = alainaTable;
                                        destinationTableName = alainaTableName.name;
                                        destinationRows = alainaTableRows.items;
                                        destinationTableRange = alainaRange;
                                        destinationHeader = alainaHeader;
                                    } else if (rowInfo.artist.value == "Joe") {
                                        destinationTable = joeTable;
                                        destinationTableName = joeTableName.name;
                                        destinationRows = joeTableRows.items;
                                        destinationTableRange = joeRange;
                                        destinationHeader = joeHeader;
                                    } else if (rowInfo.artist.value == "Sarah") {
                                        destinationTable = sarahTable;
                                        destinationTableName = sarahTableName.name;
                                        destinationRows = sarahTableRows.items;
                                        destinationTableRange = sarahRange;
                                        destinationHeader = sarahHeader;
                                    } else if (rowInfo.artist.value == "Michael") {
                                        destinationTable = michaelTable;
                                        destinationTableName = michaelTableName.name;
                                        destinationRows = michaelTableRows.items;
                                        destinationTableRange = michaelRange;
                                        destinationHeader = michaelHeader;
                                    } else if (rowInfo.artist.value == "Danny") {
                                        destinationTable = dannyTable;
                                        destinationTableName = dannyTableName.name;
                                        destinationRows = dannyTableRows.items;
                                        destinationTableRange = dannyRange;
                                        destinationHeader = dannyHeader;
                                    } else if (rowInfo.artist.value == "Josh") {
                                        destinationTable = joshTable;
                                        destinationTableName = joshTableName.name;
                                        destinationRows = joshTableRows.items;
                                        destinationTableRange = joshRange;
                                        destinationHeader = joshHeader;
                                    } else if (rowInfo.artist.value == "Luke") {
                                        destinationTable = lukeTable;
                                        destinationTableName = lukeTableName.name;
                                        destinationRows = lukeTableRows.items;
                                        destinationTableRange = lukeRange;
                                        destinationHeader = lukeHeader;
                                    } else if (rowInfo.artist.value == "Bre B.") {
                                        destinationTable = breBTable;
                                        destinationTableName = breBTableName.name;
                                        destinationRows = breBTableRows.items;
                                        destinationTableRange = breBRange;
                                        destinationHeader = breBHeader;
                                    } else if (rowInfo.artist.value == "Kristen") {
                                        destinationTable = kristenTable;
                                        destinationTableName = kristenTableName.name;
                                        destinationRows = kristenTableRows.items;
                                        destinationTableRange = kristenRange;
                                        destinationHeader = kristenHeader;
                                    } else if (rowInfo.artist.value == "Ethan") {
                                        destinationTable = ethanTable;
                                        destinationTableName = ethanTableName.name;
                                        destinationRows = ethanTableRows.items;
                                        destinationTableRange = ethanRange;
                                        destinationHeader = ethanHeader;
                                    } else if (rowInfo.artist.value == "Christian") {
                                        destinationTable = christianTable;
                                        destinationTableName = christianTableName.name;
                                        destinationRows = christianTableRows.items;
                                        destinationTableRange = christianRange;
                                        destinationHeader = christianHeader;
                                    } else if (rowInfo.artist.value == "Jessica") {
                                        destinationTable = jessicaTable;
                                        destinationTableName = jessicaTableName.name;
                                        destinationRows = jessicaTableRows.items;
                                        destinationTableRange = jessicaRange;
                                        destinationHeader = jessicaHeader;
                                    } else if (rowInfo.artist.value == "Luis") {
                                        destinationTable = luisTable;
                                        destinationTableName = luisTableName.name;
                                        destinationRows = luisTableRows.items;
                                        destinationTableRange = luisRange;
                                        destinationHeader = luisHeader;
                                    } else if (rowInfo.artist.value == "Emily") {
                                        destinationTable = emilyTable;
                                        destinationTableName = emilyTableName.name;
                                        destinationRows = emilyTableRows.items;
                                        destinationTableRange = emilyRange;
                                        destinationHeader = emilyHeader;
                                    } else if (rowInfo.artist.value == "Lisa") {
                                        destinationTable = lisaTable;
                                        destinationTableName = lisaTableName.name;
                                        destinationRows = lisaTableRows.items;
                                        destinationTableRange = lisaRange;
                                        destinationHeader = lisaHeader;
                                    } else if (rowInfo.artist.value == "Rita") {
                                        destinationTable = ritaTable;
                                        destinationTableName = ritaTableName.name;
                                        destinationRows = ritaTableRows.items;
                                        destinationTableRange = ritaRange;
                                        destinationHeader = ritaHeader;
                                    } else if (rowInfo.artist.value == "Robin") {
                                        destinationTable = robinTable;
                                        destinationTableName = robinTableName.name;
                                        destinationRows = robinTableRows.items;
                                        destinationTableRange = robinRange;
                                        destinationHeader = robinHeader;
                                    } else if (rowInfo.artist.value == "Jordan") {
                                        destinationTable = jordanTable;
                                        destinationTableName = jordanTableName.name;
                                        destinationRows = jordanTableRows.items;
                                        destinationTableRange = jordanRange;
                                        destinationHeader = jordanHeader;
                                    } else if (rowInfo.artist.value == "Todd") {
                                        destinationTable = toddTable;
                                        destinationTableName = toddTableName.name;
                                        destinationRows = toddTableRows.items;
                                        destinationTableRange = toddRange;
                                        destinationHeader = toddHeader;
                                    } else {
                                        destinationTable = null;
                                        destinationTableName = null;
                                        destinationRows = null;
                                        destinationTableRange = null;
                                        destinationHeader = null;
                                    };

                                    //For the time being, I am recreating the variables from the changed table to work with the destination table.
                                    //I am replacing the changed row index with 0 since, at this point, there is no changed row in the destination table. We just need these values to essentially return the index number of the columns we want from the destination table in future functions.

                                    if (destinationTable == null || destinationTableName == null || destinationRows == null || destinationTableRange == null || destinationHeader == null) {

                                        // for (var m = 0; m < leTable.length; m++) {

                                        //     var rowRangeSorted = changedTableRows.getItemAt(m).getRange();

                                        //     var rowValuesSorted = tableRows[m].values;

                                        //     var rowInfoSorted = new Object();

                                        //     for (var name of head[0]) {
                                        //         theGreatestFunctionEverWritten(head, name, rowValuesSorted, leTable, rowInfoSorted, m);
                                        //     };

                                        //     conditionalFormatting(rowInfoSorted, tableStart, changedWorksheet, m, completedTableChanged, rowRangeSorted, null);

                                        // };

                                        // return;

                                        console.log("I actually don't need any of these destination table variables!");

                                    // } else if (destinationRows.length == 0 && destinationTable !== null) {

                                    //     var destinationRange = destinationTableRange.values;

                                    //     var destRowValues = destinationRange;

                                    //     var destTableName = destinationTableName;

                                    //     var destTable = JSON.parse(JSON.stringify(destinationRange));

                                    //     var destHead = destinationHeader.values;

                                    //     var destRowInfo = new Object();

                                    //     for (var name of destHead[0]) {
                                    //         theGreatestFunctionEverWritten(destHead, name, destRowValues, destTable, destRowInfo, 0)
                                    //     };

                                    } else {

                                        var destinationRange = destinationTableRange.values;

                                        if (destinationRows.length == 0) {
                                            var destRowValues = destinationRange;
                                        } else {
                                            var destRowValues = destinationRows[0].values;
                                        };

                                        var destTableName = destinationTableName;

                                        //var destRow = destTableRows.getItemAt(0);

                                        var destTable = JSON.parse(JSON.stringify(destinationRange));

                                        var destHead = destinationHeader.values;

                                        var destRowInfo = new Object();

                                        for (var name of destHead[0]) {
                                            theGreatestFunctionEverWritten(destHead, name, destRowValues, destTable, destRowInfo, 0)
                                        };

                                    };



                                //#endregion ----------------------------------------------------------------------------------------------

                            // ===================================================================================================================
                            // ===================================================================================================================




                            //TRYING TO FIGURE OUT CODE THAT CANCELS MOVE BETWEEN TABLES IF HEADERS DON'T MATCH \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            // \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                                //#region LOADS ALL WORKSHEETS INTO CODE -------------------------------------------------------------------------

                                    // var worksheetNames = [];

                                    // var worksheetsInfo = {};

                                    // var oopsAllSheets = allWorksheets.items;

                                    // var oopsAllTables = allTables.items;


                                    // for (var worksheet of oopsAllSheets) {

                                    //     var nameOfSheet = worksheet.name;

                                    //     var tablesOfSheet = worksheet.tables;

                                    //     if (nameOfSheet !== "Validation") {

                                    //         if (nameOfSheet == "Unassigned Projects") {
                                    //             nameOfSheet = "Unassigned";
                                    //         };

                                    //         worksheetNames.push(nameOfSheet);

                                    //         var thisWorksheet = context.workbook.worksheets.getItem(nameOfSheet);

                                    //         var thisWorksheetsTables = thisWorksheet.tables.load("items");

                                    //         var nameOfProjectTable = nameOfSheet + "Projects";

                                    //         var worksheetValues = context.workbook.tables.getItem(nameOfProjectTable).load("worksheet/id");

                                    //         //var headerValues = worksheetValues.getHeaderRowRange().load("values");

                                    //         worksheetsInfo[nameOfSheet] = {
                                    //             nameOfProjectTable,
                                    //             thisWorksheet
                                    //             //headerValues
                                    //         };

                                    //     };
                                    // }

                                    // // for (var table of oopsAllTables) {

                                    // //     var nameOfTable = table.name;
                                    // // }

                                    // console.log("SNAILS!");

                                //#endregion ------------------------------------------------------------------------------------------------------


                                //#region ASSIGNS THE DESTINATION TABLE VALUE ---------------------------------------------------------------------


                                    // for (var sheetName of worksheetNames) {

                                    //     if (rowInfo.artist.value == sheetName) {

                                    //         if (rowInfo.artist.value == "Unassigned" && isUnassigned == true) {
                                    //             destinationTable = "null";
                                    //             destinationHeader = "null";
                                    //             return;
                                    //         };

                                    //         destinationTable = worksheetsInfo[sheetName];
                                    //         destinationTable = destinationTable.thisWorksheet;
                                    //         var destinationId = destinationTable.id;
                                    //         //destinationHeader = destinationTable.headerValues;

                                    //     } else {
                                    //         destinationTable = "null";
                                    //         //destinationHeader = "null";
                                    //     };

                                    // };

                                //#endregion -----------------------------------------------------------------------------------------------------


                                //#region CHECK TABLE HEADERS TO SEE IF THEY ARE THE SAME BEFORE MOVING DATA --------------------------------------

                                    // if (destinationTable !== "null" || destinationHeader !== "null") {

                                    //     var headerValues = headerRange.values[0];
                                    //     var destHeaderValues = destinationHeader.values[0];

                                    //     var areHeadersEqual = areArraysEqual(headerValues, destHeaderValues);

                                    //     if (areHeadersEqual == false) {
                                    //       console.log("One of the targeted tables is missing a column, therefore data was not moved.");
                                    //       return;
                                    //     };

                                    //   };

                                //#endregion ---------------------------------------------------------------------------------------------

                            // \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            // \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


                            //#region CHECK TABLE HEADERS TO SEE IF THEY ARE THE SAME BEFORE MOVING DATA ----------------------------

                                if (destinationTable !== null || destinationHeader !== null) {

                                    var headerValues = headerRange.values[0];
                                    var destHeaderValues = destinationHeader.values[0];

                                    var areHeadersEqual = areArraysEqual(headerValues, destHeaderValues);

                                    if (areHeadersEqual == false) {
                                    console.log("One of the targeted tables is missing a column, therefore data was not moved.");
                                    return;
                                    };

                                };

                            //#endregion ---------------------------------------------------------------------------------------------


                            //#region MOVE DATA BASED ON STATUS COLUMN -------------------------------------------------------------------------

                                if (changedColumnIndex == rowInfo.status.columnIndex) {

                                    //#region FINDS THE COMPLETED TABLE IN CHANGED WORKSHEET ---------------------------------------------------

                                    // worksheetTables

                                        // worksheetTables.items.forEach(function (table) { //for each table in the changed worksheet...
                                        //
                                        //     if (table.name.includes("Completed")) { //if the table name includes the word "Completed" in it...
                                        //
                                        //
                                        //
                                        //       var zeTable = table.name; //sets var to name of said completed table
                                        //       completedTable = worksheetTables.getItem(zeTable); //grabs said table's data from the worksheet
                                        //     };
                                        // });

                                    //#endregion ------------------------------------------------------------------------------------------------

                                    //#region MOVES DATA TO COMPLETED TABLE ---------------------------------------------------------------------

                                        // if ((rowInfo.status.value == "Completed" || rowInfo.status.value == "Cancelled") && completedTableChanged == false && isUnassigned == false) { //if status column = "Completed" or "Cancelled", the changedTable is not a Completed table, & the changedWorksheet is not UnassignedProjects, move data to changedWorksheet's completed table
                                        //     completedTable.rows.add(0, rowValues); //Adds empty row to bottom of the completedTable, then inserts the changed values into this empty row
                                        //     myRow.delete(); //Deletes the changed row from the original sheet
                                        //     console.log("Data was moved to the artist's Completed Projects Table!");
                                        //     return;
                                        // } else if (rowInfo.status.value == "Editing" && completedTableChanged == true) { //if status column = "Editing" & the changedTable is a Completed table, move data back to the artist's table
                                        //     if (destinationTable !== "null") {
                                        //     moveData(destinationTable, rowValues, myRow, rowInfo.artist.value);
                                        //     };
                                        // };

                                        if ((rowInfo.status.value == "Completed" || rowInfo.status.value == "Cancelled") && completedTableChanged == false && isUnassigned == false) { //if status column = "Completed" or "Cancelled", the changedTable is not a Completed table, & the changedWorksheet is not UnassignedProjects, move data to changedWorksheet's completed table

                                            //#region UPDATE DATE OF LAST EDIT ------------------------------------------------------------------

                                                //generate a new date and time based on the current date and time
                                                var dateOfLastEditTime = new Date();
                                                var dateOfLastEditTimeJS = JSDateToExcelDate(dateOfLastEditTime);

                                                leTable[changedRowTableIndex][rowInfo.dateOfLastEdit.columnIndex] = dateOfLastEditTimeJS; //write current date and time to the Date of Last Edit position within the table array

                                            //#endregion ----------------------------------------------------------------------------------------


                                            completedTable.rows.add(0, rowValues); //Adds empty row to bottom of the completedTable, then inserts the changed values into this empty row
                                            //completedTable.rows.add(0, null); //Adds empty row to bottom of the completedTable, then inserts the changed values into this empty row

                                            myRow.delete(); //Deletes the changed row from the original sheet
                                            console.log("Data was moved to the artist's Completed Projects Table!");

                                            leTable.splice(changedRowTableIndex, 1); //removes changed row from table content array

                                            var leTableSort = leSorting(rowInfo, leTable, proofToClientColumnIndex, rowValues[0]) //sorts the artist table by proof to client
                                            var bodyRangeReload = changedTable.getDataBodyRange().load("values"); //reload artist tables values after deleting a row

                                            var newCompletedRows = completedTable.rows.load("items");

                                            var newComplatedBodyValues = completedTable.getDataBodyRange().load("values");

                                            //var startOfCompletedTable = completedTable.getRange().load("columnIndex");


                                            await context.sync();


                                            bodyRangeReload.values = leTableSort; //writes sorted table content from the array to the artist table

                                            var newCompletedTableValues = newComplatedBodyValues.values

                                            // var newCompletedTableRows = newCompletedRows.items;

                                            // var newCompletedTableStart = startOfCompletedTable.columnIndex; //column index of the start of the table


                                            for (var m = 0; m < newCompletedTableValues.length; m++) {

                                                var rowRangeSortedCompleted = newCompletedRows.getItemAt(m).getRange();

                                                rowRangeSortedCompleted.format.fill.clear();
                                                rowRangeSortedCompleted.format.font.color = "black";
                                                rowRangeSortedCompleted.format.font.bold = false;

                                                // var rowValuesSortedCompleted = newCompletedTableRows[m].values;

                                                // var rowInfoSortedCompleted = new Object();

                                                // for (var name of head[0]) {
                                                //     theGreatestFunctionEverWritten(head, name, rowValuesSortedCompleted, newCompletedTableValues, rowInfoSortedCompleted, m);
                                                // };

                                                // conditionalFormatting(rowInfoSortedCompleted, newCompletedTableStart, changedWorksheet, m, completedTableChanged, rowRangeSortedCompleted, completedTable);

                                            };



                                            //return;

                                        } else if ((rowInfo.status.value == "Light Changes" && completedTableChanged == true) || (rowInfo.status.value == "Moderate Changes" && completedTableChanged == true) || (rowInfo.status.value == "Heavy Changes" && completedTableChanged == true)) { //if status column = "Editing" & the changedTable is a Completed table, move data back to the artist's table
                                            if (destinationTable !== "null") {

                                                //moveData(destinationTable, rowValues, myRow, rowInfo.artist.value);
                                                myRow.delete(); //Deletes the changed row from the original sheet
                                                destinationTable.rows.add(null);

                                                //destTable.push(rowValues[0]);

                                                moveDataTwo(destTable, rowValues, leTable, changedRowTableIndex);

                                                rowIndexInDestTable = destTable.length - 1;

                                                //adjusts proof to client turn around time
                                                var destProofToClientTime = getProofToClientTime(rowInfo, destTable, 97, rowIndexInDestTable); //since this will only ever trigger the part of the function that references ligh, moderate, and heavy changes, the pick up time value is unneeded for the most part. Therefore, the random number 97 is inserted to take up its spot, and to make sure the first if statement passes every time.

                                                //var changedRowValues = leTable[changedRowTableIndex];

                                                //sorts based on pickedUp column values and assigns priority numbers
                                                //var sortAndPrioritize = leSorting(rowInfo, leTable, pickedUpColumnIndex, changedRowValues);


                                                var destTableSort = leSorting(destRowInfo, destTable, proofToClientColumnIndex, rowValues[0]);


                                                var unassignedRange = unassignedTable.getDataBodyRange().load("values");
                                                var peterRange = peterTable.getDataBodyRange().load("values");
                                                var mattRange = mattTable.getDataBodyRange().load("values");
                                                var alainaRange = alainaTable.getDataBodyRange().load("values");
                                                var joeRange = joeTable.getDataBodyRange().load("values");
                                                var sarahRange = sarahTable.getDataBodyRange().load("values");
                                                var michaelRange = michaelTable.getDataBodyRange().load("values");
                                                var dannyRange = dannyTable.getDataBodyRange().load("values");
                                                var joshRange = joshTable.getDataBodyRange().load("values");
                                                var lukeRange = lukeTable.getDataBodyRange().load("values");
                                                var breBRange = breBTable.getDataBodyRange().load("values");
                                                var kristenRange = kristenTable.getDataBodyRange().load("values");
                                                var ethanRange = ethanTable.getDataBodyRange().load("values");
                                                var christianRange = christianTable.getDataBodyRange().load("values");
                                                var jessicaRange = jessicaTable.getDataBodyRange().load("values");
                                                var luisRange = luisTable.getDataBodyRange().load("values");
                                                var emilyRange = emilyTable.getDataBodyRange().load("values");
                                                var lisaRange = lisaTable.getDataBodyRange().load("values");
                                                var ritaRange = ritaTable.getDataBodyRange().load("values");
                                                var robinRange = robinTable.getDataBodyRange().load("values");
                                                var jordanRange = jordanTable.getDataBodyRange().load("values");
                                                var toddRange = toddTable.getDataBodyRange().load("values");



                                                if (rowInfo.artist.value == "Unassigned" && isUnassigned == false) {
                                                    var destinationStation = unassignedRange;
                                                } else if (rowInfo.artist.value == "Peter") {
                                                    var destinationStation = peterRange;
                                                } else if (rowInfo.artist.value == "Matt") {
                                                    var destinationStation = mattRange;
                                                } else if (rowInfo.artist.value == "Alaina") {
                                                    var destinationStation = alainaRange;
                                                } else if (rowInfo.artist.value == "Joe") {
                                                    var destinationStation = joeRange;
                                                } else if (rowInfo.artist.value == "Sarah") {
                                                    var destinationStation = peterRange;
                                                } else if (rowInfo.artist.value == "Michael") {
                                                    var destinationStation = peterRange;
                                                } else if (rowInfo.artist.value == "Danny") {
                                                    var destinationStation = dannyRange;
                                                } else if (rowInfo.artist.value == "Josh") {
                                                    var destinationStation = joshRange;
                                                } else if (rowInfo.artist.value == "Luke") {
                                                    var destinationStation = lukeRange;
                                                } else if (rowInfo.artist.value == "Bre B.") {
                                                    var destinationStation = breBRange;
                                                } else if (rowInfo.artist.value == "Kristen") {
                                                    var destinationStation = kristenRange;
                                                } else if (rowInfo.artist.value == "Ethan") {
                                                    var destinationStation = ethanRange;
                                                } else if (rowInfo.artist.value == "Christian") {
                                                    var destinationStation = christianRange;
                                                } else if (rowInfo.artist.value == "Jessica") {
                                                    var destinationStation = jessicaRange;
                                                } else if (rowInfo.artist.value == "Luis") {
                                                    var destinationStation = luisRange;
                                                } else if (rowInfo.artist.value == "Emily") {
                                                    var destinationStation = emilyRange;
                                                } else if (rowInfo.artist.value == "Lisa") {
                                                    var destinationStation = lisaRange;
                                                } else if (rowInfo.artist.value == "Rita") {
                                                    var destinationStation = ritaRange;
                                                } else if (rowInfo.artist.value == "Robin") {
                                                    var destinationStation = robinRange;
                                                } else if (rowInfo.artist.value == "Jordan") {
                                                    var destinationStation = jordanRange;
                                                } else if (rowInfo.artist.value == "Todd") {
                                                    var destinationStation = toddRange;
                                                } else {
                                                    var destinationStation = "null";
                                                };

                                                // //#region UPDATE DATE OF LAST EDIT ------------------------------------------------------------------

                                                //     //generate a new date and time based on the current date and time
                                                //     var dateOfLastEditTime = new Date();
                                                //     var dateOfLastEditTimeJS = JSDateToExcelDate(dateOfLastEditTime);

                                                //     destTable[rowIndexPostSort][rowInfo.dateOfLastEdit.columnIndex] = dateOfLastEditTimeJS; //write current date and time to the Date of Last Edit position within the table array

                                                // //#endregion ----------------------------------------------------------------------------------------


                                                await context.sync();

                                                destinationStation.values = destTableSort;

                                            };
                                        };

                                    //#endregion ----------------------------------------------------------------------------------------------

                                };

                            //#endregion --------------------------------------------------------------------------------------------------


                            //#region MOVE DATA BASED ON ARTIST COLUMN --------------------------------------------------------------------

                                if (changedColumnIndex == rowInfo.artist.columnIndex) {

                                    //#region MOVES DATA TO DESTINATION TABLE -----------------------------------------------------------------

                                        //if (completedTableChanged == false) {
                                        if (destinationTable !== "null") {
                                            if (destinationTable.worksheet.id !== changedWorksheet.id) { //if destination table is not in the same worksheet as the changedTable (prevents for unnecessary moving of data across tables in the same worksheet), do the following...

                                                var newStatus = statusAutofill(destTableName);

                                                rowValues[0][rowInfo.status.columnIndex] = newStatus;

                                                //var changedRowValues = leTable[changedRowTableIndex];


                                                //moveData(destinationTable, rowValues, myRow, rowInfo.artist.value);
                                                moveDataTwo(destTable, rowValues, leTable, changedRowTableIndex);

                                                if (destinationRows.length == 0) {
                                                    destTable.shift();
                                                };



                                                if (changedTable.id == unassignedTable.id) { //if data is moving from the unassigned table to an artist table, sort this way...

                                                    var leTableSort = leSorting(rowInfo, leTable, pickedUpColumnIndex, rowValues[0]); //sorts the changed unassigned table by picked up / started by
                                                    var destTableSort = leSorting(destRowInfo, destTable, proofToClientColumnIndex, rowValues[0]); //sorts the destination artist table by proof to client

                                                } else if (destinationTable.id == unassignedTable.id) { //if data is moving from an artist table to the unassigned table, sort this way...

                                                    var leTableSort = leSorting(rowInfo, leTable, proofToClientColumnIndex, rowValues[0]); //sorts the changed artist table by proof to client
                                                    var destTableSort = leSorting(destRowInfo, destTable, pickedUpColumnIndex, rowValues[0]); //sorts the destination Unassigned table by picked up / started by

                                                } else if ((destinationTable.id !== unassignedTable.id) && (changedTable.id !== unassignedTable.id)) { //if data is moving between artist tables, both will be sorted by proof to client

                                                    var leTableSort = leSorting(rowInfo, leTable, proofToClientColumnIndex, rowValues[0]); //sorts the changed artist table by proof to client
                                                    var destTableSort = leSorting(destRowInfo, destTable, proofToClientColumnIndex, rowValues[0]); //sorts the destination arist table by proof to client

                                                };

                                                myRow.delete();
                                                //bodyRange.values = leTableSort;

                                                destinationTable.rows.add(null);
                                                //destinationTableRange.values = destTableSort;

                                                // var farts = destinationRows[0];

                                                // var sharts = destinationTable.rows[0];

                                                var bodyPositivity = changedTable.getDataBodyRange().load("values");

                                                var unassignedRange = unassignedTable.getDataBodyRange().load("values");
                                                var peterRange = peterTable.getDataBodyRange().load("values");
                                                var mattRange = mattTable.getDataBodyRange().load("values");
                                                var alainaRange = alainaTable.getDataBodyRange().load("values");
                                                var joeRange = joeTable.getDataBodyRange().load("values");
                                                var sarahRange = sarahTable.getDataBodyRange().load("values");
                                                var michaelRange = michaelTable.getDataBodyRange().load("values");
                                                var dannyRange = dannyTable.getDataBodyRange().load("values");
                                                var joshRange = joshTable.getDataBodyRange().load("values");
                                                var lukeRange = lukeTable.getDataBodyRange().load("values");
                                                var breBRange = breBTable.getDataBodyRange().load("values");
                                                var kristenRange = kristenTable.getDataBodyRange().load("values");
                                                var ethanRange = ethanTable.getDataBodyRange().load("values");
                                                var christianRange = christianTable.getDataBodyRange().load("values");
                                                var jessicaRange = jessicaTable.getDataBodyRange().load("values");
                                                var luisRange = luisTable.getDataBodyRange().load("values");
                                                var emilyRange = emilyTable.getDataBodyRange().load("values");
                                                var lisaRange = lisaTable.getDataBodyRange().load("values");
                                                var ritaRange = ritaTable.getDataBodyRange().load("values");
                                                var robinRange = robinTable.getDataBodyRange().load("values");
                                                var jordanRange = jordanTable.getDataBodyRange().load("values");
                                                var toddRange = toddTable.getDataBodyRange().load("values");



                                                if (rowInfo.artist.value == "Unassigned" && isUnassigned == false) {
                                                    var destinationStation = unassignedRange;
                                                } else if (rowInfo.artist.value == "Peter") {
                                                    var destinationStation = peterRange;
                                                } else if (rowInfo.artist.value == "Matt") {
                                                    var destinationStation = mattRange;
                                                } else if (rowInfo.artist.value == "Alaina") {
                                                    var destinationStation = alainaRange;
                                                } else if (rowInfo.artist.value == "Joe") {
                                                    var destinationStation = joeRange;
                                                } else if (rowInfo.artist.value == "Sarah") {
                                                    var destinationStation = sarahRange;
                                                } else if (rowInfo.artist.value == "Michael") {
                                                    var destinationStation = michaelRange;
                                                } else if (rowInfo.artist.value == "Danny") {
                                                    var destinationStation = dannyRange;
                                                } else if (rowInfo.artist.value == "Josh") {
                                                    var destinationStation = joshRange;
                                                } else if (rowInfo.artist.value == "Luke") {
                                                    var destinationStation = lukeRange;
                                                } else if (rowInfo.artist.value == "Bre B.") {
                                                    var destinationStation = breBRange;
                                                } else if (rowInfo.artist.value == "Kristen") {
                                                    var destinationStation = kristenRange;
                                                } else if (rowInfo.artist.value == "Ethan") {
                                                    var destinationStation = ethanRange;
                                                } else if (rowInfo.artist.value == "Christian") {
                                                    var destinationStation = christianRange;
                                                } else if (rowInfo.artist.value == "Jessica") {
                                                    var destinationStation = jessicaRange;
                                                } else if (rowInfo.artist.value == "Luis") {
                                                    var destinationStation = luisRange;
                                                } else if (rowInfo.artist.value == "Emily") {
                                                    var destinationStation = emilyRange;
                                                } else if (rowInfo.artist.value == "Lisa") {
                                                    var destinationStation = lisaRange;
                                                } else if (rowInfo.artist.value == "Rita") {
                                                    var destinationStation = ritaRange;
                                                } else if (rowInfo.artist.value == "Robin") {
                                                    var destinationStation = robinRange;
                                                } else if (rowInfo.artist.value == "Jordan") {
                                                    var destinationStation = jordanRange;
                                                } else if (rowInfo.artist.value == "Todd") {
                                                    var destinationStation = toddRange;
                                                } else {
                                                    var destinationStation = "null";
                                                };

                                                await context.sync()

                                                    var newBodyRange = bodyPositivity.values;
                                                    var newDestinationTableRange = destinationStation.values;

                                                    if (leTable.length == 0) {
                                                        newBodyRange.shift();
                                                    };

                                                    // bodyPositivity.values = leTableSort;
                                                    newBodyValues = leTableSort;

                                                    destinationStation.values = destTableSort;



                                                    // return {
                                                    //     leTableSort,
                                                    //     destTableSort
                                                    // };

                                                    //commitMoveData(bodyRange, leTableSort, destinationRange, destTableSort);

                                                    console.log("I didn't fail!");

                                                //});




                                                //setStatus(destinationTable, unassignedTable, tableColumns, changedRowIndex, tableStart, changedWorksheet);
                                            };
                                        } else {
                                            console.log("No artist was assigned or updated, so no data was moved.")
                                            return;
                                        };

                                        //};

                                    //#endregion --------------------------------------------------------------------------------------------

                                };

                            //#endregion ------------------------------------------------------------------------------------------------


                        };

                    //#endregion -----------------------------------------------------------------------------------------------------------------


                    if (changedColumnIndex !== rowInfo.printDate.columnIndex || changedColumnIndex !== rowInfo.group.columnIndex) {

                        if (
                            (changedColumnIndex == rowInfo.artist.columnIndex) || (changedColumnIndex == rowInfo.status.columnIndex &&
                                (
                                    ((rowInfo.status.value == "Completed" && completedTableChanged == false) || (rowInfo.status.value == "Cancelled" && completedTableChanged == false))
                                ||
                                    ((rowInfo.status.value == "Light Changes" && completedTableChanged == true) || (rowInfo.status.value == "Moderate Changes" && completedTableChanged == true) || (rowInfo.status.value == "Heavy Changes" && completedTableChanged == true))
                                )
                            )
                        ) {

                            var newChangedTableRows = destinationTable.rows.load("items");

                            var newBodyValues = destinationTable.getDataBodyRange().load("values");

                            var destinationWorksheetId = destinationTable.worksheet.id;

                            var newChangedWorksheet = context.workbook.worksheets.getItem(destinationWorksheetId).load("name");

                            var newStartOfTable = destinationTable.getRange().load("columnIndex");


                            // var otherTableRows = changedTable.rows.load("items");

                            // var otherBodyValues = changedTable.getDataBodyRange().load("values");

                            // var otherWorksheet = changedWorksheet;

                            // var otherStartOfTable = startOfTable;


                        } else {

                            var newChangedTableRows = changedTable.rows.load("items");

                            var newBodyValues = changedTable.getDataBodyRange().load("values");

                            var newChangedWorksheet = changedWorksheet;

                            var newStartOfTable = startOfTable;


                            // var otherTableRows = destinationTable.rows.load("items");

                            // var otherBodyValues = destinationTable.getDataBodyRange().load("values");

                            // var otherWorksheetId = destinationTable.worksheet.id;

                            // var otherWorksheet = context.workbook.worksheets.getItem(otherWorksheetId).load("name");

                            // var otherStartOfTable = destinationTable.getRange().load("columnIndex");


                        };


                        await context.sync();


                        var leTableSorted = newBodyValues.values

                        var tableRowsSorted = newChangedTableRows.items;

                        var newTableStart = newStartOfTable.columnIndex; //column index of the start of the table


                        for (var m = 0; m < leTableSorted.length; m++) {

                            var rowRangeSorted = newChangedTableRows.getItemAt(m).getRange();

                            var rowValuesSorted = tableRowsSorted[m].values;

                            var rowInfoSorted = new Object();

                            for (var name of head[0]) {
                                theGreatestFunctionEverWritten(head, name, rowValuesSorted, leTableSorted, rowInfoSorted, m);
                            };

                            conditionalFormatting(rowInfoSorted, newTableStart, newChangedWorksheet, m, completedTableChanged, rowRangeSorted, destTable);

                        };

                    };

                    // didTableChangeFire = true;
                    // console.log(didTableChangeFire);

                    // for (var y = 0; y < tablesInWorksheetCount; y++) {
                    //     var aTable = tablesInWorksheet.getItemAt(y);
                    //     selectionEvent = aTable.onSelectionChanged.add(onTableSelectionChangedEvents);
                    // };

                    eventsOn(); //turns events back on
                    console.log("Events: ON  →  turned on at the end of the onTableChanged Function!");

            }).catch (err => {
                console.log(err) // <--- does this log?
                showMessage(err, "show");
                context.runtime.enableEvents = true;
            });
            return;
            // eventsOn(); //turns events back on
            // console.log("Events: ON  →  turned on at the end of the onTableChanged Function!");

        };

    //#endregion ---------------------------------------------------------------------------------------------------------------------------------


//#endregion --------------------------------------------------------------------------------------------------------------------------------


//#region FUNCTIONS ---------------------------------------------------------------------------------------------------------------------------


    //#region CONDITIONAL FORMATTING -------------------------------------------------------------------------------------------

        /**
         * Applys all the row colors and other visual formatting when data is updated
         * @param {Object} rowInfoSorted An object containing the values and column index numbers for each cell of the changed row
         * @param {Number} newTableStart A number signifying the starting position of the table (0 means it's the first table in the sheet)
         * @param {Object} changedWorksheet The changed worksheet
         * @param {Number} rowIndexPostSort A number signifying the row index of the changed row AFTER sorting
         * @param {Boolean} completedTableChanged If the current table is a Completed table, this returns true
         * @param {Range} rowRangeSorted The range of the changed row AFTER sorting
         * @param {Array} destTable An array of arrays containing all the info in the destination table
         */
        function conditionalFormatting(rowInfoSorted, newTableStart, changedWorksheet, rowIndexPostSort, completedTableChanged, rowRangeSorted, destTable) {

            /**
             * GET TABLE RANGE
             * CLEAR THE RANGE FORMATTING
             * tableRange.format.fill.clear()
             */

            /**
             * loadData
             *
             * for each row
             *  if cell.date() < now
             *      cell.range().format.fill(red)
             */

            var now = new Date();
            var justNowDate = now.getDate();
            var toSerial = Number(JSDateToExcelDate(now));

            var worksheetRowIndex = rowIndexPostSort + 1; //adjusts index post table sort to work on worksheet level

            var pickedUpWorksheetColumn = rowInfoSorted.pickedUpStartedBy.columnIndex + newTableStart;
            var proofToClientWorksheetColumn = rowInfoSorted.proofToClient.columnIndex + newTableStart;
            var printDateWorksheetColumn = rowInfoSorted.printDate.columnIndex + newTableStart;
            var groupWorksheetColumn = rowInfoSorted.group.columnIndex + newTableStart;

            var pickedUpAddress = changedWorksheet.getCell(worksheetRowIndex, pickedUpWorksheetColumn);
            var proofToClientAddress = changedWorksheet.getCell(worksheetRowIndex, proofToClientWorksheetColumn);

            var printDate = Math.trunc(rowInfoSorted.printDate.value);
            var currentDateAbsolute = Math.trunc(toSerial);

            var printDateAddress = changedWorksheet.getCell(worksheetRowIndex, printDateWorksheetColumn);
            var groupAddress = changedWorksheet.getCell(worksheetRowIndex, groupWorksheetColumn);


            if (completedTableChanged == true && destTable == null) { //if completed table was changed, clear formatting and do not do any other formatting rules

                rowRangeSorted.format.fill.clear();
                rowRangeSorted.format.font.color = "black";
                rowRangeSorted.format.font.bold = false;

            } else {

                //#region ALL ENTRIES USE CONSISTENT FONT STYLING --------------------------------------------------------------------------------

                    rowRangeSorted.format.font.name = "Calibri";
                    rowRangeSorted.format.font.size = 12;
                    rowRangeSorted.format.font.color = "#000000";
                    rowRangeSorted.format.font.bold = false;

                //#endregion ---------------------------------------------------------------------------------------------------------------------


                //#region REMOVE INVALID HIGHLIGHTING IF NO LONGER INVALID -----------------------------------------------------------------------

                    if (rowInfoSorted.pickedUpStartedBy.value !== "NO PRODUCT / PROJECT TYPE" || rowInfoSorted.proofToClient.value !== "NO PRODUCT / PROJECT TYPE") {

                        rowRangeSorted.format.fill.clear();
                        pickedUpAddress.format.font.bold = false;
                        proofToClientAddress.format.font.bold = false;

                    };

                //#endregion ---------------------------------------------------------------------------------------------------------------------


                //#region GROUP & PRINT DATE FORMATTING ------------------------------------------------------------------------------------------

                    if (printDate == currentDateAbsolute) { //if current date = print date

                        rowRangeSorted.format.font.color = "#C00000";
                        rowRangeSorted.format.font.bold = true;
                        printDateAddress.format.horizontalAlignment = "center";
                        groupAddress.format.horizontalAlignment = "center";

                    } else if (((printDate - 1) == currentDateAbsolute)) { //if current date is the day before print date

                        rowRangeSorted.format.font.color = "#C00000";
                        rowRangeSorted.format.font.bold = true;
                        printDateAddress.format.horizontalAlignment = "center";
                        groupAddress.format.horizontalAlignment = "center";

                    } else if (((printDate - 6) <= currentDateAbsolute) && ((printDate - 2) >= currentDateAbsolute)) { //if current date is in the same group lock week as print date (between 7-2 days before)

                        rowRangeSorted.format.font.color = "#C00000";
                        rowRangeSorted.format.font.bold = true;
                        printDateAddress.format.horizontalAlignment = "center";
                        groupAddress.format.horizontalAlignment = "center";
                        
                    } else if (((printDate - 13) <= currentDateAbsolute) && ((printDate - 7) >= currentDateAbsolute)) { //if current date is in the week before group lock week (between 8-14 days before)

                        rowRangeSorted.format.font.color = "70AD47";
                        rowRangeSorted.format.font.bold = true;
                        printDateAddress.format.horizontalAlignment = "center";
                        groupAddress.format.horizontalAlignment = "center";
                        
                    } else if ((printDate < currentDateAbsolute) && (printDate !== 0)) { //if current date is after print date

                        rowRangeSorted.format.fill.color = "black";
                        rowRangeSorted.format.font.color = "white";
                        rowRangeSorted.format.font.bold = true;
                        printDateAddress.format.horizontalAlignment = "center";
                        groupAddress.format.horizontalAlignment = "center";
                        
                    } else { //set cell formatting to normal

                        rowRangeSorted.format.fill.clear();
                        rowRangeSorted.format.font.color = "black";
                        rowRangeSorted.format.font.bold = false;
                        printDateAddress.format.horizontalAlignment = "center";
                        groupAddress.format.horizontalAlignment = "center";
                        
                    };

                //#endregion ---------------------------------------------------------------------------------------------------------------------

                if (rowInfoSorted.group.value == "N/A") {

                    rowRangeSorted.format.fill.clear();
                    rowRangeSorted.format.font.color = "black";
                    rowRangeSorted.format.font.bold = false;
                    printDateAddress.format.horizontalAlignment = "center";
                    groupAddress.format.horizontalAlignment = "center";
                    
                };

                if (rowInfoSorted.status.value == "Working") {

                    rowRangeSorted.format.fill.color = "#FFE699";
                    rowRangeSorted.format.font.color = "#9C5700";
                    rowRangeSorted.format.font.bold = true;
                    printDateAddress.format.horizontalAlignment = "center";
                    groupAddress.format.horizontalAlignment = "center";
                    
                };

                if ((printDate < currentDateAbsolute) && (printDate !== 0)) { //if current date is after print date

                    rowRangeSorted.format.fill.color = "black";
                    rowRangeSorted.format.font.color = "white";
                    rowRangeSorted.format.font.bold = true;
                    printDateAddress.format.horizontalAlignment = "center";
                    groupAddress.format.horizontalAlignment = "center";
                    
                };

                //#region OVERDUE HIGHLIGHTING ---------------------------------------------------------------------------------------------------


                    //#region PICKED UP / STARTED BY OVERDUE -------------------------------------------------------------------------------------

                        if (toSerial > rowInfoSorted.pickedUpStartedBy.value && changedWorksheet.name == "Unassigned Projects") {
                            //pickedUpAddress.format.fill.color = "FFC000";
                            rowRangeSorted.format.fill.color = "FFC000";
                            rowRangeSorted.format.font.color = "black";
                        } //else {
                        //     pickedUpAddress.format.fill.clear();
                        // };

                    //#endregion -----------------------------------------------------------------------------------------------------------------


                    //#region PROOF TO CLIENT OVERDUE --------------------------------------------------------------------------------------------

                        if (toSerial > rowInfoSorted.proofToClient.value && changedWorksheet.name !== "Unassigned Projects") {
                            // proofToClientAddress.format.fill.color = "FF0000";
                            // proofToClientAddress.format.font.color = "white";
                            rowRangeSorted.format.fill.color = "FF0000";
                            rowRangeSorted.format.font.color = "white";
                        } //else {
                        //     proofToClientAddress.format.fill.clear();
                        //     proofToClientAddress.format.font.color = "black";
                        // };

                    //#endregion ----------------------------------------------------------------------------------------------------------------


                //#endregion ---------------------------------------------------------------------------------------------------------------------


                if (rowInfoSorted.status.value == "On Hold") {
                    rowRangeSorted.format.fill.color = "#BFBFBF";
                    rowRangeSorted.format.font.color = "#000000";
                    rowRangeSorted.format.font.bold = false;
                };

                if (rowInfoSorted.status.value == "In Review") {
                    rowRangeSorted.format.fill.clear()
                    rowRangeSorted.format.font.color = "#757171";
                    rowRangeSorted.format.font.bold = false;
                };

                if (rowInfoSorted.status.value == "At Client") {
                    rowRangeSorted.format.fill.clear()
                    rowRangeSorted.format.font.color = "#757171";
                    rowRangeSorted.format.font.bold = false;
                };

                if (rowInfoSorted.status.value == "Waiting On Info") {
                    rowRangeSorted.format.fill.clear()
                    rowRangeSorted.format.font.color = "#757171";
                    rowRangeSorted.format.font.bold = false;
                };


                //#region ADD INVALID HIGHLIGHTING IF INVALID ------------------------------------------------------------------------------------

                    if (rowInfoSorted.pickedUpStartedBy.value == "NO PRODUCT / PROJECT TYPE" || rowInfoSorted.proofToClient.value == "NO PRODUCT / PROJECT TYPE") {

                        rowRangeSorted.format.fill.color = "FFC5BB";
                        pickedUpAddress.format.font.bold = true;
                        proofToClientAddress.format.font.bold = true;
                        // pickedUpAddress.format.fill.color = "FFC5BB";
                        // proofToClientAddress.format.fill.color = "FFC5BB";

                    };

                //#endregion ---------------------------------------------------------------------------------------------------------------------



            };

        };


    //#endregion ---------------------------------------------------------------------------------------------------------------


    function showMessage(msg, showHide) {
        if (showHide === "hide") {
            $("#message-text").empty();
            $("#message").css("display", "none");
        } else if (showHide === "show") {
            $("#message-text").text(msg);
            $("#message").css("display", "flex");
        }
    }


   // function showFissh(showHide) {
    //     if (showHide === "hide") {
    //         $("#fissh").css("display", "none");
    //     } else if (showHide === "show") {
    //         $("#fissh").css("display", "flex");
    //     };
    // };

    function showElement(element, showHide) {
        if (showHide === "hide") {
            $(element).css("display", "none");
        } else if (showHide === "show") {
            $(element).css("display", "flex");
        };
    };

    function showFisshGif() {
        $("#fissh-gif").css("display", "flex");
        console.log(":O");   //  your code here
        setTimeout(hideFisshGif, 2000);
    };

    function hideFisshGif() {
        $("#fissh-gif").css("display", "none");
    };

    function showDennis() {
        $("#dennis").css("display", "flex");
        console.log("Na-Ah-Ah!");
        var naAhAh = new Audio("assets/dennis-mock.mp3");
        naAhAh.play();
        setTimeout(hideDennis, 2000);
    };

    function hideDennis() {
        $("#dennis").css("display", "none");
    };

    //#region CHECK EVENTS -----------------------------------------------------------------------------------------------------------------------

        /**
         * Writes whether events are turned on or off to the console
         */
        async function eventsFunction() {

            await Excel.run(async (context) => {
                context.runtime.load("enableEvents");
                await context.sync();
                if (context.runtime.enableEvents == true) {
                    var eventsEnabled = "on";
                } else {
                    var eventsEnabled = "off";
                };
                console.log("Events are turned " + eventsEnabled);
            });
        };

    //#endregion ----------------------------------------------------------------------------------------------------------------------------------


    //#region EVENTS ON -------------------------------------------------------------------------------------------------------------------------

        /**
         * Manually turns events on
         */
        async function eventsOn() {
            await Excel.run(async (context) => {
                context.runtime.load("enableEvents");
                await context.sync();
                context.runtime.enableEvents = true;
                //console.log("Events are turned on!");
            });
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------


    //#region TURN AROUND TIME FUNCTIONS -------------------------------------------------------------------------------------------------------


        //#region ADJUST PICKED UP / STARTED BY TURN AROUND TIME ---------------------------------------------------------------

            /**
             * Adjusts the picked up / started by turn around time value
             * @param {Object} rowInfo An object containing the values and column indexs of each cell in the changed row
             * @param {Array} leTable An array of arrays containing all the info of the changed table
             * @param {Number} rowIndex The index number of the changed row (table level)
             * @returns Date
             */
                function getPickUpTime(rowInfo, leTable, rowIndex) {

                    if (rowInfo.product.value == "" || rowInfo.projectType.value == "") {
                        leTable[rowIndex][rowInfo.pickedUpStartedBy.columnIndex] = "NO PRODUCT / PROJECT TYPE";
                        return null;
                    };

                    //get the Project Type coded variable from the Project Type ID Data based on the value in the Project Type column of the changed row
                    var theProjectTypeCode = projectTypeIDData[rowInfo.projectType.value].projectTypeCode;

                    //returns turn around time value from the Pickup Turn Around Time table based on the Product column of the changed row and the projetc type codeed variable
                    var pickedUpTurnAroundTime = pickupData[rowInfo.product.value][theProjectTypeCode];

                    //finds the start override value of the changed row and adds it to the previous turn around time variable
                    var pickedUpHours = pickedUpTurnAroundTime + rowInfo.startOverride.value;

                    //finds the added date/time serial of the changed row and converts it to a date
                    var addedDate = convertToDate(rowInfo.added.value);

                    //adds the adjusted turn around time to the added date and adjusts to be within office hours
                    var pickupOfficeHours = officeHours(addedDate, pickedUpHours);

                    //converts date to excel date
                    var excelPickupOfficeHours = Number(JSDateToExcelDate(pickupOfficeHours));

                    //updates the pickedup turn around time value in the table array based on our calculations
                    leTable[rowIndex][rowInfo.pickedUpStartedBy.columnIndex] = excelPickupOfficeHours;

                    return pickupOfficeHours;

                };

        //#endregion ------------------------------------------------------------------------------------------------------------


        //#region ADJUST PROOF TO CLIENT TURN AROUND TIME ----------------------------------------------------------------------

            /**
             * Adjusts the proof to client turn around time values
             * @param {Object} rowInfo An object containing the values and column indexs of each cell in the changed row
             * @param {Array} leTable An array of arrays containing all the info of the changed table
             * @param {Date} lePickUpTime The date returned from the previous getPickUpTime function
             * @param {Number} rowIndex The index number of the changed row (table level)
             * @returns Date
             */
            function getProofToClientTime(rowInfo, leTable, lePickUpTime, rowIndex) {

                if (lePickUpTime == null) {
                    leTable[rowIndex][rowInfo.proofToClient.columnIndex] = "NO PRODUCT / PROJECT TYPE";
                    return null;
                };

                if (rowInfo.status.value == "Light Changes" || rowInfo.status.value == "Moderate Changes" || rowInfo.status.value == "Heavy Changes") { //if the request's status is a form of "Editing"...

                    //get the Changes coded variable from the Changes ID Data based on the value in the Status column of the changed row
                    var theChangesCode = changesIDData[rowInfo.status.value].changesCode;

                    //gets the turn around time value from the Changes Data Table based on the product and status of the row
                    var proofToClient = changesData[rowInfo.product.value][theChangesCode];

                    //Heavy Changes has creative review time factored into the table variables, so no need to calculate it. Moderate and Light chnages will almost never need creative review

                    //finds the work override value of the changed row and adds it to the proofToClient variable
                    var artTurnAround = proofToClient + rowInfo.workOverride.value;

                    //need to create a new date variable, possibly in the Date of Last Edit column

                    //#region UPDATE DATE OF LAST EDIT ------------------------------------------------------------------

                        //generate a new date and time based on the current date and time
                        var dateOfLastEditTime = new Date();
                        var dateOfLastEditTimeJS = JSDateToExcelDate(dateOfLastEditTime);

                        leTable[rowIndex][rowInfo.dateOfLastEdit.columnIndex] = dateOfLastEditTimeJS; //write current date and time to the Date of Last Edit position within the table array

                    //#endregion ----------------------------------------------------------------------------------------

                    //adds the adjusted turn around time to the new Date of Last Edit date that was just generated and adjusts to be within office hours
                    var proofToClientOfficeHours = officeHours(dateOfLastEditTime, artTurnAround);

                    //converts date to excel date
                    var excelProofToClientOfficeHours = Number(JSDateToExcelDate(proofToClientOfficeHours));

                    //updates the proof to client turn around time value in the table array based on our calculations
                    leTable[rowIndex][rowInfo.proofToClient.columnIndex] = excelProofToClientOfficeHours;

                    return proofToClientOfficeHours;
                };

                //get the Project Type coded variable from the Project Type ID Data based on the value in the Project Type column of the changed row
                var theProjectTypeCode = projectTypeIDData[rowInfo.projectType.value].projectTypeCode;

                //returns turn around time value from the Proof to Client Turn Around Time table based on the Product column of the changed row and the projetc type codeed variable
                var proofToClient = proofToClientData[rowInfo.product.value][theProjectTypeCode];

                //returns creative review process hours adjustment number from thhe creative review table based on the Product column value of the changed row
                var creativeReview = creativeProofData[rowInfo.product.value].creativeReviewProcess;

                //adds the proof to client turn around time to the creative review time
                var proofWithReview = proofToClient + creativeReview;

                //finds the work override value of the changed row and adds it to the previous turn around time variable
                var artTurnAround = proofWithReview + rowInfo.workOverride.value;

                //adds the adjusted turn around time to the pickedUpStartedBy date and adjusts to be within office hours
                var proofToClientOfficeHours = officeHours(lePickUpTime, artTurnAround);

                //converts date to excel date
                var excelProofToClientOfficeHours = Number(JSDateToExcelDate(proofToClientOfficeHours));

                //updates the proof to client turn around time value in the table array based on our calculations
                leTable[rowIndex][rowInfo.proofToClient.columnIndex] = excelProofToClientOfficeHours;

                return proofToClientOfficeHours;

            };

        //#endregion -----------------------------------------------------------------------------------------------------------



        //#region SORTING THE TABLE BY PICKED UP TURN AROUND TIME ---------------------------------------------------------------

            /**
             * Sorts the table array by the values in leColumnIndex and then assigns updated priority numbers
             * @param {Object} rowInfo An object containing the values and column indexs of each cell in the changed row
             * @param {Array} leTable An array of arrays containing all the info of the changed table
             * @param {Number} leColumnIndex The index number of the column that will be used for sorting the table
             * @param {Array} changedRowValues The values of the changed row
             * @returns Array
             */
            function leSorting(rowInfo, leTable, leColumnIndex, changedRowValues) {

                //a copy of the array containing all the table data that will be used for sorting
                var leTableSorted = JSON.parse(JSON.stringify(leTable)); //creates a duplicate of original array to be used for assigning the priority numbers, without having anything done to it affect oriignal array

                var priorityColumnIndex = rowInfo.priority.columnIndex; //index of priority column

                var pickedUpColumnIndex = rowInfo.pickedUpStartedBy.columnIndex;
                var proofToClientColumnIndex = rowInfo.proofToClient.columnIndex;

                var statusColumnIndex = rowInfo.status.columnIndex;

                var tempTable = [];
                var onHoldTable = [];
                var awaitingChangesTable = [];


                for (var i = 0; i < leTableSorted.length; i++) { //for each row in the table...

                    if (leTableSorted[i][pickedUpColumnIndex] == "NO PRODUCT / PROJECT TYPE" || leTableSorted[i][proofToClientColumnIndex] == "NO PRODUCT / PROJECT TYPE") { //removes invlaid requests from table and puts them in a temp table to be added back in after sorting
                        tempTable.push(leTableSorted[i]);
                        leTableSorted.splice(i, 1);
                        i = i - 1;
                    } else if (leTableSorted[i][statusColumnIndex] == "On Hold") { //removes on hold requests from table and puts them in an on hold table to be added back in after sorting
                        onHoldTable.push(leTableSorted[i]);
                        leTableSorted.splice(i, 1);
                        i = i - 1;
                    } else if (leTableSorted[i][statusColumnIndex] == "At Client" || leTableSorted[i][statusColumnIndex] == "In Review" || leTableSorted[i][statusColumnIndex] == "Waiting On Info") { //removes awaiting changes requests from table and puts them in an awaiting changes table to be added back in after sorting
                        awaitingChangesTable.push(leTableSorted[i]);
                        leTableSorted.splice(i, 1);
                        i = i - 1;
                    };

                };

                //#region RESOLVE DUPLICATE DATE/TIMES ------------------------------------------------------------------------------

                    //if the pickedUp array has duplicate values, this nested for statement will add 1 second to the times of each duplicate value to allow the data sorting to work properly
                    // for (var i = 0; i < leTableSorted.length; i++) {
                    //     for (var j = 0; j < leTableSorted.length; j++) {
                    //         if (i !== j) { //makes sure that the values do not equal (so the first pass will fail, naturally)
                    //             if (leTableSorted[i][leColumnIndex] == leTableSorted[j][leColumnIndex]) {
                    //                 // if (leTableSorted[j][leColumnIndex] == "NO PRODUCT / PROJECT TYPE") {
                    //                 //     leTableSorted[j][leColumnIndex] = j + i;
                    //                 // };
                    //                 console.log("A duplicate is present at index " + j + " of the array");
                    //                 leTableSorted[j][leColumnIndex] = leTableSorted[j][leColumnIndex] + 0.0000115740; //adds one second to the duplicate entry
                    //             };
                    //         };
                    //     };
                    // };

                //#endregion -------------------------------------------------------------------------------------------------------



                //sorts the parent array (a) by the number in the sub array (b) at index of the picked up column
                //leTableSorted.sort(function(a,b){return a[leColumnIndex] > b[leColumnIndex]});
                leTableSorted.sort((a, b) => (a[leColumnIndex] > b[leColumnIndex]) ? 1 : -1); //sorts

                if (awaitingChangesTable.length > 0) { //adds awaiting changes requests back into table at the bottom
                    for (var i = 0; i < awaitingChangesTable.length; i++) {
                        leTableSorted.push(awaitingChangesTable[i]);
                        awaitingChangesTable.splice(i, 1);
                        i = i - 1;
                    };
                };

                if (onHoldTable.length > 0) { //adds on hold requests back into table at the bottom, under awaiting changes requests
                    for (var i = 0; i < onHoldTable.length; i++) {
                        leTableSorted.push(onHoldTable[i]);
                        onHoldTable.splice(i, 1);
                        i = i - 1;
                    };
                };

                if (tempTable.length > 0) { //adds invalid requests back into table at the bottom, under on hold requests
                    for (var i = 0; i < tempTable.length; i++) {
                        leTableSorted.push(tempTable[i]);
                        tempTable.splice(i, 1);
                        i = i - 1;
                    };
                };

                for (var j = 0; j < leTableSorted.length; j++) { //finds the post row sort index number
                    for (var k = 0; k < leTableSorted[j].length; k++) {
                        if (changedRowValues[k] !== leTableSorted[j][k]) {
                            break;
                        } else {
                            var l = leTableSorted[j].length - 1;
                            if (k == l) {
                                rowIndexPostSort = j;
                                //break;
                            };
                        };
                    };
                };



                // for (var name of head[0]) {
                //     theGreatestFunctionEverWritten(head, name, changedRowValues, leTableSorted, sortedRowInfo, rowIndexPostSort);
                // };






                //#region ASSIGN PRIORITY NUMBERS ---------------------------------------------------------------------------------------

                    //for each item in the sorted array of table values, assign updated priority numbers to the priority column index
                    for (var n = 0; n < leTableSorted.length; n++) {
                        leTableSorted[n][priorityColumnIndex] = n + 1;
                    };

                //#endregion ------------------------------------------------------------------------------------------------------------

                return leTableSorted;

            }

        //#endregion -------------------------------------------------------------------------------------------------------------


    //#endregion ----------------------------------------------------------------------------------------------------------------------------------


    // async function poopsicle(changeEvent) {

    //     await Excel.run(async (context) => {

    //         var smells = changedTable;
    //         var kells = changedTable.name;
    //     });


    // };





    //#region MOVE DATA FUNCTION ------------------------------------------------------------------------------------------------------------------

        /**
         * moves the changed row's data to the destionation table
         * @param {Object} destinationTable the table that the data is being moved to
         * @param {Array} myRow the data, values, and attributes of the changed row
         * @param {String} artistCellValue the value of the artist cell in the changed row
         */
        function moveData(destinationTable, rowValues, myRow, artistCellValue) {
            destinationTable.rows.add(null, rowValues); //Adds empty row to bottom of the destinationTable, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to " + artistCellValue + "'s Projects Table!");
        };

    //#endregion -----------------------------------------------------------------------------------------------------------------------------------




    function moveDataTwo(destTable, rowValues, leTable, changedRowTableIndex) {

        destTable.push(rowValues[0]);
        leTable.splice(changedRowTableIndex, 1);

    };

    // async function commitMovedData(bodyRange, leTableSort, destinationRange, destTableSort) {
    //     await Excel.run(async (context) => {

    //         bodyRange.values = leTableSort;
    //         destinationRange.values = destTableSort;

    //         await context.sync();

    //     });
    // };



    //#region CHECK IF TWO ARRAYS ARE EQUAL --------------------------------------------------------------------------------------------------------

        /**
         * Compares th contents of two arrays and returns a boolean of true if they are the same and false if they differ in anyway
         * @param {Array} array1 The first array being compared
         * @param {Array} array2 The second array being compared
         * @returns Boolean
         */
        function areArraysEqual(array1, array2) {
            if (array1.length == array2.length) {
                return array1.every((element, index) => {
                    if (element == array2[index]) {
                        return true;
                    };
                    return false;
                });
            };
            return false;
        };

    //#endregion ----------------------------------------------------------------------------------------------------------------------------------


    //#region INDEX & VALUES OF CHANGED ROW (THE GREATEST FUNCTION EVER WRITTEN) ------------------------------------------------------------------

        /**
         * Using the column names, finds and writes the column index and value of each cell in the changed row to an object. Also updates a copy of the header array with the values of the changed row in the correct column indexed positions
         * @param {Array} head An array of all the header values in the table
         * @param {String} columnName The name of the column to find the index for
         * @param {Array} rowValues An array of arrays containing all the row values for the changed row
         * @param {Array} leTable A copy array of the head array that will be used to write new values to the sheet for the row
         * @param {Object} obj An empty object that will be filled with column indexs and values for each cell in the changed row
         * @param {Number} changedRowTableIndex The row index of the changed table
         */
        function theGreatestFunctionEverWritten (head, columnName, rowValues, leTable, obj, changedRowTableIndex) {

            //returns the index number of the column name based on it's position in the table header row
            var columnIndex = findColumnIndex(head, columnName);

            //returns the values of a specific cell from a specific columnn in the changed row
            var value = rowValues[0][columnIndex];

            //writes value to appropriate columnIndex in leTable array
            leTable[changedRowTableIndex][columnIndex] = value;

            var headerColumn = headersToCode(columnName); //returns a properly coded variable based on the column name

            obj[headerColumn] = { //adds a new key to the object, including it's column index and value properties
                columnIndex,
                value
            };

        };

    //#endregion -----------------------------------------------------------------------------------------------------------------------------


    //#region HEADERS TO CODE ---------------------------------------------------------------------------------------------------------------------

        /**
         * Finds the coded version of the column names
         * @param {String} name The name of the column
         * @returns String
         */
        function headersToCode(name) {
            var codedHeader;
            if (name == "Priority") {
                codedHeader = "priority";
            } else if (name == "Design Manager") {
                codedHeader = "designManager";
            } else if (name == "Queue") {
                codedHeader = "queue";
            } else if (name == "Tier") {
                codedHeader = "tier";
            } else if (name == "Subject") {
                codedHeader = "subject";
            } else if (name == "Client") {
                codedHeader = "client";
            } else if (name == "Location") {
                codedHeader = "location";
            } else if (name == "Product") {
                codedHeader = "product";
            } else if (name == "Project Type") {
                codedHeader = "projectType";
            } else if (name == "CSM") {
                codedHeader = "csm";
            } else if (name == "Added") {
                codedHeader = "added";
            } else if (name == "Print Date") {
                codedHeader = "printDate";
            } else if (name == "Group") {
                codedHeader = "group";
            } else if (name == "Picked Up / Started By") {
                codedHeader = "pickedUpStartedBy";
            } else if (name == "Proof to Client") {
                codedHeader = "proofToClient";
            } else if (name == "Date of Last Edit") {
                codedHeader = "dateOfLastEdit";
            } else if (name == "Tags") {
                codedHeader = "tags";
            } else if (name == "Status") {
                codedHeader = "status";
            } else if (name == "Code") {
                codedHeader = "code";
            } else if (name == "Artist") {
                codedHeader = "artist";
            } else if (name == "Notes") {
                codedHeader = "notes";
            } else if (name == "Start Override") {
                codedHeader = "startOverride";
            } else if (name == "Work Override") {
                codedHeader = "workOverride";
            } else {
                codedHeader = name;
            };
            return codedHeader;
        }

    //#endregion --------------------------------------------------------------------------------------------------------------------------------


    //#region ERROR HANDLING ----------------------------------------------------------------------------------------------------------------------

        //#region TRY CATCH ---------------------------------------------------------------------------------------------
            async function tryCatch(callback) {
                //console.log("Error callback type is: ");
                //console.log(typeof callback);
                //if (typeof callback === 'function') {
                    try {
                        await callback();
                    } catch (error) {
                        console.error(error);
                        showMessage(error, "show");

                    }
                //}
        }
        //#endregion ---------------------------------------------------------------------------------------------------

    //#endregion -----------------------------------------------------------------------------------------------------


//#endregion ----------------------------------------------------------------------------------------------------------------------------------











//#region ANTIQUATED FUNCTIONS (NO LONGER IN USE / ALTERNATE VERSIONS OF WORKING FUNCTIONS) ---------------------------------------------------

    // $("#subject").on("keyup", function() {

    //     var paste = $(this).val(); // Get value from pasted input

    //     if (paste.length == 0) { // If what's pasted is empty

    //            $(".warning1").hide(); // Don't show the error
    //            $(this).removeClass("warning-box")
    //            $(this).removeClass("warning-box + .label")
    //            $("#client, #location, #product, #code").val(""); // Empty all inputs

    //     } else if (!paste.includes("~/*")) { // If what's pasted does not contain "~/*"

    //            $(".warning1").show().text(`This subject does not contain "~/*"`);

    //         //    var warningCSS = {
    //         //        "border": "2px",
    //         //        "border-color": "red"
    //         //    }
    //         //    $(this).css("border", "2px solid red");

    //             $(this).addClass("warning-box")
    //             $(this).addClass("warning-box + .label")


    //         //    $(this).css("pointer-events", "none");
    //            $("#client, #location, #product, #code").val(""); // Empty all inputs

    //     } else { // Probably a valid subject (contains ~/*)

    //            $(".warning1").hide() // Hide error
    //            $(this).removeClass("warning-box")
    //            $(this).removeClass("warning-box + .label")


    //             /** ------------------------------------------------------------
    //                Parse the subject, fill the other inputs
    //             ------------------------------------------------------------ */

    //             // Split at "-"s
    //             var splitPaste = paste.split("-");

    //             var blanks = splitPaste.includes("");

    //             if (blanks == true) {

    //                     var noBlanksArr = splitPaste.filter(function(x) {
    //                             return x !== "";
    //                     });

    //             } else {

    //                     var noBlanksArr = splitPaste;

    //             };


    //             if (noBlanksArr[0].includes(":")) {

    //                     var str = noBlanksArr[0];

    //                     str = str.substring(str.indexOf(":") + 1);

    //                     noBlanksArr.splice(0, 1, str);

    //             };

    //             var hasRequest = noBlanksArr[0].includes("CREATIVE REQUEST") || noBlanksArr[0].includes("Creative Request") || noBlanksArr[0].includes("ARTIST REQUEST") || noBlanksArr[0].includes("Artist Request");

    //             if (hasRequest == true) {

    //                     noBlanksArr.shift();

    //             };

    //             var plasticS = removeFirstAndLastSpace(noBlanksArr[noBlanksArr.length - 2]);

    //             if (plasticS == "S" || plasticS == "Flat") {

    //                     var plasticSIndex = noBlanksArr.indexOf(noBlanksArr[noBlanksArr.length - 2]);

    //                     noBlanksArr.splice(plasticSIndex, 1);

    //                     if (plasticS == "Flat") {

    //                             var productPostFlatIndex = noBlanksArr.indexOf(noBlanksArr[noBlanksArr.length - 2]);

    //                             noBlanksArr[productPostFlatIndex] = noBlanksArr[noBlanksArr.length - 2] + "Flat";

    //                     };

    //             };

    //             if (noBlanksArr.length > 3) { //if the subject line includes a location, do this...

    //                     // .NET stuff at end (~/*20104,51824,2*/~)
    //                     // Remove spaces (just in case), "~/*", "*/~", then split at ","
    //                     var splitCodes = noBlanksArr[noBlanksArr.length - 1].replace(' ','').replace('~/*','').replace('*/~','').split(",");

    //                     var theClient = noBlanksArr[0];
    //                     var updatedClient = removeFirstAndLastSpace(theClient);

    //                     var theLocation = noBlanksArr[1];
    //                     var updatedLocation = removeFirstAndLastSpace(theLocation);

    //                     var theProduct = noBlanksArr[noBlanksArr.length - 2];
    //                     var updatedProduct = removeFirstAndLastSpace(theProduct);
    //                     // updatedProduct = productID(updatedProduct, 1);
    //                     productID(updatedProduct, 1).then((updatedProduct) => {
    //                         console.log("A snail was exceuted");
    //                         return updatedProduct;
    //                     });

    //                     var theCode = splitCodes[0];
    //                     var updatedCode = removeFirstAndLastSpace(theCode);




    //                     // if we can't fill the data out, error:
    //                     try {
    //                             $("#client").val(updatedClient);
    //                             $("#location").val(updatedLocation);
    //                             $("#product").val(updatedProduct).removeClass("grey-sel");
    //                             $("#code").val(updatedCode);
    //                     } catch (e) {
    //                             // Something was wrong with the subject
    //                             $(".warning1").show().text(`Something's wrong with this subject. Error: ` + e);
    //                     }

    //             } else { //if subject line does not include a location, do this...

    //                     // .NET stuff at end (~/*20104,51824,2*/~)
    //                     // Remove spaces (just in case), "~/*", "*/~", then split at ","
    //                     var splitCodes = noBlanksArr[2].replace(' ','').replace('~/*','').replace('*/~','').split(",");

    //                     var theClient = noBlanksArr[0];
    //                     var updatedClient = removeFirstAndLastSpace(theClient);

    //                     var theProduct = noBlanksArr[1];
    //                     var updatedProduct = removeFirstAndLastSpace(theProduct);
    //                     updatedProduct = productID(updatedProduct, 1);

    //                     var theCode = splitCodes[0];
    //                     var updatedCode = removeFirstAndLastSpace(theCode);



    //                     // if we can't fill the data out, error:
    //                     try {
    //                             $("#client").val(updatedClient);
    //                             $("#product").val(updatedProduct).removeClass("grey-sel");
    //                             $("#code").val(updatedCode);
    //                     } catch (e) {
    //                             // Something was wrong with the subject
    //                             $(".warning1").show().text(`Something's wrong with this subject. Error: ` + e);
    //                     }

    //             };

    //     };

    // });








    // function removeFirstAndLastSpace(splitItem) {
    //     var firstChar = splitItem.charAt(0);

    //     if (firstChar == " ") {
    //             splitItem = splitItem.slice(1);
    //     };

    //     var lastChar = splitItem.charAt(splitItem.length - 1);

    //     if (lastChar == " ") {
    //             splitItem = splitItem.slice(0, splitItem.length - 1);
    //     };

    //     return splitItem;
    // };







    // async function productID(product, option) {

    //     var relativeProduct;

    //     await Excel.run(async (context) => {
    //         var sheet = context.workbook.worksheets.getItem("Validation");
    //         var productIDValTable = sheet.tables.getItem("ProductIDTable");

    //         // Get data from the table.
    //         var productIDBodyRange = productIDValTable.getDataBodyRange().load("values");

    //         await context.sync();

    //         //#region PRODUCT ID VALUES -----------------------------------------------------------------------------------

    //         var poop = productIDBodyRange.values;

    //         if (option == 1) {
    //             for (var row of poop) {
    //                 var check = row[0].trim();
    //                 // var netProduct = removeFirstAndLastSpace(row[0]);
    //                 if (row[0].trim() == product) {
    //                     relativeProduct = row[1];
    //                     break;
    //                 };
    //             };



    //             // if (option == 1) {

    //             //     var relativeProduct;

    //             //     for (var row of productIDBodyValues) {
    //             //         var a = row;
    //             //         var netProduct = removeFirstAndLastSpace(row[0]);
    //             //         if (netProduct == product) {
    //             //             relativeProduct = row[1];
    //             //             return relativeProduct;
    //             //         }
    //             //     }

    //                 // productIDBodyValues.forEach(function(row) {

    //                 //     var netProduct = removeFirstAndLastSpace(row[0]);

    //                 //     if (netProduct == product) {
    //                 //         relativeProduct = row[1];
    //                 //         return relativeProduct;
    //                 //     };

    //                 // });

    //                 // return relativeProduct;

    //             };


    //             if (option == 2) {

    //                 var code;

    //                 productIDBodyValues.forEach(function(row) {

    //                     if (row[1] == product) {
    //                         code = row[2];
    //                     };

    //                 });

    //                 return code;

    //             }





    //             //     // Add an option to the select box
    //             //     var option = `<option product-id="${row[0]}" relative-product="${row[1]}" product-code="${row[2]}">${row[1]}</option>`;

    //             //     var x = $(`#product > option[relative-product="${row[1]}"]`).length; //finds current relative-product in current option in the product dropdown and returns how many are currently in the dropdown

    //             //     if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
    //             //         if (row[1] !== "") { //if the relative-product in option is empty, do not add to list
    //             //             $("#product").append(option);
    //             //         };
    //             //     };
    //              //});

    //         //#endregion ---------------------------------------------------------------------------------------------------

    //     });

    //     return relativeProduct;

    // };



    //#region ORIGINAL ON TABLE CHANGED PRIORITY GENERATION AND SORTATION CODE ----------------------------------------------------------------

        //#region PRIORITY GENERATION AND SORTATION ---------------------------------------------------------------------------------------

            /**
             * Generates a priority number for each row in the table based on the values in the Picked Up / Started By column. Also sorts the data by priority
             */
             async function tableChangedPriorityAndSort(rowInfo, bodyRange, priorityColumnData) {

                console.log("Priority and sorting function has fired!");

                await Excel.run(async (context) => {

                    context.runtime.load("enableEvents");

                    var pickedUpColumnIndex = rowInfo.pickedUpStartedBy.columnIndex; //returns the index number of the "Picked Up / Started By" column based on it's position in the table header row
                    var activeTableValues = bodyRange.values; //loads all values of the active table
                    var pickedUpAllValuesArr = allColumnValues(activeTableValues, pickedUpColumnIndex); //makes an array of just the values from the Picked Up / Started By column

                    //if the pickedUp array has duplicate values, this nested for statement will add 1 second to the times of each duplicate value to allow the priority number generation to work properly
                    for (var i = 0; i < pickedUpAllValuesArr.length; i++) {
                        for (var j = 0; j < pickedUpAllValuesArr.length; j++) {
                            if (i !== j) { //makes sure that the values do not equal (so the first pass will fail, naturally)
                                if (pickedUpAllValuesArr[i] == pickedUpAllValuesArr[j]) {
                                    console.log("A duplicate is present at index " + j + " of the array");
                                    pickedUpAllValuesArr[j] = pickedUpAllValuesArr[j] + 0.0000115740; //adds one second to the duplicate entry
                                };
                            };
                        };
                    };

                    var priorityNumbers = JSON.parse(JSON.stringify(pickedUpAllValuesArr)); //creates a duplicate of original array to be used for assigning the priority numbers, without having anything done to it affect oriignal array
                    var pickedUpAllValuesSorted = JSON.parse(JSON.stringify(pickedUpAllValuesArr)); //creates a duplicate of original array to be used to sort the original arrays values, without having anything done to it affect oriignal array
                    pickedUpAllValuesSorted.sort(); //sorts the array

                    for (var n = 0; n < pickedUpAllValuesSorted.length; n++) {
                        var index = pickedUpAllValuesArr.indexOf(pickedUpAllValuesSorted[n]); //finds the value at n in the sorted array, then finds that index of that value in the unsorted array
                        priorityNumbers[index] = [(n + 1)]; //in the new priority numbers array, inserts the n value (+1 to account for 0 index) at the index spot
                    };

                    priorityColumnData.values = priorityNumbers; //writes values to the priority column
                    console.log("The priority numbers are " + priorityNumbers);


                    var priorityColumnIndex = rowInfo.priority.columnIndex; //returns index number of the priority column


                    sortFort(bodyRange, priorityColumnIndex);



                    await context.sync().then(function () {


                        // bodyRange.sort.apply([ //sorts entire table based on the priority column
                        //     {
                        //         key: priorityColumnIndex,
                        //         ascending: true
                        //     }
                        // ])

                        // console.log("Priority & Sorting function is now finished!");

                        eventsOn();
                        console.log("Events: ON  →  changed in yet another legacy function priority sort function that is still around for some reason but not being used. Can you tell I had a lot of trouble getting this feature to work? If you are seeing this, you might want to try stepping through a mirror because you are in an alternate dimension. Either that or some poor fool has enebaled a setting (probably me) that has destroyed the fabric of space and time. My wife is going to kill me; she takes her patterns and materials very seriously. Good luck explaining this one!");
                        return;

                        // priorityColumnData.values.push([]);

                        // var head = headerRange.values;


                        //need a function that will pull values from "pickedUpColumnIndex" position of the bodyRange.values for each row in sheet and put them in a new array


                        // pickedUpAllValuesArr.push(excelPickupOfficeHours);










                        // context.runtime.enableEvents = true;
                        // console.log("Events are turned on");

                    });

                });
            };

            function sortFort(bodyRange, priorityColumnIndex) {

                //console.log(bodyRange.values);

                bodyRange.sort.apply([ //sorts entire table based on the priority column
                    {
                        key: priorityColumnIndex,
                        ascending: true
                    }
                ])

                console.log("Priority & Sorting function is now finished!");

               //console.log(bodyRange.values);

                // eventsOn();
                return;

            };

        //#endregion ------------------------------------------------------------------------------------------------------------------

    //#endregion ------------------------------------------------------------------------------------------------------------------------------

//#endregion ----------------------------------------------------------------------------------------------------------------------------------


async function handleChange(event) {
    await Excel.run(async (context) => {
        await context.sync();
        console.log("Address of event: " + event.address);
        // console.log("The change direction state of the event: " + event.changeDirectionState);
        console.log("Change type of event: " + event.changeType);
        console.log("The details of the event: " + event.details);
        // console.log("Source of event: " + event.source);
        // console.log("The trigger source of the event: " + event.triggerSource);
        // console.log("The worrksheet ID of the event: " + event.worksheetId);
        console.log("END OF ENTRY////////////////////////////////////////////////////////////////////////////////////");
    }).catch(errorHandlerFunction);
}
