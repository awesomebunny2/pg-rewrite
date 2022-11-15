//validation sheet password: fissh
$( async () => {

    console.log("DOCUMENT IS LOADED BABY 🚀🌕🔥")

    // Pushing some label text up if they are on chrome...
    // var isChrome = !!window.chrome && (!!window.chrome.webstore || !!window.chrome.runtime);
    var isSafari = /^((?!chrome|android).)*safari/i.test(navigator.userAgent);

    // Add a style for safari
    if (isSafari === true) {

        console.log("THIS IS SAFARI ADDING CLASS!")
        $("label").addClass("safari-style");
    }

    console.log("ARE WE USING SAFARI!!??!?!??!?!?!?!", isSafari);

    //  Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    let behavior = await Office.addin.getStartupBehavior()
    console.log("BEHAVIOR", behavior)
    if (behavior === "Load") {
        $("#auto-open").prop("checked", true)
    } else {
        $("#auto-open").prop("checked", false) 
    }

     // Office.addin.getStartupBehavior().then((curBehavior) => {
     //     if (curBehavior !== "None") {
     //         $("#auto-open").prop("checked", true);
     //         //console.log("Checkbox is checked!");
     //     } else {
     //         $("#auto-open").prop("checked", false);
     //         //console.log("Checkbox is not checked...");
     //     };
     // });
 
    //  $('#auto-open').change(function() {
    //      if (this.checked == true) {
    //          console.log("Turning auto-open ON!")
    //          Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    //          console.log("Auto-open is ON!")
    //      } else {
    //          console.log("Turning auto-open OFF!")
    //          Office.addin.setStartupBehavior(Office.StartupBehavior.none);
    //          console.log("Auto-open is OFF!")
    //      };
    //  });
 

    // let isAuto = Office.context.document.settings.get("Office.AutoShowTaskpaneWithDocument");
    // let isAuto

    
     // let isAuto = Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
 
     //console.log("The document setting for auto open taskpane is: " + isAuto);
 
 
    //  Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", false);
 
     $('#auto-open').change(function() {
         if (this.checked == true) {
             console.log("Turning auto-open ON!")
             Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
             Office.context.document.settings.saveAsync();
             console.log("Auto-open is ON!")
         } else {
             console.log("Turning auto-open OFF!")
             Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", false);
             Office.context.document.settings.saveAsync();
             console.log("Auto-open is OFF!")
         };
     });

});


// Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
// Office.context.document.settings.saveAsync();


//#region GLOBAL -------------------------------------------------------------------------------------------------------------------------------------

    //#region TEST SUBJECTS --------------------------------------------------------------------------------------------------------------------------
        //CREATIVE REQUEST -Alfredo's Pizza - West Babylon - MENU - ~/*1338,52130,1*/~
        //CREATIVE REQUEST -Bella Napoli - Canfield - Env #10 8.5x11 S2 - ~/*1837,65845,1*/~
        //Re: Artist Request - Brickhouse Pizzeria - Richfield Springs - MENU - ~/*30601,72301,1*/~
    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region I MIGHT NEED THIS BEGINNING MATERIAL SOME DAY ------------------------------------------------------------------------------------------

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

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region GLOBAL VARIABLES -----------------------------------------------------------------------------------------------------------------------

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
            rowIndex: "",
            columnIndex: ""
        };
        var didTableChangeFire = false;
        var deactivationEvent;
        var deactivatedWorksheetId;
        var activatedWorksheet;
        var valPassword = "fissh";
        var dennisHere = false;

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------------------------------------------


//#region ON READY -----------------------------------------------------------------------------------------------------------------------------------

    Office.onReady((info) => {
        
        console.log("⚠️ INFO HOST AND PLATFORM TYPE:")
        console.log(info)
        // console.log(Office.PlatformType)
        
        if (info.platform === "Mac") {
            // THIS IS ON THE DESKTOP
            console.log("You are probably using the Desktop version of Excel on a Mac")
            $("label").addClass("safari-style");

        } else if (info.platform === "OfficeOnline") {
            // THIS IS THE ONLINE
            console.log("You're currently using the online version of Excel!")
        };

        if (info.host === Office.HostType.Excel) {


            Excel.run(async (context) => {





                activationEvent = registerOnActivateHandler();

                //#region LOADING VALUES -------------------------------------------------------------------------------------------------------------

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
                    var activeProjectTable = activeSheet.tables.getItemAt(0);
                    var workbookName = context.workbook.load("name");

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

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                await context.sync()

                    //#region GRABBING DATA FROM VALIDATION AND WRITING TO CODE ----------------------------------------------------------------------

                        //#region PRODUCT ID DATA ----------------------------------------------------------------------------------------------------

                            var productIDArr = productIDBodyRange.values;

                            for (var row of productIDArr) {
                                productIDData[row[0].trim()] = {
                                    "productID":row[0].trim(),
                                    "relativeProduct":row[1].trim(),
                                    "productCode":row[2].trim()
                                };
                            };

                            // console.log(productIDData);

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region PROJECT TYPE ID DATA -----------------------------------------------------------------------------------------------

                            var projectTypeIDArr = projectTypeIDBodyRange.values;

                            for (var row of projectTypeIDArr) {
                                projectTypeIDData[row[0].trim()] = {
                                    "projectType":row[0].trim(),
                                    "projectTypeCode":row[1].trim(),
                                };
                            };

                            // console.log(projectTypeIDData);

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region DESIGN MANAGERS DATA -----------------------------------------------------------------------------------------------

                            var designManagersArr = designManagersBodyRange.values;

                            for (var row of designManagersArr) {
                                designManagersData[row[0].trim()] = {
                                    "designManager":row[0],
                                    "worksheetTabColor":row[1]
                                };
                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region PICKED UP TURN AROUND TIME DATA ------------------------------------------------------------------------------------

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

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region PROOF TO CLIENT TIME DATA ------------------------------------------------------------------------------------------

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

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region TIER LEVEL DATA ----------------------------------------------------------------------------------------------------

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

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region CREATIVE PROOF DATA ------------------------------------------------------------------------------------------------

                            var creativeProofArr = creativeProofBodyRange.values;

                            for (var row of creativeProofArr) {
                                creativeProofData[row[0].trim()] = {
                                    "creativeReviewProcess":row[1]
                                };
                            };

                            // console.log(creativeProofData);

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region OFFICE HOURS DATA --------------------------------------------------------------------------------------------------

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

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region CHANGES DATA -------------------------------------------------------------------------------------------------------

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

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region CHANGES ID DATA ----------------------------------------------------------------------------------------------------

                            var changesIDArr = changesIDBodyRange.values;

                            for (var row of changesIDArr) {
                                changesIDData[row[0].trim()] = {
                                    "changes":row[0].trim(),
                                    "changesCode":row[1].trim(),
                                };
                            };

                            // console.log(changesIDData);

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region PRINT DATE DATA ----------------------------------------------------------------------------------------------------

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

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region GROUP DATA ---------------------------------------------------------------------------------------------------------

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

                                //if group letter is not already in the data, create the object and properties for the row
                                if (isGroupAlreadyPresent == false) {

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

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                });

            // console.log(info);
            tryCatch(updateDropDowns);

            eventsOn();
            console.log("Events: ON  →  turned on in onReady function!");
            //updateDropDowns();
        };
    });

//#endregion -----------------------------------------------------------------------------------------------------------------------------------------


//#region TASKPANE -----------------------------------------------------------------------------------------------------------------------------------

    //#region STYLIZING TASKPANE ELEMENTS ------------------------------------------------------------------------------------------------------------

        //#region STYLIZE SPECIFIC CHARACTERS --------------------------------------------------------------------------------------------------------

            //this stylizes all * characeters in the container element to use different CSS than other character elements
            $("#container").each(function () {
                $(this).html($(this).html().replace(/(\*)/g, 
                '<span style="color: rgba(220, 20, 60, 0.50); font-size: 9pt; padding-left: 1px; padding-bottom: 1px;">$1</span>'));
            });

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region WHEN THESE TASKPANE ITEMS ARE NO LONGER FOCUSED ON, DO SOMETHING -------------------------------------------------------------------

            $("#client").on("focusout", function() {
                removeWarningClass("#client", "#warning2");
            });

            $("#product").on("focusout", function() {
                removeWarningClass("#product", "#warning3");
            });

            $("#project-type").on("focusout", function() {
                removeWarningClass("#project-type", "#warning4");
            });

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region ADD WARNING CLASS ------------------------------------------------------------------------------------------------------------------

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

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region REMOVE WARNING CLASS ---------------------------------------------------------------------------------------------------------------

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

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region UPDATE DROPDOWNS -----------------------------------------------------------------------------------------------------------------------

        /**
         * Populates taskpane dropdowns with items from cooresponding validation sheet tables
         */
        async function updateDropDowns() {

            await Excel.run(async (context) => {

                //#region LOAD VALUES ----------------------------------------------------------------------------------------------------------------

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

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                await context.sync();

                //#region PRODUCT ID VALUES ----------------------------------------------------------------------------------------------------------

                    var productIDBodyValues = productIDBodyRange.values;

                    $("#product").empty();
                    $("#product").append($("<option disabled selected hidden></option>").val("").text(""));

                    productIDBodyValues.forEach(function(row) {

                        // Add an option to the select box
                        var option = `<option product-id="${row[0]}" relative-product="${row[1]}" product-code="${row[2]}">${row[1]}</option>`;

                        //finds current relative-product in current option in the product dropdown and returns how many are currently in the dropdown
                        var x = $(`#product > option[relative-product="${row[1]}"]`).length;

                        if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
                            if (row[1] !== "") { //if the relative-product in option is empty, do not add to list
                                $("#product").append(option);
                            };
                        };
                    });

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region PROJECT TYPE VALUES --------------------------------------------------------------------------------------------------------

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

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region PRINT DATE & GROUP VALUES --------------------------------------------------------------------------------------------------

                    var groupDateRefValues = groupDateRefRange.values;

                    $("#print-date").empty();
                    $("#print-date").append($("<option disabled selected hidden></option>").val("").text(""));
                    $("#group").empty();
                    $("#group").append($("<option disabled selected hidden></option>").val("").text(""));

                    groupDateRefValues.forEach(function(row) {

                        // Add an option to the select box
                        var option = `<option based-on-now="${row[0]}" year-based-on-now="${row[1]}" 
                        week-based-on-now="${row[2]}" print-date="${row[3]}" weekday="${row[4]}" adjust="${row[5]}" 
                        group="${row[6]}">${row[6]}</option>`;

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

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region ARTIST LEAD VALUES ---------------------------------------------------------------------------------------------------------

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

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region QUEUE VALUES ---------------------------------------------------------------------------------------------------------------

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

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region TIER VALUES ----------------------------------------------------------------------------------------------------------------

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

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region TAGS VALUES ----------------------------------------------------------------------------------------------------------------

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

                //#endregion -------------------------------------------------------------------------------------------------------------------------

            });
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region AUTO POPULATE TASKPANE FIELDS ----------------------------------------------------------------------------------------------------------

        //#region AUTO POPULATE TASKPANE BASED ON SUBJECT --------------------------------------------------------------------------------------------

            $("#subject").keyup(() => tryCatch(subjectPasted));

            //#region SUBJECT PASTED FUNCTION --------------------------------------------------------------------------------------------------------

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
                            $(this).addClass("warning-box")
                            $(this).addClass("warning-box + .label")
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

                        var hasRequest = noBlanksArr[0].includes("CREATIVE REQUEST") || noBlanksArr[0].includes("Creative Request") || 
                        noBlanksArr[0].includes("ARTIST REQUEST") || noBlanksArr[0].includes("Artist Request")  || 
                        noBlanksArr[0].includes("Urgent") || noBlanksArr[0].includes("Urgent!") || noBlanksArr[0].includes("Urgent!!") || 
                        noBlanksArr[0].includes("URGENT") || noBlanksArr[0].includes("URGENT!") || noBlanksArr[0].includes("URGENT!!") ||

                        noBlanksArr[0].includes("Urgent Art Request") || noBlanksArr[0].includes("Urgent Art Request!") || 
                        noBlanksArr[0].includes("Urgent! Art Request") || noBlanksArr[0].includes("Urgent! Art Request!") ||
                        noBlanksArr[0].includes("Urgent Art Request!!") || noBlanksArr[0].includes("Urgent!! Art Request") ||
                        noBlanksArr[0].includes("Urgent! Art Request!!") || noBlanksArr[0].includes("Urgent!! Art Request!") ||
                        noBlanksArr[0].includes("Urgent!! Art Request!!") || 
                        
                        noBlanksArr[0].includes("URGENT Art Request") || noBlanksArr[0].includes("URGENT Art Request!") || 
                        noBlanksArr[0].includes("URGENT! Art Request") || noBlanksArr[0].includes("URGENT! Art Request!")
                        noBlanksArr[0].includes("URGENT Art Request!!") || noBlanksArr[0].includes("URGENT!! Art Request") || 
                        noBlanksArr[0].includes("URGENT! Art Request!!") || noBlanksArr[0].includes("URGENT!! Art Request!") || 
                        noBlanksArr[0].includes("URGENT!! Art Request!!") ||

                        noBlanksArr[0].includes("Urgent ART REQUEST") || noBlanksArr[0].includes("Urgent ART REQUEST!") || 
                        noBlanksArr[0].includes("Urgent! ART REQUEST") || noBlanksArr[0].includes("Urgent! ART REQUEST!") ||
                        noBlanksArr[0].includes("Urgent ART REQUEST!!") || noBlanksArr[0].includes("Urgent!! ART REQUEST") ||
                        noBlanksArr[0].includes("Urgent! ART REQUEST!!") || noBlanksArr[0].includes("Urgent!! ART REQUEST!") ||
                        noBlanksArr[0].includes("Urgent!! ART REQUEST!!") || 

                        noBlanksArr[0].includes("URGENT ART REQUEST") || noBlanksArr[0].includes("URGENT ART REQUEST!") ||
                        noBlanksArr[0].includes("URGENT! ART REQUEST") || noBlanksArr[0].includes("URGENT! ART REQUEST!") ||
                        noBlanksArr[0].includes("URGENT ART REQUEST!!") || noBlanksArr[0].includes("URGENT!! ART REQUEST") ||
                        noBlanksArr[0].includes("URGENT! ART REQUEST!!") || noBlanksArr[0].includes("URGENT!! ART REQUEST!") || 
                        noBlanksArr[0].includes("URGENT!! ART REQUEST!!") ||



                        noBlanksArr[0].includes("Urgent Artist Request") || noBlanksArr[0].includes("Urgent Artist Request!") || 
                        noBlanksArr[0].includes("Urgent! Artist Request") || noBlanksArr[0].includes("Urgent! Artist Request!") ||
                        noBlanksArr[0].includes("Urgent Artist Request!!") || noBlanksArr[0].includes("Urgent!! Artist Request") ||
                        noBlanksArr[0].includes("Urgent! Artist Request!!") || noBlanksArr[0].includes("Urgent!! Artist Request!") ||
                        noBlanksArr[0].includes("Urgent!! Artist Request!!") || 
                        
                        noBlanksArr[0].includes("URGENT Artist Request") || noBlanksArr[0].includes("URGENT Artist Request!") || 
                        noBlanksArr[0].includes("URGENT! Artist Request") || noBlanksArr[0].includes("URGENT! Artist Request!")
                        noBlanksArr[0].includes("URGENT Artist Request!!") || noBlanksArr[0].includes("URGENT!! Artist Request") || 
                        noBlanksArr[0].includes("URGENT! Artist Request!!") || noBlanksArr[0].includes("URGENT!! Artist Request!") || 
                        noBlanksArr[0].includes("URGENT!! Artist Request!!") ||

                        noBlanksArr[0].includes("Urgent ARTIST REQUEST") || noBlanksArr[0].includes("Urgent ARTIST REQUEST!") || 
                        noBlanksArr[0].includes("Urgent! ARTIST REQUEST") || noBlanksArr[0].includes("Urgent! ARTIST REQUEST!") ||
                        noBlanksArr[0].includes("Urgent ARTIST REQUEST!!") || noBlanksArr[0].includes("Urgent!! ARTIST REQUEST") ||
                        noBlanksArr[0].includes("Urgent! ARTIST REQUEST!!") || noBlanksArr[0].includes("Urgent!! ARTIST REQUEST!") ||
                        noBlanksArr[0].includes("Urgent!! ARTIST REQUEST!!") || 

                        noBlanksArr[0].includes("URGENT ARTIST REQUEST") || noBlanksArr[0].includes("URGENT ARTIST REQUEST!") ||
                        noBlanksArr[0].includes("URGENT! ARTIST REQUEST") || noBlanksArr[0].includes("URGENT! ARTIST REQUEST!") ||
                        noBlanksArr[0].includes("URGENT ARTIST REQUEST!!") || noBlanksArr[0].includes("URGENT!! ARTIST REQUEST") ||
                        noBlanksArr[0].includes("URGENT! ARTIST REQUEST!!") || noBlanksArr[0].includes("URGENT!! ARTIST REQUEST!") || 
                        noBlanksArr[0].includes("URGENT!! ARTIST REQUEST!!") ||



                        noBlanksArr[0].includes("Urgent Creative Request") || noBlanksArr[0].includes("Urgent Creative Request!") || 
                        noBlanksArr[0].includes("Urgent! Creative Request") || noBlanksArr[0].includes("Urgent! Creative Request!") ||
                        noBlanksArr[0].includes("Urgent Creative Request!!") || noBlanksArr[0].includes("Urgent!! Creative Request") ||
                        noBlanksArr[0].includes("Urgent! Creative Request!!") || noBlanksArr[0].includes("Urgent!! Creative Request!") ||
                        noBlanksArr[0].includes("Urgent!! Creative Request!!") || 
                        
                        noBlanksArr[0].includes("URGENT Creative Request") || noBlanksArr[0].includes("URGENT Creative Request!") || 
                        noBlanksArr[0].includes("URGENT! Creative Request") || noBlanksArr[0].includes("URGENT! Creative Request!")
                        noBlanksArr[0].includes("URGENT Creative Request!!") || noBlanksArr[0].includes("URGENT!! Creative Request") || 
                        noBlanksArr[0].includes("URGENT! Creative Request!!") || noBlanksArr[0].includes("URGENT!! Creative Request!") || 
                        noBlanksArr[0].includes("URGENT!! Creative Request!!") ||

                        noBlanksArr[0].includes("Urgent CREATIVE REQUEST") || noBlanksArr[0].includes("Urgent CREATIVE REQUEST!") || 
                        noBlanksArr[0].includes("Urgent! CREATIVE REQUEST") || noBlanksArr[0].includes("Urgent! CREATIVE REQUEST!") ||
                        noBlanksArr[0].includes("Urgent CREATIVE REQUEST!!") || noBlanksArr[0].includes("Urgent!! CREATIVE REQUEST") ||
                        noBlanksArr[0].includes("Urgent! CREATIVE REQUEST!!") || noBlanksArr[0].includes("Urgent!! CREATIVE REQUEST!") ||
                        noBlanksArr[0].includes("Urgent!! CREATIVE REQUEST!!") || 

                        noBlanksArr[0].includes("URGENT CREATIVE REQUEST") || noBlanksArr[0].includes("URGENT CREATIVE REQUEST!") ||
                        noBlanksArr[0].includes("URGENT! CREATIVE REQUEST") || noBlanksArr[0].includes("URGENT! CREATIVE REQUEST!") ||
                        noBlanksArr[0].includes("URGENT CREATIVE REQUEST!!") || noBlanksArr[0].includes("URGENT!! CREATIVE REQUEST") ||
                        noBlanksArr[0].includes("URGENT! CREATIVE REQUEST!!") || noBlanksArr[0].includes("URGENT!! CREATIVE REQUEST!") || 
                        noBlanksArr[0].includes("URGENT!! CREATIVE REQUEST!!");

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

                        var theCode = splitCodes[0].trim();

                        try {
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

            //#endregion -----------------------------------------------------------------------------------------------------------------------------

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region AUTO POPULATED PRINT DATE BASED ON GROUP -------------------------------------------------------------------------------------------

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
                        formattedPrintDateMatch = [('' + (leNewDate.getMonth() + 1)).slice(-2), ('' + leNewDate.getDate()).slice(-2), 
                        (leNewDate.getFullYear() % 100)].join('/');
                        $("#print-date").val(formattedPrintDateMatch);
                    } catch (e) {
                        console.log("Error with print date autofill based on group letter input. Please debug to resolve.")
                    };

                };
            };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region AUTO POPULATE GROUP BASED ON PRINT DATE --------------------------------------------------------------------------------------------

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

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region TASKPANE BUTTONS -----------------------------------------------------------------------------------------------------------------------

        //#region ON SUBMIT CLICK --------------------------------------------------------------------------------------------------------------------

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

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region ON CLEAR CLICK ---------------------------------------------------------------------------------------------------------------------

            $("#clear").on("click", function() {

                $("#subject, #client, #location, #product, #code, #project-type, #csm, #print-date, #group, #design-managers, #queue, #tier, #tags, #start-override, #work-override, #notes").val(""); // Empty all inputs
                removeWarningClass("#subject", "#warning1");
                removeWarningClass("#client", "#warning2");
                removeWarningClass("#product", "#warning3");
                removeWarningClass("#project-type", "#warning4");

            });

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region OTHER BUTTONS ----------------------------------------------------------------------------------------------------------------------

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

                $("#close-xp").on("click", function() {
                    showElement("#fissh", "hide");
                });

                $("#reload").on("click", function() {
                    // Hide the message
                    // alert("YES YOU ARE!");
                    showMessage(undefined, "hide");
                    location.reload();
                });

                $(".gotcha").on("click", function() {
                    showElement("#na-ah-ah", "hide");
                });

                $(".cs-text").on("click", function() {
                    showElement("#color-cheat-sheet", "show");
                });

                $("#cs-back").on("click", function() {
                    showElement("#color-cheat-sheet", "hide");
                });

                $( ".collapsible" ).click(function() {
                    $(this).next().slideToggle("fast");
                
                
                    const isPlus = $(this).find("i").hasClass("fa-plus")
                
                    if (isPlus) {
                        $(this).find("i")
                            .removeClass("fa-plus")
                            .addClass("fa-minus")
                            $(this).addClass("expanded");
                
                    } else {
                        $(this).find("i")
                            .removeClass("fa-minus")
                            .addClass("fa-plus")

                            setTimeout(() => {
                                $(this).removeClass("expanded");
                            }, 200)
                    };
                    
                });

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------------------------------------------


//#region EDITING THE TABLE --------------------------------------------------------------------------------------------------------------------------

    //#region ADDING A PROJECT FROM TASKPANE ---------------------------------------------------------------------------------------------------------

        //#region TURN OFF EVENTS BEFORE EXECUTING ADD A PROJECT -------------------------------------------------------------------------------------

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

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region ADD A PROJECT ----------------------------------------------------------------------------------------------------------------------

            /**
             * Generates Added date/time, turn around times for both the Picked Up / Started By and Proof To Client columns adjusted for office hours, 
             * adds these values to the table, then generates a priority number for each row based on the value in the Picked Up / Started By column, 
             * then sorts the data by priority
             */
            async function addAProject() {

                console.log("Add A Project was fired!");

                await Excel.run(async (context) => {

                    //#region LOAD VALUES ------------------------------------------------------------------------------------------------------------

                        var sheet = context.workbook.worksheets.getActiveWorksheet().load("name");
                        sheet.load("tabColor");
                        //updating this variable to work for the changedTable will not work since the taskpane doesn't trigger an onchanged event 
                        //until afterward
                        var sheetTable = sheet.tables.getItemAt(0).load("name"); //this is fine since the user will only ever be adding new projects 
                        //to the unassigned table or the artist tables, which are all the first tables in their documents
                        sheetTable.rows.add(null);

                        var sheetTableRows = sheetTable.rows.load("items");
                        var sheetTableRange = sheetTable.getDataBodyRange().load("values");
                        var sheetTableHeader = sheetTable.getHeaderRowRange().load("values");
                        context.runtime.load("enableEvents");

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region GET INPUT FROM TASKPANE ------------------------------------------------------------------------------------------------

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

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region WRITE ARRAY ------------------------------------------------------------------------------------------------------------

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

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    await context.sync(); // BOOM!

                    //#region WRITE VARIABLES --------------------------------------------------------------------------------------------------------

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

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region GENERATE ADDED DATE ----------------------------------------------------------------------------------------------------

                        var now = new Date();
                        var toSerial = JSDateToExcelDate(now);

                        write[0][tableRowInfo.added.columnIndex] = toSerial;

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region AUTO-FILL BLANK OVERRIDES ----------------------------------------------------------------------------------------------
                        
                        if (startOverrideVal == "") {
                            write[0][tableRowInfo.startOverride.columnIndex] = 0;
                        };

                        if (workOverrideVal == "") {
                            write[0][tableRowInfo.workOverride.columnIndex] = 0;
                        };
                    
                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region AUTO-FILL STATUS -------------------------------------------------------------------------------------------------------

                        var leStatus = statusAutofill(tableName);

                        write[0][tableRowInfo.status.columnIndex] = leStatus;

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region AUTO-FILL ARTIST -------------------------------------------------------------------------------------------------------

                        var leArtist = artistAutofill(tableName, leSheetName);

                        write[0][tableRowInfo.artist.columnIndex] = leArtist;

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region AUTO-FILL DESIGN MANAGERS ----------------------------------------------------------------------------------------------

                        if ((designManagersVal == "" || designManagersVal == null)
                        &&
                        (sheet.name !== "Unassigned Projects") && (sheet.name !== "Validation")) {

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

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region AUTO-FILL TIER ---------------------------------------------------------------------------------------------------------

                        //get the Project Type Coded variable from the Project Type ID Data based on the returned Project Type from the taskpane
                        var theProjectTypeCode = projectTypeIDData[projectTypeVal].projectTypeCode;

                        if((tierVal == "" || tierVal == null) && (sheet.name !== "Validation")) {
                            var defaultTier = tierLevelData[productVal][theProjectTypeCode];
                            write[0][tableRowInfo.tier.columnIndex] = defaultTier;
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region GENERATE PICKED UP / TURN AROUND TIME VALUE ----------------------------------------------------------------------------

                        //returns turn around time value from the PickedUp Turn Around Time table based on the product and project type values
                        var pickedUpTurnAroundTime = pickupData[productVal][theProjectTypeCode];

                        //add start override time to # of hours
                        var pickedUpHours = pickedUpTurnAroundTime + Number(startOverrideVal);

                        //add new time to date added, then adjust for office hours
                        var addedDate = new Date(now);
                        var pickupOfficeHours = officeHours(addedDate, pickedUpHours);

                        //converts to excel readable format
                        var excelPickupOfficeHours = Number(JSDateToExcelDate(pickupOfficeHours));

                        //#region OFFICE HOURS TESTING VARIABLES -------------------------------------------------------------------------------------

                            //BEFORE DAY: 44670.31389 (4/19/22 7:32 AM)
                            //AFTER DAY: 44670.78264 (4/19/22 6:47 PM)
                            //DURING DAY: 44670.59444 (4/19/22 2:16 PM)
                            //ON WEEKEND: 44667.29167 (4/16/22 7:00 AM)
                            //JUST BEFORE WEEKEND: 44666.45833 (4/15/22 11:00 AM)

                            // testingDate = 44667.29167;
                            // testingDate = convertToDate(testingDate);

                            // testingHours = 24;

                            // var PickupOHAdjust = officeHours(testingDate, testingHours, officeHoursData);

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        write[0][tableRowInfo.pickedUpStartedBy.columnIndex] = excelPickupOfficeHours;

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region GENERATE ART TURN AROUND TIME VALUE ------------------------------------------------------------------------------------

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

                        //#region OFFICE HOURS TESTING VARIABLES -------------------------------------------------------------------------------------

                            //BEFORE DAY: 44670.31389 (4/19/22 7:32 AM)
                            //AFTER DAY: 44670.78264 (4/19/22 6:47 PM)
                            //DURING DAY: 44670.59444 (4/19/22 2:16 PM)
                            //ON WEEKEND: 44667.29167 (4/16/22 7:00 AM)
                            //JUST BEFORE WEEKEND: 44666.45833 (4/15/22 11:00 AM)

                            // testingDate = 44667.29167;
                            // testingDate = convertToDate(testingDate);

                            // testingHours = 24;

                            // var PickupOHAdjust = officeHours(testingDate, testingHours, officeHoursData);

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        write[0][tableRowInfo.proofToClient.columnIndex] = excelProofToClientOfficeHours;

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region SORT THE TABLE ---------------------------------------------------------------------------------------------------------

                        var tablePickedUpColumnIndex = tableRowInfo.pickedUpStartedBy.columnIndex;
                        var tableProofToClientColumnIndex = tableRowInfo.proofToClient.columnIndex;

                        rangeOfTable[tableRowIndex] = write[0]; //writes content to the excel table

                        if (leSheetName == "Unassigned Projects") {
                            var gee = leSorting(tableRowInfo, rangeOfTable, tablePickedUpColumnIndex, write[0]);
                        } else {
                            var gee = leSorting(tableRowInfo, rangeOfTable, tableProofToClientColumnIndex, write[0]);
                        };

                        var kale = rowIndexPostSort;

                        sheetTableRange.values = gee;

                        console.log("Content has been added to the table through the taskpane successfully!")

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region APPLY CONDITIONAL FORMATTING -------------------------------------------------------------------------------------------

                        var newSheetTableRows = sheetTable.rows.load("items");
                        var newSheetTableRange = sheetTable.getDataBodyRange().load("values");

                        await context.sync();

                        for (var m = 0; m < rangeOfTable.length; m++) {

                            var newTableRowItems = newSheetTableRows.items;

                            var newRangeOfTable = newSheetTableRange.values;

                            var newRowValuesOfTable = newTableRowItems[m].values;

                            var newRowRange = newSheetTableRows.getItemAt(m).getRange();

                            var newTableRowInfo = new Object();

                            for (var name of headerOfTable[0]) {
                                theGreatestFunctionEverWritten(headerOfTable, name, newRowValuesOfTable, newRangeOfTable, newTableRowInfo, m);
                            };

                            console.log("ConForm is about to trigger for when a project is added to the sheet through the taskpane");

                            conditionalFormatting(newTableRowInfo, 0, sheet, m, false, newRowRange, null);

                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    await context.sync();

                    var xyz = 12;

                    if (xyz == 12) {
                        console.log("Wow, that stinks");
                    };


                    var shouldAutoLogo = true;

                    //if product is a Logo Recreation, Logo Creation, Map Creation, Media Kit, or any Marco's Product, skip logo line generation
                    // || productVal !== "Marco's Nutritional Guide - MPNutGuide") {

                    var excludeAutoLogo = ["Logo Recreation", "Logo Creation", "Map Creation", "Media Kit", 
                    "Marco's 6.8 Box Topper A La Carte - MPBTAC", "Marco's 6.8 Box Topper National Offer - MPBTNO", "Marco's 8.5x11 Poster - MPP", 
                    "Marco's 24x36 Poster - MPPO2436", "Marco's 30x40 Poster - MPPO3040", "Marco's Counter Card - MPCC", 
                    "Marco's Door Stickers - MPCMS", "Marco's Interior Sticker - MPIS", "Marco's Floor Stickers - MPSCFS", 
                    "Marco's Napkin Dispenser Insert - MPND", "Marco's Laminated Menu - MPPICMENU", "Marco's Laminated Sign - MPCVDLAM", 
                    "Marco's 8.5x11 Window Cling - MPC", "Marco's 24x36 Exterior Window Cling A La Carte - MPEXTWC2436AC", 
                    "Marco's 24x36 Exterior Window Cling National Offer - MPEXTWC2436NO", 
                    "Marco's 30x40 Exterior Window Cling A La Carte - MPEXTWC3040AC", 
                    "Marco's 30x40 Exterior Window Cling National Offer - MPEXTWC3040NO", 
                    "Marco's 24x36 Interior Window Cling A La Carte - MPINTWC2436AC", 
                    "Marco's 24x36 Interior Window Cling National Offer - MPINTWC2436NO", 
                    "Marco's 30x40 Interior Window Cling A La Carte - MPINTWC3040AC", 
                    "Marco's 30x40 Interior Window Cling National Offer - MPINTWC3040NO", "Marco's Nutritional Guide - MPNutGuide"];

                    //#region AUTO-GENERATE LOGO RECREATION LINE -------------------------------------------------------------------------------------

                        for (var item of excludeAutoLogo) {
                            if (productVal == item) {
                                shouldAutoLogo = false;
                            };
                        };

                        console.log("farts");

                        if (shouldAutoLogo == true) {

                            sheetTable.rows.add(null);
                            var reSheetTableRows = sheetTable.rows.load("items");
                            var reSheetTableRange = sheetTable.getDataBodyRange().load("values");


                            await context.sync();


                            var newTableRowIndex = sheetTable.rows.count - 1;
                            var newTableRowItems = reSheetTableRows.items;
                            var newRangeOfTable = reSheetTableRange.values
                            var newRowValuesOfTable = newTableRowItems[newTableRowIndex].values;

                            var newTableRowInfo = new Object();

                            for (var name of headerOfTable[0]) {
                                theGreatestFunctionEverWritten(headerOfTable, name, newRowValuesOfTable, newRangeOfTable, 
                                newTableRowInfo, newTableRowIndex);
                            };

                            //#region WRITE ARRAY ----------------------------------------------------------------------------------------------------

                                // Data to send to Table
                                var writeLogo = [[
                                    "", // 0 - Priority
                                    theDesignManager, // 1 - Design Manager
                                    queueVal, // 2 - Queue
                                    defaultTier, // 3 - Tier
                                    subjectVal, // 4 - Subject
                                    clientVal, // 5 - Client
                                    locationVal, // 6 - Location
                                    "Logo Recreation", // 7 - Product
                                    projectTypeVal, // 8 - Project Type
                                    csmVal, // 9 - CSM
                                    toSerial, // 10 - Added
                                    printDateVal, // 11 - Print Data
                                    groupVal, // 12 - Group
                                    excelPickupOfficeHours, // 13 - Picked Up / Started By
                                    "", // 14 - Proof to Client
                                    "", // 15 - Date of Last Edit
                                    tagsVal, // 16 - Tags
                                    "Logo Status TBD", // 17 - Status
                                    codeVal, // 18 - Code
                                    leArtist, // 19 - Artist
                                    notes, // 20 - Notes
                                    0, // 21 - Start Override
                                    0 // 22 - Work Override
                                ]];

                            //#endregion -------------------------------------------------------------------------------------------------------------

                            //use same pick up time because the logo should be assigned at the same time the product is

                            //will need the art turn around time to be the end of the work day on the group print date
                            var thePrintDate = new Date(printDateVal);

                            var datePrint = thePrintDate.getDate();
                            var monthPrint = thePrintDate.getMonth();
                            var yearPrint = thePrintDate.getFullYear();

                            var printDateDOW = thePrintDate.getDay();

                            if (printDateDOW == 0) {
                                printDateDOW = "Sunday"
                            } else if (printDateDOW == 1) {
                                printDateDOW = "Monday"
                            } else if (printDateDOW == 2) {
                                printDateDOW = "Tuesday"
                            } else if (printDateDOW == 3) {
                                printDateDOW = "Wednesday"
                            } else if (printDateDOW == 4) {
                                printDateDOW = "Thursday"
                            } else if (printDateDOW == 5) {
                                printDateDOW = "Friday"
                            } else if (printDateDOW == 6) {
                                printDateDOW = "Saturday"
                            };

                            var groupWeekdayVars = officeHoursData[printDateDOW];

                            var endOfGroupDay = convertToDate(groupWeekdayVars.endTime);

                            endOfGroupDay.setFullYear(yearPrint);
                            endOfGroupDay.setMonth(monthPrint);
                            endOfGroupDay.setDate(datePrint);

                            console.log(endOfGroupDay);

                            var groupDateExcel = Number(JSDateToExcelDate(endOfGroupDay));

                            writeLogo[0][tableRowInfo.proofToClient.columnIndex] = groupDateExcel;

                            //this edit will need to be made also to the table changed event so that when the group is changed for these specific 
                            //requests, the art turn around time will also auto-updated accordingly.

                            //#region SORT THE TABLE, AGAIN ------------------------------------------------------------------------------------------

                                newRangeOfTable[newTableRowIndex] = writeLogo[0]; //writes content to the excel table

                                if (leSheetName == "Unassigned Projects") {
                                    var geePee = leSorting(newTableRowInfo, newRangeOfTable, tablePickedUpColumnIndex, writeLogo[0]);
                                } else {
                                    var geePee = leSorting(newTableRowInfo, newRangeOfTable, tableProofToClientColumnIndex, writeLogo[0]);
                                };
            
                                reSheetTableRange.values = geePee;
        
                                console.log("Logo Recreation line has been automatically added to the table successfully!")

                            //#endregion -------------------------------------------------------------------------------------------------------------


                            //#region APPLY CONDITIONAL FORMATTING -----------------------------------------------------------------------------------

                                var newReSheetTableRows = sheetTable.rows.load("items");
                                var newReSheetTableRange = sheetTable.getDataBodyRange().load("values");

                                await context.sync();

                                for (var n = 0; n < newRangeOfTable.length; n++) {

                                    var newReTableRowItems = newReSheetTableRows.items;

                                    var newReRangeOfTable = newReSheetTableRange.values;

                                    var newReRowValuesOfTable = newReTableRowItems[n].values;

                                    var newReRowRange = newReSheetTableRows.getItemAt(n).getRange();

                                    var newReTableRowInfo = new Object();

                                    for (var name of headerOfTable[0]) {
                                        theGreatestFunctionEverWritten(headerOfTable, name, newReRowValuesOfTable, newReRangeOfTable, 
                                        newReTableRowInfo, n);
                                    };

                                    console.log("ConForm is about to trigger for when the logo recreation line is auto-generated");

                                    conditionalFormatting(newReTableRowInfo, 0, sheet, n, false, newReRowRange, null);

                                };

                            //#endregion -------------------------------------------------------------------------------------------------------------

                        } else {
                            console.log("Product does not need a logo recreation, so no extra line was generated for this project.")
                        };
                        
                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                });

                eventsOn();
                console.log("Events: ON  →  turned on in the addAProject function after a project was added to the sheet through the taskpane!");

            };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region TABLE SELECTION HIGHLIGHTING -----------------------------------------------------------------------------------------------------------

        /**
         * Adds a purple row highlight to the row of the current selection
         * @param {Object} eventArgs The event arguments, which are details about the event that was triggered
         */
        async function onTableSelectionChangedEvents(eventArgs) {
            await Excel.run(/*previousSelection,*/ async (context) => {

                var theActiveWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId);
                var activeWorksheetTables = theActiveWorksheet.tables.load("items/count");

                var currentRange = theActiveWorksheet.getRange(eventArgs.address);
                currentRange.load("columnIndex");
                var currentRow = currentRange.getRow();

                var changeType = eventArgs.changeType;

                await context.sync();

                var activeWorksheetTablesCount = activeWorksheetTables.count;

                var previousColumn = currentRange.columnIndex;

                //if user has made a selection prior to the current selection without triggering a reload, the previousSelectionObj should 
                //have arguments that will bring the user into this function to load in variables to handle the previous row highlighting
                if (previousSelectionObj.tableId !== "") {
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

                //if previousTable is undefined, then the last function never fired, meaning that thsi is the first time the user is selecting 
                //anything this run. Since there are no previous selection variables stored, we skip this function. 
                if (previousTable !== undefined) {

                    var previousTableName = previousTable.name

                    var newLeTable = previousTableRange.values;

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
                range.load(['address', 'values', 'rowIndex', 'columnIndex']);

                var selectedTableRows = selectedTable.rows.load("items");
                var selectedTableRowsCount = selectedTable.rows.load("count");


                await context.sync();

                var isTableEmpty = selectedTableRowsCount.count;

                if (isTableEmpty == 0) {
                    console.log("Table is empty, so no highlighting was applied");
                    return;
                };

                //adds formatting to current row
                if (eventArgs.address !== "") { //if the selection address is not a part of a table, this function is skipped

                    //applies border to selected row
                    var rI = range.rowIndex;
                    var cI = range.columnIndex;

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
                        console.log("Current selection address was the same as previous selection address, so previous row formatting was prevented");
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

                        // console.log("farts");
                        // console.log(rowInfoSorted.group.value);
                        // console.log(rowInfoSorted.printDate.value);

                        if (previousSelectionObj.columnIndex == rowInfoSorted.group.columnIndex) {
                            var groupUppercase = rowInfoSorted.group.value.toUpperCase();
                            var matchPrintDate = groupRefData[groupUppercase].printDate;
                            if (matchPrintDate == undefined) {
                                matchPrintDate = "N/A";
                            };
                            // newLeTable[changedRowTableIndex][rowInfoSorted.printDate.columnIndex] = matchPrintDate;
                            // newLeTable[changedRowTableIndex][rowInfoSorted.group.columnIndex] = groupUppercase;

                            rowInfoSorted.printDate.value = matchPrintDate;

                            console.log("Print Date in the Previosu Selection Obj Row Info Sorted value was updated to match Group Letter!")
                        };
          

                        // if (previousSelectionObj.columnIndex == rowInfoSorted.group.columnIndex
                        //     || previousSelectionObj.columnIndex == rowInfoSorted.printDate.columnIndex) {
                        //     console.log("A group letter or print date was updated, so the row highlight will refrain from updating the row colors");
                        //     return;
                        // }

                        // if (Number(previousColumn) !== Number(rowInfoSorted.printDate.columnIndex) || Number(previousColumn) !== Number(rowInfoSorted.group.columnIndex)) {

                        // if (previousSelectionObj.columnIndex !== rowInfoSorted.group.columnIndex
                        //     && previousSelectionObj.columnIndex !== rowInfoSorted.printDate.columnIndex) {

                            console.log("ConForm is about to trigger for when a row selection highlight changes");

                            conditionalFormatting(rowInfoSorted, tableStart, previousWorksheet, previousRowIndex, 
                                completedTableChanged, previousSelectionRange, null);
                        // };

                    };

                    if (rI !== 0) {
                        bees.format.fill.color = "#F5D9FF";
                        bees.format.font.color = "black";
            
                        previousSelectionObj.tableId = eventArgs.tableId;
            
                        previousSelectionObj.address = eventArgs.address;
            
                        previousSelectionObj.rowIndex = rI;

                        previousSelectionObj.columnIndex = cI;
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


                        console.log("ConForm is about to trigger for when a row selection highlight changes");
                        conditionalFormatting(rowInfoSorted, tableStart, previousWorksheet, previousRowIndex, 
                            completedTableChanged, previousSelectionRange, null);
                    };

                }

            }).catch (err => {
                if (dennisHere == true) {
                    return;
                };
                console.log(err) // <--- does this log?
                showMessage(err, "show");
                context.runtime.enableEvents = true;
            });
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region ON TABLE CHANGED -----------------------------------------------------------------------------------------------------------------------

        /**
         * When anything is changed in the workbook, it is handled here
         * @param {Object} eventArgs The event arguments, which are details about the event that was triggered
         */
        async function onTableChanged(eventArgs) {
            await Excel.run(async (context) => {

                //#region HANDLE REMOTE CHANGES ------------------------------------------------------------------------------------------------------

                    console.log("Source of the onTableChanged event: " + eventArgs.source);

                    if (eventArgs.source == "Remote") {
                        console.log("Content was changed by a remote user, exiting onTableChanged Event");
                        return;
                    };

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region HANDLE ILLEGAL ROW INSERT --------------------------------------------------------------------------------------------------

                    if (eventArgs.changeType == "RowInserted") {
                        handleIllegalInsert(eventArgs);
                        dennisHere = true;
                        showDennis();
                        return;
                    };

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                context.runtime.load("enableEvents"); //loads runtime events so I can turn them off and on after the context.sync()

                await context.sync();

                //turns events off
                context.runtime.enableEvents = false;
                console.log("Events: OFF - Occured in onTableChanged!");

                //#region LOAD VARIABLES FROM WORKBOOK -----------------------------------------------------------------------------------------------

                    var details = eventArgs.details;
                    var address = eventArgs.address;
                    var changeType = eventArgs.changeType;
                    //console.log(changeType);

                    var allWorksheets = context.workbook.worksheets;
                    allWorksheets.load("items/name/tables/id");
                    var changedWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId).load("name");
                    var worksheetTables = changedWorksheet.tables;
                    var valSheet = context.workbook.worksheets.getItem("Validation").load("name");

                    //Used to find the column and row index on a worksheet level
                    var changedAddress = changedWorksheet.getRange(address);
                    changedAddress.load("columnIndex");
                    changedAddress.load("rowIndex");

                    //Used to load values of the changed row/table to be used in functions & to return updated values to the table
                    var allTables = context.workbook.tables;
                    allTables.load("items/name");
                    var tablesInWorksheet = changedWorksheet.tables.load("items/count");
                    //Returns tableId of the table where the event occured
                    changedTable = context.workbook.tables.getItem(eventArgs.tableId).load("name");
                    var changedTableColumns = changedTable.columns
                    changedTableColumns.load("items/name");
                    var changedTableRows = changedTable.rows;
                    changedTableRows.load("items");
                    var startOfTable = changedTable.getRange().load("columnIndex");

                    // LEGACY CODE THAT NEEDS TO BE UPDATED TO BE MORE FLEXIBLE ======================================================================
                    // ===============================================================================================================================

                        //#region SPECIFIC TABLE VARIABLES -------------------------------------------------------------------------------------------

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

                            var michaelTable = context.workbook.tables.getItem("MichaelProjects").load("worksheet");
                            var michaelTableName = context.workbook.tables.getItem("MichaelProjects").load("name");
                            var michaelTableRows = michaelTable.rows;
                            michaelTableRows.load("items");
                            var michaelRange = michaelTable.getDataBodyRange().load("values");
                            var michaelHeader = michaelTable.getHeaderRowRange().load("values");

                            var sarahTable = context.workbook.tables.getItem("SarahProjects").load("worksheet");
                            var sarahTableName = context.workbook.tables.getItem("SarahProjects").load("name");
                            var sarahTableRows = sarahTable.rows;
                            sarahTableRows.load("items");
                            var sarahRange = sarahTable.getDataBodyRange().load("values");
                            var sarahHeader = sarahTable.getHeaderRowRange().load("values");

                            var dannyTable = context.workbook.tables.getItem("DannyProjects").load("worksheet");
                            var dannyTableName = context.workbook.tables.getItem("DannyProjects").load("name");
                            var dannyTableRows = dannyTable.rows;
                            dannyTableRows.load("items");
                            var dannyRange = dannyTable.getDataBodyRange().load("values");
                            var dannyHeader = dannyTable.getHeaderRowRange().load("values");

                            var aliTable = context.workbook.tables.getItem("AliProjects").load("worksheet");
                            var aliTableName = context.workbook.tables.getItem("AliProjects").load("name");
                            var aliTableRows = aliTable.rows;
                            aliTableRows.load("items");
                            var aliRange = aliTable.getDataBodyRange().load("values");
                            var aliHeader = aliTable.getHeaderRowRange().load("values");

                            var joshKTable = context.workbook.tables.getItem("JoshKProjects").load("worksheet");
                            var joshKTableName = context.workbook.tables.getItem("JoshKProjects").load("name");
                            var joshKTableRows = joshKTable.rows;
                            joshKTableRows.load("items");
                            var joshKRange = joshKTable.getDataBodyRange().load("values");
                            var joshKHeader = joshKTable.getHeaderRowRange().load("values");

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

                            var ethanTable = context.workbook.tables.getItem("EthanProjects").load("worksheet");
                            var ethanTableName = context.workbook.tables.getItem("EthanProjects").load("name");
                            var ethanTableRows = ethanTable.rows;
                            ethanTableRows.load("items");
                            var ethanRange = ethanTable.getDataBodyRange().load("values");
                            var ethanHeader = ethanTable.getHeaderRowRange().load("values");

                            var jessicaTable = context.workbook.tables.getItem("JessicaProjects").load("worksheet");
                            var jessicaTableName = context.workbook.tables.getItem("JessicaProjects").load("name");
                            var jessicaTableRows = jessicaTable.rows;
                            jessicaTableRows.load("items");
                            var jessicaRange = jessicaTable.getDataBodyRange().load("values");
                            var jessicaHeader = jessicaTable.getHeaderRowRange().load("values");

                            var joshCTable = context.workbook.tables.getItem("JoshCProjects").load("worksheet");
                            var joshCTableName = context.workbook.tables.getItem("JoshCProjects").load("name");
                            var joshCTableRows = joshCTable.rows;
                            joshCTableRows.load("items");
                            var joshCRange = joshCTable.getDataBodyRange().load("values");
                            var joshCHeader = joshCTable.getHeaderRowRange().load("values");

                            var emilyTable = context.workbook.tables.getItem("EmilyProjects").load("worksheet");
                            var emilyTableName = context.workbook.tables.getItem("EmilyProjects").load("name");
                            var emilyTableRows = emilyTable.rows;
                            emilyTableRows.load("items");
                            var emilyRange = emilyTable.getDataBodyRange().load("values");
                            var emilyHeader = emilyTable.getHeaderRowRange().load("values");

                            var alainaTable = context.workbook.tables.getItem("AlainaProjects").load("worksheet");
                            var alainaTableName = context.workbook.tables.getItem("AlainaProjects").load("name");
                            var alainaTableRows = alainaTable.rows;
                            alainaTableRows.load("items");
                            var alainaRange = alainaTable.getDataBodyRange().load("values");
                            var alainaHeader = alainaTable.getHeaderRowRange().load("values");

                            var ritaTable = context.workbook.tables.getItem("RitaProjects").load("worksheet");
                            var ritaTableName = context.workbook.tables.getItem("RitaProjects").load("name");
                            var ritaTableRows = ritaTable.rows;
                            ritaTableRows.load("items");
                            var ritaRange = ritaTable.getDataBodyRange().load("values");
                            var ritaHeader = ritaTable.getHeaderRowRange().load("values");

                            var dawnTable = context.workbook.tables.getItem("DawnProjects").load("worksheet");
                            var dawnTableName = context.workbook.tables.getItem("DawnProjects").load("name");
                            var dawnTableRows = dawnTable.rows;
                            dawnTableRows.load("items");
                            var dawnRange = dawnTable.getDataBodyRange().load("values");
                            var dawnHeader = dawnTable.getHeaderRowRange().load("values");

                            var joeyTable = context.workbook.tables.getItem("JoeyProjects").load("worksheet");
                            var joeyTableName = context.workbook.tables.getItem("JoeyProjects").load("name");
                            var joeyTableRows = joeyTable.rows;
                            joeyTableRows.load("items");
                            var joeyRange = joeyTable.getDataBodyRange().load("values");
                            var joeyHeader = joeyTable.getHeaderRowRange().load("values");

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

                            var kristenTable = context.workbook.tables.getItem("KristenProjects").load("worksheet");
                            var kristenTableName = context.workbook.tables.getItem("KristenProjects").load("name");
                            var kristenTableRows = kristenTable.rows;
                            kristenTableRows.load("items");
                            var kristenRange = kristenTable.getDataBodyRange().load("values");
                            var kristenHeader = kristenTable.getHeaderRowRange().load("values");

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                    // ===============================================================================================================================
                    // ===============================================================================================================================

                    //all of the data from the changed table
                    var bodyRange = changedTable.getDataBodyRange().load("values");

                    //the header data of the changed table
                    var headerRange = changedTable.getHeaderRowRange().load("values");

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                await context.sync();

                    //#region ASSIGNING VARIABLES ----------------------------------------------------------------------------------------------------

                        //#region VALIDATION EXODUS --------------------------------------------------------------------------------------------------

                            if (changedWorksheet.name == valSheet.name) { //if the change was made to the Validation sheet, exit the function
                                console.log("Validation Sheet was changed, exiting the table changed event...")
                                return;
                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region CREATING AND ASSIGNING WORKBOOK VARIABLES --------------------------------------------------------------------------

                            //#region CALL LOADED VARIABLES ------------------------------------------------------------------------------------------

                                if (changedWorksheet.name == "Unassigned Projects") {
                                    var completedTable = null;
                                } else {
                                    var completedTable = worksheetTables.getItemAt(1);
                                };

                                //index of the column where the change was made (on a worksheet level)
                                var changedColumnIndexOG = changedAddress.columnIndex;
                                var changedRowIndex = changedAddress.rowIndex; //index of the row where the change was made (on a worksheet level)

                                var tableColumns = changedTableColumns.items; //loads all the changed table's columns
                                var tableRows = changedTableRows.items; //loads all the changed table's rows
                                if (changeType == "RowDeleted") {
                                    var changedRowTableIndex = 0;
                                } else {
                                    var changedRowTableIndex = changedRowIndex - 1; //adjusts index number for table level (-1 to skip header row)
                                };
                                var rowValues = tableRows[changedRowTableIndex].values; //loads the values of the changed row in the changed table
                                //loads the changed row in the changed table as an object
                                var myRow = changedTableRows.getItemAt(changedRowTableIndex);
                                var rowRange = changedTableRows.getItemAt(changedRowTableIndex).getRange();
                                var justToCheck = rowIndexPostSort;
                                var tablesInWorksheetCount = tablesInWorksheet.count;

                                var tableContent = bodyRange.values; //all of the changed table's content
                                var head = headerRange.values; //all of the changed table's headers

                                var tableStart = startOfTable.columnIndex; //column index of the start of the table
                                //adjusts columnIndex to reflect the actual position in the table, no matter where the table is on the sheet
                                var changedColumnIndex = changedColumnIndexOG - tableStart;

                            //#endregion -------------------------------------------------------------------------------------------------------------

                            //#region RECREATES CHANGED TABLE IN CODE AND ASSIGNS COLUMN INDEX AND VALUE PROPERTIES OF THE CHANGED ROW TO AN OBJECT --

                                var leTable = JSON.parse(JSON.stringify(tableContent)); //creates a duplicate array of the entire changed tables 
                                //content to be used for making adjustments to the sheet, without having anything done to it affect oriignal array

                                var rowInfo = new Object(); //object that will contain the values and column indexs of every item in the changed row

                                for (var name of head[0]) { //for each header item in the head array...
                                    //creates keys with the header names of each column in the changed table and assigns them to the rowInfo object. 
                                    //For each key, the column index and cell values are added for the cell in that column in the changed row
                                    theGreatestFunctionEverWritten(head, name, rowValues, leTable, rowInfo, changedRowTableIndex);
                                };

                                var pickedUpColumnIndex = rowInfo.pickedUpStartedBy.columnIndex; //index of picked up column
                                var proofToClientColumnIndex = rowInfo.proofToClient.columnIndex; //index of proof to client cloumn
                                var addedColumnIndex = rowInfo.added.columnIndex;
                                var statusValue = rowInfo.status.value;

                            //#endregion -------------------------------------------------------------------------------------------------------------

                              // if (changeType == "RowInserted") {
                            //     console.log("tsk tsk tsk...Don't forget the 7th commandment of the Art Queue Add-In:");
                            //     console.log('"Thou shalt submit all requests to thy own sheet by means of the Add A Project taskpane. 
                            //     Manually adding rows of info to thyn sheet is forbidden."');
                            //     console.log("It's a simple mistake, but make sure not to do it again.");
                            //     rowRange.delete("Up");
                            //     // eventsOn();
                            //     // console.log("Events: ON  →  triggered after a row was manually inserted into the sheet by the user, 
                            //     followed by the swift removal of said row and a slap on the wrist.");
                            //     return;
                            // };

                            //#region FINDS IF CHANGE WAS MADE TO THE UNASSIGNED PROJECTS TABLE OR NOT -----------------------------------------------

                                var isUnassigned;

                                if (changedWorksheet.name == "Unassigned Projects") {
                                    isUnassigned = true;
                                } else {
                                    isUnassigned = false;
                                };

                            //#endregion -------------------------------------------------------------------------------------------------------------

                            //#region FINDS IF CHANGED TABLE IS A COMPLETED TABLE OR NOT -------------------------------------------------------------

                                var changedTableName = changedTable.name;

                                var completedTableChanged = changedTableName.includes("Completed");

                            //#endregion -------------------------------------------------------------------------------------------------------------

                            //#region FINDS IF STATUS IS MOVING DATA ---------------------------------------------------------------------------------

                                var statusMove = false;

                                if (rowInfo.status.value == "Completed" || rowInfo.status.value == "Cancelled") {
                                    statusMove = true;
                                };

                            //#endregion -------------------------------------------------------------------------------------------------------------

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region ROW DELETED ------------------------------------------------------------------------------------------------------------

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

                                console.log("ConForm is about to trigger for when a row gets deleted from the sheet");

                                conditionalFormatting(rowInfoSorted, tableStart, changedWorksheet, m, completedTableChanged, rowRangeSorted, null);

                            };

                            eventsOn();
                            console.log("Events: ON  →  turned on after a row was deleted within the onTableChanged function!");

                            return;

                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region PRINT DATE / GROUP LETTER AUTO UPDATE ----------------------------------------------------------------------------------

                        if ((changedColumnIndex == rowInfo.printDate.columnIndex) || (changedColumnIndex == rowInfo.group.columnIndex)) {

                            //update group letter if the print date is changed
                            if (changedColumnIndex == rowInfo.printDate.columnIndex) {

                                var formattedDate = convertToDate(rowInfo.printDate.value);
                                var newerDate = new Date(formattedDate);
                                formattedDate = [
                                    ('' + (newerDate.getMonth() + 1)).slice(-2), 
                                    ('' + newerDate.getDate()).slice(-2), 
                                    (newerDate.getFullYear() % 100)
                                ].join('/');

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

                                console.log("Group Letter was updated to match the Print Date!");
                            };

                            //update the print date if the group letter is changed
                            if (changedColumnIndex == rowInfo.group.columnIndex) {
                                var groupUppercase = rowInfo.group.value.toUpperCase();
                                var matchPrintDate = groupRefData[groupUppercase].printDate;
                                if (matchPrintDate == undefined) {
                                    matchPrintDate = "N/A";
                                };
                                leTable[changedRowTableIndex][rowInfo.printDate.columnIndex] = matchPrintDate;
                                leTable[changedRowTableIndex][rowInfo.group.columnIndex] = groupUppercase;
                                bodyRange.values = leTable;

                                console.log("Print Date was updated to match the Group Letter!")
                            };

                            var newChangedTableRows = changedTable.rows.load("items");

                            await context.sync();

                            var newTableRows = newChangedTableRows.items;

                            var newRowValues = newTableRows[changedRowTableIndex].values;

                            var newRowInfo = new Object(); //object that will contain the values and column indexs of every item in the changed row

                            for (var name of head[0]) { //for each header item in the head array...
                                theGreatestFunctionEverWritten(head, name, newRowValues, leTable, newRowInfo, changedRowTableIndex);
                            };

                            console.log("ConForm is about to trigger for when the print date or group was updated");

                            conditionalFormatting(newRowInfo, tableStart, changedWorksheet, changedRowTableIndex, 
                                completedTableChanged, rowRange, completedTable);

                            console.log("Conditional Formatting was applied to the Group/Print Date");


                            var groupPrintChangedTableRows = changedTable.rows.load("items");
                            var groupPrintBodyValues = changedTable.getDataBodyRange().load("values");

                            // console.log("The table's data body ranfe was re-evaluated");


                            await context.sync();


                            //#region IF PRODUCT IS A LOGO RECREATION AND HAS A LOGO SPECIFIFC STATUS, UPDATE TURN AROUND TIME -----------------------

                                //if product is a logo recreation AND it has any of the three main recreation statuses (used spcifically for 
                                //when a logo is being recreated based on another project), do the following...
                                if (rowInfo.product.value == "Logo Recreation" && (rowInfo.status.value == "Logo Status TBD" 
                                || rowInfo.status.value == "Logo Needs Recreating" || rowInfo.status.value == "Logo Needs Uploading")) {


                                    var groupPrintTabledSorted = groupPrintBodyValues.values;
                                    var groupPrintTableRowsSorted = groupPrintChangedTableRows.items;

                                    // for (var Q = 0; Q < leTable.length; Q++) {

                                        var groupPrintRowRangeSorted = groupPrintChangedTableRows.getItemAt(changedRowTableIndex).getRange();
        
                                        var groupPrintRowValuesSorted = groupPrintTableRowsSorted[changedRowTableIndex].values;
        
                                        var groupPrintRowInfo = new Object();
        
                                        for (var name of head[0]) {
                                            theGreatestFunctionEverWritten(head, name, groupPrintRowValuesSorted, leTable, groupPrintRowInfo, changedRowTableIndex);
                                        };

                                    // };

                                    var groupPrintProofToClient = getProofToClientTime(groupPrintRowInfo, leTable, 0, changedRowTableIndex);

                                    groupPrintRowInfo.proofToClient.value = groupPrintProofToClient;

                                    var groupPrintChangedRowValues = leTable[changedRowTableIndex];

                                    if (changedTable.id == unassignedTable.id) {
                                        leTable = leSorting(groupPrintRowInfo, leTable, pickedUpColumnIndex, groupPrintChangedRowValues);
                                    };
                                    if (changedTable.id !== unassignedTable.id && completedTableChanged == false) {
                                        leTable = leSorting(groupPrintRowInfo, leTable, proofToClientColumnIndex, groupPrintChangedRowValues);
                                    };

                                    bodyRange.values = leTable;

                                    await context.sync();

                                    console.log("No Poop in this SouP!!");
        
                                    var newChangedTableRows = changedTable.rows;
                                    newChangedTableRows.load("items");
        
                                    await context.sync();
        
                                    var tableRows = newChangedTableRows.items; //loads all the changed table's rows
        
                                    for (var m = 0; m < tableRows.length; m++) {
        
                                        var rowRangeSorted = newChangedTableRows.getItemAt(m).getRange();
        
                                        var rowValuesSorted = tableRows[m].values;
        
                                        var rowInfoSorted = new Object();
        
                                        for (var name of head[0]) {
                                            theGreatestFunctionEverWritten(head, name, rowValuesSorted, leTable, rowInfoSorted, m);
                                        };

                                        console.log("About to trigger ConForm for Logo Recreation line being removed!");
        
                                        conditionalFormatting(rowInfoSorted, tableStart, changedWorksheet, m, 
                                            completedTableChanged, rowRangeSorted, null);


                                        // console.log("Conditional formatting was applied to row " + m);
        
                                    };

        


                                    // var thePrintDateUpdated = convertToDate(groupPrintRowInfo.printDate.value)
                                    // var updatedPrintDate = new Date(thePrintDateUpdated);
                                    // var updatedDatePrint = updatedPrintDate.getDate();
                                    // var updatedMonthPrint = updatedPrintDate.getMonth();
                                    // var updatedYearPrint = updatedPrintDate.getFullYear();

                                    // var updatedPrintDateDOW = updatedPrintDate.getDay();

                                    // if (updatedPrintDateDOW == 0) {
                                    //     updatedPrintDateDOW = "Sunday"
                                    // } else if (updatedPrintDateDOW == 1) {
                                    //     updatedPrintDateDOW = "Monday"
                                    // } else if (updatedPrintDateDOW == 2) {
                                    //     updatedPrintDateDOW = "Tuesday"
                                    // } else if (updatedPrintDateDOW == 3) {
                                    //     updatedPrintDateDOW = "Wednesday"
                                    // } else if (updatedPrintDateDOW == 4) {
                                    //     updatedPrintDateDOW = "Thursday"
                                    // } else if (updatedPrintDateDOW == 5) {
                                    //     updatedPrintDateDOW = "Friday"
                                    // } else if (updatedPrintDateDOW == 6) {
                                    //     updatedPrintDateDOW = "Saturday"
                                    // };

                                    // var updatedWeekdayVars = officeHoursData[updatedPrintDateDOW];

                                    // var updatedGroupEndOfDay = convertToDate(updatedWeekdayVars.endTime);

                                    // updatedGroupEndOfDay.setFullYear(updatedYearPrint);
                                    // updatedGroupEndOfDay.setMonth(updatedMonthPrint);
                                    // updatedGroupEndOfDay.setDate(updatedDatePrint);

                                    // console.log(updatedGroupEndOfDay);

                                    // var updatedGroupDateExcel = Number(JSDateToExcelDate(updatedGroupEndOfDay));

                                    // leTable[changedRowTableIndex][rowInfo.proofToClient.columnIndex] = updatedGroupDateExcel;
                                    //bodyRange.values = leTable;

                                };

                            //#endregion -------------------------------------------------------------------------------------------------------------

                        };

                        // var statusMove = false;

                        // if (rowInfo.status.value == "Completed" || rowInfo.status.value == "Cancelled") {
                        //     statusMove = true;
                        // };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region ADJUST TURN AROUND TIMES, SORTING, & PRIORITY NUMBERS ------------------------------------------------------------------

                        //if any of these columns are changed (all columns that can potentially affect turn around times), 
                        //turn around times will be adjusted and the table will be sorted
                        if (changedColumnIndex == rowInfo.pickedUpStartedBy.columnIndex 
                            || changedColumnIndex == rowInfo.proofToClient.columnIndex 
                            || changedColumnIndex == rowInfo.priority.columnIndex 
                            || changedColumnIndex == rowInfo.product.columnIndex 
                            || changedColumnIndex == rowInfo.projectType.columnIndex 
                            || changedColumnIndex == rowInfo.added.columnIndex 
                            || changedColumnIndex == rowInfo.startOverride.columnIndex 
                            || changedColumnIndex == rowInfo.workOverride.columnIndex 
                            || changedColumnIndex == rowInfo.tags.columnIndex
                            || (changedColumnIndex == rowInfo.status.columnIndex //if status is adjusted & table is NOT a completed table & the status 
                            //is not moving the data to the completed table (saved for another if statement that handles moving data between tables)
                                && completedTableChanged == false 
                                && statusMove == false)
                        ) {

                            console.log("I will update the turn around times, priority numbers, and sort the sheet before turning events back on!");


                            //#region REMOVE LOGO RECREATION LINE WHEN STATUS IS SET TO "NO LOGO RECREATION NEEDED" ----------------------------------
                                
                                if (changedColumnIndex == rowInfo.status.columnIndex && rowInfo.product.value == "Logo Recreation" 
                                && rowInfo.status.value == "No Logo Recreation Needed") {

                                    myRow.delete();
                                    // console.log("row was deleted from table, affective after context.sync()");
                                    leTable.splice(changedRowTableIndex, 1);
                                    // console.log("row removed from leTable array");

                                    bodyRange = changedTable.getDataBodyRange().load("values");
                                    // console.log("bodyRange has been evaluated");


                                    await context.sync();

                                    if (changedTable.id == unassignedTable.id) {
                                        leTable = leSorting(rowInfo, leTable, pickedUpColumnIndex, rowValues[0]);
                                    };
                                    if (changedTable.id !== unassignedTable.id && completedTableChanged == false) {
                                        leTable = leSorting(rowInfo, leTable, proofToClientColumnIndex, rowValues[0]);
                                    };
                                    // console.log("leTable has been resorted");
        
                                    bodyRange.values = leTable;

                                    // console.log("bodyRange has now been updated to the sorted table");
        
                                    await context.sync();

                                    console.log("How bout dat soup, man??");
        
                                    var newChangedTableRows = changedTable.rows;
                                    newChangedTableRows.load("items");
        
                                    await context.sync();
        
                                    var tableRows = newChangedTableRows.items; //loads all the changed table's rows
        
                                    for (var m = 0; m < tableRows.length; m++) {
        
                                        var rowRangeSorted = newChangedTableRows.getItemAt(m).getRange();
        
                                        var rowValuesSorted = tableRows[m].values;
        
                                        var rowInfoSorted = new Object();
        
                                        for (var name of head[0]) {
                                            theGreatestFunctionEverWritten(head, name, rowValuesSorted, leTable, rowInfoSorted, m);
                                        };

                                        console.log("About to trigger ConForm for Logo Recreation line being removed!");
        
                                        conditionalFormatting(rowInfoSorted, tableStart, changedWorksheet, m, 
                                            completedTableChanged, rowRangeSorted, null);


                                        // console.log("Conditional formatting was applied to row " + m);
        
                                    };
        
                                    eventsOn();
                                    console.log("Events: ON  →  turned on after a row was deleted within the onTableChanged function!");
        
                                    return;

                                };

                            //#endregion -------------------------------------------------------------------------------------------------------------

                            //adjusts picked up / started by turn around time
                            var lePickUpTime = getPickUpTime(rowInfo, leTable, changedRowTableIndex);

                            //adjusts proof to client turn around time
                            var leProofToClientTime = getProofToClientTime(rowInfo, leTable, lePickUpTime, changedRowTableIndex);

                            var changedRowValues = leTable[changedRowTableIndex];

                            if (changedTable.id == unassignedTable.id) { //if changedTable is Unassigned Table...

                                //sorts based on pickedUp column values and assigns priority numbers
                                var sortAndPrioritize = leSorting(rowInfo, leTable, pickedUpColumnIndex, changedRowValues);

                            } else { //if changed table is any other table...

                                //sorts based on proof to client column values and assigns priority numbers
                                var sortAndPrioritize = leSorting(rowInfo, leTable, proofToClientColumnIndex, changedRowValues);

                            };

                            //writes updated values to the table
                            bodyRange.values = sortAndPrioritize; //overwrite changed table data with the new data from the sorted array

                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region MOVE DATA BETWEEN TABLES -----------------------------------------------------------------------------------------------

                        //if either the artist column or the status columns are updated, data is (probably) going to be moved between tables/sheets
                        if (changedColumnIndex == rowInfo.artist.columnIndex || changedColumnIndex == rowInfo.status.columnIndex) {
                            console.log("Here is where all the complex move functions will take place!")

                            //#region ASSIGNS THE DESTINATION TABLE VALUE ----------------------------------------------------------------------------

                                //LEGACY CODE THAT WORKS BUT NEEDS TO BE UPDATED TO BE MORE FLEXIBLE =================================================
                                // ===================================================================================================================

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
                                    } else if (rowInfo.artist.value == "Michael") {
                                        destinationTable = michaelTable;
                                        destinationTableName = michaelTableName.name;
                                        destinationRows = michaelTableRows.items;
                                        destinationTableRange = michaelRange;
                                        destinationHeader = michaelHeader;
                                    } else if (rowInfo.artist.value == "Sarah") {
                                        destinationTable = sarahTable;
                                        destinationTableName = sarahTableName.name;
                                        destinationRows = sarahTableRows.items;
                                        destinationTableRange = sarahRange;
                                        destinationHeader = sarahHeader;
                                    } else if (rowInfo.artist.value == "Danny") {
                                        destinationTable = dannyTable;
                                        destinationTableName = dannyTableName.name;
                                        destinationRows = dannyTableRows.items;
                                        destinationTableRange = dannyRange;
                                        destinationHeader = dannyHeader;
                                    } else if (rowInfo.artist.value == "Ali") {
                                        destinationTable = aliTable;
                                        destinationTableName = aliTableName.name;
                                        destinationRows = aliTableRows.items;
                                        destinationTableRange = aliRange;
                                        destinationHeader = aliHeader;
                                    } else if (rowInfo.artist.value == "JoshK") {
                                        destinationTable = joshKTable;
                                        destinationTableName = joshKTableName.name;
                                        destinationRows = joshKTableRows.items;
                                        destinationTableRange = joshKRange;
                                        destinationHeader = joshKHeader;
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
                                    } else if (rowInfo.artist.value == "Ethan") {
                                        destinationTable = ethanTable;
                                        destinationTableName = ethanTableName.name;
                                        destinationRows = ethanTableRows.items;
                                        destinationTableRange = ethanRange;
                                        destinationHeader = ethanHeader;
                                    } else if (rowInfo.artist.value == "Jessica") {
                                        destinationTable = jessicaTable;
                                        destinationTableName = jessicaTableName.name;
                                        destinationRows = jessicaTableRows.items;
                                        destinationTableRange = jessicaRange;
                                        destinationHeader = jessicaHeader;
                                    } else if (rowInfo.artist.value == "JoshC") {
                                        destinationTable = joshCTable;
                                        destinationTableName = joshCTableName.name;
                                        destinationRows = joshCTableRows.items;
                                        destinationTableRange = joshCRange;
                                        destinationHeader = joshCHeader;
                                    } else if (rowInfo.artist.value == "Emily") {
                                        destinationTable = emilyTable;
                                        destinationTableName = emilyTableName.name;
                                        destinationRows = emilyTableRows.items;
                                        destinationTableRange = emilyRange;
                                        destinationHeader = emilyHeader;
                                    } else if (rowInfo.artist.value == "Alaina") {
                                        destinationTable = alainaTable;
                                        destinationTableName = alainaTableName.name;
                                        destinationRows = alainaTableRows.items;
                                        destinationTableRange = alainaRange;
                                        destinationHeader = alainaHeader;
                                    } else if (rowInfo.artist.value == "Rita") {
                                        destinationTable = ritaTable;
                                        destinationTableName = ritaTableName.name;
                                        destinationRows = ritaTableRows.items;
                                        destinationTableRange = ritaRange;
                                        destinationHeader = ritaHeader;
                                    } else if (rowInfo.artist.value == "Dawn") {
                                        destinationTable = dawnTable;
                                        destinationTableName = dawnTableName.name;
                                        destinationRows = dawnTableRows.items;
                                        destinationTableRange = dawnRange;
                                        destinationHeader = dawnHeader;
                                    } else if (rowInfo.artist.value == "Joey") {
                                        destinationTable = joeyTable;
                                        destinationTableName = joeyTableName.name;
                                        destinationRows = joeyTableRows.items;
                                        destinationTableRange = joeyRange;
                                        destinationHeader = joeyHeader;
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
                                    } else if (rowInfo.artist.value == "Kristen") {
                                        destinationTable = kristenTable;
                                        destinationTableName = kristenTableName.name;
                                        destinationRows = kristenTableRows.items;
                                        destinationTableRange = kristenRange;
                                        destinationHeader = kristenHeader;
                                    } else {
                                        destinationTable = null;
                                        destinationTableName = null;
                                        destinationRows = null;
                                        destinationTableRange = null;
                                        destinationHeader = null;
                                    };

                                    //For the time being, I am recreating the variables from the changed table to work with the destination table.
                                    //I am replacing the changed row index with 0 since, at this point, there is no changed row in the destination 
                                    //table. We just need these values to essentially return the index number of the columns we want from the 
                                    //destination table in future functions.

                                    //if any destination variables are null, do not evaluate further destination variables & objects
                                    if (destinationTable == null 
                                        || destinationTableName == null 
                                        || destinationRows == null 
                                        || destinationTableRange == null 
                                        || destinationHeader == null
                                    ) {

                                        console.log("I actually don't need any of these destination table variables!");

                                    } else { 

                                        var destinationRange = destinationTableRange.values;

                                        if (destinationRows.length == 0) {
                                            var destRowValues = destinationRange;
                                        } else {
                                            var destRowValues = destinationRows[0].values;
                                        };

                                        var destTableName = destinationTableName;

                                        var destTable = JSON.parse(JSON.stringify(destinationRange));

                                        var destHead = destinationHeader.values;

                                        var destRowInfo = new Object();

                                        for (var name of destHead[0]) {
                                            theGreatestFunctionEverWritten(destHead, name, destRowValues, destTable, destRowInfo, 0)
                                        };

                                    };

                                // ===================================================================================================================
                                // ===================================================================================================================

                            //#endregion ---------------------------------------------------------------------------------------------------------

                            //#region FINDING DESTINATION TABLE VARIABLES DYNAMICALLY: BROKEN --------------------------------------------------------
                                
                                //TRYING TO FIGURE OUT CODE THAT CANCELS MOVE BETWEEN TABLES IF HEADERS DON'T MATCH \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                // \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

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

                                    //#endregion -----------------------------------------------------------------------------------------------------


                                    //#region ASSIGNS THE DESTINATION TABLE VALUE --------------------------------------------------------------------


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


                                    //#region CHECK TABLE HEADERS TO SEE IF THEY ARE THE SAME BEFORE MOVING DATA -------------------------------------

                                        // if (destinationTable !== "null" || destinationHeader !== "null") {

                                        //     var headerValues = headerRange.values[0];
                                        //     var destHeaderValues = destinationHeader.values[0];

                                        //     var areHeadersEqual = areArraysEqual(headerValues, destHeaderValues);

                                        //     if (areHeadersEqual == false) {
                                        //       console.log("One of the targeted tables is missing a column, therefore data was not moved.");
                                        //       return;
                                        //     };

                                        //   };

                                    //#endregion -----------------------------------------------------------------------------------------------------

                                // \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                // \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                            //#endregion -------------------------------------------------------------------------------------------------------------

                            //#region CHECK TABLE HEADERS TO SEE IF THEY ARE THE SAME BEFORE MOVING DATA ---------------------------------------------

                                if (destinationTable !== null || destinationHeader !== null) {

                                    var headerValues = headerRange.values[0];
                                    var destHeaderValues = destinationHeader.values[0];

                                    var areHeadersEqual = areArraysEqual(headerValues, destHeaderValues);

                                    if (areHeadersEqual == false) {
                                    console.log("One of the targeted tables is missing a column, therefore data was not moved.");
                                    return;
                                    };

                                };

                            //#endregion -------------------------------------------------------------------------------------------------------------

                            //#region MOVE DATA BASED ON STATUS COLUMN -------------------------------------------------------------------------------

                                if (changedColumnIndex == rowInfo.status.columnIndex) {

                                    //#region MOVE DATA TO COMPLETED TABLE ---------------------------------------------------------------------------

                                        //if status column = "Completed" or "Cancelled", the changedTable is not a Completed table, 
                                        //& the changedWorksheet is not UnassignedProjects, move data to changedWorksheet's completed table
                                        if (
                                            (rowInfo.status.value == "Completed" || rowInfo.status.value == "Cancelled") 
                                            && completedTableChanged == false 
                                            && isUnassigned == false
                                        ) {

                                            //#region UPDATE DATE OF LAST EDIT -----------------------------------------------------------------------

                                                //generate a new date and time based on the current date and time
                                                var dateOfLastEditTime = new Date();
                                                var dateOfLastEditTimeJS = JSDateToExcelDate(dateOfLastEditTime);

                                                //write current date and time to the Date of Last Edit position within the table array
                                                rowValues[0][rowInfo.dateOfLastEdit.columnIndex] = dateOfLastEditTimeJS; 

                                            //#endregion ---------------------------------------------------------------------------------------------

                                            rowValues[0][rowInfo.tags.columnIndex] = ""; //clear tags

                                            //Adds empty row to bottom of the completedTable, then inserts the changed values into this empty row
                                            completedTable.rows.add(0, rowValues);

                                            myRow.delete(); //Deletes the changed row from the original sheet
                                            console.log("Data was moved to the artist's Completed Projects Table!");

                                            leTable.splice(changedRowTableIndex, 1); //removes changed row from table content array

                                            //sorts the artist table by proof to client
                                            var leTableSort = leSorting(rowInfo, leTable, proofToClientColumnIndex, rowValues[0]);

                                            //reload artist tables values after deleting a row
                                            var bodyRangeReload = changedTable.getDataBodyRange().load("values");

                                            var newCompletedRows = completedTable.rows.load("items");

                                            var newComplatedBodyValues = completedTable.getDataBodyRange().load("values");


                                            await context.sync();


                                            bodyRangeReload.values = leTableSort; //writes sorted table content from the array to the artist table

                                            var newCompletedTableValues = newComplatedBodyValues.values

                                            //#region REMOVE CONDITIONAL FORMATTING FROM COMPLETED TABLE ---------------------------------------------

                                                //removes conditional formatting from completed table projects
                                                for (var m = 0; m < newCompletedTableValues.length; m++) {

                                                    var rowRangeSortedCompleted = newCompletedRows.getItemAt(m).getRange();

                                                    rowRangeSortedCompleted.format.fill.clear();
                                                    rowRangeSortedCompleted.format.font.color = "black";
                                                    rowRangeSortedCompleted.format.font.bold = false;

                                                };

                                            //#endregion ---------------------------------------------------------------------------------------------

                                    //#endregion -----------------------------------------------------------------------------------------------------

                                    //#region MOVE DATA FROM COMPLETED TO ARTIST TABLE ---------------------------------------------------------------

                                        //if status column = "Editing" & the changedTable is a Completed table, move data back to the artist's table
                                        } else if ( 
                                            (rowInfo.status.value == "Light Changes" && completedTableChanged == true) 
                                            || (rowInfo.status.value == "Moderate Changes" && completedTableChanged == true) 
                                            || (rowInfo.status.value == "Heavy Changes" && completedTableChanged == true)
                                        ) {

                                            if (destinationTable !== "null") { //as long as there is a destination value, do the following

                                                myRow.delete(); //Deletes the changed row from the original sheet
                                                destinationTable.rows.add(null); //adds a blank row to the destination artist table

                                                rowValues[0][rowInfo.tags.columnIndex] = ""; //clear tags

                                                //moves the data between the two table array, to be written to the sheet later
                                                moveDataTwo(destTable, rowValues, leTable, changedRowTableIndex);

                                                //#region NOT SURE WHAT THESE DO BUT I DON'T WANT TO REMOVE THEM -------------------------------------

                                                    rowIndexInDestTable = destTable.length - 1;

                                                    //adjusts proof to client turn around time
                                                    var destProofToClientTime = getProofToClientTime(rowInfo, destTable, 97, rowIndexInDestTable); 
                                                    //since this will only ever trigger the part of the function that references light, 
                                                    //moderate, and heavy changes, the pick up time value is unneeded for the most part. 
                                                    //Therefore, the random number 97 is inserted to take up its spot, and to make sure 
                                                    //the first if statement passes every time.

                                                //#endregion -----------------------------------------------------------------------------------------

                                                //sorts the artist table array
                                                var destTableSort = leSorting(destRowInfo, destTable, proofToClientColumnIndex, rowValues[0]);

                                                //#region RELOAD THE DATA BODY RANGE FOR EACH ARTIST TABLE -------------------------------------------

                                                    var unassignedRange = unassignedTable.getDataBodyRange().load("values");
                                                    var peterRange = peterTable.getDataBodyRange().load("values");
                                                    var mattRange = mattTable.getDataBodyRange().load("values");
                                                    var michaelRange = michaelTable.getDataBodyRange().load("values");
                                                    var sarahRange = sarahTable.getDataBodyRange().load("values");
                                                    var dannyRange = dannyTable.getDataBodyRange().load("values");
                                                    var aliRange = aliTable.getDataBodyRange().load("values");
                                                    var joshKRange = joshKTable.getDataBodyRange().load("values");
                                                    var lukeRange = lukeTable.getDataBodyRange().load("values");
                                                    var breBRange = breBTable.getDataBodyRange().load("values");
                                                    var ethanRange = ethanTable.getDataBodyRange().load("values");
                                                    var jessicaRange = jessicaTable.getDataBodyRange().load("values");
                                                    var joshCRange = joshCTable.getDataBodyRange().load("values");
                                                    var emilyRange = emilyTable.getDataBodyRange().load("values");
                                                    var alainaRange = alainaTable.getDataBodyRange().load("values");
                                                    var ritaRange = ritaTable.getDataBodyRange().load("values");
                                                    var dawnRange = dawnTable.getDataBodyRange().load("values");
                                                    var joeyRange = joeyTable.getDataBodyRange().load("values");
                                                    var jordanRange = jordanTable.getDataBodyRange().load("values");
                                                    var toddRange = toddTable.getDataBodyRange().load("values");
                                                    var kristenRange = kristenTable.getDataBodyRange().load("values");

                                                //#endregion -----------------------------------------------------------------------------------------

                                                //#region ASSIGNS THE UPDATED ARTIST TABLE RANGE AS THE DESTINATION ----------------------------------

                                                    if (rowInfo.artist.value == "Unassigned" && isUnassigned == false) {
                                                        var destinationStation = unassignedRange;
                                                    } else if (rowInfo.artist.value == "Peter") {
                                                        var destinationStation = peterRange;
                                                    } else if (rowInfo.artist.value == "Matt") {
                                                        var destinationStation = mattRange;
                                                    } else if (rowInfo.artist.value == "Michael") {
                                                        var destinationStation = michaelRange;
                                                    } else if (rowInfo.artist.value == "Sarah") {
                                                        var destinationStation = sarahRange;
                                                    } else if (rowInfo.artist.value == "Danny") {
                                                        var destinationStation = dannyRange;
                                                    } else if (rowInfo.artist.value == "Ali") {
                                                        var destinationStation = aliRange;
                                                    } else if (rowInfo.artist.value == "JoshK") {
                                                        var destinationStation = joshKRange;
                                                    } else if (rowInfo.artist.value == "Luke") {
                                                        var destinationStation = lukeRange;
                                                    } else if (rowInfo.artist.value == "Bre B.") {
                                                        var destinationStation = breBRange;
                                                    } else if (rowInfo.artist.value == "Ethan") {
                                                        var destinationStation = ethanRange;
                                                    } else if (rowInfo.artist.value == "Jessica") {
                                                        var destinationStation = jessicaRange;
                                                    } else if (rowInfo.artist.value == "JoshC") {
                                                        var destinationStation = joshCRange;
                                                    } else if (rowInfo.artist.value == "Emily") {
                                                        var destinationStation = emilyRange;
                                                    } else if (rowInfo.artist.value == "Alaina") {
                                                        var destinationStation = alainaRange;
                                                    } else if (rowInfo.artist.value == "Rita") {
                                                        var destinationStation = ritaRange;
                                                    } else if (rowInfo.artist.value == "Dawn") {
                                                        var destinationStation = dawnRange;
                                                    } else if (rowInfo.artist.value == "Joey") {
                                                        var destinationStation = joeyRange;
                                                    } else if (rowInfo.artist.value == "Jordan") {
                                                        var destinationStation = jordanRange;
                                                    } else if (rowInfo.artist.value == "Todd") {
                                                        var destinationStation = toddRange;
                                                    } else if (rowInfo.artist.value == "Kristen") {
                                                        var destinationStation = kristenRange;
                                                    } else {
                                                        var destinationStation = "null";
                                                    };

                                                //#endregion -----------------------------------------------------------------------------------------

                                                await context.sync();

                                                destinationStation.values = destTableSort; //writes artist table values to the worksheet

                                            };
                                        };

                                    //#endregion -----------------------------------------------------------------------------------------------------

                                };

                            //#endregion -------------------------------------------------------------------------------------------------------------

                            //#region MOVE DATA BASED ON ARTIST COLUMN -------------------------------------------------------------------------------

                                if (changedColumnIndex == rowInfo.artist.columnIndex) {

                                    //#region MOVES DATA TO DESTINATION TABLE ------------------------------------------------------------------------

                                        if (destinationTable !== "null") {
                                            //if destination table is not in the same worksheet as the changedTable (prevents for unnecessary 
                                            //moving of data across tables in the same worksheet), do the following...
                                            if (destinationTable.worksheet.id !== changedWorksheet.id) {

                                                //#region SETS STATUS AUTOFILL -----------------------------------------------------------------------

                                                    var newStatus = statusAutofill(destTableName); //updates the status to default autofill value
                                                    rowValues[0][rowInfo.status.columnIndex] = newStatus;

                                                //#endregion -----------------------------------------------------------------------------------------

                                                //#region CLEAR TAGS ---------------------------------------------------------------------------------

                                                    rowValues[0][rowInfo.tags.columnIndex] = ""; //clear tags

                                                //#endregion -----------------------------------------------------------------------------------------

                                                //#region MOVES DATA BETWEEN TABLE ARRAYS ------------------------------------------------------------

                                                    //moves data from leTable array to destTable array, whihc will be written to workbook later
                                                    moveDataTwo(destTable, rowValues, leTable, changedRowTableIndex);

                                                    //#region PREVENT CODE FROM ERRORING BECAUSE TABLE LENGTH !== DESTTABLE ARRAY LENGTH -------------

                                                        //removes empty row from destTable array if the destination body range is 0
                                                        if (destinationRows.length == 0) {
                                                            destTable.shift();
                                                        };

                                                    //#endregion -------------------------------------------------------------------------------------

                                                //#endregion -----------------------------------------------------------------------------------------

                                                //#region SORTS TABLE ARRAYS -------------------------------------------------------------------------

                                                    //if data is moving from the unassigned table to an artist table, sort this way...
                                                    if (changedTable.id == unassignedTable.id) {

                                                        //sorts the changed unassigned table by picked up / started by
                                                        var leTableSort = leSorting(rowInfo, leTable, pickedUpColumnIndex, rowValues[0]);

                                                        //sorts the destination artist table by proof to client
                                                        var destTableSort = leSorting(destRowInfo, destTable, proofToClientColumnIndex, rowValues[0]); 

                                                    //if data is moving from an artist table to the unassigned table, sort this way...
                                                    } else if (destinationTable.id == unassignedTable.id) {

                                                        //sorts the changed artist table by proof to client
                                                        var leTableSort = leSorting(rowInfo, leTable, proofToClientColumnIndex, rowValues[0]);

                                                        //sorts the destination Unassigned table by picked up / started by
                                                        var destTableSort = leSorting(destRowInfo, destTable, pickedUpColumnIndex, rowValues[0]);

                                                    //if data is moving between artist tables, both will be sorted by proof to client
                                                    } else if (
                                                        (destinationTable.id !== unassignedTable.id) 
                                                        && (changedTable.id !== unassignedTable.id)
                                                    ) {

                                                        //sorts the changed artist table by proof to client
                                                        var leTableSort = leSorting(rowInfo, leTable, proofToClientColumnIndex, rowValues[0]);

                                                        //sorts the destination arist table by proof to client
                                                        var destTableSort = leSorting(destRowInfo, destTable, proofToClientColumnIndex, rowValues[0]); 
                                                    };

                                                //#endregion -----------------------------------------------------------------------------------------

                                                //#region ADJUSTS TABLE RANGES TO RECIEVE THE DATA ---------------------------------------------------

                                                    myRow.delete();

                                                    destinationTable.rows.add(null);

                                                    //#region RELOAD ARTIST TABLES BODY RANGE VALUES -------------------------------------------------

                                                        var bodyPositivity = changedTable.getDataBodyRange().load("values");

                                                        var unassignedRange = unassignedTable.getDataBodyRange().load("values");
                                                        var peterRange = peterTable.getDataBodyRange().load("values");
                                                        var mattRange = mattTable.getDataBodyRange().load("values");
                                                        var michaelRange = michaelTable.getDataBodyRange().load("values");
                                                        var sarahRange = sarahTable.getDataBodyRange().load("values");
                                                        var dannyRange = dannyTable.getDataBodyRange().load("values");
                                                        var aliRange = aliTable.getDataBodyRange().load("values");
                                                        var joshKRange = joshKTable.getDataBodyRange().load("values");
                                                        var lukeRange = lukeTable.getDataBodyRange().load("values");
                                                        var breBRange = breBTable.getDataBodyRange().load("values");
                                                        var ethanRange = ethanTable.getDataBodyRange().load("values");
                                                        var jessicaRange = jessicaTable.getDataBodyRange().load("values");
                                                        var joshCRange = joshCTable.getDataBodyRange().load("values");
                                                        var emilyRange = emilyTable.getDataBodyRange().load("values");
                                                        var alainaRange = alainaTable.getDataBodyRange().load("values");
                                                        var ritaRange = ritaTable.getDataBodyRange().load("values");
                                                        var dawnRange = dawnTable.getDataBodyRange().load("values");
                                                        var joeyRange = joeyTable.getDataBodyRange().load("values");
                                                        var jordanRange = jordanTable.getDataBodyRange().load("values");
                                                        var toddRange = toddTable.getDataBodyRange().load("values");
                                                        var kristenRange = kristenTable.getDataBodyRange().load("values");

                                                    //#endregion -------------------------------------------------------------------------------------

                                                    //#region ASSIGNS THE UPDATED ARTIST RANGE AS THE DESTINATION ------------------------------------

                                                        if (rowInfo.artist.value == "Unassigned" && isUnassigned == false) {
                                                            var destinationStation = unassignedRange;
                                                        } else if (rowInfo.artist.value == "Peter") {
                                                            var destinationStation = peterRange;
                                                        } else if (rowInfo.artist.value == "Matt") {
                                                            var destinationStation = mattRange;
                                                        } else if (rowInfo.artist.value == "Michael") {
                                                            var destinationStation = michaelRange;
                                                        } else if (rowInfo.artist.value == "Sarah") {
                                                            var destinationStation = sarahRange;
                                                        } else if (rowInfo.artist.value == "Danny") {
                                                            var destinationStation = dannyRange;
                                                        } else if (rowInfo.artist.value == "Ali") {
                                                            var destinationStation = aliRange;
                                                        } else if (rowInfo.artist.value == "JoshK") {
                                                            var destinationStation = joshKRange;
                                                        } else if (rowInfo.artist.value == "Luke") {
                                                            var destinationStation = lukeRange;
                                                        } else if (rowInfo.artist.value == "Bre B.") {
                                                            var destinationStation = breBRange;
                                                        } else if (rowInfo.artist.value == "Ethan") {
                                                            var destinationStation = ethanRange;
                                                        } else if (rowInfo.artist.value == "Jessica") {
                                                            var destinationStation = jessicaRange;
                                                        } else if (rowInfo.artist.value == "JoshC") {
                                                            var destinationStation = joshCRange;
                                                        } else if (rowInfo.artist.value == "Emily") {
                                                            var destinationStation = emilyRange;
                                                        } else if (rowInfo.artist.value == "Alaina") {
                                                            var destinationStation = alainaRange;
                                                        } else if (rowInfo.artist.value == "Rita") {
                                                            var destinationStation = ritaRange;
                                                        } else if (rowInfo.artist.value == "Dawn") {
                                                            var destinationStation = dawnRange;
                                                        } else if (rowInfo.artist.value == "Joey") {
                                                            var destinationStation = joeyRange;
                                                        } else if (rowInfo.artist.value == "Jordan") {
                                                            var destinationStation = jordanRange;
                                                        } else if (rowInfo.artist.value == "Todd") {
                                                            var destinationStation = toddRange;
                                                        } else if (rowInfo.artist.value == "Kristen") {
                                                            var destinationStation = kristenRange;
                                                        } else {
                                                            var destinationStation = "null";
                                                        };

                                                    //#endregion -------------------------------------------------------------------------------------

                                                //#endregion -----------------------------------------------------------------------------------------

                                                await context.sync()

                                                    //#region LOAD NEWLY CHANGED BODY RANGE AND DESTINATION BODY RANGE -------------------------------

                                                        var newBodyRange = bodyPositivity.values;
                                                        var newDestinationTableRange = destinationStation.values;

                                                    //#endregion -------------------------------------------------------------------------------------

                                                    //#region PREVENT CODE FROM ERRORING BECAUSE TABLE LENGTH !== LETABLE ARRAY LENGTH ---------------

                                                        //removes empty row from leTable array if the leTable body range is now 0
                                                        if (leTable.length == 0) {
                                                            newBodyRange.shift();
                                                        };

                                                    //#endregion -------------------------------------------------------------------------------------

                                                    //#region WRTIE NEW VALUES TO THE UPDATED TABLES -------------------------------------------------

                                                        newBodyValues = leTableSort;

                                                        destinationStation.values = destTableSort;

                                                    //#endregion -------------------------------------------------------------------------------------

                                            };
                                        } else {
                                            console.log("No artist was assigned or updated, so no data was moved.")
                                            return;
                                        };

                                    //#endregion -----------------------------------------------------------------------------------------------------

                                };

                            //#endregion -------------------------------------------------------------------------------------------------------------

                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region CONDITIONAL FORMATTING HANDLER -----------------------------------------------------------------------------------------

                        //only do the following if the change was not made to a Print Date or Group column
                        if (
                                (
                                    //change not made to Print Date column or Group column
                                    (changedColumnIndex !== rowInfo.printDate.columnIndex && changedColumnIndex !== rowInfo.group.columnIndex) 
                                    || //OR
                                    (
                                        //change is made to Print Date or Group columns AND...
                                        (changedColumnIndex == rowInfo.printDate.columnIndex || changedColumnIndex !== rowInfo.group.columnIndex) 
                                        && (
                                            rowInfo.product.value == "Logo Recreation" //...& is a logo recreation product AND...
                                            && (
                                                rowInfo.status.value == "Logo Status TBD" || rowInfo.status.value == "Logo Needes Recreating"
                                                || rowInfo.status.value == "Logo Needs Uploading"
                                            )
                                        ) //...has a logo recreation status
                                    )
                                )
                            ) {

                            //Applys different values to the variables if data is moving between tables (more details below):

                            //if the changed column was the artist column OR the status column with these conditions:
                                //status is either Completed or Cancelled and is NOT in a Completed Table
                                //status is a change and is from a Completed Table
                            if (
                                (changedColumnIndex == rowInfo.artist.columnIndex) || (changedColumnIndex == rowInfo.status.columnIndex 
                                    && (
                                        (
                                            (rowInfo.status.value == "Completed" && completedTableChanged == false) 
                                            || (rowInfo.status.value == "Cancelled" && completedTableChanged == false)
                                        )
                                    ||
                                        (
                                            (rowInfo.status.value == "Light Changes" && completedTableChanged == true) 
                                            || (rowInfo.status.value == "Moderate Changes" && completedTableChanged == true) 
                                            || (rowInfo.status.value == "Heavy Changes" && completedTableChanged == true)
                                        )
                                    )
                                )
                            ) {

                                var newChangedTableRows = destinationTable.rows.load("items");

                                var newBodyValues = destinationTable.getDataBodyRange().load("values");

                                var destinationWorksheetId = destinationTable.worksheet.id;

                                var newChangedWorksheet = context.workbook.worksheets.getItem(destinationWorksheetId).load("name");

                                var newStartOfTable = destinationTable.getRange().load("columnIndex");

                            } else { //data is not moving to another table, so no need for destination variables

                                var newChangedTableRows = changedTable.rows.load("items");

                                var newBodyValues = changedTable.getDataBodyRange().load("values");

                                var newChangedWorksheet = changedWorksheet;

                                var newStartOfTable = startOfTable;

                            };

                            await context.sync();

                            var leTableSorted = newBodyValues.values

                            var tableRowsSorted = newChangedTableRows.items;

                            var newTableStart = newStartOfTable.columnIndex; //column index of the start of the table

                            //applies conditional formatting to each row of the table
                            for (var m = 0; m < leTableSorted.length; m++) {

                                var rowRangeSorted = newChangedTableRows.getItemAt(m).getRange();

                                var rowValuesSorted = tableRowsSorted[m].values;

                                var rowInfoSorted = new Object();

                                for (var name of head[0]) {
                                    theGreatestFunctionEverWritten(head, name, rowValuesSorted, leTableSorted, rowInfoSorted, m);
                                };

                                console.log("ConForm is about to trigger to recolor the whole sheet!");
                                conditionalFormatting(rowInfoSorted, newTableStart, newChangedWorksheet, m, 
                                    completedTableChanged, rowRangeSorted, destTable);
                            };
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                eventsOn(); //turns events back on
                console.log("Events: ON  →  turned on at the end of the onTableChanged Function!");

            }).catch (err => { //error catcher
                if (dennisHere == true) { //if row was inserted illegally, do not return error
                    return;
                };
                console.log(err) // <--- does this log?
                showMessage(err, "show");
                context.runtime.enableEvents = true;
            });

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------------------------------------------


//#region FUNCTIONS ----------------------------------------------------------------------------------------------------------------------------------

    //#region ACTIVATION HANDLERS --------------------------------------------------------------------------------------------------------------------

        //#region REGISTER ON WORKSHEET ACTIVATE HANDLER & BINDS OTHER HANDLERS TO CURRENT SHEET -----------------------------------------------------

            /**
             * Enables onSelectionChanged event upon inital load and any reloads of the taskpane
             */
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

                                    console.log("ConForm is about to trigger for the activition handler (so for the taskpane reloading");
                                    conditionalFormatting(rowInfoSorted, tableStart, theWorksheet, iRow, completedTableChanged, rowRange, null)
                                };
                            };
                        };
                    };

                    for (var y = 0; y < worksheetTablesCount; y++) {
                        var bonTable = worksheetTables.getItemAt(y);

                        changeEvent = bonTable.onChanged.add(onTableChanged);

                        selectionEvent = bonTable.onSelectionChanged.add(onTableSelectionChangedEvents);

                    };

                    sheets.onActivated.add(onActivate);
                    sheets.onDeactivated.add(onDeactivate);

                    console.log("A handler has been registered for the OnActivate event.");

                    eventsOn();
                    console.log("Events: ON  →  turned on in the registerOnActivateHandler function, typically triggered by a reload");

                }).catch (err => {
                    console.log(err) // <--- does this log?
                    showMessage(err, "show");
                    context.runtime.enableEvents = true;
                });
            };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region ON NEW WORKSHEET ACTIVATE ----------------------------------------------------------------------------------------------------------

            /**
             * When the worksheet changes, this fires and binds the events to the first table that is selected in the sheet
             * @param {Object} eventArgs The event arguments, which are details about the event that was triggered
             */
            async function onActivate(eventArgs) {
                await Excel.run(async (context) => {

                    // console.log("Worksheet change Selection Event: ");
                    // console.log(selectionEvent);
                    // removeSelectionEvent();
                    console.log("Source of the onActivate event: " + eventArgs.source);

                    console.log("Worksheet Switched (onActivate) function fired");

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

                        //ignore all tables in Validation sheet cycles through each row in the table
                        if (oldTablesWorksheet.name !== "Validation" && oldUsedDataRange.isNullObject !== true) { 
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

                                    console.log("ConForm is about to trigger for the whole sheet because the user moved to a new worksheet");
                                    conditionalFormatting(oldRowInfoSorted, oldTableStart, theOldWorksheet, 
                                        aRow, oldCompletedTableChanged, oldRowRange, null)
                                };
                            };
                        };
                    };

                    location.reload();

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

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region DEACTIVATION EVENT -----------------------------------------------------------------------------------------------------------------

            /**
             * Simply writes the worksheetId to a global deactivatedWorksheetId variable to be used in other functions
             * @param {Object} eventArgs The event arguments, which are details about the event that was triggered
             */
            async function onDeactivate(eventArgs) {
                await Excel.run(async(context) => {
                    console.log("Source of the onDeactivate event: " + eventArgs.source);

                    console.log("The worksheet Id that was deactivated was: " + eventArgs.worksheetId);

                    deactivatedWorksheetId = eventArgs.worksheetId
                });
            };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region AUTO-FILLING ---------------------------------------------------------------------------------------------------------------------------
        
        //#region AUTOFILL STATUS COLUMN -------------------------------------------------------------------------------------------------------------

            /**
             * If the table name is "UnassignedProjects", sets status to "Awaiting Artist", otherwise sets status to "Not Working"
             * @param {String} tableName The name of the table
             * @returns String
             */    
             function statusAutofill(tableName) {
                //if the table the row was inserted into is "UnassignedProjects", set status column to "Awaiting Artist"
                if (tableName == "UnassignedProjects") {
                    var status = "Awaiting Artist";
                };
                //if the table the row was inserted into is not "UnassaignedProjects"...
                if (tableName !== "UnassignedProjects") {
                    var status = "Not Working";
                };
                return status;
            };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region AUTOFILL ARTIST COLUMN -------------------------------------------------------------------------------------------------------------

            /**
             * If table is "UnassignedProjects", sets artist to "Unassigned", otherwise sets artist to the name of the worksheet
             * @param {String} tableName The name of the table
             * @param {String} leSheetName The name of the worksheet
             * @returns String
             */
            function artistAutofill(tableName, leSheetName) {
                if (tableName == "UnassignedProjects") {
                    var artist = "Unassigned";
                } else if (tableName !== "UnassignedProjects") {
                    var artist = leSheetName;
                };
                return artist;
            };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region THE GREATEST FUNCTION EVER WRITTEN (AND GUEST) -----------------------------------------------------------------------------------------

        //#region INDEX & VALUES OF CHANGED ROW (THE GREATEST FUNCTION EVER WRITTEN) -----------------------------------------------------------------

            /**
             * Using the column names, finds and writes the column index and value of each cell in the changed row to an object. 
             * Also updates a copy of the header array with the values of the changed row in the correct column indexed positions
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

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region HEADERS TO CODE --------------------------------------------------------------------------------------------------------------------

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

        //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region TURN AROUND TIME FUNCTIONS, SORTING, & OFFICE HOURS ------------------------------------------------------------------------------------

        //#region ADJUST PICKED UP / STARTED BY TURN AROUND TIME -------------------------------------------------------------------------------------

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

                //get the Project Type coded variable from the Project Type ID Data based on the value in Project Type column of the changed row
                var theProjectTypeCode = projectTypeIDData[rowInfo.projectType.value].projectTypeCode;

                //returns turn around time value from the Pickup Turn Around Time table based on the Product column of the 
                //changed row and the projetc type codeed variable
                var pickedUpTurnAroundTime = pickupData[rowInfo.product.value][theProjectTypeCode];

                if (rowInfo.status.value == "Light Changes" || rowInfo.status.value == "Moderate Changes" || rowInfo.status.value == "Heavy Changes") {
                    //do not add start override to picked up turn around time
                    var pickedUpHours = pickedUpTurnAroundTime;
                } else {
                    //finds the start override value of the changed row and adds it to the previous turn around time variable
                    var pickedUpHours = pickedUpTurnAroundTime + rowInfo.startOverride.value;
                };

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

    //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region ADJUST PROOF TO CLIENT TURN AROUND TIME --------------------------------------------------------------------------------------------

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

                //if the request's status is a form of "Editing"...
                if (
                    rowInfo.status.value == "Light Changes" 
                    || rowInfo.status.value == "Moderate Changes" 
                    || rowInfo.status.value == "Heavy Changes"
                ) {

                    //get the Changes coded variable from the Changes ID Data based on the value in the Status column of the changed row
                    var theChangesCode = changesIDData[rowInfo.status.value].changesCode;
                    // var snailsss = theChangesCode.changesCode;

                    //gets the turn around time value from the Changes Data Table based on the product and status of the row
                    var proofToClient = changesData[rowInfo.product.value][theChangesCode];

                    if (rowInfo.status.value == "Light Changes" || rowInfo.status.value == "Moderate Changes" || rowInfo.status.value == "Heavy Changes") {
                        //do not add work override to proof to client turn around time
                        var artTurnAround = proofToClient;
                    } else {
                        //finds the work override value of the changed row and adds it to the proofToClient variable
                        var artTurnAround = proofToClient + rowInfo.workOverride.value;
                    };

                    //#region UPDATE DATE OF LAST EDIT -----------------------------------------------------------------------------------------------

                        //generate a new date and time based on the current date and time
                        var dateOfLastEditTime = new Date();
                        var dateOfLastEditTimeJS = JSDateToExcelDate(dateOfLastEditTime);

                        //write current date and time to the Date of Last Edit position within the table array
                        leTable[rowIndex][rowInfo.dateOfLastEdit.columnIndex] = dateOfLastEditTimeJS;

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //adds the adjusted turn around time to new Date of Last Edit date that was just generated and adjusts to be within office hours
                    var proofToClientOfficeHours = officeHours(dateOfLastEditTime, artTurnAround);

                    //converts date to excel date
                    var excelProofToClientOfficeHours = Number(JSDateToExcelDate(proofToClientOfficeHours));

                    //updates the proof to client turn around time value in the table array based on our calculations
                    leTable[rowIndex][rowInfo.proofToClient.columnIndex] = excelProofToClientOfficeHours;

                    return proofToClientOfficeHours;

               
                //#region IF PRODUCT IS A LOGO RECREATION AND HAS A LOGO SPECIFIFC STATUS, UPDATE TURN AROUND TIME -----------------------------------

                    //if product is a logo recreation AND it has any of the three main recreation statuses (used spcifically for 
                    //when a logo is being recreated based on another project), do the following...
                    } else if (rowInfo.product.value == "Logo Recreation" && (rowInfo.status.value == "Logo Status TBD" 
                    || rowInfo.status.value == "Logo Needs Recreating" || rowInfo.status.value == "Logo Needs Uploading")) {

                        var thePrintDateUpdated = convertToDate(rowInfo.printDate.value)
                        var updatedPrintDate = new Date(thePrintDateUpdated);
                        var updatedDatePrint = updatedPrintDate.getDate();
                        var updatedMonthPrint = updatedPrintDate.getMonth();
                        var updatedYearPrint = updatedPrintDate.getFullYear();

                        var updatedPrintDateDOW = updatedPrintDate.getDay();

                        if (updatedPrintDateDOW == 0) {
                            updatedPrintDateDOW = "Sunday"
                        } else if (updatedPrintDateDOW == 1) {
                            updatedPrintDateDOW = "Monday"
                        } else if (updatedPrintDateDOW == 2) {
                            updatedPrintDateDOW = "Tuesday"
                        } else if (updatedPrintDateDOW == 3) {
                            updatedPrintDateDOW = "Wednesday"
                        } else if (updatedPrintDateDOW == 4) {
                            updatedPrintDateDOW = "Thursday"
                        } else if (updatedPrintDateDOW == 5) {
                            updatedPrintDateDOW = "Friday"
                        } else if (updatedPrintDateDOW == 6) {
                            updatedPrintDateDOW = "Saturday"
                        };

                        var updatedWeekdayVars = officeHoursData[updatedPrintDateDOW];

                        var updatedGroupEndOfDay = convertToDate(updatedWeekdayVars.endTime);

                        updatedGroupEndOfDay.setFullYear(updatedYearPrint);
                        updatedGroupEndOfDay.setMonth(updatedMonthPrint);
                        updatedGroupEndOfDay.setDate(updatedDatePrint);

                        console.log(updatedGroupEndOfDay);

                        var updatedGroupDateExcel = Number(JSDateToExcelDate(updatedGroupEndOfDay));

                        leTable[rowIndex][rowInfo.proofToClient.columnIndex] = updatedGroupDateExcel;
                        
                        return updatedGroupDateExcel;

                //#endregion -------------------------------------------------------------------------------------------------------------------------


                    } else {


                        //get the Project Type coded variable from the Project Type ID Data based on value in the Project Type column of changed row
                        var theProjectTypeCode = projectTypeIDData[rowInfo.projectType.value].projectTypeCode;

                        //returns turn around time value from the Proof to Client Turn Around Time table based on the Product column of the 
                        //changed row and the projetc type codeed variable
                        var proofToClient = proofToClientData[rowInfo.product.value][theProjectTypeCode];

                        //returns creative review process hours adjustment number from thhe creative review table based on 
                        //the Product column value of the changed row
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
            };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region SORTING THE TABLE BY PICKED UP TURN AROUND TIME ------------------------------------------------------------------------------------

            /**
             * Sorts the table array by the values in leColumnIndex and then assigns updated priority numbers
             * @param {Object} rowInfo An object containing the values and column indexs of each cell in the changed row
             * @param {Array} leTable An array of arrays containing all the info of the changed table
             * @param {Number} leColumnIndex The index number of the column that will be used for sorting the table
             * @param {Array} changedRowValues The values of the changed row
             * @returns Array
             */
            function leSorting(rowInfo, leTable, leColumnIndex, changedRowValues) {

                //#region ASSIGNING VARIABLES --------------------------------------------------------------------------------------------------------

                    //a copy of the array containing all the table data that will be used for sorting
                    var leTableSorted = JSON.parse(JSON.stringify(leTable)); //creates a duplicate of original array to be used for 
                    //assigning the priority numbers, without having anything done to it affect oriignal array

                    var priorityColumnIndex = rowInfo.priority.columnIndex; //index of priority column

                    var pickedUpColumnIndex = rowInfo.pickedUpStartedBy.columnIndex;
                    var proofToClientColumnIndex = rowInfo.proofToClient.columnIndex;

                    var statusColumnIndex = rowInfo.status.columnIndex;
                    
                    var tagsColumnIndex = rowInfo.tags.columnIndex;

                    var tempTable = [];
                    var tagOrder = [0, 0, 0, 0, 0, 0, 0, 0];
                    var tagItems = [0, 0, 0, 0, 0, 0, 0, 0];
                    var veryUrgentArr = [];
                    var urgentArr = [];
                    var semiUrgentArr = [];
                    var eventualArr = [];
                    var onHoldTable = [];
                    var awaitingChangesTable = [];

                //#endregion -------------------------------------------------------------------------------------------------------------------------


                //#region WITHHOLDS ITEMS NOT TO BE SORTED FROM TABLE ARRAY ----------------------------------------------------------------------------

                    for (var i = 0; i < leTableSorted.length; i++) { //for each row in the table...

                        //removes invlaid requests from table and puts them in a temp table to be added back in after sorting
                        if (
                            leTableSorted[i][pickedUpColumnIndex] == "NO PRODUCT / PROJECT TYPE" 
                            || leTableSorted[i][proofToClientColumnIndex] == "NO PRODUCT / PROJECT TYPE"
                        ) {
                            tempTable.push(leTableSorted[i]);
                            leTableSorted.splice(i, 1);
                            i = i - 1;
                        //removes on hold requests from table and puts them in an on hold table to be added back in after sorting
                        } else if (
                            leTableSorted[i][statusColumnIndex] == "On Hold"
                            ) { 
                            onHoldTable.push(leTableSorted[i]);
                            leTableSorted.splice(i, 1);
                            i = i - 1;
                        //removes awaiting changes requests from table and puts them in an awaiting changes table to be added back in after sorting
                        } else if (
                            leTableSorted[i][statusColumnIndex] == "At Client" 
                            || leTableSorted[i][statusColumnIndex] == "In Review" 
                            || leTableSorted[i][statusColumnIndex] == "Waiting On Info"
                            ) {
                            awaitingChangesTable.push(leTableSorted[i]);
                            leTableSorted.splice(i, 1);
                            i = i - 1;
                        };
                    };


                    for (var i = 0; i < leTableSorted.length; i++) {

                        //removes tagged projects from table and puts them in order in the tagsOrder array to be added back in after sorting
                        if (leTableSorted[i][tagsColumnIndex] == "VERY URGENT") {
                            veryUrgentArr.push(leTableSorted[i]);
                            leTableSorted.splice(i, 1);
                            i = i - 1;
                        } else if (leTableSorted[i][tagsColumnIndex] == "URGENT") {
                            urgentArr.push(leTableSorted[i]);
                            leTableSorted.splice(i, 1);
                            i = i - 1;
                        } else if (leTableSorted[i][tagsColumnIndex] == "SEMI-URGENT") {
                            semiUrgentArr.push(leTableSorted[i]);
                            leTableSorted.splice(i, 1);
                            i = i - 1;
                        } else if (leTableSorted[i][tagsColumnIndex] == "EVENTUAL") {
                            eventualArr.push(leTableSorted[i]);
                            leTableSorted.splice(i, 1);
                            i = i - 1;
                        };
                        
                    };

                //#endregion -------------------------------------------------------------------------------------------------------------------------


                //#region SORTS THE ARRAYS CONTAINING THE TAGGED ITEMS -------------------------------------------------------------------------------
                    if (veryUrgentArr.length > 1) {
                        veryUrgentArr.sort((a,b) => (a[leColumnIndex] > b[leColumnIndex]) ? 1 : -1);
                    };

                    if (urgentArr.length > 1) {
                        urgentArr.sort((a,b) => (a[leColumnIndex] > b[leColumnIndex]) ? 1 : -1);
                    };

                    if (semiUrgentArr.length > 1) {
                        semiUrgentArr.sort((a,b) => (a[leColumnIndex] > b[leColumnIndex]) ? 1 : -1);
                    };

                    if (eventualArr.length > 1) {
                        eventualArr.sort((a,b) => (a[leColumnIndex] > b[leColumnIndex]) ? 1 : -1);
                    };

                //#endregion -------------------------------------------------------------------------------------------------------------------------


                //#region SORTS THE TABLE ARRAY ------------------------------------------------------------------------------------------------------

                    //sorts the parent array (a) by the number in the sub array (b) at index of the picked up column
                    //leTableSorted.sort(function(a,b){return a[leColumnIndex] > b[leColumnIndex]});
                    leTableSorted.sort((a, b) => (a[leColumnIndex] > b[leColumnIndex]) ? 1 : -1); //sorts

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region ADDS THE WITHHELD AND SORTED TAGGED ITEMS BACK INTO THE TABLE --------------------------------------------------------------

                    var veryUrgentLength;

                    for (var z = 0; z < veryUrgentArr.length; z++) {
                        // if (veryUrgentArr[z] !== null) {
                        //     leTableSorted.splice(z, 0, tagOrder[z]);
                        // };
                        leTableSorted.splice(z, 0, veryUrgentArr[z]);
                        veryUrgentLength = z + 1;
                    };

                    var urgentLength;

                    for (var zz = 0; zz < urgentArr.length; zz++) {
                        if (veryUrgentLength == undefined) {
                            var urgentAdjustedPosition = zz;
                        } else {
                            var urgentAdjustedPosition = veryUrgentLength + zz;
                        };
                        leTableSorted.splice(urgentAdjustedPosition, 0, urgentArr[zz]);
                        urgentLength = urgentAdjustedPosition + 1;
                    };

                    var semiUrgentLength;

                    for (var zzz = 0; zzz < semiUrgentArr.length; zzz++) {
                        if (urgentLength == undefined) {
                            var semiUrgentAdjustedPosition = veryUrgentLength + zzz;
                        } else {
                            var semiUrgentAdjustedPosition = urgentLength + zzz;
                        };
                        leTableSorted.splice(semiUrgentAdjustedPosition, 0, semiUrgentArr[zzz]);
                        semiUrgentLength = semiUrgentAdjustedPosition + 1;
                    };

                    for (var zzzz = 0; zzzz < eventualArr.length; zzzz++) {
                        leTableSorted.push(eventualArr[zzzz]);
                    };

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region ADDS WITHHELD INFO BACK INTO THE TABLE AT THE BOTTOM -----------------------------------------------------------------------

                    //#region ADDED AT CLIENT, IN REVIEW, AND WAITING ON INFO FIRST ------------------------------------------------------------------

                        if (awaitingChangesTable.length > 0) { //adds awaiting changes requests back into table at the bottom
                            for (var i = 0; i < awaitingChangesTable.length; i++) {
                                leTableSorted.push(awaitingChangesTable[i]);
                                awaitingChangesTable.splice(i, 1);
                                i = i - 1;
                            };
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region ADDS ON HOLD NEXT ------------------------------------------------------------------------------------------------------

                        if (onHoldTable.length > 0) { //adds on hold requests back into table at the bottom, under awaiting changes requests
                            for (var i = 0; i < onHoldTable.length; i++) {
                                leTableSorted.push(onHoldTable[i]);
                                onHoldTable.splice(i, 1);
                                i = i - 1;
                            };
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region ADDS INVALID REQUEST AT THE VERY END -----------------------------------------------------------------------------------

                        if (tempTable.length > 0) { //adds invalid requests back into table at the bottom, under on hold requests
                            for (var i = 0; i < tempTable.length; i++) {
                                leTableSorted.push(tempTable[i]);
                                tempTable.splice(i, 1);
                                i = i - 1;
                            };
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region FINDS POST SORT INDEX NUMBER OF ROW ----------------------------------------------------------------------------------------

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

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region ASSIGN PRIORITY NUMBERS ----------------------------------------------------------------------------------------------------

                    //for each item in the sorted array of table values, assign updated priority numbers to the priority column index
                    for (var n = 0; n < leTableSorted.length; n++) {
                        leTableSorted[n][priorityColumnIndex] = n + 1;
                    };

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                return leTableSorted;

            }

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region GENERATE OFFICE HOURS --------------------------------------------------------------------------------------------------------------

            //#region OFFICE HOURS FUNCTION ----------------------------------------------------------------------------------------------------------

                /**
                 * Adds adjustment hours to the date and adjusts to fit within office hours
                 * @param {Date} date the added date
                 * @param {Number} hoursToAdd The humber of adjustment hours to add to the added date
                 * @returns Date
                 */
                function officeHours(date, hoursToAdd) {

                    //#region FUNCTION VARIABLES -----------------------------------------------------------------------------------------------------

                        //date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
                        //gets the day of the week
                        var theDay = date.getDay();
                        if (theDay == 0) {
                            theDay = "Sunday"
                        } else if (theDay == 1) {
                            theDay = "Monday"
                        } else if (theDay == 2) {
                            theDay = "Tuesday"
                        } else if (theDay == 3) {
                            theDay = "Wednesday"
                        } else if (theDay == 4) {
                            theDay = "Thursday"
                        } else if (theDay == 5) {
                            theDay = "Friday"
                        } else if (theDay == 6) {
                            theDay = "Saturday"
                        };

                        var adjustmentMinutes = hoursToAdd * 60; // 12.5 hours = 750 minutes
                        var includesWeekends = false;

                        var current = new Date(date); //clone of the date variable that calculations will be made to

                        //#region SET DATES WITH 0 TIME ----------------------------------------------------------------------------------------------

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

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region CREATE START TIME AND END TIME VARIABLES FOR THE DATE --------------------------------------------------------------

                            var weekdayVars = officeHoursData[theDay] //returns all the info for the weekday that date lands on

                            //#region CREATE START TIME DATE -----------------------------------------------------------------------------------------

                                //this varibale will have the correct start time but will still be using the ground 0 date for serial numbers. 
                                //Will be adjusted up next
                                var theStart = convertToDate(weekdayVars.startTime); //converts serial number to JSDate for start of work day
                                //sets that date of theStart to be at ground 0 for JSDates
                                theStart.setFullYear(1970);
                                theStart.setMonth(0);
                                theStart.setDate(1);
                                //removes time zone offset to bring all dates to the same level
                                theStart.setMinutes(theStart.getMinutes() - theStart.getTimezoneOffset());
                                //gives us the milliseconds between 0 and this time
                                var fartTime = theStart.getTime();
                                workDayStart.setMilliseconds(fartTime); //adds the startTime to the correct date variable from eariler

                            //#endregion -------------------------------------------------------------------------------------------------------------

                            //#region CREATE END TIME DATE -------------------------------------------------------------------------------------------

                                //this varibale will have the correct end time but will still be using the ground 0 date for serial numbers. 
                                //Will be adjusted up next
                                var theEnd = convertToDate(weekdayVars.endTime); //converts serial number to JSDate for end of work day
                                //sets that date of theEnd to be at ground 0 for JSDates
                                theEnd.setFullYear(1970);
                                theEnd.setMonth(0);
                                theEnd.setDate(1);
                                //removes time zone offset to bring all dates to the same level
                                theEnd.setMinutes(theEnd.getMinutes() - theEnd.getTimezoneOffset());
                                //gives us the milliseconds between 0 and this time
                                var shartTime = theEnd.getTime();
                                workDayEnd.setMilliseconds(shartTime); //adds the endTime to the correct date variable from eariler

                            //#endregion -------------------------------------------------------------------------------------------------------------

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region WHILE ADJUSTMENT NUMBER REMAINS POSITIVE -------------------------------------------------------------------------------

                        while(adjustmentMinutes > 0) {

                            //#region RECALCULATE START AND END TIMES IF DATE ADVANCES ---------------------------------------------------------------

                                var currentInfo = shortDate(current);
                                var dateInfo = shortDate(date);

                                // if (current.toLocaleDateString('en-US') !== date.toLocaleDateString('en-US')) { //if we go on into another day, 
                                //recalculate start and end time dates
                                //if (current.getDay() !== date.getDay()) { //if we go on into another day, recalculate start and end time dates
                                if (currentInfo !== dateInfo) {

                                    //gets the day of the week
                                    var theDay = current.getDay();
                                    if (theDay == 0) {
                                        theDay = "Sunday"
                                    } else if (theDay == 1) {
                                        theDay = "Monday"
                                    } else if (theDay == 2) {
                                        theDay = "Tuesday"
                                    } else if (theDay == 3) {
                                        theDay = "Wednesday"
                                    } else if (theDay == 4) {
                                        theDay = "Thursday"
                                    } else if (theDay == 5) {
                                        theDay = "Friday"
                                    } else if (theDay == 6) {
                                        theDay = "Saturday"
                                    };

                                    //#region SET DATES WITH 0 TIME (CURRENT) ------------------------------------------------------------------------

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

                                    //#endregion -----------------------------------------------------------------------------------------------------

                                    //#region CREATE START TIME AND END TIME VARIABLES FOR THE DATE (CURRENT) ----------------------------------------

                                        //gets start and end times of date's work day
                                        weekdayVars = officeHoursData[theDay] //returns all the info for the weekday that date lands on

                                        //#region CREATE START TIME DATE (CURRENT) -------------------------------------------------------------------

                                            theStart = convertToDate(weekdayVars.startTime); //converts serial number to JSDate for start of work day
                                            theStart.setFullYear(1970);
                                            theStart.setMonth(0);
                                            theStart.setDate(1);
                                            //removes time zone offset to bring all dates to the same level
                                            theStart.setMinutes(theStart.getMinutes() - theStart.getTimezoneOffset()); 
                                            fartTime = theStart.getTime();
                                            workDayStart.setMilliseconds(fartTime);

                                        //#endregion -------------------------------------------------------------------------------------------------

                                        //#region CREATE END TIME DATE (CURRENT) ---------------------------------------------------------------------

                                            theEnd = convertToDate(weekdayVars.endTime); //converts serial number to JSDate for end of work day
                                            theEnd.setFullYear(1970);
                                            theEnd.setMonth(0);
                                            theEnd.setDate(1);
                                            //removes time zone offset to bring all dates to the same level
                                            theEnd.setMinutes(theEnd.getMinutes() - theEnd.getTimezoneOffset()); 
                                            shartTime = theEnd.getTime();
                                            workDayEnd.setMilliseconds(shartTime);

                                        //#endregion -------------------------------------------------------------------------------------------------

                                    //#endregion -----------------------------------------------------------------------------------------------------

                                };

                            //#endregion -------------------------------------------------------------------------------------------------------------

                            //#region INCREMENT ------------------------------------------------------------------------------------------------------

                                //if current is still within the workday and not on a weekend, subtract 1 minute from the adjustment number
                                if (current > workDayStart 
                                    && current < workDayEnd 
                                    && (includesWeekends ? current.getDay() !== 0 
                                    && current.getDay() !== 6 : true)) {
                                        adjustmentMinutes--;
                                };
                                current.setTime(current.getTime() + 1000 * 60); //adds 1 minute to current time

                            //#endregion -------------------------------------------------------------------------------------------------------------

                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    return current;

                };

            //#endregion -----------------------------------------------------------------------------------------------------------------------------

            //#region FUNCTIONS USED IN OFFICE HOURS -------------------------------------------------------------------------------------------------

                //#region SHORT DATE -----------------------------------------------------------------------------------------------------------------

                    /**
                     * Takes the date, month, and year from the input and outputs it as month, date, year date object (or sting)
                     * @param {Date} aDate A date object
                     * @returns Date? Or maybe a String?
                     */
                    function shortDate(aDate) {
                        var day = aDate.getDate();
                        var month = aDate.getMonth();
                        var year = aDate.getFullYear();
                        var output = `${month} ${day} ${year}`;
                        return output;
                    };

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region CONVERT DATE TO SERIAL -----------------------------------------------------------------------------------------------------

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

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region CONVERT SERIAL TO DATE -----------------------------------------------------------------------------------------------------

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

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region ALL COLUMN VALUES ----------------------------------------------------------------------------------------------------------

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

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region FIND COLUMN INDEX ----------------------------------------------------------------------------------------------------------

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
                            //if the item matches the columnName input, return the value of i, otherwise increment i & continue through rest of array
                            if (column == columnName) { 
                                jelly = i;
                                return jelly;
                            }
                            i++;
                        };
                    };

                //#endregion -------------------------------------------------------------------------------------------------------------------------

            //#endregion -----------------------------------------------------------------------------------------------------------------------------

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region MOVING DATA FUNCTIONS ------------------------------------------------------------------------------------------------------------------

        //#region MOVE DATA TWO ----------------------------------------------------------------------------------------------------------------------

            /**
             * Pushes the rowValues to the destTable array and removes it from the leTable array
             * @param {Array} destTable the table array that the data is being moved to
             * @param {Array} rowValues an array of all the values of the changed row
             * @param {Array} leTable the changed table array
             * @param {Number} changedRowTableIndex the index number of the changed table row
             */
             function moveDataTwo(destTable, rowValues, leTable, changedRowTableIndex) {

                destTable.push(rowValues[0]);
                leTable.splice(changedRowTableIndex, 1);

            };

        //#endregion -------------------------------------------------------------------------------------------------------------------------------------

        //#region CHECK IF TWO ARRAYS ARE EQUAL ------------------------------------------------------------------------------------------------------

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

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region CONDITIONAL FORMATTING -----------------------------------------------------------------------------------------------------------------

        //#region MAIN CONDITIONAL FORMATTING FUNCTION -----------------------------------------------------------------------------------------------
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
            function conditionalFormatting(rowInfoSorted, newTableStart, changedWorksheet, 
                rowIndexPostSort, completedTableChanged, rowRangeSorted, destTable) {

                //#region DEFINING VARIABLES ---------------------------------------------------------------------------------------------------------

                    var now = new Date();
                    var justNowDate = now.getDate();
                    var toSerial = Number(JSDateToExcelDate(now));

                    var worksheetRowIndex = rowIndexPostSort + 1; //adjusts index post table sort to work on worksheet level

                    var pickedUpWorksheetColumn = rowInfoSorted.pickedUpStartedBy.columnIndex + newTableStart;
                    var proofToClientWorksheetColumn = rowInfoSorted.proofToClient.columnIndex + newTableStart;
                    var printDateWorksheetColumn = rowInfoSorted.printDate.columnIndex + newTableStart;
                    var groupWorksheetColumn = rowInfoSorted.group.columnIndex + newTableStart;
                    var tagsWorksheetColumn = rowInfoSorted.tags.columnIndex + newTableStart;

                    var pickedUpAddress = changedWorksheet.getCell(worksheetRowIndex, pickedUpWorksheetColumn);
                    var proofToClientAddress = changedWorksheet.getCell(worksheetRowIndex, proofToClientWorksheetColumn);

                    var printDate = Math.trunc(rowInfoSorted.printDate.value);
                    // console.log(convertToDate(printDate));
                    var currentDateAbsolute = Math.trunc(toSerial);

                    var printDateAddress = changedWorksheet.getCell(worksheetRowIndex, printDateWorksheetColumn);
                    var groupAddress = changedWorksheet.getCell(worksheetRowIndex, groupWorksheetColumn);

                    var tagsAddress = changedWorksheet.getCell(worksheetRowIndex, tagsWorksheetColumn);

                    var logoRecreationStatus = ["Logo Status TBD", "Logo Needs Recreating", "Logo Needs Uploading", "No Logo Recreation Needed"];

                    console.log(changedWorksheet.name);

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                if (changedWorksheet.name == "Unassigned Projects" && rowInfoSorted.subject.value == "TEMPORARY DATA VALIDATION PLACEHOLDER") {
                    rowRangeSorted.format.fill.color = "#BFBFBF";
                    rowRangeSorted.format.font.color = "#808080";

                    return;
                };

                //#region CLEAR COMPLETED TABLE FORMATTING IF IT WAS CHANGED -------------------------------------------------------------------------

                    //if completed table was changed, clear formatting and do not do any other formatting rules
                    if (completedTableChanged == true && destTable == null) {
                        rowRangeSorted.format.fill.clear();
                        rowRangeSorted.format.font.color = "black";
                        rowRangeSorted.format.font.bold = false;

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                } else {

                    //#region ALL ENTRIES USE CONSISTENT FONT STYLING --------------------------------------------------------------------------------

                        rowRangeSorted.format.font.name = "Calibri";
                        rowRangeSorted.format.font.size = 12;
                        rowRangeSorted.format.font.color = "#000000";
                        rowRangeSorted.format.font.bold = false;

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region REMOVE INVALID HIGHLIGHTING IF NO LONGER INVALID -----------------------------------------------------------------------

                        if (
                            rowInfoSorted.pickedUpStartedBy.value !== "NO PRODUCT / PROJECT TYPE" 
                            || rowInfoSorted.proofToClient.value !== "NO PRODUCT / PROJECT TYPE"
                        ) {

                            rowRangeSorted.format.fill.clear();
                            pickedUpAddress.format.font.bold = false;
                            proofToClientAddress.format.font.bold = false;

                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region GROUP & PRINT DATE FORMATTING ------------------------------------------------------------------------------------------

                        if (printDate == currentDateAbsolute) { //if current date = print date

                            logoRecreationStatus.forEach(
                                leStatus =>  logoInsertHighlighting(rowInfoSorted, rowRangeSorted, printDateAddress, groupAddress, leStatus)
                            );

                            // rowRangeSorted.format.font.color = "#C00000";
                            // rowRangeSorted.format.font.bold = true;
                            printDateAddress.format.fill.color = "#FFBBB8"; //light red
                            printDateAddress.format.font.color = "black";
                            printDateAddress.format.font.bold = true;

                            groupAddress.format.fill.color = "#FFBBB8"; //light red
                            groupAddress.format.font.color = "black";
                            groupAddress.format.font.bold = true;

                            printDateAddress.format.horizontalAlignment = "center";
                            groupAddress.format.horizontalAlignment = "center";

                            // if (rowInfoSorted.tags.value !== "") {
                            //     rowRangeSorted.format.font.color = "#ED7D31"; //#9BC2E6
                            //     rowRangeSorted.format.font.bold = true;
                            //     printDateAddress.format.font.color = "white";
                            //     groupAddress.format.font.color = "white";
                            // };

                            console.log("Print Date & Group cells were colored red");

                        } else if (((printDate - 1) == currentDateAbsolute)) { //if current date is the day before print date

                            logoRecreationStatus.forEach(
                                leStatus =>  logoInsertHighlighting(rowInfoSorted, rowRangeSorted, printDateAddress, groupAddress, leStatus)
                            );

                            // rowRangeSorted.format.font.color = "#C00000";
                            // rowRangeSorted.format.font.bold = true;
                            printDateAddress.format.fill.color = "#FFBBB8"; //light red
                            printDateAddress.format.font.color = "black";
                            printDateAddress.format.font.bold = true;

                            groupAddress.format.fill.color = "#FFBBB8"; //light red
                            groupAddress.format.font.color = "black";
                            groupAddress.format.font.bold = true;

                            printDateAddress.format.horizontalAlignment = "center";
                            groupAddress.format.horizontalAlignment = "center";

                            // if (rowInfoSorted.tags.value !== "") {
                            //     rowRangeSorted.format.font.color = "#ED7D31"; //#9BC2E6
                            //     rowRangeSorted.format.font.bold = true;
                            //     printDateAddress.format.font.color = "white";
                            //     groupAddress.format.font.color = "white";
                            // };

                            // for (var leStatus in logoRecreationStatus) {
                            //     logoInsertHighlighting(rowInfoSorted, rowRangeSorted, printDateAddress, groupAddress, leStatus);
                            // };

                            console.log("Print Date & Group cells were colored red");

                        
                            //if current date is in the same group lock week as print date (between 7-2 days before)
                        } else if (((printDate - 6) <= currentDateAbsolute) && ((printDate - 2) >= currentDateAbsolute)) { 

                            logoRecreationStatus.forEach(
                                leStatus =>  logoInsertHighlighting(rowInfoSorted, rowRangeSorted, printDateAddress, groupAddress, leStatus)
                            );

                            // rowRangeSorted.format.font.color = "#C00000";
                            // rowRangeSorted.format.font.bold = true;
                            printDateAddress.format.fill.color = "#FFBBB8"; //light red
                            printDateAddress.format.font.color = "black";
                            printDateAddress.format.font.bold = true;

                            groupAddress.format.fill.color = "#FFBBB8"; //light red
                            groupAddress.format.font.color = "black";
                            groupAddress.format.font.bold = true;

                            printDateAddress.format.horizontalAlignment = "center";
                            groupAddress.format.horizontalAlignment = "center";

                            // if (rowInfoSorted.tags.value !== "") {
                            //     rowRangeSorted.format.font.color = "#ED7D31"; //#9BC2E6
                            //     rowRangeSorted.format.font.bold = true;
                            //     printDateAddress.format.font.color = "white";
                            //     groupAddress.format.font.color = "white";
                            // };

                            console.log("Print Date & Group cells were colored red");


                        //if current date is in the week before group lock week (between 8-14 days before)
                        } else if (((printDate - 13) <= currentDateAbsolute) && ((printDate - 7) >= currentDateAbsolute)) { 

                            logoRecreationStatus.forEach(
                                leStatus =>  logoInsertHighlighting(rowInfoSorted, rowRangeSorted, printDateAddress, groupAddress, leStatus)
                            );

                            // rowRangeSorted.format.font.color = "#C00000";
                            // rowRangeSorted.format.font.bold = true;
                            printDateAddress.format.fill.color = "#A9D08E"; //light green
                            printDateAddress.format.font.color = "black";
                            printDateAddress.format.font.bold = true;

                            groupAddress.format.fill.color = "#A9D08E"; //light green
                            groupAddress.format.font.color = "black";
                            groupAddress.format.font.bold = true;

                            printDateAddress.format.horizontalAlignment = "center";
                            groupAddress.format.horizontalAlignment = "center";

                            // if (rowInfoSorted.tags.value !== "") {
                            //     rowRangeSorted.format.font.color = "#ED7D31"; //#9BC2E6
                            //     rowRangeSorted.format.font.bold = true;
                            //     printDateAddress.format.font.color = "white";
                            //     groupAddress.format.font.color = "white";
                            // };

                            console.log("Print Date & Group cells were colored green");

                            
                        } else if ((printDate < currentDateAbsolute) && (printDate !== 0)) { //if current date is after print date

                            logoRecreationStatus.forEach(
                                leStatus =>  logoInsertHighlighting(rowInfoSorted, rowRangeSorted, printDateAddress, groupAddress, leStatus)
                            );

                            printDateAddress.format.fill.color = "black";
                            printDateAddress.format.font.color = "white";
                            printDateAddress.format.font.bold = true;

                            groupAddress.format.fill.color = "black";
                            groupAddress.format.font.color = "white";
                            groupAddress.format.font.bold = true;

                            printDateAddress.format.horizontalAlignment = "center";
                            groupAddress.format.horizontalAlignment = "center";

                            // if (rowInfoSorted.tags.value !== "") {
                            //     rowRangeSorted.format.font.color = "#ED7D31";
                            //     //rowRangeSorted.format.font.bold = true;
                            // };

                            console.log("Print Date & Group cells were colored black");

                            
                        } else { //set cell formatting to normal

                            if (changedWorksheet.name == "Unassigned Projects" && rowInfoSorted.subject.value == "Test for Artist Data Validation") {
                                return;
                            };

                            rowRangeSorted.format.fill.clear();
                            rowRangeSorted.format.font.color = "black";
                            rowRangeSorted.format.font.bold = false;
                            printDateAddress.format.horizontalAlignment = "center";
                            groupAddress.format.horizontalAlignment = "center";

                            // if (rowInfoSorted.tags.value !== "") {
                            //     rowRangeSorted.format.font.color = "#ED7D31";
                            //     rowRangeSorted.format.font.bold = true;
                            // };

                            console.log("Color formatting was removed from row");

                            logoRecreationStatus.forEach(
                                leStatus =>  logoInsertHighlighting(rowInfoSorted, rowRangeSorted, printDateAddress, groupAddress, leStatus)
                            );
                            
                        };

                        if (rowInfoSorted.group.value == "N/A") { //if group column in blank

                            rowRangeSorted.format.fill.clear();
                            rowRangeSorted.format.font.color = "black";
                            rowRangeSorted.format.font.bold = false;
                            printDateAddress.format.horizontalAlignment = "center";
                            groupAddress.format.horizontalAlignment = "center";

                            // if (rowInfoSorted.tags.value !== "") {
                            //     rowRangeSorted.format.font.color = "#ED7D31";
                            //     rowRangeSorted.format.font.bold = true;
                            // };

                            console.log("Color formatting was removed from row");

                            logoRecreationStatus.forEach(
                                leStatus =>  logoInsertHighlighting(rowInfoSorted, rowRangeSorted, printDateAddress, groupAddress, leStatus)
                            );

                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region IF STATUS IS WORKING ---------------------------------------------------------------------------------------------------

                        if (rowInfoSorted.status.value == "Working") {

                            rowRangeSorted.format.fill.color = "#FFE699";
                            rowRangeSorted.format.font.color = "#9C5700";
                            rowRangeSorted.format.font.bold = true;
                            printDateAddress.format.horizontalAlignment = "center";
                            groupAddress.format.horizontalAlignment = "center";

                            // if (rowInfoSorted.tags.value !== "") {
                            //     rowRangeSorted.format.font.color = "#ED7D31";
                            //     //rowRangeSorted.format.font.bold = true;
                            // };

                            console.log("Row was highlighted in yellow for the working status");
                            
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region TAGS FORMATTING --------------------------------------------------------------------------------------------------------

                        if (rowInfoSorted.tags.value == "VERY URGENT") {
                            tagsAddress.format.font.color = "#C00000";
                            tagsAddress.format.font.bold = true;
                            tagsAddress.format.horizontalAlignment = "center";
                        };

                        if (rowInfoSorted.tags.value == "URGENT") {
                            tagsAddress.format.font.color = "#BF8F00";
                            tagsAddress.format.font.bold = true;
                            tagsAddress.format.horizontalAlignment = "center";
                        };

                        if (rowInfoSorted.tags.value == "SEMI-URGENT") {
                            tagsAddress.format.font.color = "#548235";
                            tagsAddress.format.font.bold = true;
                            tagsAddress.format.horizontalAlignment = "center";
                        };

                        if (rowInfoSorted.tags.value == "EVENTUAL") {
                            tagsAddress.format.font.color = "#7030A0";
                            tagsAddress.format.font.bold = true;
                            tagsAddress.format.horizontalAlignment = "center";
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region OVERDUE HIGHLIGHTING ---------------------------------------------------------------------------------------------------

                        //#region PRINT DATE OVERDUE (OUTSIDE OF GROUP & PRINT REGION BECAUSE IT NEEDS TO OVERRIDE WORKING STATUS) -------------------

                            if ((printDate < currentDateAbsolute) && (printDate !== 0)) { //if current date is after print date

                                printDateAddress.format.fill.color = "black";
                                printDateAddress.format.font.color = "white";
                                printDateAddress.format.font.bold = true;

                                groupAddress.format.fill.color = "black";
                                groupAddress.format.font.color = "white";
                                groupAddress.format.font.bold = true;

                                printDateAddress.format.horizontalAlignment = "center";
                                groupAddress.format.horizontalAlignment = "center";

                                // if (rowInfoSorted.tags.value !== "") {
                                //     rowRangeSorted.format.font.color = "#ED7D31";
                                //     //rowRangeSorted.format.font.bold = true;
                                // };

                                console.log("Print Date & Group cells were colored black");

                                logoRecreationStatus.forEach(
                                    leStatus =>  logoInsertHighlighting(rowInfoSorted, rowRangeSorted, printDateAddress, groupAddress, leStatus)
                                );
                                
                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region PICKED UP / STARTED BY OVERDUE -------------------------------------------------------------------------------------

                            if (toSerial > rowInfoSorted.pickedUpStartedBy.value && changedWorksheet.name == "Unassigned Projects") {
                                //pickedUpAddress.format.fill.color = "FFC000";
                                rowRangeSorted.format.fill.color = "FFC000";
                                rowRangeSorted.format.font.color = "black";

                                // if (rowInfoSorted.tags.value !== "") {
                                //     rowRangeSorted.format.font.color = "#ED7D31";
                                //     rowRangeSorted.format.font.bold = true;
                                // };

                                console.log("Row was colored yellow because it is overdue to be assigned");
                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region PROOF TO CLIENT OVERDUE --------------------------------------------------------------------------------------------

                            if (toSerial > rowInfoSorted.proofToClient.value && changedWorksheet.name !== "Unassigned Projects") {
                                // proofToClientAddress.format.fill.color = "FF0000";
                                // proofToClientAddress.format.font.color = "white";
                                rowRangeSorted.format.fill.color = "FF0000";
                                rowRangeSorted.format.font.color = "white";
                                rowRangeSorted.format.font.bold = true;

                                // if (rowInfoSorted.tags.value !== "") {
                                //     rowRangeSorted.format.font.color = "#FFFF00";
                                //     rowRangeSorted.format.font.bold = true;
                                // };

                                console.log("Row was colored red because it is overdue");

                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region STATUS OVERRIDE FORMATTING (FINAL SAY) ---------------------------------------------------------------------------------

                        //#region ON HOLD STATUS -----------------------------------------------------------------------------------------------------
                                    
                            if (rowInfoSorted.status.value == "On Hold") {
                                rowRangeSorted.format.fill.color = "#BFBFBF";
                                rowRangeSorted.format.font.color = "#000000";
                                rowRangeSorted.format.font.bold = false;
                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region IN REVIEW STATUS ---------------------------------------------------------------------------------------------------

                            if (rowInfoSorted.status.value == "In Review") {
                                rowRangeSorted.format.fill.clear()
                                rowRangeSorted.format.font.color = "#757171";
                                rowRangeSorted.format.font.bold = false;
                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region AT CLIENT STATUS ---------------------------------------------------------------------------------------------------

                            if (rowInfoSorted.status.value == "At Client") {
                                rowRangeSorted.format.fill.clear()
                                rowRangeSorted.format.font.color = "#757171";
                                rowRangeSorted.format.font.bold = false;
                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region WAITING ON INFO STATUS ---------------------------------------------------------------------------------------------

                            if (rowInfoSorted.status.value == "Waiting On Info") {
                                rowRangeSorted.format.fill.clear()
                                rowRangeSorted.format.font.color = "#757171";
                                rowRangeSorted.format.font.bold = false;
                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    //#region ADD INVALID HIGHLIGHTING IF INVALID ------------------------------------------------------------------------------------

                        if (
                            rowInfoSorted.pickedUpStartedBy.value == "NO PRODUCT / PROJECT TYPE" 
                            || rowInfoSorted.proofToClient.value == "NO PRODUCT / PROJECT TYPE"
                        ) {

                            rowRangeSorted.format.fill.color = "FFC5BB";
                            rowRangeSorted.format.font.color = "black";
                            rowRangeSorted.format.font.bold = false;

                            pickedUpAddress.format.font.bold = true;
                            proofToClientAddress.format.font.bold = true;
                            // pickedUpAddress.format.fill.color = "FFC5BB";
                            // proofToClientAddress.format.fill.color = "FFC5BB";

                            console.log("Row is invalid, so it is red");

                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------

                    // console.log(groupAddress.format.fill.color)
                    // console.log(groupAddress.format.font.color);
                    // console.log(groupAddress.format.font.bold);
                };

            };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region LOGO INSERT HIGHLIGHT FUNCTION -----------------------------------------------------------------------------------------------------

            /**
             * Determines wheather or not the row is a logo recreation row with the proper status to get the orange text highlight
             * @param {Object} rowInfo An object containing the values and column indexs of each cell in the changed row
             * @param {Range} rowRange The range of the changed row
             * @param {Range} printDateAddress The cell address of the print date
             * @param {Range} groupAddress The cell address of the group
             * @param {String} input The qualifying status that will award the row an orange text highlight
             */
            function logoInsertHighlighting (rowInfo, rowRange, printDateAddress, groupAddress, input) {
                if (rowInfo.product.value == "Logo Recreation" && rowInfo.status.value == input) {
                    rowRange.format.font.color = "#ED7D31"; //#9BC2E6
                    rowRange.format.font.bold = true;
                    console.log("Logo Insert Row was highlighted!")
                    // printDateAddress.format.font.color = "white";
                    // groupAddress.format.font.color = "white";
                };
            };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region TOGGLE AND CHECK EVENTS ----------------------------------------------------------------------------------------------------------------

        //#region EVENTS ON --------------------------------------------------------------------------------------------------------------------------

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

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

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

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region HANDLE CHANGE EVENT FOR DEBUGGING --------------------------------------------------------------------------------------------------

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

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region SHOW AND HIDE MESSAGES / ERRORS --------------------------------------------------------------------------------------------------------

    function showMessage(msg, showHide) {
        if (showHide === "hide") {
            $("#message-text").empty();
            $("#message").css("display", "none");
        } else if (showHide === "show") {
            $("#message-text").text(msg);
            $("#message").css("display", "flex");
        }
    };

    function showElement(element, showHide) {
        if (showHide === "hide") {
            $(element).css("display", "none");
        } else if (showHide === "show") {
            $(element).css("display", "flex");
        };
    };

    function showFisshGif() {
        $("#fissh-gif").css("display", "flex");
        setTimeout(hideFisshGif, 2000);
    };

    function hideFisshGif() {
        $("#fissh-gif").css("display", "none");
    };


    // Illegal insert dennis gif
    function showDennis() {
        $("#dennis").css("display", "flex");
        console.log("Na-Ah-Ah!");
        var naAhAh = new Audio("assets/dennis-mock.mp3");
        naAhAh.play();

        $("#dennis").append(`
            <img id="dennis-gif" src="assets/dennis-crop.gif" />
        `);

        // Wait 1.5 seconds, hide Dennis
        setTimeout(() => {
        $("#dennis").css("display", "none");
        $("#dennis-gif").remove();
        $("#na-ah-ah").css("display", "flex");
        }, 1700);

        return;

    };

//#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region REMOVE ROW -----------------------------------------------------------------------------------------------------------------------------

        /**
         * Removes current row from table
         * @param {Object} eventArgs The event arguments, which are details about the event that was triggered
         */
         async function removeRow(eventArgs) {
        
            await Excel.run(async (context) => {

                var changedWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId).load("name");

                //Returns tableId of the table where the event occured
                var changedTable = context.workbook.tables.getItem(eventArgs.tableId).load("name"); 

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

                rowRange.delete("Up");

                console.log("A logo Recreation row has been removed");

                // eventsOn();
                // console.log(`Events: ON  →  triggered after a row was manually inserted into the sheet by the user, 
                // followed by the swift removal of said row and a slap on the wrist.`);
                
                return;

            });

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region HANDLING ILLEGAL ROW INSERTS -----------------------------------------------------------------------------------------------------------

        /**
         * If a row is inserted to the sheet manually, it will be removed and the user chastised
         * @param {Object} eventArgs The event arguments, which are details about the event that was triggered
         */
        async function handleIllegalInsert(eventArgs) {
        
            await Excel.run(async (context) => {

                var changedWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId).load("name");

                //Returns tableId of the table where the event occured
                var changedTable = context.workbook.tables.getItem(eventArgs.tableId).load("name"); 

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
                console.log(`"Thou shalt submit all requests to thy own sheet by means of the Add A Project taskpane. 
                Manually adding rows of info to thyn sheet beith forbidden."`);
                console.log("It's a simple mistake, but make sure not to do it again.");

                rowRange.delete("Up");

                eventsOn();
                console.log(`Events: ON  →  triggered after a row was manually inserted into the sheet by the user, 
                followed by the swift removal of said row and a slap on the wrist.`);
                
                return;

            });

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region ERROR HANDLING -------------------------------------------------------------------------------------------------------------------------

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
        }

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

    //#region UNUSED FUNCTIONS / WORK IN PROGRESS ----------------------------------------------------------------------------------------------------

        //#region MOVE DATA FUNCTION (LEGACY, NOT USED ANYMORE) --------------------------------------------------------------------------------------

            /**
             * moves the changed row's data to the destionation table
             * @param {Object} destinationTable the table that the data is being moved to
             * @param {Array} myRow the data, values, and attributes of the changed row
             * @param {String} artistCellValue the value of the artist cell in the changed row
             */
             function moveData(destinationTable, rowValues, myRow, artistCellValue) {
                //Adds empty row to bottom of the destinationTable, then inserts the changed values into this empty row
                destinationTable.rows.add(null, rowValues); 
                myRow.delete(); //Deletes the changed row from the original sheet
                console.log("Data was moved to " + artistCellValue + "'s Projects Table!");
            };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

        //#region DATA PROTECTION (UNUSED) -----------------------------------------------------------------------------------------------------------

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

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------------------------------------------


