//#region TEST SUBJECTS -------------------------------------------------------------------------------------------------------------
    //CREATIVE REQUEST -Alfredo's Pizza - West Babylon - MENU - ~/*1338,52130,1*/~
    //CREATIVE REQUEST -Bella Napoli - Canfield - Env #10 8.5x11 S2 - ~/*1837,65845,1*/~
    //Re: Artist Request - Brickhouse Pizzeria - Richfield Springs - MENU - ~/*30601,72301,1*/~
//#endregion ------------------------------------------------------------------------------------------------------------------------

//#region I MIGHT NEED THIS BEGINNING MATERIAL SOME DAY -----------------------------------------------------------------------------

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
var loop = true;



//#region ON READY ------------------------------------------------------------------------------------------------------------

    //#region LOADS VALIDATION VALUES AND UPDATES DROPDOWN VALUES IN TASKPANE ------------------------------------------------

        Office.onReady((info) => {
            if (info.host === Office.HostType.Excel) {


                Excel.run(async (context) => {

                    //#region LOADING VALUES ---------------------------------------------------------------------------------

                        //load up the validation tables being referenced
                        var sheet = context.workbook.worksheets.getItem("Validation");
                        var productIDValTable = sheet.tables.getItem("ProductIDTable");
                        var projectTypeIDTable = sheet.tables.getItem("ProjectTypeIDTable")
                        var pickedUpValTable = sheet.tables.getItem("PickupTurnaroundTime");
                        var proofToClientValTable = sheet.tables.getItem("ArtTurnaroundTime");
                        var creativeProofTable = sheet.tables.getItem("CreativeProofAdjust");
                        var officeHoursTable = sheet.tables.getItem("OfficeHours");



                        //get data from the tables
                        var productIDBodyRange = productIDValTable.getDataBodyRange().load("values");
                        var projectTypeIDBodyRange = projectTypeIDTable.getDataBodyRange().load("values");
                        var pickedUpBodyRange = pickedUpValTable.getDataBodyRange().load("values");
                        var proofToClientBodyRange = proofToClientValTable.getDataBodyRange().load("values");
                        var creativeProofBodyRange = creativeProofTable.getDataBodyRange().load("values");
                        var officeHoursBodyRange = officeHoursTable.getDataBodyRange().load("values");

                    //#endregion ----------------------------------------------------------------------------------------------

                    await context.sync();

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

                            // console.log(proofToClientData);

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


                    //#endregion --------------------------------------------------------------------------------------------------

                });
                // console.log(info);
                tryCatch(updateDropDowns);
                
            };
        });

    //#endregion ---------------------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------------------



//#region STYLIZING TASKPANE ELEMENTS ----------------------------------------------------------------------------------------


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



//#region UPDATE DROPDOWNS ----------------------------------------------------------------------------------------------------

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
                var artistLeadValTable = sheet.tables.getItem("ArtistLeadTable");
                var queueValTable = sheet.tables.getItem("QueueTable");
                var tierValTable = sheet.tables.getItem("TierTable");
                var tagsValTable = sheet.tables.getItem("TagsTable");

                // Get data from the table.
                var productIDBodyRange = productIDValTable.getDataBodyRange().load("values");
                var projectTypeBodyRange = projectTypeValTable.getDataBodyRange().load("values");
                var groupPrintBodyRange = groupPrintValTable.getDataBodyRange().load("values");
                var artistLeadBodyRange = artistLeadValTable.getDataBodyRange().load("values");
                var queueBodyRange = queueValTable.getDataBodyRange().load("values");
                var tierBodyRange = tierValTable.getDataBodyRange().load("values");
                var tagsBodyRange = tagsValTable.getDataBodyRange().load("values");

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

            //#region PRINT DATE & GROUP VALUES -----------------------------------------------------------------------------------

                var groupPrintBodyValues = groupPrintBodyRange.values;

                $("#print-date").empty();
                $("#print-date").append($("<option disabled selected hidden></option>").val("").text(""));
                $("#group").empty();
                $("#group").append($("<option disabled selected hidden></option>").val("").text(""));

                groupPrintBodyValues.forEach(function(row) {

                    // Add an option to the select box
                    var option = `<option group-id="${row[0]}" print-date-id="${row[1]}">${row[0]}</option>`;

                    var x = $(`#print-date > option[print-date-id="${row[1]}"]`).length;
                    var y = $(`#group > option[group-id="${row[0]}"]`).length;


                    if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
                        var leDate = convertToDate(`${row[1]}`);

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

                var artistLeadBodyValues = artistLeadBodyRange.values;

                $("#artist-lead").empty();
                $("#artist-lead").append($("<option disabled selected hidden></option>").val("").text(""));

                artistLeadBodyValues.forEach(function(row) {

                    // Add an option to the select box
                    var option = `<option artist-lead-id="${row[0]}">${row[0]}</option>`;

                    var x = $(`#artist-lead > option[artist-lead-id="${row[0]}"]`).length;

                    if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
                        $("#artist-lead").append(option);
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



//#region AUTO POPULATE TASKPANE BASED ON SUBJECT ----------------------------------------------------------------------------


    $("#subject").keyup(() => tryCatch(subjectPasted));


    //#region SUBJECT PASTED FUNCTION ----------------------------------------------------------------------------------------------

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


//#endregion -----------------------------------------------------------------------------------------------------------------



//#region TASKPANE BUTTONS ---------------------------------------------------------------------------------------------------

    //#region ON SUBMIT CLICK ------------------------------------------------------------------------------------------------

        $("#submit").on("click", function() {

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
                addAProject();
            };
        });

    //#endregion -------------------------------------------------------------------------------------------------------------

    //#region ON CLEAR CLICK ----------------------------------------------------------------------------------------------------

    $("#clear").on("click", function() {

        $("#subject, #client, #location, #product, #code, #project-type, #csm, #print-date, #group, #artist-lead, #queue, #tier, #tags, #start-override, #work-override").val(""); // Empty all inputs
        removeWarningClass("#subject", "#warning1");
        removeWarningClass("#client", "#warning2");
        removeWarningClass("#product", "#warning3");
        removeWarningClass("#project-type", "#warning4");

    });

    //#endregion ---------------------------------------------------------------------------------------------------------------

//#endregion ------------------------------------------------------------------------------------------------------------------



//#region ADDING A PROJECT FROM TASKPANE ---------------------------------------------------------------------------------------------------


    //#region ADD A PROJECT ----------------------------------------------------------------------------------------------------------------

        /**
         * Generates Added date/time, turn around times for both the Picked Up / Started By and Proof To Client columns adjusted for office hours, adds these values to the table, then generates a priority number for each row based on the value in the Picked Up / Started By column, then sorts the data by priority
         */
        async function addAProject() {

            await Excel.run(async (context) => {

                //#region LOAD VALUES ------------------------------------------------------------------------------------------------------

                    var sheet = context.workbook.worksheets.getActiveWorksheet();
                    var sheetTable = sheet.tables.getItemAt(0);

                //#endregion ---------------------------------------------------------------------------------------------------------------


                //#region GET INPUT FROM TASKPANE ------------------------------------------------------------------------------------------

                    // Data from DOM
                    var artistLeadVal = $("#artist-lead").val();
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

                //#endregion --------------------------------------------------------------------------------------------------------------


                //#region WRITE ARRAY -----------------------------------------------------------------------------------------------------

                    // Data to send to Table
                    var write = [[
                        "", // 1 - Priority
                        artistLeadVal, // 2 - Artist Lead
                        queueVal, // 3 - Queue
                        tierVal, // 4 - Tier
                        subjectVal, // 5 - Subject
                        clientVal, // 6 - Client
                        locationVal, // 7 - Location
                        productVal, // 8 - Product
                        projectTypeVal, // 9 - Project Type
                        csmVal, // 10 - CSM
                        "", // 11 - Added
                        printDateVal, // 12 - Print Data
                        groupVal, // 13 - Group
                        "", // 14 - Picked Up / Started By
                        "", // 15 - Proof to Client
                        "", // 16 - Date of Last Edit
                        tagsVal, // 17 - Tags
                        "", // 18 - Status
                        codeVal, // 19 - Code
                        "", // 20 - Artist
                        "", // 21 - Notes
                        startOverrideVal, // 22 - Start Override
                        workOverrideVal // 23 - Work Override
                    ]];

                //#endregion -----------------------------------------------------------------------------------------------------------


                //#region GENERATE ADDED DATE -----------------------------------------------------------------------------------------

                    var now = new Date();
                    var toSerial = JSDateToExcelDate(now);

                    write[0][10] = toSerial;

                //#endregion ---------------------------------------------------------------------------------------------------------


                //#region GENERATE PICKED UP / TURN AROUND TIME VALUE -----------------------------------------------------------------

                    //get the Project Type Coded variable from the Project Type ID Data based on the returned Project Type from the taskpane
                    var theProjectTypeCode = projectTypeIDData[projectTypeVal].projectTypeCode;

                    //returns turn around time value from the PickedUp Turn Around Time table based on the product and project type values
                    var pickedUpTurnAroundTime = pickupData[productVal][theProjectTypeCode];

                    //add start override time to # of hours
                    var pickedUpHours = pickedUpTurnAroundTime + startOverrideVal;
                
                    //add new time to date added, then adjust for office hours
                    var addedDate = new Date(now);
                    var pickupOfficeHours = officeHours(addedDate, pickedUpHours, officeHoursData);

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

                    write[0][13] = excelPickupOfficeHours;

                //#endregion --------------------------------------------------------------------------------------------------------


            //#region GENERATE ART TURN AROUND TIME VALUE --------------------------------------------------------------------------

                    //returns turn around time value from the Proof To Client Turn Around Time table based on the product and project type values
                    var proofToClient = proofToClientData[productVal][theProjectTypeCode];

                    //returns the Creative Review Process value from said table based on the product
                    var creativeReview = creativeProofData[productVal].creativeReviewProcess;

                    //adds proof to client value to the creative review turn around time
                    var proofWithReview = proofToClient + creativeReview;

                    //add work override time to # of hours
                    var artTurnAround = proofWithReview + workOverrideVal;
                
                    //add new time to the value previouskly found in the pickUpOfficeHours variable, then adjust for office hours
                    var proofToClientOfficeHours = officeHours(pickupOfficeHours, artTurnAround, officeHoursData);

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

                    write[0][14] = excelProofToClientOfficeHours;

                //#endregion --------------------------------------------------------------------------------------------------------


                //writes the write array to the table
                sheetTable.rows.add(null /*add rows to the end of the table*/, write);

                await context.sync(); // BOOM!


                //#region PRIORITY NUMBER GENERATION AND SORTING -------------------------------------------------------------------------

                    //assign priority numbers and sorts table
                    priorityGenerationAndSortation();

                //#endregion -------------------------------------------------------------------------------------------------------------


                console.log("I randed!");

            });
        };

    //#endregion ----------------------------------------------------------------------------------------------------------------------------


    //#region ADD A PROJECT SUB-FUNCTIONS -------------------------------------------------------------------------------------------------


        //#region OFFICE HOURS ---------------------------------------------------------------------------------------

            /**
             * Sets weekday variables and loops through the withinOfficeHours function, which adjusts the date to be within office hours
             * @param {Date} date Date to be adjusted to be within office hours
             * @param {Number} number Number of adjustment hours to add to date
             * @returns Date
             */
            function officeHours(day, number) { 

                while (loop == true) { //loops through the office hours function until the value returns within office hours
                    var officeHours = withinOfficeHours(day, number);
                    day = officeHours.date;
                    number = officeHours.adjustmentNumber;
                    loop = officeHours.loop;
                };
                //console.log("The correct date & time is: " + day);
                loop = true;
                // console.log(day);
                return day;

            };


            //#region OFFICE HOURS FUNCTIONS -----------------------------------------------------------------------------------------------------------


                //#region WITHIN OFFICE HOURS FUNCTION -------------------------------------------------------------------------------------------------

                    /**
                     * Adjusts date to be within office hours while maintaining an accurate turn around time variable for the adjustment number
                     * @param {Date} date Date to be adjusted to be within office hours 
                     * @param {Number} adjustmentNumber Number of adjustment hours to add to date
                     * @returns An object with properties (date, adjustment number, and loop)
                     */
                    function withinOfficeHours(date, adjustmentNumber) {

                        //#region VARIABLES ------------------------------------------------------------------------------------------------------------

                            //#region SETS DATE VARIABLES ----------------------------------------------------------------------------------------------

                                var dateSerial = Number(JSDateToExcelDate(date)); //converts date to excel serial for calculations
                                var adjusted = parseFloat(adjustmentNumber); //converts adjustment number from String to Number for calculations
                                var numberMinutes = adjusted * 60;
                                var adjustmentNumberSerial = minutesToSerial(numberMinutes);

                                //gets day of the week attributes for the date variable
                                var dateDayOfWeek = date.getDay(); //returns a dayID (0-6) for the day of the week of the date object
                                var dayTitle = titleDOW(dateDayOfWeek); //returns a day title based on the dayID of the dateDayOfWeek variable
                                var theWeekdayVar = officeHoursData[dayTitle];

                                var startOfWorkDay = setToStartOfDay(date, theWeekdayVar);

                                var endOfWorkDay = setToEndOfDay(date, theWeekdayVar);

                            //#endregion -------------------------------------------------------------------------------------------------------------

                            //#region ADJUSTS DATES IN CASE REQUEST WAS SUBMITTED OUTSIDE OF OFFICE HOURS ---------------------------------------

                                if (dateSerial < startOfWorkDay) { //if date is between 12AM and start time, adjust hours to be the start time
                                    dateSerial = startOfWorkDay;
                                    date = convertToDate(dateSerial);
                                    //   dateSerial = Number(JSDateToExcelDate(date)); //converts date to excel serial for calculations

                                    //   dateMilli = date.getTime();
                                    //   bookendVars = startEndMidnight(date, theWeekdayVar);
                                };

                                if (dateSerial > endOfWorkDay) { //if date is after end time and before 12AM, go to next day and adjust hours to be the start time of that next day
                                    date.setDate(date.getDate() + 1);
                                    dateDayOfWeek = date.getDay();
                                    dayTitle = titleDOW(dateDayOfWeek);
                                    theWeekdayVar = officeHoursData[dayTitle];
                                    startOfWorkDay = setToStartOfDay(date, theWeekdayVar);
                                    endOfWorkDay = setToEndOfDay(date, theWeekdayVar);
                                    dateSerial = startOfWorkDay;
                                    date = convertToDate(dateSerial); //converts date to excel serial for calculations

                                    //   dateMilli = date.getTime();
                                    //   bookendVars = startEndMidnight(date, theWeekdayVar);
                                };
                            
                            //#endregion ------------------------------------------------------------------------------------------------------------

                            //#region ADJUSTS DATES IN CASE REQUEST WAS SUBMITTED ON WEEKEND ----------------------------------------------------

                            if ((dateDayOfWeek == 6) || (dateDayOfWeek == 0)) { //if date was submitted on a weekend...
                                date = weekendAdjust(date, dateDayOfWeek);
                                dateDayOfWeek = date.getDay();
                                dayTitle = titleDOW(dateDayOfWeek);
                                theWeekdayVar = officeHoursData[dayTitle];
                                startOfWorkDay = setToStartOfDay(date, theWeekdayVar);
                                endOfWorkDay = setToEndOfDay(date, theWeekdayVar);
                                dateSerial = startOfWorkDay;
                                date = convertToDate(dateSerial); //converts date to excel serial for calculations

                                // dateMilli = date.getTime();
                                // bookendVars = startEndMidnight(date, theWeekdayVar);
                            };
                    
                        //#endregion ------------------------------------------------------------------------------------------------------------

                            //#region SETS ADJUSTMENT DATE VARIABLES -----------------------------------------------------------------------------------

                                //adds adjustmentNumber to date to get an adjustedDate value that will be used in later checks and calculations
                                var adjustedDate = new Date(date);
                                var adjustedDateSerial = Number(JSDateToExcelDate(adjustedDate));
                                adjustedDateSerial = adjustedDateSerial + adjustmentNumberSerial;
                                adjustedDate = convertToDate(adjustedDateSerial);

                            //#endregion ---------------------------------------------------------------------------------------------------------------
                        
                            //#region SETS ADD A DAY VARIABLES -----------------------------------------------------------------------------------------

                                //gets day of the week attributes for the day after the date variable
                                var nextDay = new Date(date);

                                var newNextDay = getNextDay(nextDay); //also sets this variable to the start time of the next day
                                var addADaySerial = newNextDay.nextDay;
                                var addADay = convertToDate(addADaySerial);
                                var addADayTitle = newNextDay.nextDayTitle;
                                var addADayWeekdayVar = officeHoursData[addADayTitle];
                                var addADayEnd = setToEndOfDay(addADay, addADayWeekdayVar);

                            //#endregion ----------------------------------------------------------------------------------------------------------------

                        //#endregion ----------------------------------------------------------------------------------------------------------------

                        //#region ACTION: SETS ADJUSTED DATE TO BE WITHIN OFFICE HOURS ------------------------------------------------------------------

                            //if adjustedDate falls outside of office hours, do this...
                            if (adjustedDateSerial < startOfWorkDay || adjustedDateSerial > endOfWorkDay) { //since the bookendVars is in reference to the date variable, this function will still trigger if adjustedDate is technically within office hours, but on a different day

                                //#region SETS ADJUSTMENT NUMBER VALUES ---------------------------------------------------------------------------------

                                    var dayRemainder = (endOfWorkDay - dateSerial) // / 1000) / 60) / 60; //time between end of work day and the original date time
                                    var remainingAdjust = adjustmentNumberSerial - dayRemainder; //gives us the remaining adjustment hours based off of what was already used to get to the end of the work day
                                    // var remainingAdjustMilli = remainingAdjust * 3600000;

                                //#endregion ------------------------------------------------------------------------------------------------------------

                                //#region NEW DAY CALCULATIONS ------------------------------------------------------------------------------------------

                                    var newDay = new Date(addADay);
                                    var newDaySerial = Number(JSDateToExcelDate(newDay));

                                    //adds remaining adjustment hours to the beginning of the work day the next day after date (addADay)
                                    var dateTimeAdjusted = newDaySerial + remainingAdjust;

                                    var dateTimeAdjustedConvert = convertToDate(dateTimeAdjusted); //convert serial number to date object

                                    date = dateTimeAdjustedConvert; //not sure if it should be date or something else yet. Need to make sure that the function works with this

                                //#endregion ------------------------------------------------------------------------------------------------------------

                                //#region SET LOOP VARIABLES IF STILL NOT WITHIN OFFICE HOURS OR EXCEEDS OFFICE HOURS OF NEXT DAY -----------------------

                                    //if the new date exceeds the office hours of addADay, then do this...
                                    if (dateTimeAdjusted > addADayEnd) {
                                    var addADayWorkDayLength = parseFloat(addADayWeekdayVar.workDay); //converts adjustment number from String to Number for calculations
                                    var addADayLengthMinutes = addADayWorkDayLength * 60;
                                    var addADayLengthSerial = minutesToSerial(addADayLengthMinutes);
                                    adjustmentNumber = (remainingAdjust - addADayLengthSerial) //subtracts remainingAdjust hours from the total workDay hours in the addADay variable
                                    var dayAfterTomorrow = new Date(addADay);
                                    var newDayAfterTomorrow = getNextDay(dayAfterTomorrow);
                                    dateSerial = newDayAfterTomorrow.nextDay;
                                    date = convertToDate(dateSerial);
                                    loop = true;
                                    var newAdjustmentNumber = convertToDate(adjustmentNumber);
                                    var cheesey = newAdjustmentNumber.getHours();
                                    var squeezy = newAdjustmentNumber.getMinutes();
                                    var truMinutes = squeezy/60;
                                    adjustmentNumber = cheesey + truMinutes;
                                    return {
                                        date,
                                        adjustmentNumber,
                                        loop
                                    };
                                    } else {
                                    loop = false;
                                    return {
                                        date,
                                        adjustmentNumber,
                                        loop
                                    };
                                    };

                                //#endregion -------------------------------------------------------------------------------------------------------------
                            
                            } else {
                                date = adjustedDate;
                                loop = false;
                                return {
                                date,
                                adjustmentNumber,
                                loop
                                };
                            };
                        
                        //#endregion --------------------------------------------------------------------------------------------------------------------

                    };

                //#endregion ---------------------------------------------------------------------------------------------------------------------------


                //#region TITLE DAY OF WEEK FUNCTION ---------------------------------------------------------------------------------------------------

                    /**
                     * Returns the weekday variable, with all it's associated properties, from the weekday index input value
                     * @param {Number} d The indexed number (0-6) of the weekday
                     * @returns An object with properties
                     */
                    function titleDOW(d) { //returns the day of the week (refered to directly in another variable) based on the dayID index number
                        if (d == 0) {
                        return "Sunday";
                        } else if (d == 1) {
                        return "Monday";
                        } else if (d == 2) {
                        return "Tuesday";
                        } else if (d == 3) {
                        return "Wednesday";
                        } else if (d == 4) {
                        return "Thursday";
                        } else if (d == 5) {
                        return "Friday";
                        } else if (d == 6) {
                        return "Saturday";
                        };
                    };

                //#endregion ----------------------------------------------------------------------------------------------------------------------------------


                //#region START/END/MIDNIGHT FUNCTIONS --------------------------------------------------------------------------------------------------

                    
                    //I used to use Milliseconds to do my calculations, but since I am loading in date serial #'s from the excel sheet that could chnage at any time,
                    //it makes more since it instead work within the Excel Serial Number and do all my calculations as serial instead of milliseconds that I then later convert to serial

                    //I also decided to break these apart into separate functions so I can reference them one at a time later on in the code
                
                    
                    //#region SET TO START OF THE WORK DAY --------------------------------------------------------------------------------------------------

                        /**
                         * Set date to the start of the workday based on the weekday
                         * @param {Date} date The date variable
                         * @param {Object} theWeekdayVar The object associated with the specific weekday including all of its properties
                         * @returns Date
                         */    
                        function setToStartOfDay(date, theWeekdayVar) {

                            var theDateBlank = new Date(date);
                            theDateBlank.setHours(0);
                            theDateBlank.setMinutes(0);
                            theDateBlank.setSeconds(0);
                            var theDateBlankSerial = Number(JSDateToExcelDate(theDateBlank));
                            //   var theDateBlankMilli = theDateBlank.getTime();

                            var startOfWorkDay = theDateBlankSerial + theWeekdayVar.startTime;

                            var startWorkDayReadable = convertToDate(startOfWorkDay);

                            return startOfWorkDay;

                        };

                    //#endregion ----------------------------------------------------------------------------------------------------------------------------


                    //#region SET TO END OF THE WORK DAY ----------------------------------------------------------------------------------------------------

                        /**
                         * Set the date to the end of the workday based on the weekday
                         * @param {Date} date The date variable
                         * @param {Object} theWeekdayVar The object associated with the specific weekday including all of its properties
                         * @returns Date
                         */
                        function setToEndOfDay(date, theWeekdayVar) {

                            var theDateBlank = new Date(date);
                            theDateBlank.setHours(0);
                            theDateBlank.setMinutes(0);
                            theDateBlank.setSeconds(0);
                            var theDateBlankSerial = Number(JSDateToExcelDate(theDateBlank));
                            //   var theDateBlankMilli = theDateBlank.getTime();

                            var endOfWorkDay = theDateBlankSerial + theWeekdayVar.endTime;

                            var endWorkDayReadable = convertToDate(endOfWorkDay);

                            return endOfWorkDay;

                        };

                    //#endregion ----------------------------------------------------------------------------------------------------------------------------


                    //#region SET TO MIDNIGHT ---------------------------------------------------------------------------------------------------------------

                        /**
                         * Sets date to serial number of the next day at midnight (very beginning of the day)
                         * @param {Date} date The date variable
                         * @returns Number
                         */
                        function setToMidnight(date) {

                            var midnight = new Date(date);
                            midnight.setDate(midnight.getDate() + 1);
                            midnight.setHours(0);
                            midnight.setMinutes(0);
                            midnight.setSeconds(0);
                            var midnightSerial = Number(JSDateToExcelDate(midnight));

                            return midnightSerial;

                        };

                    //#endregion -----------------------------------------------------------------------------------------------------------------------------


                //#endregion ----------------------------------------------------------------------------------------------------------------------------------


                //#region GET NEXT DAY FUNCTION --------------------------------------------------------------------------------------------------------

                    /**
                     * Adds a day to the date variable and sets it to the start time of that new day's day of the week. Also adjusts for weekends if needed.
                     * @param {Date} date A date object
                     * @returns An object with properties
                     */
                    function getNextDay(date) {

                        var nextDay = new Date(date);
                        var newNextDay = nextDay.setDate(nextDay.getDate() + 1); //returns the day after the original date
                        nextDay = new Date(newNextDay);
                        var nextDayDayOfWeek = nextDay.getDay();
                        var nextDayTitle = titleDOW(nextDayDayOfWeek); //returns a day title based on the dayID of the addADay variable
                        var theWeekdayVar = officeHoursData[nextDayTitle];

                        if ((nextDayDayOfWeek == 6) || (nextDayDayOfWeek == 0)) { //checks if nextDay falls on a weekend
                            nextDay = weekendAdjust(nextDay, nextDayDayOfWeek); //adjusts nextDay output to not fall on a weekend
                            nextDayDayOfWeek = nextDay.getDay();
                            nextDayTitle = titleDOW(nextDayDayOfWeek);
                            theWeekdayVar = officeHoursData[nextDayTitle];
                        };

                        nextDay = setToStartOfDay(nextDay, theWeekdayVar);

                        return {
                            nextDay,
                            nextDayTitle
                        };
                    };

                //#endregion ----------------------------------------------------------------------------------------------------------------------------------


                //#region MINUTES TO SERIAL ------------------------------------------------------------------------------------------------------------

                    /**
                     * Converts from a time in minutes to an Excel serial number, starting from the beginning of time (otherwise known as Dec 31, 1899)
                     * @param {Number} minutes A time in minutes
                     * @returns Number
                     */
                    function minutesToSerial(minutes) {
                    //   var date = new Date();
                    //   date.setDate(0);
                        var date = 0;
                        date = convertToDate(date);
                        date.setMinutes(minutes);
                        var numberSerial = Number(JSDateToExcelDate(date));
                        return numberSerial;
                    };

                //#endregion ----------------------------------------------------------------------------------------------------------------------------


                //#region WEEKEND ADJUST FUNCTION ------------------------------------------------------------------------------------------------------
                
                    /**
                     * If input date falls on a weekend, returns a new date adjusted to start on the next upcoming Monday
                     * @param {Date} date A date variable
                     * @param {Number} dateWeekday A number indexed 0-6 representing the weekday of the date variable
                     * @returns Date
                     */
                    function weekendAdjust(date, dateWeekday) {
                        if (dateWeekday == 6) {
                            var weekend = new Date(date);
                            weekend.setDate(weekend.getDate() + 2);
                            return weekend;
                        } else if (dateWeekday == 0) {
                            var weekend = new Date(date);
                            weekend.setDate(weekend.getDate() + 1);
                            return weekend;
                        };
                    };

                //#endregion ------------------------------------------------------------------------------------------------------------------------------


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


            //#endregion -------------------------------------------------------------------------------------------------------------------------------


        //#endregion -------------------------------------------------------------------------------------------------------------------------------------


        //#region PRIORITY GENERATION AND SORTATION ---------------------------------------------------------------------------------------

            /**
             * Generates a priority number for each row in the table based on the values in the Picked Up / Started By column. Also sorts the data by priority
             */
            async function priorityGenerationAndSortation() {

                await Excel.run(async (context) => {

                    var sheet = context.workbook.worksheets.getActiveWorksheet();
                    var sheetTable = sheet.tables.getItemAt(0);
                    var priorityColumnData = sheetTable.columns.getItem("Priority").getDataBodyRange().load("values");
                    var bodyRange = sheetTable.getDataBodyRange().load("values");
                    var headerRange = sheetTable.getHeaderRowRange().load("values");

                    await context.sync().then(function () {

                        // priorityColumnData.values.push([]);

                        var head = headerRange.values;

                        var pickedUpColumnIndex = findColumnIndex(head, "Picked Up / Started By"); //returns the index number of the "Picked Up / Started By" column based on it's position in the table header row

                        //need a function that will pull values from "pickedUpColumnIndex" position of the bodyRange.values for each row in sheet and put them in a new array

                        var activeTableValues = bodyRange.values; //loads all values of the active table

                        var pickedUpAllValuesArr = allColumnValues(activeTableValues, pickedUpColumnIndex); //makes an array of just the values from the Picked Up / Started By column
                        // pickedUpAllValuesArr.push(excelPickupOfficeHours);

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



//#region ERROR HANDLING ------------------------------------------------------------------------------------------

    //#region TRY CATCH ---------------------------------------------------------------------------------------------
    async function tryCatch(callback) {
        try {
            await callback();
        } catch (error) {
            console.error(error);
        }
    }
    //#endregion ---------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------



//#region ANTIQUATED FUNCTIONS (NO LONGER IN USE / ALTERNATE VERSIONS OF WORKING FUNCTIONS) -------------------------------

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

//#endregion ----------------------------------------------------------------------------------------------------------------------------------
