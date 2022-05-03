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
    var loop = true;
    var changeEvent;
    var snailPoop = {};

    var changedTable;

    var destinationTable;
    var destinationRows;
    var destinationTableRange;
    var destinationHeader;

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


//#region ON READY ---------------------------------------------------------------------------------------------------------------------------

    //#region LOADS VALIDATION VALUES AND UPDATES DROPDOWN VALUES IN TASKPANE ------------------------------------------------

        Office.onReady((info) => {

            // eventsOn();

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
                        var activeSheet = context.workbook.worksheets.getActiveWorksheet();
                        //activeSheet.onChanged.add(handleChange);
                        // context.runtime.load("enableEvents");





                        //get data from the tables
                        var productIDBodyRange = productIDValTable.getDataBodyRange().load("values");
                        var projectTypeIDBodyRange = projectTypeIDTable.getDataBodyRange().load("values");
                        var pickedUpBodyRange = pickedUpValTable.getDataBodyRange().load("values");
                        var proofToClientBodyRange = proofToClientValTable.getDataBodyRange().load("values");
                        var creativeProofBodyRange = creativeProofTable.getDataBodyRange().load("values");
                        var officeHoursBodyRange = officeHoursTable.getDataBodyRange().load("values");

                    //#endregion ----------------------------------------------------------------------------------------------

                    //eventsFunction();
                    changeEvent = context.workbook.tables.onChanged.add(onTableChanged);
                    //tryCatch(changeEvent);

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
                //updateDropDowns();
            };
        });

    //#endregion ---------------------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------------------


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



    //#region AUTO POPULATE TASKPANE BASED ON SUBJECT ---------------------------------------------------------------------------------------


        $("#subject").keyup(() => tryCatch(subjectPasted));
        //$("#subject").keyup(() => subjectPasted());



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



    //#region TASKPANE BUTTONS --------------------------------------------------------------------------------------------------------------

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
                    addAProjectEvents();
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


//#endregion --------------------------------------------------------------------------------------------------------------------------------


//#region EDITING THE TABLE ------------------------------------------------------------------------------------------------------------------


    //#region ADDING A PROJECT FROM TASKPANE -------------------------------------------------------------------------------------------------


        //#region TURN OFF EVENTS BEFORE EXECUTING ADD A PROJECT -------------------------------------------------------------------------------

            /**
             * Turns events off, then executes the addAProject function
             */
            async function addAProjectEvents() {
                await Excel.run(async (context) => {
                    context.runtime.load("enableEvents");
                    await context.sync();

                    //turns events off
                    context.runtime.enableEvents = false;
                    console.log("Events are turned off");
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

                        var sheet = context.workbook.worksheets.getActiveWorksheet();
                        //updating this variable to work for the changedTable will not work since the taskpane doesn't trigger an onchanged event until afterward
                        var sheetTable = sheet.tables.getItemAt(0); //this is fine since the user will only ever be adding new projects to the unassigned table or the artist tables, which are all the first tables in their documents. 
                        context.runtime.load("enableEvents");

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

                        write[0][14] = excelProofToClientOfficeHours;

                    //#endregion --------------------------------------------------------------------------------------------------------


                    //writes the write array to the table
                    sheetTable.rows.add(null /*add rows to the end of the table*/, write);

                    await context.sync(); // BOOM!


                    //#region PRIORITY NUMBER GENERATION AND SORTING -------------------------------------------------------------------------

                        //assign priority numbers and sorts table
                        priorityGenerationAndSortation();

                    //#endregion -------------------------------------------------------------------------------------------------------------


                    console.log("Add a project function has completed, but the sub-function for priority & sorting is still running asyncronously");

                    // context.runtime.enableEvents = true;
                    // console.log("Events are turned on");

                });

                // eventsOn();

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
                            return;
                        });

                    })
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

        // /**
        //  * Turns events off, then executes the onTableChanged function
        //  */
        // async function onTableChangedEvents(eventArgs) {

        //     console.log("I have been TRIGGERED!");

        //     await Excel.run(async (context) => {
        //         context.runtime.load("enableEvents");
        //         await context.sync();

        //         //turns events off
        //         context.runtime.enableEvents = false;
        //         console.log("Events are turned off");
        //         tryCatch(onTableChanged(eventArgs));
        //         // var result = await onTableChanged(eventArgs).then(tableChangedPriorityAndSort(poop.rowInfo, poop.bodyRange, poop.priorityColumnData));

        //     })//.then(() => {
        //     //     eventsOn();
        //     //     return;
        //     // });
        // };

    //#endregion ---------------------------------------------------------------------------------------------------------------------------


    //#region ON TABLE CHANGED --------------------------------------------------------------------------------------------------------------

        async function onTableChanged(eventArgs) {

            // console.log(eventArgs);

            await Excel.run(async (context) => {


                //#region LOAD VARIABLES FROM WORKBOOK -----------------------------------------------------------------------------------------

                    var details = eventArgs.details;
                    var address = eventArgs.address;
                    var changeType = eventArgs.changeType;
                    console.log(changeType);

                    var allWorksheets = context.workbook.worksheets;
                    allWorksheets.load("items/name/tables/id");
                    var changedWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId).load("name");
                    var worksheetTables = changedWorksheet.tables.load("items/name");
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
                            var unassignedTableRows = unassignedTable.rows.load("items");
                            var unassignedRange = unassignedTable.getDataBodyRange().load("values");
                            var unassignedHeader = unassignedTable.getHeaderRowRange().load("values");

                            var mattTable = context.workbook.tables.getItem("MattProjects").load("worksheet");
                            var mattTableRows = mattTable.rows.load("items");
                            var mattRange = mattTable.getDataBodyRange().load("values");
                            var mattHeader = mattTable.getHeaderRowRange().load("values");


                            // var alainaTable = context.workbook.tables.getItem("AlainaProjects").load("worksheet");
                            // var alainaHeader = alainaTable.getHeaderRowRange().load("values");


                            // var sarahTable = context.workbook.tables.getItem("SarahProjects").load("worksheet");
                            // var sarahHeader = sarahTable.getHeaderRowRange().load("values");


                            // var breBTable = context.workbook.tables.getItem("BreBProjects").load("worksheet");
                            // var breBHeader = breBTable.getHeaderRowRange().load("values");


                            // var christianTable = context.workbook.tables.getItem("ChristianProjects").load("worksheet");
                            // var christianHeader = christianTable.getHeaderRowRange().load("values");


                            // var emilyTable = context.workbook.tables.getItem("EmilyProjects").load("worksheet");
                            // var emilyHeader = emilyTable.getHeaderRowRange().load("values");


                            // var ianTable = context.workbook.tables.getItem("IanProjects").load("worksheet");
                            // var ianHeader = ianTable.getHeaderRowRange().load("values");


                            // var michaelTable = context.workbook.tables.getItem("MichaelProjects").load("worksheet");
                            // var michaelHeader = michaelTable.getHeaderRowRange().load("values");


                            // var joshTable = context.workbook.tables.getItem("JoshProjects").load("worksheet");
                            // var joshHeader = joshTable.getHeaderRowRange().load("values");


                            // var kristenTable = context.workbook.tables.getItem("KristenProjects").load("worksheet");
                            // var kristenHeader = kristenTable.getHeaderRowRange().load("values");


                            // var robinTable = context.workbook.tables.getItem("RobinProjects").load("worksheet");
                            // var robinHeader = robinTable.getHeaderRowRange().load("values");


                            // var lukeTable = context.workbook.tables.getItem("LukeProjects").load("worksheet");
                            // var lukeHeader = lukeTable.getHeaderRowRange().load("values");


                            // var lisaTable = context.workbook.tables.getItem("LisaProjects").load("worksheet");
                            // var lisaHeader = lisaTable.getHeaderRowRange().load("values");


                            // var luisTable = context.workbook.tables.getItem("LuisProjects").load("worksheet");
                            // var luisHeader = luisTable.getHeaderRowRange().load("values");


                            // var peterTable = context.workbook.tables.getItem("PeterProjects").load("worksheet");
                            // var peterHeader = peterTable.getHeaderRowRange().load("values");


                            // var ritaTable = context.workbook.tables.getItem("RitaProjects").load("worksheet");
                            // var ritaHeader = ritaTable.getHeaderRowRange().load("values");


                            // var ethanTable = context.workbook.tables.getItem("EthanProjects").load("worksheet");
                            // var ethanHeader = ethanTable.getHeaderRowRange().load("values");


                            // var breZTable = context.workbook.tables.getItem("BreZProjects").load("worksheet");
                            // var breZHeader = breZTable.getHeaderRowRange().load("values");


                            // var joeTable = context.workbook.tables.getItem("JoeProjects").load("worksheet");
                            // var joeHeader = joeTable.getHeaderRowRange().load("values");


                            // var jordanTable = context.workbook.tables.getItem("JordanProjects").load("worksheet");
                            // var jordanHeader = jordanTable.getHeaderRowRange().load("values");


                            // var hazelTable = context.workbook.tables.getItem("HazelProjects").load("worksheet");
                            // var hazelHeader = hazelTable.getHeaderRowRange().load("values");


                            // var toddTable = context.workbook.tables.getItem("ToddProjects").load("worksheet");
                            // var toddHeader = toddTable.getHeaderRowRange().load("values");

                        //#endregion ----------------------------------------------------------------------------------------------------

                    // ===================================================================================================================
                    // ===================================================================================================================




                    //TRYING TO FIGURE OUT CODE THAT CANCELS MOVE BETWEEN TABLES IF HEADERS DON'T MATCH \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    // \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    // for (var table of allTables) {
                    //     var headerVal = table.getHeaderRowRange().load("values")

                    //     snails[table.name] = {
                    //         headerVal
                    //     }
                    // }

                    // \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    // \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\



                    //all of the data from the changed table
                    var bodyRange = changedTable.getDataBodyRange().load("values");

                    //the header data of the changed table
                    var headerRange = changedTable.getHeaderRowRange().load("values");


                    context.runtime.load("enableEvents"); //turn on and off events

                //#endregion --------------------------------------------------------------------------------------------------------------


                await context.sync()

                    console.log("!!!!!!!");


                    //#region ASSIGNING VARIABLES -----------------------------------------------------------------------------------------------


                        //#region VALIDATION EXODUS AND TURNING OFF EVENTS -------------------------------------------------------------------------


                            //#region DON'T DO ANYTHING IF CHANGE WAS MADE TO VALIDATION SHEET -----------------------------------------------------

                                if (changedWorksheet.name == valSheet.name) { //if the change was made to the Validation sheet, exit the function
                                    console.log("Validation Sheet was changed, exiting the table changed event...")
                                    return;
                                };

                            //#endregion -----------------------------------------------------------------------------------------------------------


                            //#region TURN EVENTS OFF ----------------------------------------------------------------------------------------------

                                context.runtime.enableEvents = false; //turns events off
                                console.log("Events are turned off!!");

                            //#endregion -----------------------------------------------------------------------------------------------------------
                

                        //#endregion ---------------------------------------------------------------------------------------------------------------


                        //#region CREATING AND ASSIGNING WORKBOOK VARIABLES ------------------------------------------------------------------------

                            
                            //#region CALL LOADED VARIABLES ---------------------------------------------------------------------------------------

                                var changedColumnIndexOG = changedAddress.columnIndex; //index of the column where the change was made (on a worksheet level)
                                var changedRowIndex = changedAddress.rowIndex; //index of the row where the change was made (on a worksheet level)

                                var tableColumns = changedTableColumns.items; //loads all the changed table's columns
                                var tableRows = changedTableRows.items; //loads all the changed table's rows
                                var changedRowTableIndex = changedRowIndex - 1; //adjusts index number for table level (-1 to skip header row)
                                var rowValues = tableRows[changedRowTableIndex].values; //loads the values of the changed row in the changed table
                                var myRow = changedTableRows.getItemAt(changedRowTableIndex); //loads the changed row in the changed table as an object



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

                                //console.log("I am a fart"); //hehe

                            //#endregion ------------------------------------------------------------------------------------------------------------


                        //#endregion ----------------------------------------------------------------------------------------------------------------
                    

                    //#endregion ---------------------------------------------------------------------------------------------------------------


                    //#region ADJUST TURN AROUND TIMES, SORTING, & PRIORITY NUMBERS ------------------------------------------------------------

                        //if any of these columns are changed, turn around times will be adjusted and the table will be sorted
                        if (changedColumnIndex == rowInfo.pickedUpStartedBy.columnIndex || changedColumnIndex == rowInfo.proofToClient.columnIndex || changedColumnIndex == rowInfo.priority.columnIndex || changedColumnIndex == rowInfo.product.columnIndex || changedColumnIndex == rowInfo.projectType.columnIndex || changedColumnIndex == rowInfo.added.columnIndex || changedColumnIndex == rowInfo.startOverride.columnIndex || changedColumnIndex == rowInfo.workOverride.columnIndex) {
                            
                            console.log("I will update the turn around times, priority numbers, and sort the sheet before turning events back on!")

                            //adjusts picked up / started by turn around time
                            var lePickUpTime = getPickUpTime(rowInfo, leTable, changedRowTableIndex);

                            //adjusts proof to client turn around time
                            var leProofToClientTime = getProofToClientTime(rowInfo, leTable, lePickUpTime, changedRowTableIndex);

                            //sorts based on pickedUp column values and assigns priority numbers
                            var sortAndPrioritize = leSorting(rowInfo, leTable, pickedUpColumnIndex);

                            //writes updated values to the table
                            bodyRange.values = sortAndPrioritize; //overwrite changed table data with the new data from the sorted array 

                        };

                    //#endregion ----------------------------------------------------------------------------------------------------------------------


                    //#region MOVE DATA BETWEEN TABLES -----------------------------------------------------------------------------------------
                    
                        if (changedColumnIndex == rowInfo.artist.columnIndex || changedColumnIndex == rowInfo.status.columnIndex) {
                            console.log("Here is where all the complex move functions will take place!")

                            //#region FINDS IF CHANGE WAS MADE TO THE UNASSIGNED PROJECTS TABLE OR NOT ----------------------------------------

                                var isUnassigned;

                                if (changedWorksheet.name == "Unassigned Projects") {
                                    isUnassigned = true;
                                } else {
                                    isUnassigned = false;
                                };

                            //#endregion ------------------------------------------------------------------------------------------------------


                            //#region FINDS IF CHANGED TABLE IS A COMPLETED TABLE OR NOT ------------------------------------------------------

                                var listOfCompletedTables = [];

                                allTables.items.forEach(function (table) { //for each table in the workbook...
                                    if (table.name.includes("Completed")) { //if the table name includes the word "Completed" in it...
                                        listOfCompletedTables.push(table.name); //push the name of that table into an array
                                    };
                                });

                                //returns true if the changedTable is a completed table from the array previously made, false if it is anything else
                                var includesCompletedTables = listOfCompletedTables.includes(changedTable.name);

                            //#endregion ------------------------------------------------------------------------------------------------------

                            //LEGACY CODE THAT WORKS BUT NEEDS TO BE UPDATED TO BE MORE FLEXIBLE =================================================
                            // ===================================================================================================================

                                //#region ASSIGNS THE DESTINATION TABLE VALUE ---------------------------------------------------------

                                    if (rowInfo.artist.value == "Unassigned" && isUnassigned == false) {
                                        destinationTable = unassignedTable;
                                        destinationRows = unassignedTableRows.items;
                                        destinationTableRange = unassignedRange;
                                        destinationHeader = unassignedHeader;
                                    } else if (rowInfo.artist.value == "Matt") {
                                        destinationTable = mattTable;
                                        destinationRows = mattTableRows.items;
                                        destinationTableRange = mattRange;
                                        destinationHeader = mattHeader;
                                    // } else if (rowInfo.artist.value == "Alaina") {
                                    //     destinationTable = alainaTable;
                                    //     destinationHeader = alainaHeader;
                                    // } else if (rowInfo.artist.value == "Sarah") {
                                    //     destinationTable = sarahTable;
                                    //     destinationHeader = sarahHeader;
                                    // } else if (rowInfo.artist.value == "Bre B.") {
                                    //     destinationTable = breBTable;
                                    //     destinationHeader = breBHeader;
                                    // } else if (rowInfo.artist.value == "Christian") {
                                    //     destinationTable = christianTable;
                                    //     destinationHeader = christianHeader;
                                    // } else if (rowInfo.artist.value == "Emily") {
                                    //     destinationTable = emilyTable;
                                    //     destinationHeader = emilyHeader;
                                    // } else if (rowInfo.artist.value == "Ian") {
                                    //     destinationTable = ianTable;
                                    //     destinationHeader = ianHeader;
                                    // } else if (rowInfo.artist.value == "Michael") {
                                    //     destinationTable = michaelTable;
                                    //     destinationHeader = michaelHeader;
                                    // } else if (rowInfo.artist.value == "Josh") {
                                    //     destinationTable = joshTable;
                                    //     destinationHeader = joshHeader;
                                    // } else if (rowInfo.artist.value == "Kristen") {
                                    //     destinationTable = kristenTable;
                                    //     destinationHeader = kristenHeader;
                                    // } else if (rowInfo.artist.value == "Robin") {
                                    //     destinationTable = robinTable;
                                    //     destinationHeader = robinHeader;
                                    // } else if (rowInfo.artist.value == "Luke") {
                                    //     destinationTable = lukeTable;
                                    //     destinationHeader = lukeHeader;
                                    // } else if (rowInfo.artist.value == "Lisa") {
                                    //     destinationTable = lisaTable;
                                    //     destinationHeader = lisaHeader;
                                    // } else if (rowInfo.artist.value == "Luis") {
                                    //     destinationTable = luisTable;
                                    //     destinationHeader = luisHeader;
                                    // } else if (rowInfo.artist.value == "Peter") {
                                    //     destinationTable = peterTable;
                                    //     destinationHeader = peterHeader;
                                    // } else if (rowInfo.artist.value == "Rita") {
                                    //     destinationTable = ritaTable;
                                    //     destinationHeader = ritaHeader;
                                    // } else if (rowInfo.artist.value == "Ethan") {
                                    //     destinationTable = ethanTable;
                                    //     destinationHeader = ethanHeader;
                                    // } else if (rowInfo.artist.value == "Bre Z.") {
                                    //     destinationTable = breZTable;
                                    //     destinationHeader = breZHeader;
                                    // } else if (rowInfo.artist.value == "Joe") {
                                    //     destinationTable = joeTable;
                                    //     destinationHeader = joeHeader;
                                    // } else if (rowInfo.artist.value == "Jordan") {
                                    //     destinationTable = jordanTable;
                                    //     destinationHeader = jordanHeader;
                                    // } else if (rowInfo.artist.value == "Hazel-Rah") {
                                    //     destinationTable = hazelTable;
                                    //     destinationHeader = hazelHeader;
                                    // } else if (rowInfo.artist.value == "Todd") {
                                    //     destinationTable = toddTable;
                                    //     destinationHeader = toddHeader;
                                    } else {
                                        destinationTable = "null";
                                        destinationTableRange = "null";
                                        destinationHeader = "null";
                                    };

                                    //For the time being, I am recreating the variables from the changed table to work with the destination table.
                                    //I am replacing the changed row index with 0 since, at this point, there is no changed row in the destination table. We just need these values to essentially return the index number of the columns we want from the destination table in future functions.
                                    
                                    var destinationRange = destinationTableRange.values;

                                    var destRowValues = destinationRows[0].values;

                                    //var destRow = destTableRows.getItemAt(0);

                                    var destTable = JSON.parse(JSON.stringify(destinationRange));

                                    var destHead = destinationHeader.values;

                                    var destRowInfo = new Object();

                                    for (var name of destHead[0]) {
                                        theGreatestFunctionEverWritten(destHead, name, destRowValues, destTable, destRowInfo, 0)
                                    }


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

                                if (destinationTable !== "null" || destinationHeader !== "null") {

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

                                        var completedTable;

                                        worksheetTables.items.forEach(function (table) { //for each table in the changed worksheet...
                                            if (table.name.includes("Completed")) { //if the table name includes the word "Completed" in it...
                                            var leTable = table.name; //sets var to name of said completed table
                                            completedTable = worksheetTables.getItem(leTable); //grabs said table's data from the worksheet
                                            };
                                        });

                                    //#endregion ------------------------------------------------------------------------------------------------

                                    //#region MOVES DATA TO COMPLETED TABLE ---------------------------------------------------------------------

                                        if ((rowInfo.status.value == "Completed" || rowInfo.status.value == "Cancelled") && includesCompletedTables == false && isUnassigned == false) { //if status column = "Completed" or "Cancelled", the changedTable is not a Completed table, & the changedWorksheet is not UnassignedProjects, move data to changedWorksheet's completed table
                                            completedTable.rows.add(0, rowValues); //Adds empty row to bottom of the completedTable, then inserts the changed values into this empty row
                                            myRow.delete(); //Deletes the changed row from the original sheet
                                            console.log("Data was moved to the artist's Completed Projects Table!");
                                            return;
                                        } else if (rowInfo.status.value == "Editing" && includesCompletedTables == true) { //if status column = "Editing" & the changedTable is a Completed table, move data back to the artist's table
                                            if (destinationTable !== "null") {
                                            moveData(destinationTable, rowValues, myRow, rowInfo.artist.value);
                                            };
                                        };

                                    //#endregion ----------------------------------------------------------------------------------------------

                                };

                            //#endregion --------------------------------------------------------------------------------------------------


                            //#region MOVE DATA BASED ON ARTIST COLUMN --------------------------------------------------------------------

                                if (changedColumnIndex == rowInfo.artist.columnIndex) {

                                    //#region MOVES DATA TO DESTINATION TABLE -----------------------------------------------------------------

                                        //if (includesCompletedTables == false) {
                                        if (destinationTable !== "null") {
                                            if (destinationTable.worksheet.id !== changedWorksheet.id) { //if destination table is not in the same worksheet as the changedTable (prevents for unnecessary moving of data across tables in the same worksheet), do the following...
                                                //moveData(destinationTable, rowValues, myRow, rowInfo.artist.value);
                                                moveDataTwo(destTable, rowValues, leTable, changedRowTableIndex);

                                                if (changedTable.id == unassignedTable.id) { //if data is moving from the unassigned table to an artist table, sort this way...

                                                    var leTableSort = leSorting(rowInfo, leTable, pickedUpColumnIndex); //sorts the changed unassigned table by picked up / started by
                                                    var destTableSort = leSorting(destRowInfo, destTable, proofToClientColumnIndex); //sorts the destination artist table by proof to client

                                                } else if (destinationTable.id == unassignedTable.id) { //if data is moving from an artist table to the unassigned table, sort this way...

                                                    var leTableSort = leSorting(rowInfo, leTable, proofToClientColumnIndex); //sorts the changed artist table by proof to client
                                                    var destTableSort = leSorting(destRowInfo, destTable, pickedUpColumnIndex); //sorts the destination Unassigned table by picked up / started by

                                                } else if ((destinationTable.id !== unassignedTable.id) && (changedTable.id !== unassignedTable.id)) { //if data is moving between artist tables, both will be sorted by proof to client

                                                    var leTableSort = leSorting(rowInfo, leTable, proofToClientColumnIndex); //sorts the changed artist table by proof to client
                                                    var destTableSort = leSorting(destRowInfo, destTable, proofToClientColumnIndex); //sorts the destination arist table by proof to client

                                                };

                                                myRow.delete();
                                                //bodyRange.values = leTableSort;

                                                destinationTable.rows.add(null);
                                                //destinationTableRange.values = destTableSort;

                                                var bodyPositivity = changedTable.getDataBodyRange().load("values");

                                                var unassignedRange = unassignedTable.getDataBodyRange().load("values");
                                                var mattRange = mattTable.getDataBodyRange().load("values");

                                                if (rowInfo.artist.value == "Unassigned" && isUnassigned == false) {
                                                    var destinationStation = unassignedRange;
                                                } else if (rowInfo.artist.value == "Matt") {
                                                    var destinationStation = mattRange;
                                                } else {
                                                    var destinationStation = "null";
                                                };

                                                await context.sync()

                                                    var newBodyRange = bodyPositivity.values;
                                                    var newDestinationTableRange = destinationStation.values;
    
                                                    newBodyRange.values = leTableSort;
    
                                                    newDestinationTableRange.values = destTableSort;
    
    
    
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


                });

                eventsOn(); //turns events back on

            //});

        };

    //#endregion ---------------------------------------------------------------------------------------------------------------------------------


//#endregion --------------------------------------------------------------------------------------------------------------------------------


//#region FUNCTIONS ---------------------------------------------------------------------------------------------------------------------------


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
                console.log("Events are turned on!");
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
             * @returns Array
             */
                function leSorting(rowInfo, leTable, leColumnIndex) {

                //a copy of the array containing all the table data that will be used for sorting
                var leRowSorted = JSON.parse(JSON.stringify(leTable)); //creates a duplicate of original array to be used for assigning the priority numbers, without having anything done to it affect oriignal array
                
                var priorityColumnIndex = rowInfo.priority.columnIndex; //index of priority column

                //#region RESOLVE DUPLICATE DATE/TIMES ------------------------------------------------------------------------------

                    //if the pickedUp array has duplicate values, this nested for statement will add 1 second to the times of each duplicate value to allow the data sorting to work properly
                    for (var i = 0; i < leRowSorted.length; i++) {
                        for (var j = 0; j < leRowSorted.length; j++) {
                            if (i !== j) { //makes sure that the values do not equal (so the first pass will fail, naturally)
                                if (leRowSorted[i][leColumnIndex] == leRowSorted[j][leColumnIndex]) {
                                    console.log("A duplicate is present at index " + j + " of the array");
                                    leRowSorted[j][leColumnIndex] = leRowSorted[j][leColumnIndex] + 0.0000115740; //adds one second to the duplicate entry
                                };
                            };
                        };
                    };

                //#endregion -------------------------------------------------------------------------------------------------------


                //sorts the parent array (a) by the number in the sub array (b) at index of the picked up column
                leRowSorted.sort(function(a,b){return a[leColumnIndex] > b[leColumnIndex]});


                                            
                //#region ASSIGN PRIORITY NUMBERS ---------------------------------------------------------------------------------------

                    //for each item in the sorted array of table values, assign updated priority numbers to the priority column index
                    for (var n = 0; n < leRowSorted.length; n++) {
                        leRowSorted[n][priorityColumnIndex] = n + 1;
                    };

                //#endregion ------------------------------------------------------------------------------------------------------------

                return leRowSorted;

            };
        
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
        var leMovedTable = leTable.splice(changedRowTableIndex, 1);

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
            } else if (name == "Artist Lead") {
                codedHeader = "artistLead";
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
                try {
                    await callback();
                } catch (error) {
                    console.error(error);
                }
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