//CREATIVE REQUEST -Alfredo's Pizza - West Babylon - MENU - ~/*1338,52130,1*/~
//CREATIVE REQUEST -Bella Napoli - Canfield - Env #10 8.5x11 S2 - ~/*1837,65845,1*/~
//Re: Artist Request - Brickhouse Pizzeria - Richfield Springs - MENU - ~/*30601,72301,1*/~


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

// var lookup = new Object();
var productIDData = {};
var projectTypeIDData = {};
var pickupData = {};
var proofToClientData = {};
var creativeProofData = {};
var officeHoursData = {};


Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {


        Excel.run(async (context) => {

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



            await context.sync();


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
                        "endTime":row[2]
                    };
                };

                console.log(officeHoursData);

            //#endregion -----------------------------------------------------------------------------

        });
        // console.log(info);
        tryCatch(updateDropDowns);
        
    };
});








$("#container").each(function () {
    $(this).html($(this).html().replace(/(\*)/g, '<span style="color: rgba(220, 20, 60, 0.50); font-size: 9pt; padding-left: 1px; padding-bottom: 1px;">$1</span>'));
});



$("#submit").on("click", function() {
    // var clientWarningText = `Client name required to submit request`;
    // var productWarningText = `Product required to submit request`;
    // var projectTypeWarningText = `Project Type required to submit request`;

    if ((($("#client").val()) == "") || (($("#product").val()) == null) || (($("#project-type").val()) == null)) {
        addWarningClass("#client", "#warning2");
        addWarningClass("#product", "#warning3");
        addWarningClass("#project-type", "#warning4");
        return;
    };

    if ((($("#client").val()) !== "") || (($("#product").val()) !== null) || (($("#project-type").val()) !== null)) {
        removeWarningClass("#client", "#warning2");
        removeWarningClass("#product", "#warning3");
        removeWarningClass("#project-type", "#warning4");
        ugh();
    };
});



$("#client").on("focusout", function() {

    removeWarningClass("#client", "#warning2");

});

$("#product").on("focusout", function() {

    removeWarningClass("#product", "#warning3");

});

$("#project-type").on("focusout", function() {

    removeWarningClass("#project-type", "#warning4");

});










function removeWarningClass(object, warning) {

    if ((($(object).val()) !== "") || (($(object).val()) !== null)) {

        $(warning).hide(); // Don't show the error
        $(object).removeClass("warning-box")
        $(object).removeClass("warning-box + .label")
    
    }
    
};

function addWarningClass(object, warning) {

    var cheese = $(object).val();

    if ((($(object).val()) == "") || (($(object).val()) == null)) {

        $(warning).show().text(`Required`); //show error
        $(object).addClass("warning-box")
        $(object).addClass("warning-box + .label")
    
    };

};



$("#clear").on("click", function() {

    $("#subject, #client, #location, #product, #code, #project-type, #csm, #print-date, #group, #artist-lead, #queue, #tier, #tags, #start-override, #work-override").val(""); // Empty all inputs
    removeWarningClass("#subject", "#warning1");
    removeWarningClass("#client", "#warning2");
    removeWarningClass("#product", "#warning3");
    removeWarningClass("#project-type", "#warning4");

});



async function ugh() {

    await Excel.run(async (context) => {

        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var sheetTable = sheet.tables.getItemAt(0);
        // let bodyRange = sheetTable.getDataBodyRange().load("values");

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





        //Generate Date Added
        var now = new Date();
        var toSerial = JSDateToExcelDate(now);

        write[0][10] = toSerial;





        //get the Pickup Turn Around Time Values based on product and project type returned from arrays

        let theProjectTypeCode = projectTypeIDData[projectTypeVal].projectTypeCode;

        //this is the code that doesn't work that I am trying to have return the value from the pickupturnaroundtime table
        let pickedUpTurnAroundTime = pickupData[productVal].theProjectTypeCode;


        //but this works, just not on the level I need it to
        let newPickedUpTurnAroundTime = pickupData[productVal];
        console.log(pickedUpTurnAroundTime);

        //This is the value I want to get, but it only works when ".brandNewBuild" is not a string
        var test = pickedUpTurnAroundTime.brandNewBuild;





        //add start override time to # of hours
        var overrideTime = pickedUpTurnAroundTime + startOverrideVal;
    



        //add new time to date added, then adjust for office hours
        var adjustedPUTime = toSerial + overrideTime;






        sheetTable.rows.add(null /*add rows to the end of the table*/, write);



        await context.sync(); // BOOM!
    });
}


$("#subject").keyup(() => tryCatch(subjectPasted));

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

        var plasticS = removeFirstAndLastSpace(noBlanksArr[noBlanksArr.length - 2]);

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

function removeFirstAndLastSpace(splitItem) {
    var firstChar = splitItem.charAt(0);

    if (firstChar == " ") {
            splitItem = splitItem.slice(1);
    };

    var lastChar = splitItem.charAt(splitItem.length - 1);

    if (lastChar == " ") {
            splitItem = splitItem.slice(0, splitItem.length - 1);
    };

    return splitItem;
};


async function updateDropDowns() {
    await Excel.run(async (context) => {
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

// productID(product, option).then((updatedProduct) => {
//     // Async functions return promises
//     console.log(updatedProduct);
// });
// });


async function productID(product, option) {
    
    var relativeProduct;

    await Excel.run(async (context) => {
        var sheet = context.workbook.worksheets.getItem("Validation");
        var productIDValTable = sheet.tables.getItem("ProductIDTable");

        // Get data from the table.
        var productIDBodyRange = productIDValTable.getDataBodyRange().load("values");

        await context.sync();

        //#region PRODUCT ID VALUES -----------------------------------------------------------------------------------

        var poop = productIDBodyRange.values;

        if (option == 1) {
            for (var row of poop) {
                var check = row[0].trim();
                // var netProduct = removeFirstAndLastSpace(row[0]);
                if (row[0].trim() == product) {
                    relativeProduct = row[1];
                    break;
                };
            };



            // if (option == 1) {

            //     var relativeProduct;

            //     for (var row of productIDBodyValues) {
            //         var a = row;
            //         var netProduct = removeFirstAndLastSpace(row[0]);
            //         if (netProduct == product) {
            //             relativeProduct = row[1];
            //             return relativeProduct;
            //         }
            //     }

                // productIDBodyValues.forEach(function(row) {

                //     var netProduct = removeFirstAndLastSpace(row[0]);

                //     if (netProduct == product) {
                //         relativeProduct = row[1];
                //         return relativeProduct;
                //     };
    
                // });

                // return relativeProduct;

            };


            if (option == 2) {

                var code;

                productIDBodyValues.forEach(function(row) {

                    if (row[1] == product) {
                        code = row[2];
                    };
    
                });

                return code;

            }





            //     // Add an option to the select box
            //     var option = `<option product-id="${row[0]}" relative-product="${row[1]}" product-code="${row[2]}">${row[1]}</option>`;

            //     var x = $(`#product > option[relative-product="${row[1]}"]`).length; //finds current relative-product in current option in the product dropdown and returns how many are currently in the dropdown

            //     if (x == 0) { // Meaning, it's not there yet, because it's length count is 0
            //         if (row[1] !== "") { //if the relative-product in option is empty, do not add to list
            //             $("#product").append(option);
            //         };
            //     };
             //});

        //#endregion ---------------------------------------------------------------------------------------------------

    });

    return relativeProduct;

};

    //#region CONVERT SERIAL TO DATE ----------------------------------------------------------------------------------------------

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