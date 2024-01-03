
await Excel.run(async () => {
    var worksheetTables = context.workbook.worksheets.getActiveWorksheet().tables.load("items/count");
    await context.sync();
    console.log(worksheetTables.getItemAt(0).onChanged)
})
