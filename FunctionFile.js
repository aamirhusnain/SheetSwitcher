// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

function Add(first, second) {


    var values = [
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
    ];

    // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {
        // Create a proxy object for the active sheet
        var sheet = ctx.workbook.worksheets.getActiveWorksheet();
        // Queue a command to write the sample data to the worksheet
        sheet.getRange("B3:D5").values = values;


       
        // Run the queued-up commands, and return a promise to indicate task completion
        return ctx.sync();
    })
   
  return first + second;
}
CustomFunctions.associate("ADD", Add);


function ShowTaskPane(inputValue) {
    if (inputValue == "show") {
        Office.addin.showAsTaskpane();
    }
    else {
        Office.addin.hide();
    }
    var taskpane = inputValue;
    return taskpane;
}
CustomFunctions.associate("ShowTaskPane", ShowTaskPane);
