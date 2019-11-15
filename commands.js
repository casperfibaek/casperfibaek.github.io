Office.onReady(() => {
    // If needed, Office.js is ready to be called
});
  
function toggleProtection(args) {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load('protection/protected');
        return context.sync()
            .then(function () {
            if (sheet.protection.protected) {
                sheet.protection.unprotect();
            }
            else {
                sheet.protection.protect();
            }
        })
            .then(context.sync);
    })["catch"](function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
    args.completed();
}

function openDialog(event) {
    Office.context.ui.displayDialogAsync('localhost', {
        height: 40,
        width: 30
    });

    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

function openDialog1(event) {
    Office.context.ui.displayDialogAsync('https://localhost/index.html', {
        height: 40,
        width: 30
    });

    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

function openDialog2(event) {
    Office.context.ui.displayDialogAsync('https://google.com', {
        height: 40,
        width: 30
    });

    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

function openDialog3(event) {
    Office.context.ui.displayDialogAsync('about:blank', {
        height: 40,
        width: 30
    });

    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
        ? global
        : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;
g.openDialog = openDialog;
