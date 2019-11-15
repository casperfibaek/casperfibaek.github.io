/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }

  Excel.run(function (context) {        
    context.workbook.customFunctions.addAll();
    return context.sync().then(function(){});
  }).catch(function(error){});

    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
  if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
    console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
  }

  var dialog = null;

  function openDialog() {
    Office.context.ui.displayDialogAsync(
      'https://localhost/SATF/addin/dialog/popup.html',
      {height: 45, width: 55},

      function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      }
    );
  }

  function processMessage(arg) {
    document.getElementById("user-name").innerHTML = arg.message;
    dialog.close();
  }


  // Assign event handlers and other initialization logic.
  document.getElementById("run").onclick = run;
});
