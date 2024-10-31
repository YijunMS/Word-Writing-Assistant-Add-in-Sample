/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

async function insertParagraph(event) {
  // Implement your custom code here. The following code is a simple Excel example.
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph("Hello World", Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    //console.error(error);
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

async function insertTable(event) {
  // Implement your custom code here. The following code is a simple Excel example.
  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.getItemOrNullObject("Sample").delete();
      const sheet = context.workbook.worksheets.add("Sample");

      let expensesTable = sheet.tables.add("A1:E1", true);
      expensesTable.name = "SalesTable";
      expensesTable.getHeaderRowRange().values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"]];

      expensesTable.rows.add(null, [
        ["Frames", 5000, 7000, 6544, 4377],
        ["Saddles", 400, 323, 276, 651],
        ["Brake levers", 12000, 8766, 8456, 9812],
        ["Chains", 1550, 1088, 692, 853],
        ["Mirrors", 225, 600, 923, 544],
        ["Spokes", 6005, 7634, 4589, 8765],
      ]);

      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();

      sheet.activate();
      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    //console.error(error);
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

async function addTextBox(event) {
  // Implement your custom code here. The following code is a simple Excel example.
  try {
    await PowerPoint.run(async (context) => {
      const shapes: PowerPoint.ShapeCollection = context.presentation.slides.getItemAt(0).shapes;
      const shapeOptions: PowerPoint.ShapeAddOptions = {
        left: 100,
        top: 300,
        height: 300,
        width: 450,
      };
      const textbox: PowerPoint.Shape = shapes.addTextBox("Hello!", shapeOptions);

      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    //console.error(error);
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
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

const g = getGlobal() as any;

Office.actions.associate("insertParagraph", insertParagraph);
Office.actions.associate("insertTable", insertTable);
Office.actions.associate("addTextBox", addTextBox);
