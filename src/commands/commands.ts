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
