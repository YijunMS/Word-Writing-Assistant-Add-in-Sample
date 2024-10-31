/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

async function addTextBox(event) {
  // Implement your custom code here. The following code is a simple Excel example.
  try {
    await PowerPoint.run(function (context) {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const textbox = shapes.addTextBox("Hello!", {
        left: 100,
        top: 300,
        height: 300,
        width: 450,
      });
      textbox.name = "Textbox";

      return context.sync();
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

const g = getGlobal();

// The add-in command functions need to be available in global scope

Office.actions.associate("addTextBox", addTextBox);
