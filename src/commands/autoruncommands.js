/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

async function insertParagraph(event) {
  Word.run(async (context) => {
    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph(
      "ExecuteFunction works. Button ID=" /*+ event.source.id*/,
      Word.InsertLocation.end
    );

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();

    const paragraphs = context.document.body.paragraphs;
    paragraphs.load();
    await context.sync();
    context.document.body.insertParagraph("current paragraph count: " + paragraphs.items.length, "End");
  });

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

Office.actions.associate("insertParagraph", insertParagraph);
