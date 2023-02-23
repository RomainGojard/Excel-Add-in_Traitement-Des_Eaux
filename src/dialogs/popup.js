/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  document.getElementById("ok-button").onclick = () => tryCatch(sendStringToParentPage);
});

function sendStringToParentPage() {
  const userName = Document.getElementById("name-box").value;
  Office.context.ui.messageParent(userName);
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    Console.error(error);
  }
}

function getSheetNames() {
  const sheets = Excel.Sheets;
  const sheetNames = [];
  for (let i = 0; i < sheets.count; i++) {
    sheetNames.push(sheets.getItemAt(i).name);
  }
  return sheetNames;
}
