/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    document.getElementById("open-dialog").onclick = openDialog;
    document.getElementById("app").addEventListener("click", app);
  }
});

export async function app() {
  Excel.run(async (context) => {
    try {
      const BDD = await getBDD();
      console.log(BDD);
      await context.sync();
      const baseEtapes = BDD[0];
      const baseParents = BDD[1];
      const worksheetsEtapes = await initSheets(baseEtapes);
      //console.log(worksheetsEtapes);
      return context;
    } catch (error) {
      console.error(error);
      return context.sync();
    }
  });
}

async function initSheets(baseEtapes) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const worksheetsEtapes = [];
      console.log(baseEtapes);
      baseEtapes.forEach(async (etape) => {
        const id = etape[0]; //premier élément de la rangée est le nom de l'étape
        const nomEtape = etape[1]; //deuxième élément de la rangée est le nom de la feuille
        const nomFeuille = nomEtape + "|" + id;
        const sheetExists = await worksheetExists(nomFeuille);
        if (!sheetExists) {
          //obtenir le modèle de la feuille
          const sheetModel = context.workbook.worksheets.getItem("MODEL_" + nomEtape);
          await context.sync();
          //dupliquer la feuille de base
          if (worksheetsEtapes.length === 0) {
            const newSheet = sheetModel.copy();
            await context.sync();
            //set name of new sheet
            newSheet.name = nomFeuille;
            newSheet.visibility = "Visible";
            await context.sync();
            // ajouter la feille à la liste de feuilles
            worksheetsEtapes.push(newSheet);
          } else {
            const newSheet = sheetModel.copy(
              Excel.WorksheetPositionType.after,
              worksheetsEtapes[worksheetsEtapes.length - 1]
            );
            await context.sync();
            //set name of new sheet
            newSheet.name = nomFeuille;
            newSheet.visibility = "Visible";
            await context.sync();
            // ajouter la feille à la liste de feuilles
            worksheetsEtapes.push(newSheet);
          }
        }
      });
      await context.sync();
      resolve(worksheetsEtapes);
    }).catch((error) => reject(error));
  });
}

async function getBDD() {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const SHEET_NAME_BDD = "_BDD";
      const sheetBDD = context.workbook.worksheets.getItem(SHEET_NAME_BDD);
      const baseEtapes = sheetBDD.getRange("A2:B150").getUsedRange();
      const baseParents = sheetBDD.getRange("F2:H150").getUsedRange();
      baseEtapes.load("values");
      baseParents.load("values");
      await context.sync();
      resolve([baseEtapes.values, baseParents.values]);
    }).catch((error) => reject(error));
  });
}

async function worksheetExists(worksheetName) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      const worksheet = worksheets.getItemOrNullObject(worksheetName);
      await context.sync();
      resolve(worksheet.isNullObject ? false : true);
    }).catch((error) => reject(error));
  });
}

async function createTable() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company", "Communications", "120"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      ["1/11/2017", "Bellows College", "Education", "350.1"],
      ["1/15/2017", "Trey Research", "Other", "135"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"],
    ]);

    expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

async function filterTable() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
    const categoryFilter = expensesTable.columns.getItem("Category").filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);

    await context.sync();
  });
}

async function sortTable() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
    const sortFields = [
      {
        key: 1, // Merchant column
        ascending: false,
      },
    ];

    expensesTable.sort.apply(sortFields);
    await context.sync();
  });
}

let dialog = null;

function openDialog() {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/popup.html",
    { height: 45, width: 55 },

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
