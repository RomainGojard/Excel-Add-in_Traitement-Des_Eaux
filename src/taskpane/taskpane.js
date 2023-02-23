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
    const SHEET_NAME_BDD = "_BDD";
    const sheetBDD = context.workbook.worksheets.getItem(SHEET_NAME_BDD);
    const baseEtapes = sheetBDD.getRange("A2:B150").getUsedRange();
    baseEtapes.load("values");
    await context.sync();
    baseEtapes.values.forEach(async (etape) => {
      const id = etape[0]; //premier élément de la rangée est le nom de l'étape
      const nomEtape = etape[1]; //deuxième élément de la rangée est le nom de la feuille
      const nomFeuille = nomEtape + "|" + id;
      //vérifier s'il n'y a pas déjà une feuille avec ce nom
      const sheetExists = context.workbook.worksheets.getItemOrNullObject(nomFeuille);
      await context.sync();
      console.log(sheetExists.isNullObject);
      if (sheetExists.isNullObject) {
        //obtenir le modèle de la feuille
        const sheetModel = context.workbook.worksheets.getItem("MODEL_" + nomEtape);
        //dupliquer la feuille de base
        const newSheet = sheetModel.copy(Excel.WorksheetPositionType.before, sheetBDD);
        await context.sync();
        //set name of new sheet
        newSheet.name = nomFeuille;
        newSheet.visibility = "Visible";
        await context.sync();
      }
    });
    return context;
  });
}

/*const baseEtapes = [
    [1, "Acidification"],
    [2, "Clarif"],
    [3, "FAS"],
    [4, "Acidification"],
    [5, "Désinf"],
  ];*/
//const baseEtapes = sheetBDD.getRange("A2:B150");
//baseEtapes.load("values");
/*
    baseEtapes.values.forEach((etape) => {
      const id = etape[0]; //premier élément de la rangée est le nom de l'étape
      const nomEtape = etape[1]; //deuxième élément de la rangée est le nom de la feuille
      const nomFeuille = nomEtape + "|" + id;
      //dupliquer la feuille de base
      sheetBDD.copy(nomFeuille);
    });
    */

// await context.sync(); //synchroniser le contexte avec le serveur Excel après la création de toutes les feuilles

async function duplicateSheet(sheetName, newSheetName) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.copy(newSheetName, Excel.WorksheetPositionType.after);
    await context.sync();
  });
}

export async function getBDD() {
  const SHEET_NAME_BDD = "_BDD";
  Excel.run(async (context) => {
    const sheetBDD = context.workbook.worksheets.getItem(SHEET_NAME_BDD);
    const allBaseEtapes = sheetBDD.getRange("A2:B150");
    allBaseEtapes.load("values");
    const ctx = await context.sync();
    const baseEtapes = [];
    baseEtapes.values.forEach((etape) => {
      if (etape[0] !== "") {
        baseEtapes.push(etape);
      }
    });
    console.log(baseEtapes);
    return ctx;
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
