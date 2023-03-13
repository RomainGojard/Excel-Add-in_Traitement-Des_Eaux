/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

//test

var SHEET_NAME_BDD = "_BDD";
var BASE_ETAPES_NOM_TABLE = "baseEtapes";
var BASE_PARENTS_NOM_TABLE = "baseParents";
var SHEET_NAME_DONNEES_ENTREE = "Données_entrée";
var SHEET_NAME_TABLE_CONFIG = "Configuration - Entrées Sorties";
var TABLE_CONFIG_NOM = "tableConfig";

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
      await sortBDD();
      const BDD = await getBDD();
      await context.sync();
      const baseEtapes = BDD[0];
      const baseParents = BDD[1];
      const worksheetsEtapes = await initSheets(baseEtapes);
      baseEtapes.forEach(async (etape) => {
        if (etape[0] == 1) {
          EtapeUne(etape, worksheetsEtapes[0]);
        } else {
          const idEtape = etape[0];
          // obtenir les parents de l'étape en cours depuis la base parents
          const parents = baseParents.filter((ligne) => ligne[1] == idEtape);
          EtapeN(etape, worksheetsEtapes[idEtape - 1], parents, baseEtapes, worksheetsEtapes);
        }
      });
      await context.sync();
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
      let worksheetsEtapes = [];
      let newSheet = null;
      const sheetBDD = context.workbook.worksheets.getItem(SHEET_NAME_BDD);
      for (const etape of baseEtapes) {
        const id = etape[0]; //premier élément de la rangée est le nom de l'étape
        const nomEtape = etape[1]; //deuxième élément de la rangée est le nom de la feuille
        const nomFeuille = nomEtape + "|" + id;
        const sheetExists = await worksheetExists(nomFeuille);
        if (!sheetExists) {
          //obtenir le modèle de la feuille
          const sheetModel = context.workbook.worksheets.getItem("MODEL_" + nomEtape);
          await context.sync();
          //dupliquer la feuille de base
          let newSheet = sheetModel.copy("Before", sheetBDD);
          await context.sync();
          //set name of new sheet
          newSheet.name = nomFeuille;
          newSheet.visibility = "Visible";
        }
        await context.sync();
        worksheetsEtapes.push(nomFeuille);
      }
      await context.sync();
      deleteOldEtapes(worksheetsEtapes);
      resolve(worksheetsEtapes);
    }).catch((error) => reject(error));
  });
}

async function getBDD() {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const sheetBDD = context.workbook.worksheets.getItem(SHEET_NAME_BDD);
      const baseEtapes = sheetBDD.tables.getItem(BASE_ETAPES_NOM_TABLE).getRange().getUsedRange();
      // eslint-disable-next-line prettier/prettier
      const baseParents = sheetBDD.tables.getItem(BASE_PARENTS_NOM_TABLE).getRange().getUsedRange();
      baseEtapes.load("values");
      baseParents.load("values");
      await context.sync();
      //enlevr la première ligne qui contient les noms des colonnes
      baseEtapes.values.shift();
      baseParents.values.shift();
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

async function sortBDD() {
  await Excel.run(async (context) => {
    const sheetBDD = context.workbook.worksheets.getItem(SHEET_NAME_BDD);
    const baseEtapes = sheetBDD.tables.getItem(BASE_ETAPES_NOM_TABLE);
    const baseParents = sheetBDD.tables.getItem(BASE_PARENTS_NOM_TABLE);
    const sortFieldsEtapes = [
      {
        key: 0, // colonne id_etapes
        ascending: true,
      },
    ];

    const sortFieldsParents = [
      {
        key: 0, // colonne id_etape_parent
        ascending: true,
      },
    ];

    baseEtapes.sort.apply(sortFieldsEtapes);
    baseParents.sort.apply(sortFieldsParents);
    await context.sync();
  });
}

async function deleteOldEtapes(tabNomWorksheet) {
  await Excel.run(async (context) => {
    // Récupération de tous les worksheets du classeur
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();
    // On parcourt tous les worksheets et on supprime ceux qui ne sont pas dans la base (et qui ne sont pas le modèle, ou la table de configuration)
    for (let i = 0; i < worksheets.items.length; i++) {
      const worksheet = worksheets.items[i];
      if (worksheet.name.includes("|") && !tabNomWorksheet.includes(worksheet.name)) {
        worksheet.delete();
      }
    }
    await context.sync();
  });
}

async function EtapeUne(etape, nomFeuille) {
  console.log("CGVUBHIUONUGCFXGCYGVUBHIJNOK");
  const NOM_DONNEES_ENTREE = "DONNEES_ENTREES";
  await Excel.run(async (context) => {
    //nom de l'étape
    const nomEtape = etape[1];
    //get worksheet
    const worksheet = context.workbook.worksheets.getItem(nomFeuille);
    //get données d'entrées
    const donneesEntrees = context.workbook.worksheets.getItem(SHEET_NAME_DONNEES_ENTREE);
    //charger les values
    worksheet.load("values");
    donneesEntrees.load("values");
    await context.sync();
    const colonnesDonnesEntrees = await obtenirColonnesParNomEnTete(
      SHEET_NAME_TABLE_CONFIG,
      TABLE_CONFIG_NOM,
      NOM_DONNEES_ENTREE
    );
    // eslint-disable-next-line prettier/prettier
    const colonnesEtapeUneEntree = await obtenirColonnesParNomEnTete(SHEET_NAME_TABLE_CONFIG, TABLE_CONFIG_NOM, nomEtape + "_Entrée");
    //vérifier que les colonnes fdonnes entrees et colonnesEtapeUneEntree ont la même longueur
    if (colonnesDonnesEntrees[0].length !== colonnesEtapeUneEntree[0].length) {
      console.log(colonnesDonnesEntrees[0] + " vs " + colonnesEtapeUneEntree[0]);
      throw new Error("Les colonnes données d'entrées et colonnesEtapeUneEntree n'ont pas la même longueur : \n");
    }
    // parcourir pour i allant de 1 à la longueur de la colonne des données d'entrées
    for (let i = 1; i < colonnesDonnesEntrees[0].length; i++) {
      // for de 0 à 3 pour les 4 colonnes de la table de configuration
      for (let j = 0; j < 4; j++) {
        const targetCell = worksheet.getRange(colonnesEtapeUneEntree[j][i][0]);
        // mettre l'addresse contenu dans donnees d'entrées dans la cellule en cours
        targetCell.values = [[`=${SHEET_NAME_DONNEES_ENTREE}!${colonnesDonnesEntrees[j][i][0]}`]];
      }
    }
    await context.sync();
  });
}

async function EtapeN(etape, nomFeuilleTarget, parents, baseEtapes) {
  const NOM_COLONNE_TYPE_DE_CHAMP = "TYPE_DE_CHAMP";
  await Excel.run(async (context) => {
    //nom de l'étape
    const nomEtapeTarget = etape[1];
    //id de l'étape
    const idEtape = etape[0];
    //get worksheet
    const worksheetTarget = context.workbook.worksheets.getItem(nomFeuilleTarget);
    await context.sync();
    // charger values de worksheetTarget
    worksheetTarget.load("values");
    await context.sync();
    await getTableDataByPrefix(nomFeuilleTarget, nomEtapeTarget + "_Entree");
    // pour chaque parent de l'étape, on met dans tabSources les colonnes de données de sorties
    const tabSources = [];
    parents.forEach(async (parent) => {
      // obtenir le nomEtape du parent
      const nomEtapeParent = baseEtapes.find((row) => row[0] === parent[0])[1];
      //obtenir le nom du worksheet du parent
      const nomWorksheetParent = nomEtapeParent + "|" + parent[0];
      const result = await obtenirColonnesParNomEnTete(
        SHEET_NAME_TABLE_CONFIG,
        TABLE_CONFIG_NOM,
        nomEtapeParent + "_Sortie"
      );
      result.push(nomWorksheetParent);
      tabSources.push(result);
    });
    const colonnesTarget = await obtenirColonnesParNomEnTete(
      SHEET_NAME_TABLE_CONFIG,
      TABLE_CONFIG_NOM,
      nomEtapeTarget + "_Entrée"
    );
    const colonneTypeChamp = await obtenirColonnesParNomEnTete(
      SHEET_NAME_TABLE_CONFIG,
      TABLE_CONFIG_NOM,
      NOM_COLONNE_TYPE_DE_CHAMP
    );

    //vérifier que les colonnes target colonnes de tabSources et type champ ont la même longueur
    if (
      colonnesTarget[0].length !== tabSources[0][0].length &&
      colonnesTarget[0].length !== colonneTypeChamp[0].length &&
      colonneTypeChamp[0].length !== tabSources[0][0].length
    ) {
      console.log(colonnesTarget[0] + " vs " + tabSources[0][0]);
      throw new Error("Les colonnes target et colonnes de tabSources n'ont pas la même longueur : \n");
    }
    console.log(tabSources);
    console.log(colonnesTarget);
    console.log(colonneTypeChamp[0]);

    //parcourir pour i allant de 1 à la longueur de la colonne de target
    for (let i = 1; i < colonnesTarget[0].length; i++) {
      // for de 0 à 3 pour les 4 colonnes de la table de configuration
      for (let j = 0; j < 4; j++) {
        const targetCell = worksheetTarget.getRange(colonnesTarget[j][i][0]);
        switch (colonneTypeChamp[0][i][0]) {
          case "Débit": {
            targetCell.formulas = [[calculeCelluleDebit(tabSources, j, i, parents)]];
            break;
          }
          case "Concentration": {
            targetCell.formulas = [[calculeCelluleConcentration(tabSources, j, i, parents)]];
            break;
          }
          case "Température": {
            targetCell.formulas = [[calculeCelluleTemperature(tabSources, j, i, parents)]];
            break;
          }
          case "PH": {
            targetCell.formulas = [[calculeCellulePH(tabSources, j, i, parents)]];
            break;
          }
        }
      }
    }
    await context.sync();
  });
}

function calculeCelluleDebit(tabSources, j, i, parents) {
  let result = "=";
  //boucle for sur les sources ou les parents
  for (let k = 0; k < parents.length; k++) {
    result += `('${tabSources[k].slice(-1)}'!${tabSources[k][j][i][0]}*${parents[k][2]}/100)+`;
  }
  result = result.slice(0, -1);
  return result;
}

function calculeCelluleConcentration(tabSources, j, i, parents) {
  let result = "=(";
  for (let k = 0; k < parents.length; k++) {
    result += `('${tabSources[k].slice(-1)}'!${tabSources[k][j][1][0]}*'${tabSources[k].slice(-1)}'!${tabSources[k][j][i][0]}*${parents[k][2]}/100)+`;
  }
  result = result.slice(0, -1);
  result += ") / (";
  for (let k = 0; k < parents.length; k++) {
    result += `('${tabSources[k].slice(-1)}'!${tabSources[k][j][1][0]}*${parents[k][2]}/100)+`;
  }
  result = result.slice(0, -1);
  result += ")";
  return result;
}
function calculeCelluleTemperature(tabSources, j, i, parents) {
  let result = "=";
  //boucle for sur les sources ou les parents
  for (let k = 0; k < parents.length; k++) {
    result += `('${tabSources[k].slice(-1)}'!${tabSources[k][j][i][0]}*${parents[k][2]}/100)+`;
  }
  result = result.slice(0, -1);
  return result;
}

function calculeCellulePH(tabSources, j, i, parents) {
  return `='${tabSources[0].slice(-1)}'!${tabSources[0][j][i][0]}`;
}

async function obtenirColonnesParNomEnTete(nomFeuille, nomTableau, debutEnTete) {
  const result = [];
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(nomFeuille);
    const table = worksheet.tables.getItem(nomTableau).getRange().getUsedRange();
    table.load("values");
    await context.sync();
    const headers = table.values[0];
    const indices = headers
      .map((header, index) => (header.startsWith(debutEnTete) ? index : -1))
      .filter((index) => index >= 0);
    const columnsObjects = await Promise.all(
      indices.map((index) => table.getColumn(index).getUsedRange().load("values"))
    );
    await context.sync();
    const columns = columnsObjects.map((columnValues) => {
      return columnValues.values;
    });
    result.push(...columns);
  });
  return result;
}

async function getTableDataByPrefix(nomWorksheet, tablePrefix) {
  try {
    // Charger l'API Excel
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(nomWorksheet);
      sheet.load("tables");
      await context.sync();
      // Récupérer les tables dans le workbook
      const tables = sheet.tables;
      tables.load("items/name");
      // Exécuter les requêtes
      await context.sync();
      // Trouver la première table dont le nom commence par la chaîne spécifiée
      const table = tables.items.find((t) => t.name.startsWith(tablePrefix));
      if (!table) {
        console.error(`Table with prefix "${tablePrefix}" not found`);
        return null;
      }
      // Récupérer les données du tableau
      const range = table.getDataBodyRange();
      range.load("values");
      // Exécuter les requêtes
      await context.sync();
      // Retourner les données du tableau
      console.log(range.values);
    });
  } catch (error) {
    console.error(error);
  }
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

let dialog = null;

function openDialog() {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/", //"https://csb10032001a6800cf9.z6.web.core.windows.net/popup.html",
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
