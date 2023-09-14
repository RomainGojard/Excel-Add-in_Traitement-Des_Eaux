/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var SHEET_NAME_BDD = "_BDD";
var BASE_ETAPES_NOM_TABLE = "baseEtapes";
var BASE_PARENTS_NOM_TABLE = "baseParents";
var SHEET_NAME_DONNEES_ENTREE = "Données_entrée";
var SHEET_NAME_TABLE_CONFIG = "Configuration - Entrées Sorties";
var TABLE_CONFIG_NOM = "tableConfig";
var LIGNE_PERTES_EN_EAU_SUR_CETTE_ETAPES_TABLEAU_SORTIE = 4;
var ADRESSES_PERTES_EN_EAU_DONNES_ENTREE = ["!C10", "!D10", "!E10", "!F10"];
var ADRESSES_DEBIT_JOURNALIER_EB_DONNES_ENTREE = ["!C12", "!D12", "!E12", "!F12"];

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    //document.getElementById("open-dialog").onclick = openDialog;
    document.getElementById("app").addEventListener("click", app);
  }
});

/**
 * Application principale qui est appelée lorsqu'on clique sur le bouton "Lancer le processus". Envoie une erreur si le processus ne fonctionne pas
 */
async function app() {
  Excel.run(async (context) => {
    try {
      await runProcess().catch((error) => {
        throw error;
      });
      openDialog("Processus terminé avec succès");
      return context;
    } catch (error) {
      // afficher le message d'erreur dans une popup
      errorPopUp(error);
      console.error(error);
      return context.sync();
    }
  });
}

/**
 * Fonction pour exécuter le processus
 */
async function runProcess() {
  return new Promise(async (resolve, reject) => {
    try {
      await sortBDD().catch((error) => {
        throw (
          error +
          " - La base de données est introuvable. Veuillez vérifier que vous utilisez l'application depuis le modèle prévu"
        );
      });
      const BDD = await getBDD();
      const baseEtapes = BDD[0];
      const baseParents = BDD[1];
      const worksheetsEtapes = await initSheets(baseEtapes);
      let allTables = await initAllTables();
      for (let i = 0; i < baseEtapes.length; i++) {
        const etape = baseEtapes[i];
        if (etape[0] == 1) {
          await EtapeUne(etape, worksheetsEtapes[0], allTables);
        } else {
          const idEtape = etape[0];
          const parents = baseParents.filter((ligne) => ligne[1] == idEtape);
          await EtapeN(etape, worksheetsEtapes[idEtape - 1], parents, baseEtapes, allTables);
        }
      }
      calculPertesEnEau(allTables, baseEtapes);
      resolve();
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * focntion qui initialise les feuilles d'étapes
 * @param {string[]} baseEtapes base d'étapes
 * @returns {Promise} une promesse qui renvoie un tableau contenant les noms des feuilles d'étapes
 */
async function initSheets(baseEtapes) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      let worksheetsEtapes = [];
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
      deleteOldEtapes(worksheetsEtapes);
      await context.sync();
      resolve(worksheetsEtapes);
      await context.sync();
    }).catch((error) => reject(error));
  });
}

/**
 * fonction qui rennvoie la base de données
 * @returns {Promise} une promesse qui renvoie la base de données
 */
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
      // si une erreur est renvoyée dans checkBDD, on arrête le processus
      checkBDD(baseEtapes.values, baseParents.values);
      resolve([baseEtapes.values, baseParents.values]);
    }).catch((error) => reject(error));
  });
}

/**
 * fonction qui vérifie que la base de données est correcte, renvoie une erreur sinon
 * @param {string[][]} baseEtapes base d'étapes
 * @param {*} baseParents base de parents
 */
function checkBDD(baseEtapes, baseParents) {
  if (baseEtapes.length == 0) {
    throw new Error("La base d'étapes est vide");
  }
  if (baseParents.length == 0) {
    throw new Error("La base de parents est vide");
  }

  //check que pour chaque ligne, toutes les colonnes sont remplies
  baseEtapes.forEach((etape) => {
    etape.forEach((colonne) => {
      if (colonne == null || colonne == "") {
        throw new Error("Une ligne de la base d'étapes est incomplète");
      }
    });
  });
  baseParents.forEach((parent) => {
    parent.forEach((colonne) => {
      if (colonne == null || colonne == "") {
        throw new Error("Une ligne de la base de parents est incomplète");
      }
    });
  });

  // check que id_etapes est unique
  const idEtapes = baseEtapes.map((etape) => etape[0]);
  const idEtapesUnique = [...new Set(idEtapes)];
  if (idEtapes.length != idEtapesUnique.length) {
    throw new Error("Les id des étapes ne sont pas uniques");
  }
  // check que le doublet (id_etape_parent, id_etape_enfant) est unique
  const idEtapesParents = baseParents.map((parent) => parent[0]);
  const idEtapesEnfants = baseParents.map((parent) => parent[1]);
  const idEtapesParentsEnfants = idEtapesParents.map((id, index) => [id, idEtapesEnfants[index]]);
  const idEtapesParentsEnfantsUnique = [...new Set(idEtapesParentsEnfants.map((id) => id.join()))];
  if (idEtapesParentsEnfants.length != idEtapesParentsEnfantsUnique.length) {
    throw new Error("Les id des parents et enfants ne sont pas uniques");
  }
  // check que la somme des flux (parent[2]) pour le même id_etape_parent est égale à 100
  const idEtapesParentsUnique = [...new Set(idEtapesParents)];
  idEtapesParentsUnique.forEach((id) => {
    const flux = baseParents.filter((parent) => parent[0] == id).reduce((acc, curr) => acc + curr[2], 0);
    if (flux != 100) {
      throw new Error(`La somme des flux pour l'étape ${id} est différente de 100`);
    }
  });
  // check qu'il y a au moins une étape associée à chaque id_etape_parent et id_etape_enfant
  idEtapesParentsEnfants.forEach((row) => {
    const idParent = row[0];
    const idEnfant = row[1];
    if (!idEtapes.includes(idParent)) {
      throw new Error(`L'étape ${idParent} n'existe pas`);
    }
    if (!idEtapes.includes(idEnfant)) {
      throw new Error(`L'étape ${idEnfant} n'existe pas`);
    }
  });

  // check que chaque id_etape apparaît dans la base de parents
  idEtapes.forEach((id) => {
    if (!idEtapesParents.includes(id) && !idEtapesEnfants.includes(id)) {
      throw new Error(`L'étape ${id} n'est pas associée à d'autres étapes`);
    }
  });
}

/**
 * Fonction qui renvoie un booléen indiquant si la feuille de nom worksheetName existe
 * @param {string} worksheetName nom de la feuille à vérifier
 * @returns {boolean} true si la feuille existe, false sinon
 */
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

/**
 * Fonction qui trie les tables de la base de données par ordre croissant de id_etapes et id_etape_parent
 */
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

/**
 * Fonction qui supprime les worksheets qui ne vont pas être utilisés
 * @param {string} tabNomWorksheet tableau des noms des worksheets qui vont être utilisés
 */
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

/**
 * génère les formules pour la première étape selon les données d'entrées
 * @param {string[]} etape ligne de la base d'étapes
 * @param {string} nomFeuille nom de la feuille cible
 */
async function EtapeUne(etape, nomFeuille, allTables) {
  const NOM_DONNEES_ENTREE = "DONNEES_ENTREES";

  return Excel.run(async (context) => {
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
    //const colonnesEtapeUneEntree = await obtenirColonnesParNomEnTete(SHEET_NAME_TABLE_CONFIG, TABLE_CONFIG_NOM, nomEtape + "_Entrée");
    const colonnesEtapeUneEntree = allTables[nomFeuille + "|" + nomEtape + "_Entree"];
    //vérifier que les colonnes fdonnes entrees et colonnesEtapeUneEntree ont la même longueur
    if (colonnesDonnesEntrees[0].length !== colonnesEtapeUneEntree.length) {
      console.log(colonnesDonnesEntrees[0] + " vs " + colonnesEtapeUneEntree);
      throw new Error("Les colonnes données d'entrées et colonnesEtapeUneEntree n'ont pas la même longueur : \n");
    }
    // parcourir pour i allant de 1 à la longueur de la colonne des données d'entrées
    for (let i = 1; i < colonnesDonnesEntrees[0].length; i++) {
      // for de 0 à 3 pour les 4 colonnes de la table de configuration
      for (let j = 0; j < 4; j++) {
        if (colonnesDonnesEntrees[j][i][0] !== "") {
          const parts = colonnesEtapeUneEntree[i][j].split("!");
          const targetCell = worksheet.getRange(parts[1]);
          // mettre l'addresse contenu dans donnees d'entrées dans la cellule en cours
          targetCell.values = [[`=${SHEET_NAME_DONNEES_ENTREE}!${colonnesDonnesEntrees[j][i][0]}`]];
        }
      }
    }
    await context.sync();
  });
}

/**
 * génère les formules pour les étapes qui ne sont pas la première selon les parents
 * @param {string[]} etape ligne de la base d'étapes
 * @param {string} nomFeuilleTarget nom de la feuille cible
 * @param {string[][]} parents tableau contenant les lignes de la base de parents qui sont les parents de l'étape
 * @param {string[][]} baseEtapes base d'étapes
 */
async function EtapeN(etape, nomFeuilleTarget, parents, baseEtapes, allTables) {
  const NOM_COLONNE_TYPE_DE_CHAMP = "TYPE_DE_CHAMP";
  await Excel.run(async (context) => {
    //nom de l'étape
    const nomEtapeTarget = etape[1];
    //get worksheet
    const worksheetTarget = context.workbook.worksheets.getItem(nomFeuilleTarget);
    await context.sync();
    // charger values de worksheetTarget
    worksheetTarget.load("values");
    await context.sync();
    // pour chaque parent de l'étape, on met dans tabSources les colonnes de données de sorties
    const tabSources = [];
    parents.forEach(async (parent) => {
      // obtenir le nomEtape du parent
      const nomEtapeParent = baseEtapes.find((row) => row[0] === parent[0])[1];
      //obtenir le nom du worksheet du parent
      const nomWorksheetParent = nomEtapeParent + "|" + parent[0];
      const result = allTables[nomWorksheetParent + "|" + nomEtapeParent + "_Sortie"];
      tabSources.push(result);
    });
    console.log(tabSources);
    const colonnesTarget = allTables[nomFeuilleTarget + "|" + nomEtapeTarget + "_Entree"];
    const colonneTypeChamp = await obtenirColonnesParNomEnTete(
      SHEET_NAME_TABLE_CONFIG,
      TABLE_CONFIG_NOM,
      NOM_COLONNE_TYPE_DE_CHAMP
    );

    //vérifier que les colonnes target colonnes de tabSources et type champ ont la même longueur
    if (!(colonnesTarget.length == tabSources[0].length && colonnesTarget.length == colonneTypeChamp[0].length)) {
      console.log(colonnesTarget[0] + " vs " + tabSources[0][0]);
      throw new Error("Les colonnes target et colonnes de tabSources n'ont pas la même longueur : \n");
    }

    //parcourir pour i allant de 1 à la longueur de la colonne de target
    for (let i = 1; i < colonnesTarget.length; i++) {
      // for de 0 à 3 pour les 4 colonnes de la table de configuration
      for (let j = 0; j < 4; j++) {
        const parts = colonnesTarget[i][j].split("!");
        const targetCell = worksheetTarget.getRange(parts[1]);
        switch (colonneTypeChamp[0][i][0]) {
          case "Débit": {
            targetCell.formulas = [[calculeCelluleDebit(tabSources, i, j, parents)]];
            break;
          }
          case "Concentration": {
            targetCell.formulas = [[calculeCelluleConcentration(tabSources, i, j, parents)]];
            break;
          }
          case "Température": {
            targetCell.formulas = [[calculeCelluleTemperature(tabSources, i, j, parents)]];
            break;
          }
          case "PH": {
            targetCell.formulas = [[calculeCellulePH(tabSources, i, j, parents)]];
            break;
          }
        }
      }
    }
    await context.sync();
  });
}

/**
 * fonction qui calcule le débit pour une cellule
 * @param {string[][][]} tabSources tableau qui contient les adresses des cellules des parents de l'étape [parent][ligne][colonne]
 * @param {int} i numéro de la ligne de la cellule
 * @param {int} j numéro de la colonne de la cellule
 * @returns {string} formule de la cellule selon la formule de débit (voir doc)
 */
function calculeCelluleDebit(tabSources, i, j, parents) {
  let result = "=";
  //boucle for sur les sources ou les parents
  for (let k = 0; k < parents.length; k++) {
    result += `(${tabSources[k][i][j]}*${parents[k][2]}/100)+`;
  }
  result = result.slice(0, -1);
  return result;
}

/**
 * fonction qui calcule la cooncentration pour une cellule
 * @param {string[][][]} tabSources tableau qui contient les adresses des cellules des parents de l'étape [parent][ligne][colonne]
 * @param {int} i numéro de la ligne de la cellule
 * @param {int} j numéro de la colonne de la cellule
 * @returns {string} formule de la cellule selon la formule de concentration (voir doc)
 */
function calculeCelluleConcentration(tabSources, i, j, parents) {
  let result = "=(";
  for (let k = 0; k < parents.length; k++) {
    result += `(${tabSources[k][1][j]}*${tabSources[k][i][j]}*${parents[k][2]}/100)+`;
  }
  result = result.slice(0, -1);
  result += ") / (";
  for (let k = 0; k < parents.length; k++) {
    result += `(${tabSources[k][1][j]}*${parents[k][2]}/100)+`;
  }
  result = result.slice(0, -1);
  result += ")";
  return result;
}

/**
 * fonction qui calcule la température pour une cellule
 * @param {string[][][]} tabSources tableau qui contient les adresses des cellules des parents de l'étape [parent][ligne][colonne]
 * @param {int} i numéro de la ligne de la cellule
 * @param {int} j numéro de la colonne de la cellule
 * @returns {string} formule de la cellule selon le calcul de la température (voir doc)
 */
function calculeCelluleTemperature(tabSources, i, j, parents) {
  let result = "=(";
  for (let k = 0; k < parents.length; k++) {
    result += `(${tabSources[k][1][j]}*${tabSources[k][i][j]}*${parents[k][2]}/100)+`;
  }
  result = result.slice(0, -1);
  result += ") / (";
  for (let k = 0; k < parents.length; k++) {
    result += `(${tabSources[k][1][j]}*${parents[k][2]}/100)+`;
  }
  result = result.slice(0, -1);
  result += ")";
  return result;
}

/**
 * fonction qui calcule le PH pour une cellule
 * @param {string[][][]} tabSources tableau qui contient les adresses des cellules des parents de l'étape [parent][ligne][colonne]
 * @param {int} i numéro de la ligne de la cellule
 * @param {int} j numéro de la colonne de la cellule
 * @returns {string} formule de la cellule
 */
function calculeCellulePH(tabSources, i, j) {
  return `=${tabSources[0][i][j]}`;
}

/**
 * fonction qui retourne les colonnes d'une table dont le nom de colonne commence par un préfixe
 * @param {string} nomFeuille nom du worksheet
 * @param {string} nomTableau nom de la table
 * @param {*} debutEnTete préfixe du nom de colonne
 * @returns {Promise<Array<Array<string>>>} tableau de colonnes
 */
async function obtenirColonnesParNomEnTete(nomFeuille, nomTableau, debutEnTete) {
  let result = [];
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

/**
 * Récupère les adresses de toutes les cellules de la plage de données d'une table dont le nom commence par une chaîne spécifiée
 * @param {string} nomWorksheet nom de la feuille
 * @param {string} tablePrefix prefixe du nom de la table
 * @returns {string[]} tableau des adresses des cellules de la plage de données de la table
 */
async function getTableAddressesByPrefix(nomWorksheet, tablePrefix) {
  try {
    // Charger l'API Excel
    const addresses = [];
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
      // Récupérer les adresses de chaque cellule de la plage de données
      const range = table.getDataBodyRange();
      range.load(["address", "values/length"]);
      await context.sync();
      // Stocker les adresses dans un tableau
      const rowCount = range.values.length;
      const colCount = range.values[0].length;
      for (let row = 0; row < rowCount; row++) {
        const rowAdresses = [];
        for (let col = 0; col < colCount; col++) {
          const cell = range.getCell(row, col);
          cell.load("address");
          await context.sync();
          rowAdresses.push(cell.address);
        }
        addresses.push(rowAdresses);
      }
    });
    // Retourner le tableau d'adresses
    return addresses;
  } catch (error) {
    console.error(error);
    return null;
  }
}

async function initAllTables() {
  let tablesDict = {};
  try {
    // Charger l'API Excel
    await Excel.run(async (context) => {
      await context.sync();
      var tables = context.workbook.tables;
      // chargement des tables
      tables.load("items/name");
      await context.sync();
      // itération à travers toutes les tables
      for (let i = 0; i < tables.items.length; i++) {
        let table = tables.items[i];
        let tableName = table.name;
        // si le nom de la table ne contient pas "Entree" ou "Sortie" on passe à la table suivante
        if (!tableName.includes("Entree") && !tableName.includes("Sortie")) {
          continue;
        }
        // obtenir le nom du classeur de la table
        let worksheet = table.worksheet;
        worksheet.load("name");
        const range = table.getDataBodyRange();
        range.load("address, values/length");
        await context.sync();
        const rowCount = range.values.length;
        let worksheetName = worksheet.name;
        let index = tableName.lastIndexOf("e");
        let tableShortName = worksheetName + "|" + tableName.slice(0, index + 1);
        const addresses = range.address;
        //
        const prefixAdresse = addresses.slice(0, addresses.indexOf("!") + 1);
        const premiereColonne = addresses[addresses.indexOf("!") + 1];
        const premiereLigne = Number(addresses.slice(addresses.indexOf("!") + 2, addresses.indexOf(":")));
        // construire tableau de 4 colonnes possibles commençant à la première colonne
        const colonnes = [
          premiereColonne,
          String.fromCharCode(premiereColonne.charCodeAt(0) + 1),
          String.fromCharCode(premiereColonne.charCodeAt(0) + 2),
          String.fromCharCode(premiereColonne.charCodeAt(0) + 3),
        ];
        // on construit un tableau contenant chaque adresse de cellule à partir de la range 'adresses' en parcourant les lignes de peremiereLigne à premiereLigne + range.values.length et les colonnes de colonnes
        const constTabAddresses = [tableShortName];
        for (let row = premiereLigne; row < premiereLigne + rowCount; row++) {
          const rowAdresses = [];
          for (let col = 0; col < colonnes.length; col++) {
            rowAdresses.push(prefixAdresse + colonnes[col] + row);
          }
          constTabAddresses.push(rowAdresses);
        }
        tablesDict[tableShortName] = constTabAddresses;
      }
    });
    return tablesDict;
  } catch (error) {
    console.error(error);
    throw error;
  }
}

/**
 * Fonction qui permet d'afficher une popup d'erreur
 * @param {string} error message d'erreur
 */
function errorPopUp(error) {
  Office.context.ui.displayDialogAsync(
    // eslint-disable-next-line no-undef
    window.location.origin + "/error.html?erreur=" + error,
    { height: 25, width: 35 },
    (result) => {
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const data = JSON.parse(arg.message);
        dialog.close();
      });
    }
  );
}

async function calculPertesEnEau(allTables, baseEtapes) {
  let resultCol0 = "=( ";
  let resultCol1 = "=( ";
  let resultCol2 = "=( ";
  let resultCol3 = "=( ";
  console.log(allTables);
  baseEtapes.forEach((etape) => {
    const nomTable = etape[1] + "|" + etape[0].toString() + "|" + etape[1] + "_Sortie";
    const table = allTables[nomTable];
    const ligne = table[LIGNE_PERTES_EN_EAU_SUR_CETTE_ETAPES_TABLEAU_SORTIE];
    resultCol0 += ligne[0] + "+";
    resultCol1 += ligne[1] + "+";
    resultCol2 += ligne[2] + "+";
    resultCol3 += ligne[3] + "+";
  });
  // on enlève le dernier +
  resultCol0 = resultCol0.slice(0, resultCol0.length - 1);
  resultCol1 = resultCol1.slice(0, resultCol1.length - 1);
  resultCol2 = resultCol2.slice(0, resultCol2.length - 1);
  resultCol3 = resultCol3.slice(0, resultCol3.length - 1);

  resultCol0 += " ) / '" + SHEET_NAME_DONNEES_ENTREE + "'" + ADRESSES_DEBIT_JOURNALIER_EB_DONNES_ENTREE[0];
  resultCol1 += " ) / '" + SHEET_NAME_DONNEES_ENTREE + "'" + ADRESSES_DEBIT_JOURNALIER_EB_DONNES_ENTREE[1];
  resultCol2 += " ) / '" + SHEET_NAME_DONNEES_ENTREE + "'" + ADRESSES_DEBIT_JOURNALIER_EB_DONNES_ENTREE[2];
  resultCol3 += " ) / '" + SHEET_NAME_DONNEES_ENTREE + "'" + ADRESSES_DEBIT_JOURNALIER_EB_DONNES_ENTREE[3];

  console.log(resultCol0, resultCol1, resultCol2, resultCol3);

  Excel.run(async function (context) {
    //get données d'entrées
    const donneesEntrees = context.workbook.worksheets.getItem(SHEET_NAME_DONNEES_ENTREE);
    // Obtenez la plage de cellules qui contient la cellule à modifier
    var range0 = donneesEntrees.getRange(
      "'" + SHEET_NAME_DONNEES_ENTREE + "'" + ADRESSES_PERTES_EN_EAU_DONNES_ENTREE[0]
    );
    var range1 = donneesEntrees.getRange(
      "'" + SHEET_NAME_DONNEES_ENTREE + "'" + ADRESSES_PERTES_EN_EAU_DONNES_ENTREE[1]
    );
    var range2 = donneesEntrees.getRange(
      "'" + SHEET_NAME_DONNEES_ENTREE + "'" + ADRESSES_PERTES_EN_EAU_DONNES_ENTREE[2]
    );
    var range3 = donneesEntrees.getRange(
      "'" + SHEET_NAME_DONNEES_ENTREE + "'" + ADRESSES_PERTES_EN_EAU_DONNES_ENTREE[3]
    );
    await context.sync();

    // Modifiez la valeur de la cellule à "nouvelleValeur"
    range0.formulas = [[resultCol0]];
    range1.formulas = [[resultCol1]];
    range2.formulas = [[resultCol2]];
    range3.formulas = [[resultCol3]];
    // Exécutez les modifications de la feuille de calcul
    return context.sync();
  }).catch(function (error) {
    console.log("Une erreur est survenue : " + error);
  });
}

function openDialog(message) {
  Office.context.ui.displayDialogAsync(
    window.location.origin + "/popup.html?message=" + message,
    { height: 25, width: 35 },
    (result) => {
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const data = JSON.parse(arg.message);
        dialog.close();
      });
    }
  );
}
