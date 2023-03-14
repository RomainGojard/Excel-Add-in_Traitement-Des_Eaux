/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(() => {
  $("#my-form").submit((event) => {
    event.preventDefault();
    const id_etape = $("#id_etape").val();
    const typeEtape = $("#typeEtape").val();
    Office.context.ui.messageParent(JSON.stringify({ id_etape, typeEtape }));
  });

  $("#cancel-form").click(() => {
    Office.context.ui.messageParent("cancel");
  });
});

var sheet = null;

Office.onReady(async () => {
  await Excel.run(async (context) => {
    sheet = context.workbook.worksheets.getItem("_BDD");
    const baseEtapes = sheet.tables.getItem("baseEtapes");
    const range = baseEtapes.getDataBodyRange();
    range.load("values");
    await context.sync();
    const data = range.values;
    remplirFormulaire(data);
  });
});

async function initForm() {
  await Excel.run(async (context) => {
    sheet = context.workbook.worksheets.getItem("_BDD");
    await context.sync();
    const baseEtapes = sheet.tables.getItem("baseEtapes");
    await context.sync();
    const range = baseEtapes.getDataBodyRange();
    range.load("values");
    await context.sync();
    console.log("Données de la table chargées !");
    console.log(range.values);
    remplirFormulaire(range.values);
    await context.sync();
  });
}

function remplirFormulaire(data) {
  var id = data.length + 1;
  document.getElementById("id").value = id;
  var etape = data[0][1];
  document.getElementById("etape").value = etape;
}
