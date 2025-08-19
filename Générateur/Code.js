const SPREADSHEET_ID = '13ZZ3wdnhMtnZN1p2mGjj0g_5AhEvHZOZZ5MrgoanZiI';
const SHEET_NAME = 'Templates';

function doGet(e = {}) {
  const action = e.parameter?.action || 'html';

  if (action === 'json') {
    const data = chargerMessages();
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } else {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Marjanemall - Générateur de Messages')
      .setFaviconUrl('https://www.marjanemall.ma/favicon.ico')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function chargerMessages() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return {
        error: 'Feuille introuvable : ' + SHEET_NAME,
        debug: 'Feuilles disponibles : ' + ss.getSheets().map(s => s.getName()).join(', ')
      };
    }

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) {
      return {
        error: 'Aucune donnée trouvée dans la feuille.',
        debug: 'Lignes présentes : ' + values.length
      };
    }

    const headers = values.shift();

    const colIndices = {
      categorie: headers.indexOf('Catégorie'),
      objet: headers.indexOf('Objet'),
      titre: headers.indexOf('Titre'),
      message_fr: headers.indexOf('Message_FR'),
      message_ar: headers.indexOf('Message_AR')
    };

    const templates = values
      .filter(row => row[colIndices.categorie] && (row[colIndices.message_fr] || row[colIndices.message_ar]))
      .map(row => ({
        categorie: row[colIndices.categorie],
        objet: row[colIndices.objet],
        titre: row[colIndices.titre] || row[colIndices.objet],
        message_fr: row[colIndices.message_fr],
        message_ar: row[colIndices.message_ar]
      }));

    const categoriesMap = {};
    templates.forEach(tpl => {
      if (!categoriesMap[tpl.categorie]) {
        categoriesMap[tpl.categorie] = [];
      }
      categoriesMap[tpl.categorie].push(tpl);
    });

    return {
      templates,
      categories: Object.keys(categoriesMap),
      categoriesMap
    };
  } catch (error) {
    return {
      error: 'Erreur lors du chargement des modèles : ' + error.message,
      debug: 'SPREADSHEET_ID: ' + SPREADSHEET_ID + ', SHEET_NAME: ' + SHEET_NAME
    };
  }
}

function extractVariables(message) {
  const regex = /\[([^\[\]]+)\]/g;
  const variables = [];
  let match;
  while ((match = regex.exec(message)) !== null) {
    if (match[1]) {
      variables.push(match[1].trim());
    }
  }
  return [...new Set(variables)];
}

function testConnection() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return {
        success: false,
        error: "Feuille '" + SHEET_NAME + "' non trouvée",
        sheetsDisponibles: ss.getSheets().map(s => s.getName()).join(', ')
      };
    }
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];

    return {
      success: true,
      message: "Connexion réussie !",
      nomFeuille: sheet.getName(),
      nombreLignes: values.length,
      enTetes: headers,
      aperçuDonnées: values.slice(0, 3)
    };
  } catch (e) {
    return {
      success: false,
      error: "Erreur de connexion : " + e.toString(),
      spreadsheetId: SPREADSHEET_ID,
      sheetName: SHEET_NAME
    };
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Feuille "' + SHEET_NAME + '" introuvable.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    sheet.appendRow([
      data.categorie || '',
      data.objet || '',
      data.titre || '',
      data.message_fr || '',
      data.message_ar || '',
      new Date()
    ]);

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Données ajoutées avec succès.'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
