/*
* Project Name: Pavora
* Copyright (c) 2025 Nicola Rainiero
*
* This software is released under the MIT License.
* Please refer to the LICENSE file for the full license text.
*/
function trygetRealEmail() {
  Logger.log('La modalità alias è '+aliasEmail());
  var email = Session.getEffectiveUser().getEmail();
  Logger.log('Email=' + email);
  Logger.log('RealEmail=' + getRealEmail(email));
  Logger.log('AliasEmail=' + getAliasEmail(email));
  var alias = 'administrator.user@gmail.com';
  Logger.log('Alias=' + alias);
  Logger.log('RealEmail=' + getRealEmail(alias));
  Logger.log('AliasEmail=' + getAliasEmail(alias));
  var real = users()[0][0];
  Logger.log('Real=' + real);
  Logger.log('RealEmail=' + getRealEmail(real));
  Logger.log('AliasEmail=' + getAliasEmail(real));  
}

function getRealEmail(input, idSheet) {
  try {
    // Se aliasEmail è spento (0), restituisce sempre l'input
    if (!aliasEmail()) {
      return input;
    }
    idSheet = idSheet || IDAliasEmail; // oppure: mode = mode || "compact"; 
    const ss = SpreadsheetApp.openById(IDAliasEmail);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    
    const searchInput = String(input).trim();
    
    // Cerca in entrambe le colonne
    for (let i = 1; i < data.length; i++) {
      const alias = String(data[i][0]).trim();
      const realEmail = String(data[i][1]).trim();
      
      // Se l'input corrisponde all'alias, restituisce l'email reale
      if (alias === searchInput) {
        return realEmail;
      }
      
      // Se l'input corrisponde già all'email reale, la restituisce
      if (realEmail === searchInput) {
        return realEmail;
      }
    }
    
    // Se non trovato in nessuna colonna, restituisce l'input originale
    return input;
    
  } catch (error) {
    console.error("Errore in getRealEmail:", error);
    return input; // In caso di errore, restituisce l'input originale
  }
}

function getAliasEmail(input, idSheet) {
  try {
    // Se aliasEmail è spento (0), restituisce sempre l'input
    if (!aliasEmail()) {
      return input;
    }
    idSheet = idSheet || IDAliasEmail; // oppure: mode = mode || "compact"; 
    const ss = SpreadsheetApp.openById(IDAliasEmail);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    
    const searchInput = String(input).trim();
    
    // Cerca in entrambe le colonne
    for (let i = 1; i < data.length; i++) {
      const alias = String(data[i][0]).trim();
      const realEmail = String(data[i][1]).trim();
      
      // Se l'input corrisponde all'email reale, restituisce l'alias
      if (realEmail === searchInput) {
        return alias;
      }
      
      // Se l'input corrisponde già all'alias, lo restituisce
      if (alias === searchInput) {
        return alias;
      }
    }
    
    // Se non trovato in nessuna colonna, restituisce l'input originale
    return input;
    
  } catch (error) {
    console.error("Errore in getAliasEmail:", error);
    return input; // In caso di errore, restituisce l'input originale
  }
}

function exportActiveSheetToXlsx() {
  try {
    // Ottieni il foglio di calcolo attivo e il foglio attivo
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var activeSheet = SpreadsheetApp.getActiveSheet();
    var sheetName = activeSheet.getName();

    // Crea una copia del foglio di calcolo
    var copySpreadsheet = spreadsheet.copy(sheetName + " (Esportazione)");
    var copySpreadsheetId = copySpreadsheet.getId();
    var copySheets = copySpreadsheet.getSheets();

    // Elimina tutti i fogli eccetto quello attivo nella copia
    copySheets.forEach(function (sheet) {
      if (sheet.getName() !== sheetName) {
        copySpreadsheet.deleteSheet(sheet);
      }
    });

    // Esporta il file con solo il foglio attivo
    var url = "https://docs.google.com/spreadsheets/d/" + copySpreadsheetId + "/export?format=xlsx&id=" + copySpreadsheetId;
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });

    // Crea un file temporaneo su Google Drive
    var blob = response.getBlob();
    var today = new Date();
    blob.setName(formatDateMaster(today).dataXfile + '_' + sheetName + ".xlsx");
    var tempFile = DriveApp.createFile(blob);

    // Elimina la copia temporanea
    DriveApp.getFileById(copySpreadsheetId).setTrashed(true);

    // Restituisce il link per il download
    return tempFile.getDownloadUrl();

  } catch (e) {
    Logger.log("Errore durante l'esportazione: " + e.message);
    throw new Error("Errore durante l'esportazione del foglio attivo in formato XLSX: " + e.message);
  }
}

// filter id refOp names based on the value in the fourth column (group)
// groupName(refCom(), 2)
function groupName(matrix, index) {
  return matrix
    .filter(row => row[3] === index) // check if the row[3] (=group) is equal to 'index'
    .map(row => row[1]) // Return the row[1] (=id) element 
}

function eliminaFogliNonPresentiNelVettore() {
  var nomiConsentiti = []; // Sostituisci con i nomi dei fogli che vuoi mantenere
  for (let i = 0; i < sheetsList().length; i++) {
    nomiConsentiti.push(sheetsList()[i][0]);
  }
  for (let i = 0; i < users().length; i++) {
    nomiConsentiti.push(users()[i][0]);
  }
  //Logger.log(nomiConsentiti);

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var fogli = spreadsheet.getSheets();

  fogli.forEach(function (foglio) {
    var nomeFoglio = foglio.getName();
    if (!nomiConsentiti.includes(nomeFoglio)) {
      spreadsheet.deleteSheet(foglio); // Elimina il foglio se non è nel vettore
    }
  });

}

function formatEmail(email) {

  if (aliasEmail()) {
    email = getAliasEmail(email);
  }
  // Controlla se l'email contiene tre punti
  const parts = email.split('@')[0].split('.');

  let formattedName = '';

  if (parts.length === 3) {
    // Se ci sono tre parti, formattare come Nome Cognome
    formattedName = `${capitalizeFirstLetter(parts[0])} ${capitalizeFirstLetter(parts[1])}`;
  } else if (parts.length === 2) {
    formattedName = `${capitalizeFirstLetter(parts[0])}`;
  } else {
    // Se non ci sono tre parti, prendi solo il primo segmento
    formattedName = capitalizeFirstLetter(parts[0].split('@')[0].substring(0, 8));
  }

  return formattedName;
}

// Funzione ausiliaria per capitalizzare la prima lettera di una stringa
function capitalizeFirstLetter(string) {
  return string.charAt(0).toUpperCase() + string.slice(1).toLowerCase();
}

// Sharing a Calendar in Google App Script
// https://stackoverflow.com/questions/27094575/sharing-a-calendar-in-google-app-script
// Add user as reader
//var rule = shareCalendar( 'gobbledygook@group.calendar.google.com',
//                          'user@example.com');
/**
 * Set up calendar sharing for a single user. Refer to 
 * https://developers.google.com/google-apps/calendar/v3/reference/acl/insert.
 *
 * @param {string} calId   Calendar ID
 * @param {string} user    Email address to share with
 * @param {string} role    Optional permissions, default = "reader":
 *                         "none, "freeBusyReader", "reader", "writer", "owner"
 *
 * @returns {aclResource}  See https://developers.google.com/google-apps/calendar/v3/reference/acl#resource
 */

function shareCalendar(calId, user, role) {
  role = role || "reader";

  try {
    // Verifica che l'indirizzo email sia nel formato corretto
    if (!validateEmail(user)) {
      //Logger.log("Indirizzo email non valido: " + user);
      return null;
    }

    // Assicurati che il ruolo sia valido
    // I ruoli validi sono: "none", "freeBusyReader", "reader", "writer", "owner"
    if (!["none", "freeBusyReader", "reader", "writer", "owner"].includes(role)) {
      //Logger.log("Ruolo non valido: " + role);
      return null;
    }

    // Crea l'oggetto ACL con la struttura corretta
    var acl = {
      scope: {
        type: "user",
        value: user
      },
      role: role
    };

    // Debug: mostra l'oggetto che stiamo per inviare
    //Logger.log("Invio ACL: " + JSON.stringify(acl));

    // Esegui l'inserimento
    var newRule = Calendar.Acl.insert(acl, calId);
    //Logger.log("Accesso concesso a " + user + " con ruolo " + role);
    return newRule;
  } catch (e) {
    Logger.log("Errore dettagliato: " + e.toString());
    return null;
  }
}

// Funzione per validare il formato dell'email
function validateEmail(email) {
  var regex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return regex.test(email);
}


function randomID(length) {
  var result = '';
  var characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var charactersLength = characters.length;
  for (let i = 0; i < length; i++) {
    result += characters.charAt(Math.floor(Math.random() * charactersLength));
  }
  return result;
}
// console.log(randomID(2)); --> C0

//Convert Hex to r g b
//console.log(hexToRgb("#0033ff").g); // "51";
function hexToRgb(hex) {
  var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    r: parseInt(result[1], 16) * 0.0039215,
    g: parseInt(result[2], 16) * 0.0039215,
    b: parseInt(result[3], 16) * 0.0039215
  } : null;
}

// email --> destinatario email, fromail --> chi la spedisce, subject, messaggio,
// mandaEmail(data, utenteEmail, emailTarget()[0][0], emailTarget()[0][1], evento, subject, tipoAggiunta, JSON.stringify(array));
function mandaEmail(data, email, fromMail, sender, evento, subject, tipoAggiunta, testoMatrice) {

  var message = translate('admin.emailAuto');
  var message = message + '<br><table><tr><th>Date</th><th>Action</th><th>ID Event</th><th>User Email</th><th>Details</th></tr>';
  var message = message + '<tr><td>' + data + '</td><td>' + tipoAggiunta + '</td><td>' + evento + '</td><td>' + fromMail + '</td><td>' + testoMatrice + '</td></tr>';

  var message = message + '</table>';

  var message = message + '<br>';

  var message = message + translate('admin.emailCheck') + sender + '</b>';

  var message = message + "<br><br>*<a href=\"mailto:" + fromMail + "\">" + fromMail + "</a><br>";

  var message = message + "<br><br><small><i>This e-mail and any attachment here to are intended only for the person or entity to which is addressed and may contain information that is privileged, confidential or otherwise protected from disclosure. Copying, dissemination or use of this e-mail or the information herein by anyone other than the intended recipient is prohibited. If you have received this e-mail by mistake, please notify us immediately by telephone or fax and delete it from your system.</i></small>";

  GmailApp.sendEmail(
    email,
    subject,
    message,
    {
      from: fromMail,
      name: sender,
      htmlBody: message
    });
}

function deleteSlidesUntil(number, presentationId) {
  /*
  if (SlidesApp.getActivePresentation().getId() != null) {
    var presentationId = SlidesApp.getActivePresentation().getId();
    var presentation = SlidesApp.openById(presentationId);
  } else {
    var presentation = SlidesApp.openById(presID());
  }
  */
  // ID della presentazione esistente
  var presentationId = presentationId;

  // Step 1: Aprire la presentazione esistente
  var presentation = SlidesApp.openById(presentationId);

  var slides = presentation.getSlides();

  //Logger.log(typeof slides);
  //Logger.log('Created slides: %s', slides);
  //Logger.log('Il numero di slides presenti è: %s', slides.length);
  for (let i = 0; i < number; i += 1) {
    slides[i].remove(); // per togliere le slides
  }
  return slides.length;
}

// Return the minimum between two numbers
function min(a, b) {
  if (a <= b) {
    return a;
  } else { return b }
}

function findImage(key, array, mode) {
  var index = -1;
  for (let i = 0, len = array.length; i < len; i++) {
    if (array[i][3] === key) {
      index = i;
      break;
    }
  }
  if (index > -1) {
    return array[index][mode];
  } else { return ' ' }
}

// Manage images
// ['url link','filename in Drive', 'symbol', 'corresponding letters on locations', 'description']
// put the images in this folder: driveIDFolder()[2][0] 
// '🚫''👷''🚷''🚧''🚗''🚘''🧍''🚶'
function imageFinder() {
  var images = [
    ['https://website/Download/pedonale.png', 'pedonale.png', '🚶', 'P', 'The image used to display a pedestrian'],
    ['https://website/Download/carrabile.png', 'carrabile.png', '🚗', 'C', 'The image used to display a gate for cars'],
    ['https://website/Download/dipendenti.png', 'dipendenti.png', '🅁', 'D', 'The image used to display a gate reserved for employees'],
    ['https://website/Download/vip.png', 'vip.png', '🅅', 'V', 'The image used to display a gate reserved for VIP'],
    ['https://website/Download/park.png', 'park.png', '🄿', 'PA', 'The image used to display a parking area'],
    ['https://website/Download/wip.png', 'wip.png', '🚷', 'wip', 'The image used to display an area working in progress'],
    ['https://website/Download/QUARTIERE_A4_1a2000.jpg', 'background.jpg', ' ', 'PLAN', 'The image used to cover the whole page'],
    ['0', '0', ' ', 0, 'The image used to cover the whole page']
  ];
  return images
}

// Select the fill according to type of event
function selectFill(val) {
  var answer = "";
  switch (val) {
    case categories()[1][1]:
    case categories()[2][1]:
      answer = "DASH";
      break;
    case categories()[0][1]:
      answer = "SOLID";
      break;
    case categories()[4][1]:
      answer = "DASH_DOT";
      break;
    case categories()[3][1]:
      answer = "DOT";
      break;
    default:
      answer = "SOLID";
  }
  return answer;
}

// Select the position of text according to type of line
function selectPos(val) {
  var answer = "";
  switch (val) {
    case "DASH":
      answer = 'TOP';
      break;
    case "SOLID":
      answer = 'MIDDLE';
      break;
    case "DASH_DOT":
      answer = 'BOTTOM';
      break;
    case "DOT":
      answer = 'MIDDLE';
      break;
    default:
      answer = 'MIDDLE';
  }
  return answer;
}

// Select the size of text according to type of line
function selectSize(val) {
  var answer = "";
  switch (val) {
    case "DASH":
      answer = 0.95;
      break;
    case "SOLID":
      answer = 1;
      break;
    case "DASH_DOT":
      answer = 0.90;
      break;
    case "DOT":
      answer = 1;
      break;
    default:
      answer = 1;
  }
  return answer;
}

// function to allow app loading variable in the sheets
//function viewCalendar() {
//  var structures = onlyStrcturesSelect();
//  SpreadsheetApp.getUi()
//    .showSidebar(doGet(structures, 'viewCalendarPage', 'Visualizza gli eventi a calendario'));
//}
/* 
function doGet(array, htmlPage, titlePage) {
  var htmlTemplate = HtmlService.createTemplateFromFile(htmlPage);
  htmlTemplate.dataFromServerTemplate = array;
  //htmlTemplate.dataFromServerTemplate = { first: "hello", last: "world" };
  var htmlOutput = htmlTemplate.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle(titlePage);
  return htmlOutput;
}
*/
function doGet(array, htmlPage, titlePage) {
  var htmlTemplate = HtmlService.createTemplateFromFile(htmlPage);

  // Aggiungi i dati da passare al template
  htmlTemplate.dataFromServerTemplate = array; // Dati specifici
  htmlTemplate.currentLanguage = getCurrentLanguage(); // Variabile currentLanguage
  htmlTemplate.translations = getTranslations(); // Dizionari delle traduzioni
  htmlTemplate.minutesPermitted = minutesPermitted(); // time

  // Valuta il template e restituisci l'output HTML
  var htmlOutput = htmlTemplate.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle(titlePage);

  return htmlOutput;
}

// START: Functions to let many users works on the same sheets, creating a new sheet with the name of the user online
function getUserEmail() {
  //var userEmail = PropertiesService.getUserProperties().getProperty("userEmail");
  var userEmail = Session.getEffectiveUser().getEmail();
  if (!userEmail) {
    var protection = SpreadsheetApp.getActive().getRange("A1").protect();
    // tric: the owner and user can not be removed
    protection.removeEditors(protection.getEditors());
    var editors = protection.getEditors();
    if (editors.length === 2) {
      var owner = SpreadsheetApp.getActive().getOwner();
      editors.splice(editors.indexOf(owner), 1); // remove owner, take the user
    }
    userEmail = editors[0];
    protection.remove();
    // saving for better performance next run
    PropertiesService.getUserProperties().setProperty("userEmail", userEmail);
  }
  //Logger.log(userEmail);
  return userEmail;
}

function switchToSheet(sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(sheetName));
}

function createUserSheet() {
  // Obtain user email
  //var user = Session.getActiveUser().getEmail();
  var user = getAliasEmail(getUserEmail());
  //Logger.log(user);
  // Check if the sheet exist
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUser = sheet.getSheetByName(user);
  // If it doesn't exist, create it
  if (!sheetUser) {
    yourNewUserSheet = sheet.insertSheet().setName(user);
  }
  // Active sheet
  //sheet.setActiveSheet(yourNewUserSheet);
  sheetName = getAliasEmail(getUserEmail());
  switchToSheet(sheetName);
  //protectSheets();
}
// END: Functions to let many users works on the same sheets, creating a new sheet with the name of the user online


function testLoadVariables() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var categories = readVariables('structures', 'DataStructures');
  //Logger.log('La variabile contiene ' + categories.length + ' elementi ed è la seguente:\n' + categories);
  //Logger.log('\n \n');
  //Logger.log('La variabile contiene ' + categories.length + ' elementi ed è la seguente:\n' + categories[1]);
}

// Info: readVariablesExt('name of the variable to find', 'sheet name')
function readVariablesExt(nameVar, sheetName, idSheet) {
  idSheet = idSheet || IDPavoraCustomSettings;
  var ss = SpreadsheetApp.openById(idSheet);
  var sheet = ss.getSheetByName(sheetName);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var startRow = 1;
  var data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  
  if (findKey(nameVar, data, 0) >= 0) {
    // Starting row of data
    var sR = findKey(nameVar, data, 0) + 2;
    
    // Finishing column of data
    var fC = data[findKey(nameVar, data, 0) + 1].length;
    
    // Finishing row of data
    var fR = 0;
    for (let i = sR, len = data.length; i < len; i++) {
      // Correzione: controllo più preciso per celle vuote
      if (data[i] && (data[i][0] === '' || data[i][0] === null || data[i][0] === undefined)) {
        break;
      } else {
        fR += 1;
      }
    }
    
    // Debug: aggiungi logging
    console.log('fR:', fR, 'fC:', fC);
    
    // Assicurati che ci sia almeno una riga di dati
    if (fR === 0) {
      fR = 1;
    }
    
    var variable = makeArray(fR, fC);
    var k = 0;
    
    for (let i = sR, len = data.length; i < len && k < fR; i++) {
      // Controllo più robusto per la presenza della riga
      if (data[i] && data[i].length > 0) {
        // Controllo migliorato: FALSE è un valore valido!
        if (data[i][0] !== '' && data[i][0] !== null && data[i][0] !== undefined) {
          for (let j = 0; j < fC; j++) {
            // Gestisce tutti i tipi di valori inclusi FALSE
            if (data[i][j] !== null && data[i][j] !== undefined) {
              variable[k][j] = data[i][j];
            } else {
              variable[k][j] = '';
            }
          }
          k += 1;
        } else if (data[i][0] === false || data[i][0] === 0) {
          // Gestione specifica per FALSE e 0
          for (let j = 0; j < fC; j++) {
            if (data[i][j] !== null && data[i][j] !== undefined) {
              variable[k][j] = data[i][j];
            } else {
              variable[k][j] = '';
            }
          }
          k += 1;
        } else {
          break;
        }
      } else {
        break;
      }
    }
    
    // Debug: verifica che variable sia stato creato correttamente
    console.log('variable length:', variable.length);
    if (variable.length > 0) {
      console.log('variable[0] length:', variable[0].length);
    }
    
    return variable;
  }
  
  // Restituisce un array vuoto se la chiave non viene trovata
  return [];
}

function readSheet(container, internal) {
  // With this I get this error: Exception: Non disponi dell'autorizzazione per chiamare SpreadsheetApp.openById. Autorizzazioni richieste: https://www.googleapis.com/auth/spreadsheets
  // Tried to edit appscript.json with this code found here: https://stackoverflow.com/questions/30587331/you-do-not-have-permission-to-call-openbyid
  if (internal) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetsList()[1][0]);
    if (!sheet) {
      // Crea un nuovo foglio
      sheet = ss.insertSheet(sheetsList()[1][0]);

      // Imposta i nomi delle colonne
      sheet.getRange('A1').setValue(translate('main.user'));
      sheet.getRange('A2').setValue('email');
      sheet.getRange('B2').setValue(translate('main.lastOnline'));

      // Imposta lo stile della prima riga
      var headerRange = sheet.getRange('A2:B2');
      headerRange.setBackground('#D3D3D3'); // Grigio chiaro
      headerRange.setFontWeight('bold');

      // Imposta la larghezza delle colonne
      sheet.setColumnWidth(1, 250); // Colonna A
      sheet.setColumnWidth(2, 200); // Colonna B

      var currentTime = new Date();
      currentTime.setMinutes(currentTime.getMinutes() - 20);
      for (let i = 0; i < users().length; i += 1) {
        if (users()[i][1] != 'reader') {
          sheet.getRange(i + 3, 1).setValue(users()[i][0]);
          sheet.getRange(i + 3, 2).setValue(currentTime).setNumberFormat('dd/MM/yy - HH:mm');
        }
      }
    }

  } else {
    var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/' + IDPavoraCustomSettings + '/edit');
  }
  var sheet = ss.getSheetByName(container);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var startRow = 1;
  var data = sheet.getRange(1, 1, lastRow, lastColumn).getValues(); //getRange(starting Row, starting column, number of rows, number of columns)
  return data
}

// Info: readVariables('name of the variable to find', container)
function readVariables(nameVar, container) {
  var data = container;
  // categories
  // findkey('61', pippo, 2) --> i
  if (findKey(nameVar, data, 0) >= 0) {
    //Logger.log(findKey(nameVar, data, 0));
    // Starting row of data
    var sR = findKey(nameVar, data, 0) + 2;
    //Logger.log('Starting row of data is '+sR);
    // Finishing column of data
    var sC = data[findKey(nameVar, data, 0) + 2].length;
    //Logger.log('sC è lungo ' + data[findKey(nameVar, data, 0) + 2].length);
    var fC = 0;
    for (let i = 0, len = sC; i < len; i++) {
      //Logger.log('data è = ' + data[sR][i]);
      if (data[sR - 1][i] == '') {
        break;
      } else { fC += 1 }
    }
    //Logger.log('fC è = ' + fC);
    // Finishing row of data
    var fR = 0;
    for (let i = sR, len = data.length; i < len; i++) {
      if (data[i][0] == '') {
        break;
      } else { fR += 1 }
    }
    //Logger.log('Finishing row of data is '+fR);
    var variable = makeArray(fR, fC)
    var k = 0;
    for (let i = sR, len = data.length; i < len; i++) {
      if (typeof (data[i][0]) != undefined) {
        if (data[i][0] != '') {
          for (let j = 0; j < fC; j++) {
            if (typeof (data[i][j]) != undefined) {
              variable[k][j] = data[i][j];
            } else {
              variable[k][j] = '';
            }
          }
        } else { break }
        k += 1;
      } else { break }
    }
  }
  return variable
}

// END: Functions to load the settings froma another external sheet 16DShugZ5HPj3wc65U7Gv-mF2FIolq064WOuw2q8TQOc --> Settings_and_Variables_for_SpacesOccupation

//
//START: Useful and general functions
//
// Find all the index of a 2D array looking for a key in the first position of the rows
// Example:
//
// pippo[i] = ["r1","7up","61","Albertsons"];
// pippo[j] = ["r1","8up","71","Sons"];
// findAllKey('r1', pippo) --> [i, j]
function findAllKey(key, array) {
  var index = [-1];
  var k = 0;
  for (let i = 0, len = array.length; i < len; i++) {
    if (array[i][0] === key) {
      index[k] = i;
      k = k + 1;
    }
  }
  if (index[0] > -1) {
    return index
  } else { return index }
}

// Delete all
function ClearAll() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Ripristina impostazioni di visualizzazione
  sh.setHiddenGridlines(false);
  //sh.setFrozenRows(6);
  //sh.setFrozenColumns(1);

  // Rimuove menu a tendina con dati dinamici
  const fullRange = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  fullRange.setDataValidation(null);

  // Rimuove eventuali filtri
  if (sh.getFilter()) sh.getFilter().remove();

  // Rimuove immagini
  sh.getImages().forEach(image => image.remove());

  // Rimuove formattazioni condizionali
  sh.clearConditionalFormatRules();

  // Mostra tutte le righe e colonne nascoste
  sh.unhideColumn(sh.getRange("1:1"));
  sh.unhideRow(sh.getRange("A:A"));

  // Inserisce un punto per mantenere la tabella strutturata
  sh.getRange(90, 1500).setValue('.');

  // Pulisce i contenuti e resetta lo stile
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const range = sh.getRange(1, 1, lastRow, lastCol);

  range.clearContent()
    .clearFormat()
    .setBackground('#FFFFFF')
    .setBorder(false, false, false, false, false, false, "grey", SpreadsheetApp.BorderStyle.DASHED)
    .setNote('')
    .setVerticalAlignment("middle")
    .setFontSize(16);

  // Impostazioni dimensioni colonne e righe
  //sh.setColumnWidths(1, 1, 250);   // Prima colonna larga 250
  //sh.setColumnWidths(2, 368 - 1, 40); // Colonne successive larghe 40
  //sh.setRowHeights(1, lastRow - 1, 21); // Altezza standard righe
  //sh.setRowHeight(3, 40);  // Terza riga più alta

  // Evita errori di confine con una cella vuota
  //sh.getRange(500, 370).setValue('');
}

//
//END: Useful and general functions
//

//
// START: standard functions
//
/////////////////////////////////////////////
// Split sheet and export in XLSX         //
///////////////////////////////////////////
//
// Function to split the original sheet in four-month and two-month period
//
function splitInSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Set the initial sheet to "main"
  var main = ss.getSheetByName('Recap');

  // Retrieve the year from the sheet in order to decide if it is leap or not
  var year = main.getRange(4, 1).getValue();
  isLeap = !(new Date(year, 1, 29).getMonth() - 1)
  if (isLeap) {
    bis = 1;
  } else {
    bis = 0;
  }

  // Create an array to split the year into four-month period (whole numbers)
  var quadrimestre = [
    [1, 120 + bis, 'I_Quad'],
    [121 + bis, 243 + bis, 'II_Quad'],
    [244 + bis, 365 + bis, 'III_Quad']
  ];
  // Create an array to split the year into two-month period (whole numbers)  
  var bimestre = [
    [1, 59 + bis, 'Gen-Feb'],
    [60 + bis, 120 + bis, 'Mar-Apr'],
    [121 + bis, 181 + bis, 'Mag-Giu'],
    [182 + bis, 243 + bis, 'Lug-Ago'],
    [244 + bis, 304 + bis, 'Set-Ott'],
    [305 + bis, 365 + bis, 'Nov-Dic']
  ];


  // Add the corresponding three sheets to hold the year split in four-month period
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
  splitInTimePeriodArr(quadrimestre[0]);
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
  splitInTimePeriodArr(quadrimestre[1]);
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
  splitInTimePeriodArr(quadrimestre[2]);
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
  // Add the corresponding three sheets to hold the year split in four-month period
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
  splitInTimePeriodArr(bimestre[0]);
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
  splitInTimePeriodArr(bimestre[1]);
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
  splitInTimePeriodArr(bimestre[2]);
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
  splitInTimePeriodArr(bimestre[3]);
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
  splitInTimePeriodArr(bimestre[4]);
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
  splitInTimePeriodArr(bimestre[5]);

  // Re-Order the sheets based on date
  // Calculate the day of the year (1 - 366)
  var now = new Date();
  var start = new Date(year, 0, 0);
  var diff = (now - start) + ((start.getTimezoneOffset() - now.getTimezoneOffset()) * 60 * 1000);
  var oneDay = 1000 * 60 * 60 * 24;
  var day = Math.floor(diff / oneDay);
  //console.log('Day of year: ' + day);
  // Source: https://stackoverflow.com/questions/8619879/javascript-calculate-the-day-of-the-year-1-366

  ss.setActiveSheet(ss.getSheetByName('Recap'));
  ss.moveActiveSheet(4);

  // Make decision according to the value of "day"
  if (day <= 0) {
    ss.setActiveSheet(ss.getSheetByName('I_Quad'));
    // Do nothing because the main sheet is in the future
  } else if (day > 365 + bis) {
    ss.setActiveSheet(ss.getSheetByName('I_Quad'));
    // Do nothing because the main sheet is in the past
    ss.setActiveSheet(ss.getSheetByName('Recap'));
    ss.moveActiveSheet(1);
  } else if (day > 243 + bis) {
    ss.setActiveSheet(ss.getSheetByName('III_Quad'));
    ss.moveActiveSheet(1);
    checkToday();
  } else if (day > 120 + bis) {
    ss.setActiveSheet(ss.getSheetByName('II_Quad'));
    ss.moveActiveSheet(1);
    ss.setActiveSheet(ss.getSheetByName('III_Quad'));
    ss.moveActiveSheet(2);
    ss.setActiveSheet(ss.getSheetByName('II_Quad'));
    checkToday();
  } else if (day > 0) {
    ss.setActiveSheet(ss.getSheetByName('I_Quad'));
    checkToday();
  }
}

//
// Function to split a year into many sheets given a correct matrix
//
function splitInTimePeriod(matrix) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var main = ss.getSheetByName('Recap');
  var lastRow = main.getLastRow();

  for (let i = 0; i < matrix.length; i += 1) {
    // Insert a new sheet labeled with the correct name inside the matrix
    var currentSheet = ss.insertSheet(matrix[i][2]);
    var currentId = currentSheet.getSheetId();

    //Utilities.sleep(3000);
    //SpreadsheetApp.flush();
    //var main = ss.getSheetByName('Recap');
    var dataRange = main.getRange(1, 1, lastRow + 1, 1);
    var myData = dataRange.getValues();
    currentSheet.getRange(1, 1, lastRow + 1, 1).setValues(myData);
    dataRange.copyFormatToRange(currentId, 1, 1, 1, lastRow + 1);
    //Logger.log(dataRange + ' ' + currentId + ' ' + main.getLastRow() + ' ' + matrix[i][1] + '  ');

    // Copy data in the main sheet
    // getRange(row, column, rowEnd, columnEnd)
    //Logger.log('1,' + matrix[i][0] + '+1,' + main.getLastRow() + ',' + matrix[i][1] + '+1');
    var dataRange = main.getRange(1, matrix[i][0] + 1, main.getLastRow() + 1, matrix[i][1] + 1 - matrix[i][0]);
    var myData = dataRange.getValues();
    var notes = dataRange.getNotes();
    currentSheet.getRange(1, 2, main.getLastRow() + 1, matrix[i][1] + 1 - matrix[i][0]).setValues(myData).setNotes(notes);
    // Paste data and formatin the new sheet
    //copyFormatToRange(gridId, column, columnEnd, row, rowEnd)
    dataRange.copyFormatToRange(currentId, 2, matrix[i][1] + 1 - matrix[i][0], 1, main.getLastRow());
    currentSheet.getRange(main.getLastRow(), matrix[i][1] + 2 - matrix[i][0]).setValue('.');

    // Remove empty rows and columns
    var lr = currentSheet.getLastRow();
    var mr = currentSheet.getMaxRows();
    var lc = currentSheet.getLastColumn();
    var mc = currentSheet.getMaxColumns();
    currentSheet.setColumnWidths(1, 1, 188.0);
    currentSheet.setColumnWidths(2, lc - 1, 35.0);
    if (mr - lr != 0) {
      currentSheet.deleteRows(lr + 1, mr - lr);
    }
    if (mc - lc != 0) {
      currentSheet.deleteColumns(lc + 1, mc - lc - 1);
    }
    currentSheet.setRowHeights(1, lr - 1, 21);
    currentSheet.setRowHeight(3, 40);
    currentSheet.setRowHeight(6 + struttureScelte().length + 1, 60);
    ss.setActiveSheet(ss.getSheetByName(matrix[i][2]));
    checkToday();
    createGroup();
  }
}

//
// Function to split a year into many sheets given a array of data
//
function splitInTimePeriodArr(array) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var main = ss.getSheetByName('Recap');
  var lastRow = main.getLastRow();
  var dataRange = main.getRange(1, 1, main.getLastRow() + 1, 1);
  var myData = dataRange.getValues();

  // Insert a new sheet labeled with the correct name inside the array
  var currentSheet = ss.insertSheet(array[2]);
  var currentId = currentSheet.getSheetId();

  //Utilities.sleep(3000);
  //SpreadsheetApp.flush();
  //var main = ss.getSheetByName('Recap');
  //var dataRange = main.getRange(1, 1, lastRow+1, 1);
  //var myData = dataRange.getValues();
  currentSheet.getRange(1, 1, lastRow + 1, 1).setValues(myData);
  dataRange.copyFormatToRange(currentId, 1, 1, 1, lastRow + 1);
  //Logger.log(dataRange+' '+currentId+' '+main.getLastRow()+' '+array[1]+'  ');

  // Copy data in the main sheet
  // getRange(row, column, rowEnd, columnEnd)
  //Logger.log('1,' + array[0] + '+1,' + main.getLastRow() + ',' + array[1] + '+1');
  var dataRange = main.getRange(1, array[0] + 1, main.getLastRow() + 1, array[1] + 1 - array[0]);
  var myData = dataRange.getValues();
  var notes = dataRange.getNotes();
  currentSheet.getRange(1, 2, main.getLastRow() + 1, array[1] + 1 - array[0]).setValues(myData).setNotes(notes);
  // Paste data and formatin the new sheet
  //copyFormatToRange(gridId, column, columnEnd, row, rowEnd)
  dataRange.copyFormatToRange(currentId, 2, array[1] + 1 - array[0], 1, main.getLastRow());
  currentSheet.getRange(main.getLastRow(), array[1] + 2 - array[0]).setValue('.');

  // Remove empty rows and columns
  var lr = currentSheet.getLastRow();
  var mr = currentSheet.getMaxRows();
  var lc = currentSheet.getLastColumn();
  var mc = currentSheet.getMaxColumns();
  currentSheet.setColumnWidths(1, 1, 188.0);
  currentSheet.setColumnWidths(2, lc - 1, 40.0);
  if (mr - lr != 0) {
    currentSheet.deleteRows(lr + 1, mr - lr);
  }
  if (mc - lc != 0) {
    currentSheet.deleteColumns(lc + 1, mc - lc - 1);
  }
  currentSheet.setRowHeights(1, lr - 1, 21);
  currentSheet.setRowHeight(3, 40);
  currentSheet.setRowHeight(6 + struttureScelte().length + 1, 60);
  ss.setActiveSheet(ss.getSheetByName(array[2]));
  checkToday();
  createGroup();
}


/////////////////////////////////////////////
// Secondary and basic functions          //
///////////////////////////////////////////
// Print to letter A -> Z, AA -> ZZ
//https://stackoverflow.com/questions/36129721/convert-number-to-alphabet-letter
// Better
//https://stackoverflow.com/questions/45787459/convert-number-to-alphabet-string-javascript
function printToLetter(num) {
  var s = '', t;

  while (num > 0) {
    if (((num - 1) % 61) == 31) { //31 = '
      num++;
    };
    t = (num - 1) % 61; //26 A..Z
    s = String.fromCharCode(65 + t) + s;
    num = (num - t) / 61 | 0; //26 A..Z
  }
  return s || undefined;
}

// Conditional rule to find and highlight current day in the active sheet
function checkToday() {
  var sheet = SpreadsheetApp.getActiveSheet();

  var numRows = sheet.getLastRow();
  var numColumns = sheet.getLastColumn();
  var rangeToHighlight = sheet.getRange(3, 2, 4, numColumns);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=B$5=today()')
    .setBackground("#d9d9d9")
    //.setBorder(true, true, true, true, true, true, "red", SpreadsheetApp.BorderStyle.SOLID) // Borders cannot be formatted in scope of conditional formatting
    .setRanges([rangeToHighlight])
    .build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

// Return the last day of the month
// console.log(LastDayOfMonth(2020, 2)) // => Sat Feb 29 2020
function LastDayOfMonth(Year, Month) {
  return new Date((new Date(Year, Month, 1)) - 1);
}

// Remove the accent marks from a given string
function RemoveAccents(string) {
  var strAccents = string.split('');
  var strAccentsOut = new Array();
  var strAccentsLen = strAccents.length;
  var accents = 'ÀÁÂÃÄÅàáâãäåÒÓÔÕÕÖØòóôõöøÈÉÊËèéêëðÇçÐÌÍÎÏìíîïÙÚÛÜùúûüÑñŠšŸÿýŽž';
  var accentsOut = "AAAAAAaaaaaaOOOOOOOooooooEEEEeeeeeCcDIIIIiiiiUUUUuuuuNnSsYyyZz";
  for (let y = 0; y < strAccentsLen; y++) {
    if (accents.indexOf(strAccents[y]) != -1) {
      strAccentsOut[y] = accentsOut.substr(accents.indexOf(strAccents[y]), 1);
    }
    else {
      strAccentsOut[y] = strAccents[y];
    }
  }
  strAccentsOut = strAccentsOut.join('');
  return strAccentsOut;
}

// Choice the row and column of the color array based on description of the type (set up, event, break down, work).--> 2 Color Method
function selectHigh(num, type) {
  //var colonna = (num%10)+0;
  var column = (num % (methodMcolors()[0].length - 1)) + 0; // the lenght of the array is 10 but the last column is used for optioned events
  var answer = type;
  switch (type) {
    case categories()[1][1]:
    case categories()[2][1]:
      answer = 0;
      break;
    case categories()[0][1]:
      answer = 1;
      break;
    case categories()[4][1]:
      answer = 0;
      break;
    case categories()[3][1]:
      answer = 1;
      break;
    default:
      answer = 1;
  }
  return [answer, column];
}

// Choice the row and column of the color array based on description of the type (set up, event, break down, work).--> Letter Method
function selectHighW(num, tipologia) {
  //var colonna = (num%10)+0;
  var colonna = (num % (methodMcolors()[0].length - 1)) + 0; // the lenght of the array is 10 but the last column is used for optioned events
  var answer = tipologia;
  switch (tipologia) {
    case categories()[1][1]:
      answer = 2;
      break;
    case categories()[0][1]:
      answer = 2;
      break;
    case categories()[2][1]:
      answer = 2;
      break;
    case categories()[3][1]:
      answer = 2;
      break;
    default:
      answer = 2;
  }
  return [answer, colonna];
}

// Find the index of a 2D array looking for a key in the first position of the rows
// Example:
//
// pippo[i] = ["r1","7up","61","Albertsons"];
// findkey('61', pippo, 2) --> i
function findKey(key, array, pos) {
  return array.findIndex(row => row[pos] === key);
}


// Finish the first letter of type with the correct word
function selectType(val) {
  var answer = "";
  switch (val) {
    case categories()[1][1]:
      answer = categories()[1][2];
      break;
    case categories()[0][1]:
      answer = categories()[0][2];
      break;
    case categories()[2][1]:
      answer = categories()[2][2];
      break;
    case categories()[3][1]:
      answer = categories()[3][2];
      break;
    case categories()[4][1]:
      answer = categories()[4][2];
      break;
    default:
      answer = "";
  }
  return answer;
}

// Select the fill according to type of event
function selectFill(val) {
  var answer = "";
  switch (val) {
    case categories()[1][1]:
    case categories()[2][1]:
      answer = "DASH";
      break;
    case categories()[0][1]:
      answer = "SOLID";
      break;
    case categories()[4][1]:
      answer = "DASH_DOT";
      break;
    case categories()[3][1]:
      answer = "DOT";
      break;
    default:
      answer = "SOLID";
  }
  return answer;
}

// Create an empty array with r rows and c columns
function makeArray(row, column) {
  var arr = new Array(row)
  for (let i = 0; i < row; i++)
    arr[i] = new Array(column)
  return arr
}
//console.log(makeArray(4,4))

// Convert a CSV string in the corresponding array, removing empty spaces
// example = 'a, b,  c,  d';
// string2array(example); --> ['a','b','c','d']
function string2array(inputFormat) {
  inputFormat = inputFormat.replace(/\s+/g, ''); // remove spaces
  return inputFormat.split(",");
}

// Return the subtraction of two dates in months
function monthDiff(dateFrom, dateTo) {
  return (dateTo.getMonth() + 1) - dateFrom.getMonth() + (12 * (dateTo.getFullYear() - dateFrom.getFullYear()))
}

// Return the number of days in a given month and year
function getDaysInMonth(month, year) {
  return new Date(year, month + 1, 0).getDate();
}

// Format a date for header
//console.log(convertDate('Mon Nov 19 13:29:40 2012')) // => "Monday 19/11/2012"
function convertDate(inputFormat) {
  function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat)
  //var giorno = ["dom", "lun", "mar", "mer", "gio", "ven", "sab"][d.getDay()];
  var stringDay = translate('planPage.threeDayAbb').split(',');
  var giorno = stringDay[d.getDay()];
  return giorno + ' ' + [pad(d.getDate()), pad(d.getMonth() + 1), d.getFullYear().toString().substr(-2)].join('/')
}

// Convertire orario Thu Jun 20 2024 08:00:00 GMT+0200 (Central European Summer Time) --> 20/06/2024 08:00
function convertDateFormat(dateString) {
  var date = new Date(dateString);

  var day = ('0' + date.getDate()).slice(-2);
  var month = ('0' + (date.getMonth() + 1)).slice(-2); // I mesi vanno da 0 a 11
  var year = date.getFullYear();

  var hours = ('0' + date.getHours()).slice(-2);
  var minutes = ('0' + date.getMinutes()).slice(-2);

  return `${day}/${month}/${year} ${hours}:${minutes}`;
}


// Funzione per convertire le date in formato gg/MM/aa HH:mm
function formatDateMaster(dateString) {
  var date = new Date(dateString);

  var day = ('0' + date.getDate()).slice(-2);
  var month = ('0' + (date.getMonth() + 1)).slice(-2); // I mesi vanno da 0 a 11
  var year = date.getFullYear();

  var hours = ('0' + date.getHours()).slice(-2);
  var minutes = ('0' + date.getMinutes()).slice(-2);
  var seconds = ('0' + date.getSeconds()).slice(-2);

  return {
    giorno: `${day}/${month}/${year}`,
    ora: `${hours}:${minutes}`,
    dataXfile: `${year}${month}${day}_${hours}${minutes}${seconds}`,
    dataXweb: `${year}-${month}-${day}`
  }
}

// 
// Formatta la data nel formato YYYY-MM-DD
function convertDateInputHtml(dateInput) {
  var date = new Date(dateInput);
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);
  return `${year}-${month}-${day}`;
}

// Clean the date removing "/" in DDMMYY
//console.log(convertDateClean('Mon Nov 19 13:29:40 2012')) // => "191112"
function convertDateClean(inputFormat) {
  function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat)
  return [pad(d.getDate()), pad(d.getMonth() + 1), d.getFullYear().toString().substr(-2)].join('')
}

// Clean the date removing "/" in YYMMDD
//console.log(convertDateUSAClean('Mon Nov 19 13:29:40 2012')) // => "121119"
function convertDateUSAClean(inputFormat) {
  function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat)
  return [d.getFullYear().toString().substr(-2), pad(d.getMonth() + 1), pad(d.getDate())].join('')
}

// Clean the date with "/" in YYYY/MM/DD
//console.log(convertDateUSA('Mon Nov 19 13:29:40 2012')) // => "2012/11/19"
function convertDateUSA(inputFormat) {
  function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat)
  return [d.getFullYear(), pad(d.getMonth() + 1), pad(d.getDate())].join('/')
}

// Clean the date with "/" in DD/MM/YYYY
//console.log(convertDateBar('Mon Nov 19 13:29:40 2012')) // => "19/11/2012"
function convertDateBar(inputFormat) {
  function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat)
  return [pad(d.getDate()), pad(d.getMonth() + 1), d.getFullYear()].join('/')
}


// Clean the date with "/" in DD/MM
//console.log(convertDayMonthBar('Mon Nov 19 13:29:40 2012')) // => "19/11"
function convertDayMonthBar(inputFormat) {
  function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat)
  // var giorno = ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì","sabato"][d.getDay()];
  return [pad(d.getDate()), pad(d.getMonth() + 1)].join('/')
}

// Extract time from a date
//console.log(convertHour('Mon Nov 19 13:29:40 2012')) // => "13:29"
function convertHour(inputFormat) {
  function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat)
  return [pad(d.getHours()), pad(d.getMinutes())].join(':')
}

// Added from April 2021
// Example: columnToLetter(100) --> "CV"
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// Example: rowColumn2cell(1,1) --> 'A1'
function rc2cell(r, c) {
  return columnToLetter(c) + r;
}

function myFunction() {
  ClearAll();
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var dA = [100, 90, 80, 70, 60];
  //sheet.getRangeList(['A1','B2','C3','D4','E5']).getRanges().forEach(function(rg,i){
  sheet.getRangeList([rc2cell(1, 1), rc2cell(2, 2), rc2cell(3, 3), rc2cell(4, 4), rc2cell(5, 5)]).getRanges().forEach(function (rg, i) {
    rg.setValue(dA[i]);
  });
}

function findLastKey(key, array) {
  for (let i = array.length - 1; i >= 0; i--) {
    if (array[i][0] === key) {
      return i;
    }
  }
  return -1;
}

/** https://stackoverflow.com/questions/30367547/convert-all-sheets-to-pdf-with-google-apps-script
 * Export one or all sheets in a spreadsheet as PDF files on user's Google Drive,
 * in same folder that contained original spreadsheet.
 *
 * Adapted from https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579#c25
 *
 * @param {String}  optSSId       (optional) ID of spreadsheet to export.
 *                                If not provided, script assumes it is
 *                                sheet-bound and opens the active spreadsheet.
 * @param {String}  optSheetId    (optional) ID of single sheet to export.
 *                                If not provided, all sheets will export.
 */
function savePDFs(optSSId, optSheetId) {

  // If a sheet ID was provided, open that sheet, otherwise assume script is
  // sheet-bound, and open the active spreadsheet.
  var ss = (optSSId) ? SpreadsheetApp.openById(optSSId) : SpreadsheetApp.getActiveSpreadsheet();
  //var ss = SpreadsheetApp.getActive();
  //var sheet = ss.getSheetByName('Gen-Feb');

  // Get folder containing spreadsheet, for later export
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  if (parents.hasNext()) {
    var folder = parents.next();
  }
  else {
    folder = DriveApp.getRootFolder();
  }

  // Use another and fixed folder to store it
  var folder = DriveApp.getFolderById(driveIDFolder()[0][0]);

  //additional parameters for exporting the sheet as a pdf
  var url_ext = 'export?exportFormat=pdf&format=pdf'   //export as pdf

    // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
    + (optSheetId ? ('&gid=' + ss.getSheetId()) : ('&id=' + ss.getId()))
    //+    '&gid=' + sheet.getSheetId()
    //+    '&gid=' + ss.getSheetByName('Mar-Apr') // the sheet's Id

    // following parameters are optional...
    + '&size=A4'      // paper size
    + '&portrait=false'    // orientation, false for landscape
    + '&fitw=true'        // fit to width, false for actual size
    + '&scale=2'          //1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page
    + '&top_margin=0.25'              //All four margins must be set!
    + '&bottom_margin=0.25'           //All four margins must be set!
    + '&left_margin=0.25'             //All four margins must be set!
    + '&right_margin=0.25'            //All four margins must be set!
    //+ '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
    + '&gridlines=true'  // hide gridlines
    + '&printnotes=false' // don't show notes
    + '&pagenum=CENTER'
    + '&sheetnames=true&printtitle=true' // hide optional headers and footers
    + '&fzr=false';       // do not repeat row headers (frozen rows) on each page

  var options = {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  }
  var response = UrlFetchApp.fetch("https://docs.google.com/spreadsheets/" + url_ext, options);
  // Variable to create today date
  var today = new Date();
  var todayReadable = convertDateClean(today);
  var todayUSAReadable = convertDateUSAClean(today);
  // Store the year from sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var year = sheet.getRange(4, 1).getValue();
  //Logger.log(year);
  //var blob = folder.createFile(blobs).setName('2020'+'_Pianificazione_Quartiere_'+todayReadable+'.pdf')
  var blob = response.getBlob().setName(year + '_Pianificazione_Quartiere_' + todayReadable + '.pdf');

  //from here you should be able to use and manipulate the blob to send and email or create a file per usual.
  //In this example, I save the pdf to drive
  folder.createFile(blob);

}

// Comment from https://www.andrewroberts.net/2017/03/apps-script-create-pdf-multi-sheet-google-sheet/
function hideSheets() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var visible = ['Gen-Feb', 'Mar-Apr', 'Mag-Giu', 'Lug-Ago', 'Set-Ott', 'Nov-Dic']; //sheet names

  sheets.forEach(function (sheet) {
    if (visible.indexOf(sheet.getName()) == -1) {
      sheet.hideSheet();
    }
  })
};

function showSheets() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(function (sheet) {
    sheet.showSheet()
  })
};

function createPDF() {
  hideSheets();
  savePDFs();
  showSheets();
}


//
// END: standard functions
//