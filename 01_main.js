/**
 * Project: Pavora (v0.9)
 * Created by: Nicola Rainiero
 * Creation Date: 2025-01-02
 * Description: Google Apps Script for reserving shared facilities in an exhibition area, using HTML menus and interacting with a Google Sheet
 * 
 * License: MIT License
 * 
 * Copyright (c) 2025 Nicola Rainiero
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

// Global variables
const IDPavoraCustomSettings = "18j_d2ApLsIHOnTBbThxKV3u61VtOKzndCXT6Vlb36bw"; // CHANGE MANUALLY WITH THE REAL AND YOUR ID of the CustomSettings Sheet
const IDAliasEmail = "1dW8ys39MeujUlt-eoeJ-RvGHPl7kLGqCF-K5h5UB7aY"; // CHANGE MANUALLY WITH THE REAL AND YOUR ID of the Alias Email Sheet (only if you have selected aliasEmail in the Settings)
const DataSettings = readSheet('DataSettings', 0);
const DataStructures = readSheet('DataStructures', 0);
const DataEventSheet = readSheet('DataEventSheet', 0);
const UsersOnline = readSheet(sheetsList()[1][0], 1); // the sheet name is 'online'

// Test user permissions
function checkUserWritePermission(calendarId) {
  try {
    // Try to create a temporary event
    var event = {
      'summary': 'Temporary Event for Permission Check',
      'start': {
        'dateTime': '2020-01-01T00:00:00.000Z'
      },
      'end': {
        'dateTime': '2020-01-01T01:00:00.000Z'
      }
    };
    var createdEvent = Calendar.Events.insert(event, calendarId);

    // Remove the event
    Calendar.Events.remove(calendarId, createdEvent.id);

    return true; // User has write permissions
  } catch (e) {
    return false; // User has not write permissions
  }
}

//
// Every minute it checks if a user acording to the settings sheet and minutesPermitted(), has the permission to create or edit/delete an event
// a user through updateTimeUser() refresh the time and mantain the rights on
// in this case the user has the write access to the shared calendar
//
function userWriteReadCalendar() { // funzione in uso ed ufficiale!!!!
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  //Logger.log(user);
  var timeInterval = minutesPermitted() * 60 * 1000; // 5 minuti bisogna modificarlo nella tabella variabili
  if (users()[findKey(user, users(), 0)][1] === 'admin') {
    //Logger.log('L\'utente ' + user + ' può modificare i permessi al calendario');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var fogli = ss.getSheets(); // Ottieni tutti i fogli esistenti
    var nomiFogli = fogli.map(function (foglio) { return foglio.getName(); }); // Ottieni i nomi dei fogli

    // Verifica e crea il foglio "logs" se non esiste
    if (!nomiFogli.includes(sheetsList()[0][0])) {
      ss.insertSheet(sheetsList()[0][0]);
      var newSheet = ss.getSheetByName(sheetsList()[0][0]);
      var range = newSheet.getRange(1, 1, 1, 5);
      var range1 = newSheet.getRange(2, 1, 1, 5);
      //var header = [["Data", "AzioneID", "Evento", "Email Utente", "Dettagli"]];
      var header = [translate('main.logCreation').split(',')];
      var firstRow = [[".", ".", ".", ".", "."]];
      // Stili per la tabella
      var headerColor = "#999999"; // #EDD400 giallo per l'intestazione della tabella 002d62 blu sito
      var textHeaderColor = "#000000";
      range.setValues(header).setFontColor(textHeaderColor).setFontSize(12).setBackground(headerColor).setHorizontalAlignment("center");
      range1.setValues(firstRow);
      addFilterRegisterPage();
      newSheet.deleteRow(newSheet.getLastRow());
    }

    // Verifica e crea il foglio "instructions" se non esiste
    if (!nomiFogli.includes(sheetsList()[2][0])) {
      ss.insertSheet(sheetsList()[2][0]);
      var newSheet = ss.getSheetByName(sheetsList()[2][0]);
      var countSheet = ss.getSheets().length;
      ss.moveActiveSheet(1);
      newSheet.getRange("G2").setValue(translate('main.instruction') + ' ➡️').setHorizontalAlignment('right').setFontSize(20);
      insertFloatingImage(newSheet, driveIDFiles()[3][0], 2, 3, 667 * 0.75, 738 * 0.75);

      // Rimuove colonne e righe extra
      var maxColumns = 7;
      var maxRows = 30;

      // Rimuove colonne extra
      var totalColumns = newSheet.getMaxColumns();
      if (totalColumns > maxColumns) {
        newSheet.deleteColumns(maxColumns + 1, totalColumns - maxColumns);
      }

      // Rimuove righe extra
      var totalRows = newSheet.getMaxRows();
      if (totalRows > maxRows) {
        newSheet.deleteRows(maxRows + 1, totalRows - maxRows);
      }

      // Nasconde la griglia
      newSheet.setHiddenGridlines(true);
    }
    eliminaFogliNonPresentiNelVettore();
    /* NEW CODE FOR OPTIMIZE CALL TO SHARECALENDAR */
sh = ss.getSheetByName(sheetsList()[1][0]);
var lr = sh.getLastRow();
var currentTime = new Date();

for (let i = 4; i <= lr; i += 1) {
  var userEmail = sh.getRange(i, 1).getValue();
  var lastOnline = new Date(sh.getRange(i, 2).getValue()).getTime();
  var activation = sh.getRange(i, 3).getValue(); // Valore attuale di Attivazione
  var newActivation = (currentTime.getTime() <= (lastOnline + timeInterval)) ? 1 : 0;
  var newRole = newActivation === 1 ? "writer" : "reader"; // Scrittura solo se attivo

  // Recupera il ruolo attuale dell'utente per evitare modifiche inutili
  try {
    if (aliasEmail()) {userEmail = getRealEmail(userEmail)}
    var currentAcl = Calendar.Acl.get(myCalID()[0][0], "user:" + userEmail);
    var currentRole = currentAcl.role;
  } catch (e) {
    var currentRole = null; // Se l'utente non ha ancora permessi
  }

  // Se il ruolo deve cambiare, lo aggiorniamo
  if (currentRole !== newRole) {
    shareCalendar(myCalID()[0][0], userEmail, newRole);

    if (newRole === "writer") {
      addEditorToProtectedSheet(sheetsList()[0][0], userEmail);
    } else {
      removeEditorFromProtectedSheet(sheetsList()[0][0], userEmail);
    }
  }

  // Aggiorniamo il valore in "Attivazione" solo se è cambiato
  if (activation !== newActivation) {
    sh.getRange(i, 3).setValue(newActivation);
  }
}
    protectSheets();
  } else { Logger.log(translate('alert.userPermission', { user: user })) }
}

function updateTimeUser() {
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sh = ss.getSheetByName(sheetsList()[1][0]);

  if (!sh) {
    // Crea un nuovo foglio
    sh = ss.insertSheet(sheetsList()[1][0]);

    // Imposta i nomi delle colonne
    sh.getRange('A1').setValue(translate('main.user'));
    sh.getRange('A2').setValue(translate('main.email'));
    sh.getRange('B2').setValue(translate('main.lastOnline'));

    // Imposta lo stile della prima riga
    var headerRange = sh.getRange('A2:B2');
    headerRange.setBackground('#D3D3D3'); // Grigio chiaro
    headerRange.setFontWeight('bold');

    // Imposta la larghezza delle colonne
    sh.setColumnWidth(1, 300); // Colonna A
    sh.setColumnWidth(2, 200); // Colonna B
    sh.hideSheet();
  }

  var lr = sh.getLastRow();
  var currentTime = new Date();
  // Create and update user online time
  //Logger.log('------------------' + user + ' -//||\\- ' + findKey(user, users(), 0) + ' ' + findKey(user, usersOnline(), 0));
  if ((findKey(user, users(), 0) >= 0) && (findKey(user, usersOnline(), 0) < 0)) {
    lr += 1;
    sh.getRange(lr, 1).setValue(user);
    //Logger.log('Added user ' + user + ' to the online sheet');
  }
  for (let i = 1; i <= lr; i += 1) {
    //Logger.log(sh.getRange(i, 1).getValue());
    if (user == sh.getRange(i, 1).getValue()) {
      sh.getRange(i, 2).setValue(currentTime).setNumberFormat('dd/MM/yy - HH:mm');
    }
  }
}

function topMenu() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.  
    .createMenu(translate('menu.title'))
    .addItem(translate('menu.completeMenu'), 'completeMenu')
    .addItem(translate('menu.specialEvent'), 'specialEvent')
    .addItem(translate('menu.specialDailyEvent'), 'specialDailyEvent')    
    .addItem(translate('menu.viewCalendar'), 'viewCalendar')
    .addItem(translate('menu.viewDailyCalendar'), 'viewDailyCalendar')
    .addItem(translate('menu.viewListCalendar'), 'viewListCalendar')
    .addItem(translate('menu.manageAccess'), 'manageAccess')
    .addToUi();
  SpreadsheetApp.getActive().getSheetByName(sheetsList()[1][0]).hideSheet();
  SpreadsheetApp.getActive().getSheetByName(sheetsList()[0][0]).hideSheet();
  SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName(sheetsList()[2][0]));
  completeMenu();
}

function protectSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var ownerEmail = Session.getEffectiveUser().getEmail();

  sheets.forEach(function (sheet) {
    var sheetName = sheet.getName();
    if (isValidEmail(sheetName)) {
      var protection = sheet.protect();
      var editors = [ownerEmail, getRealEmail(sheetName)]; // Proprietario e utente corrispondente

      protection.addEditors(editors);
      protection.removeEditors(protection.getEditors().filter(function (editor) {
        return editors.indexOf(editor.getEmail()) === -1;
      }));

      protection.setWarningOnly(false); // Impedisce altre modifiche
    }
  });
  protectSheetByName(sheetsList()[2][0], users()[0][0]);
}

function addEditorToProtectedSheet(sheetName, email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    //Logger.log("Foglio non trovato: " + sheetName);
    return;
  }

  // Ottieni tutte le protezioni del foglio
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);

  // Controlla se esiste una protezione per il foglio
  if (protections.length > 0) {
    var protection = protections[0]; // Prendi la prima protezione trovata

    // Ottieni la lista degli editor autorizzati
    var editors = protection.getEditors();

    // Verifica se l'utente è già presente nella lista degli editor autorizzati
    for (let i = 0; i < editors.length; i++) {
      if (editors[i].getEmail() === email) {
        //Logger.log("L'utente " + email + " è già presente nella lista degli editor autorizzati.");
        return;
      }
    }

    // Aggiungi l'utente specificato agli editor autorizzati
    protection.addEditor(email);
    //Logger.log("Permesso di modifica aggiunto per l'utente: " + email);
  } else {
    //Logger.log("Nessuna protezione trovata per il foglio: " + sheetName);
  }
}

function removeEditorFromProtectedSheet(sheetName, email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    //Logger.log("Foglio non trovato: " + sheetName);
    return;
  }

  // Ottieni tutte le protezioni del foglio
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);

  // Controlla se esiste una protezione per il foglio
  if (protections.length > 0) {
    var protection = protections[0]; // Prendi la prima protezione trovata

    // Ottieni la lista degli editor autorizzati
    var editors = protection.getEditors();

    // Verifica se l'utente è presente nella lista degli editor autorizzati
    for (let i = 0; i < editors.length; i++) {
      if (editors[i].getEmail() === email) {
        // Rimuovi l'utente specificato dagli editor autorizzati
        protection.removeEditor(email);
        //Logger.log("Permesso di modifica revocato per l'utente: " + email);
        return;
      }
    }

    //Logger.log("L'utente " + email + " non è presente nella lista degli editor autorizzati.");
  } else {
    //Logger.log("Nessuna protezione trovata per il foglio: " + sheetName);
  }
}

function isValidEmail(email) {
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

function protectSheetByName(sheetName, userEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    var protection = sheet.protect().setDescription('Automatic protection');

    var ownerEmail = Session.getEffectiveUser().getEmail();

    // Let write permission only to owner and admin
    protection.addEditor(ownerEmail);
    protection.addEditor(getRealEmail(userEmail));

    // Remove others authorizations
    protection.removeEditors(protection.getEditors());

    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  } else {
    //Logger.log('Sheet not found: ' + sheetName);
  }
}

function completeMenu() {
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  if (findKey(user, users(), 0) >= 0) {
    var permission = users()[findKey(user, users(), 0)][1];
    SpreadsheetApp.getUi()
      .showSidebar(doGet(permission, '05_completeMenuPage', translate('viewCalendar.mainCell') + ' Sidebar'));
  } else {
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    ui.alert(translate('alert.userPermission', { user: user }));
  }
}

function specialEvent() {
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  if ((users()[findKey(user, users(), 0)][1] == 'admin') || (users()[findKey(user, users(), 0)][1] == 'writer')) {
    createUserSheet();
    var structures = onlyStrcturesSelect(strutture());
    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '8_specailEventPage', translate('main.specialEvent')));
    //Utilities.sleep(3000);
    //SpreadsheetApp.flush();
    //showMonths(primo, ultimo, '', '', '');
  } else {
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    ui.alert(translate('alert.userPermission', { user: user }));
  }
}

function specialDailyEvent() {
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  if ((users()[findKey(user, users(), 0)][1] == 'admin') || (users()[findKey(user, users(), 0)][1] == 'writer')) {
    createUserSheet();
    //updateTimeUser();
    /*
    var today = convertDateInputHtml(new Date());
    var primo = convertDateInputHtml(text2monthDays(today)[0]);
    var ultimo = convertDateInputHtml(text2monthDays(today)[2]);
    if (Number(preloadSheet()[0])) {showMonths(primo, ultimo, '', '', '');}
    */
    var structures = onlyStrcturesSelect(strutture());
    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '6E_specialDailyEvent', translate('menu.specialDailyEvent')));
    //Utilities.sleep(3000);
    //SpreadsheetApp.flush();
    //showMonths(primo, ultimo, '', '', '');
  } else {
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    ui.alert(translate('alert.userPermission', { user: user }));
  }
}

function modifyEvent() {
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  if ((users()[findKey(user, users(), 0)][1] == 'admin') || (users()[findKey(user, users(), 0)][1] == 'writer')) {
    createUserSheet();
    var today = convertDateInputHtml(new Date());
    var primo = convertDateInputHtml(text2monthDays(today)[0]);
    var ultimo = convertDateInputHtml(text2monthDays(today)[2]);
    if (preloadSheet()) { showMonths(primo, ultimo, '', '', ''); }
    var structures = onlyStrcturesSelect(strutture());
    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '2A_modifyEventPage', translate('main.modifyEventPage')));
  } else {
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    ui.alert(translate('alert.userPermission', { user: user }));
  }
}

function viewCalendar() {
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  if (findKey(user, users(), 0) >= 0) {
    createUserSheet();
    //updateTimeUser();
    var today = convertDateInputHtml(new Date());
    var primo = convertDateInputHtml(text2monthDays(today)[0]);
    var ultimo = convertDateInputHtml(text2monthDays(today)[2]);
    if (preloadSheet()) { showOldMonths(primo, ultimo, '', '', ''); }
    var structures = onlyStrcturesSelect(strutture());
    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '3_viewCalendarPage', translate('main.viewCalendarPage')));
    //showMonths(primo, ultimo, '', '', '');
  } else {
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    ui.alert(translate('alert.userPermission', { user: user }));
  }
}

function viewDailyCalendar() {
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  if (findKey(user, users(), 0) >= 0) {
    createUserSheet();
    var structures = onlyStrcturesSelect(strutture());
    var today = new Date();
    var oggi = formatDateMaster(today).dataXweb;
    if (preloadSheet()) { createSlideAndExportToSheet(oggi, oggi, 'quartiere', '', 'NO'); } // cc o quartiere
    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '4_viewDaySlidePage', translate('main.viewDaySlidePage')));
    //createSlideAndExportToSheet(oggi, oggi, 'quartiere', '', 'NO'); // cc o quartiere    
  } else {
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    ui.alert(translate('alert.userPermission', { user: user }));
  }
}

function viewListCalendar() {
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  if (findKey(user, users(), 0) >= 0) {
    createUserSheet();
    //updateTimeUser();
    var today = convertDateInputHtml(new Date());
    var primo = convertDateInputHtml(text2monthDays(today)[0]);
    var oggi = convertDateInputHtml(text2monthDays(today)[1]);
    var ultimo = convertDateInputHtml(text2monthDays(today)[2]);
    var permission = users()[findKey(user, users(), 0)][1];
    var structures = onlyStrcturesSelect(strutture());
    structures.push(permission);
    //Logger.log('Primo è '+primo+ ' ultimo è '+ultimo);
    if (preloadSheet()) { createListEvent(oggi, ultimo, 'E', '', '', 'agg'); }
    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '5_viewListPage', translate('main.viewListPage')));
    //createListEvent(oggi, ultimo, 'E', '', '', 'agg');
  } else {
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    ui.alert(translate('alert.userPermission', { user: user }));
  }
}

function manageSmallRoom() {
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  if (findKey(user, users(), 0) >= 0) {
    createUserSheet();
    var permission = users()[findKey(user, users(), 0)][1];
    specialDailyEvent();
  } else {
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    ui.alert(translate('alert.userPermission', { user: user }));
  }
}

function viewMSRData() {
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  if (findKey(user, users(), 0) >= 0) {
    createUserSheet();
    //updateTimeUser();
    var structures = onlyStrcturesSelect(strutture());
    var permission = users()[findKey(user, users(), 0)][1];
    var today = convertDateInputHtml(new Date());
    var oggi = convertDateInputHtml(text2monthDays(today)[1]);
    //Logger.log('Primo è '+primo+ ' ultimo è '+ultimo);
    if (preloadSheet()) { createDailyScheduleFromCalendar(oggi, '60', '', ''); }
    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '6B_viewMSRPage', translate('main.viewMSRPage')));
    //createDailyScheduleFromCalendar(oggi, '60', '', '');
  } else {
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    ui.alert(translate('alert.userPermission', { user: user }));
  }
}

function editMSRData() {
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  if ((users()[findKey(user, users(), 0)][1] == 'admin') || (users()[findKey(user, users(), 0)][1] == 'writer')) {
    createUserSheet();
    var structures = onlyStrcturesSelect(strutture());
    var today = convertDateInputHtml(new Date());
    var oggi = convertDateInputHtml(text2monthDays(today)[1]);
    if (preloadSheet()) { createDailyScheduleFromCalendar(oggi, '60', '', ''); }
    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '6D1_editAskMSRPage', 'Edit a daily event'));
    //createDailyScheduleFromCalendar(oggi, '60', '', '');
  } else {
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    ui.alert(translate('alert.userPermission', { user: user }));
  }
}