/*
* Project Name: Pavora
* Copyright (c) 2025 Nicola Rainiero
*
* This software is released under the MIT License.
* Please refer to the LICENSE file for the full license text.
*/
// insertFloatingImage(sheet, fileId, numRow, numCol, imgWidth, imgHeight)
function insertFloatingImage(sheet, fileId, numRow, numCol, imgWidth, imgHeight) {
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();

  // Inserisci l'immagine come oggetto fluttuante
  var image = sheet.insertImage(blob, numRow, numCol); // Righe e colonne di partenza

  // Imposta la dimensione naturale dell'immagine
  image.setWidth(imgWidth).setHeight(imgHeight);
}

function addFilterRegisterPage() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetsList()[0][0]); // Cambia "Sheet1" con il nome del tuo foglio

  // Rimuove eventuali filtri esistenti
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  // Aggiungi filtri a tutte le colonne
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRow.setFontWeight('bold');
  headerRow.setBackground('grey');

  var filterRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  var filter = filterRange.createFilter();

  // Ottieni i dati dalla colonna "Strutture"
  var dataRange = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1); // Dalla seconda riga, colonna 3 (Strutture)
  var data = dataRange.getValues();

  // Estrai tutte le strutture uniche
  var uniqueStructures = new Set();
  data.forEach(row => {
    var structures = row[0].split(',').map(s => s.trim());
    structures.forEach(structure => {
      if (structure) {
        uniqueStructures.add(structure);
      }
    });
  });

  // Converti il Set in Array e ordina le strutture
  var uniqueStructuresArray = Array.from(uniqueStructures).sort();


  // Applica il filtro personalizzato alla colonna "Strutture"
  var criteria = SpreadsheetApp.newFilterCriteria()
    .build();
  filter.setColumnFilterCriteria(3, criteria);
}

// -------------------------------------------------------------------------------------
// Gestione permessi utenti (admin, writer, reader, deleted) --> basta verificare la tabella e lanciare manageAccess
// -------------------------------------------------------------------------------------
function manageAccess() {
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  if (users()[findKey(user, users(), 0)][1] == 'admin') {
    // Dati di esempio: array di utenti e ruoli
    var utenti = users();

    // ID del calendario Google
    var calendarId = myCalID()[0][0];
    //var calendarIdLav = myCalID()[1][0];
    // ID dei file Google Sheets
    var actionFileId = driveIDFiles()[0][0];
    var variabiliFileId = driveIDFiles()[1][0];
    var aliasEmail = driveIDFiles()[2][0];    
    var slideQuartiere = templateSlides()[0][0];
    var slideCC = templateSlides()[1][0];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    sh = ss.getSheetByName(sheetsList()[1][0]);
    var lr = sh.getLastRow();
    var currentTime = new Date();

    // Gestisci le condivisioni per ciascun utente
    utenti.forEach(function (user) {
      var email = getRealEmail(user[0]);
      var role = user[1];
      Logger.log(email + '   ' + role);

      if (role === 'deleted') {
        // Rimuove l'accesso al calendario e ai file per gli utenti con ruolo "rimosso"
        //removeAccess(email, calendarId, calendarIdLav, variabiliFileId, actionFileId, slideQuartiere, slideCC);
        removeAccess(email, calendarId, variabiliFileId, aliasEmail, actionFileId, slideQuartiere, slideCC);
      } else {
        // Gestisci l'accesso al calendario Google
        manageCalendarAccess(email, calendarId);
        //manageCalendarAccess(email, calendarIdLav);

        // Gestisci l'accesso ai file Google Sheets
        manageFileAccess(email, role, variabiliFileId, aliasEmail, actionFileId, slideQuartiere, slideCC);

        // Gestisci la registrazione nella scheda utenti online
        //Logger.log(usersOnline().length);
        if ((role === 'writer') && (findKey(getAliasEmail(email), usersOnline(), 0) < 0)) {
          lr += 1;
          sh.getRange(lr, 1).setValue(getAliasEmail(email));
          SpreadsheetApp.getUi().alert(translate('admin.addUser', { user: user[0] }));
          sh.getRange(lr, 2).setValue(currentTime).setNumberFormat('dd/MM/yy - HH:mm');
        }

      }
    });
  } else {
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    ui.alert(translate('alert.userPermission', { user: user }));
  }
}

// Funzione per gestire l'accesso al calendario
function manageCalendarAccess(email, calendarId) {
  //try {
    // Recupera l'elenco ACL del calendario
    var acl = Calendar.Acl.list(calendarId);
    var alreadyShared = acl.items.some(function (entry) {
      return entry.scope.type === 'user' && entry.scope.value === email;
    });

    var message = '';

    if (!alreadyShared) {
      // Aggiunge l'utente come visualizzatore del calendario (freeBusyReader)
      Calendar.Acl.insert({
        'scope': {
          'type': 'user',
          'value': email
        },
        'role': 'reader'
      }, calendarId);
      message = message + translate('admin.allowUser') + getAliasEmail(email) + '\n';
    } else {
      message = message + translate('admin.allowUser') + getAliasEmail(email) + '\n';
    }
    SpreadsheetApp.getUi().alert(message);
    /*
  } catch (e) {
    SpreadsheetApp.getUi().alert(translate('admin.errorManage') + e.toString());
  }
  */
}

/**
 * Verifica se l'utente che esegue lo script è il proprietario di un file.
 * @param {string} fileId L'ID del file di Google Drive.
 * @returns {boolean} True se l'utente è il proprietario del file, altrimenti false.
 */
function isMine(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const ownerEmail = file.getOwner().getEmail();
    const myEmail = Session.getActiveUser().getEmail();
    
    return ownerEmail === myEmail;
  } catch (e) {
    // Se il file non esiste o non hai i permessi per vederlo,
    // la funzione fallisce e restituisce false.
    //Logger.log('Errore nella verifica del file: ' + e.message);
    return false;
  }
}

// Funzione per gestire l'accesso ai file Google Sheets
function manageFileAccess(email, role, variabiliFileId, aliasEmailId, actionFileId, slideQuartiereId, slideCCId) {
  //try {
    var variabiliFile = DriveApp.getFileById(variabiliFileId);
    var actionFile = DriveApp.getFileById(actionFileId);
    var aliasFile = DriveApp.getFileById(aliasEmailId);
    var slideQuartiere = DriveApp.getFileById(slideQuartiereId);
    var slideCC = DriveApp.getFileById(slideCCId);

    // Gestisci l'accesso al file "action" (scrittura per tutti)
    var actionPermissions = actionFile.getEditors();
    //Logger.log(user.getEmail());
    var alreadyEditor = actionPermissions.some(user => user.getEmail() === email);

    var message = '';

    if (!alreadyEditor) {
      actionFile.addEditor(email);
      message = message + translate('admin.allowWrite', { file: '\'Pavora\'' }) + getAliasEmail(email) + '\n';
    } else {
      message = message + email + translate('admin.alreadyWrite', { file: '\'Pavora\'' }) + getAliasEmail(email) + '\n';
    }

    // Gestisci l'accesso al file "variabili" (in base al ruolo)
    var variabiliPermissions = variabiliFile.getEditors();
    var variabiliViewers = variabiliFile.getViewers();
    var alreadyViewer = variabiliViewers.some(user => user.getEmail() === email);
    var alreadyVariabiliEditor = variabiliPermissions.some(user => user.getEmail() === email);

    if (role === 'admin') {
      if (!alreadyVariabiliEditor) {
        variabiliFile.addEditor(email);
        if (isMine(slideQuartiere)) {slideQuartiere.addEditor(email)}
        if (isMine(slideCC)) {slideCC.addEditor(email)}
        message = message + translate('admin.allowWrite', { file: '\'PavoraCustomSettings\'' }) + getAliasEmail(email) + '\n';
      } else {
        message = message + email + translate('admin.alreadyWrite', { file: '\'PavoraCustomSettings\'' }) + '\n';
      }
    } else if (role === 'writer') {
      if (!alreadyVariabiliEditor) {
        aliasFile.addViewer(email);
        variabiliFile.addViewer(email);
        if (isMine(slideQuartiere)) {slideQuartiere.addViewer(email)}
        if (isMine(slideCC)) {slideCC.addViewer(email)}
        message = message + translate('admin.allowRead', { file: '\'PavoraCustomSettings\'' }) + getAliasEmail(email) + '\n';
      } else {
        message = message + email + translate('admin.alreadyRead', { file: '\'PavoraCustomSettings\'' }) + '\n';
      }
    } else if (role === 'reader') {
      if (!alreadyViewer) {
        aliasFile.addViewer(email);
        variabiliFile.addViewer(email);
        if (isMine(slideQuartiere)) {slideQuartiere.addViewer(email)}
        if (isMine(slideCC)) {slideCC.addViewer(email)}
        message = message + translate('admin.allowRead', { file: '\'PavoraCustomSettings\'' }) + getAliasEmail(email) + '\n';
      } else {
        message = message + email + translate('admin.alreadyRead', { file: '\'PavoraCustomSettings\'' }) + '\n';
      }
      SpreadsheetApp.getUi().alert(message);
    }
    /*
  } catch (e) {
    SpreadsheetApp.getUi().alert(translate('admin.errorManage') + e.toString());
  }
  */
}

// Funzione per rimuovere l'accesso a calendario e file
//function removeAccess(email, calendarId, calendarIdLav, variabiliFileId, actionFileId, slideQuartiereId, slideCCId) {
function removeAccess(email, calendarId, variabiliFileId, aliasEmailId, actionFileId, slideQuartiereId, slideCCId) {
  //try {
    // Rimuove l'accesso al calendario
    var message = '';
    var acl = Calendar.Acl.list(calendarId);
    acl.items.forEach(function (entry) {
      if (entry.scope.type === 'user' && entry.scope.value === email) {
        Calendar.Acl.remove(calendarId, entry.id);
        message = message + translate('admin.remCalendar') + getAliasEmail(email);
      }
    });

    // Rimuove l'accesso al file "variabili" e "slide"
    var variabiliFile = DriveApp.getFileById(variabiliFileId);
    var aliasFile = DriveApp.getFileById(aliasEmailId);
    if (variabiliFile.getEditors().some(user => user.getEmail() === email)) {
      variabiliFile.removeEditor(email);
      message = message + translate('admin.removeWrite', { file: '\'PavoraCustomSettings\'' }) + getAliasEmail(email) + '\n';
    }
    if (variabiliFile.getViewers().some(user => user.getEmail() === email)) {
      variabiliFile.removeViewer(email);
      message = message + translate('admin.removeRead', { file: '\'PavoraCustomSettings\'' }) + getAliasEmail(email) + '\n';
    }
    if (aliasFile.getViewers().some(user => user.getEmail() === email)) {
      aliasFile.removeViewer(email);
      message = message + translate('admin.removeRead', { file: '\'PavoraCustomSettings\'' }) + getAliasEmail(email) + '\n';
    }
    if (isMine(slideQuartiere)) {
    var slideQuartiere = DriveApp.getFileById(slideQuartiereId);
    if (slideQuartiere.getViewers().some(user => user.getEmail() === email)) {
      slideQuartiere.removeViewer(email);
      message = message + translate('admin.removeRead', { file: '\'slideQ\'' }) + getAliasEmail(email) + '\n';
    }
    }
    if (isMine(slideCC)) {
    var slideCC = DriveApp.getFileById(slideCCId);
    if (slideCC.getViewers().some(user => user.getEmail() === email)) {
      slideCC.removeViewer(email);
      message = message + translate('admin.removeRead', { file: '\'slideCC\'' }) + getAliasEmail(email) + '\n';
    }
    }
    // Rimuove l'accesso al file "action"
    var actionFile = DriveApp.getFileById(actionFileId);
    if (actionFile.getEditors().some(user => user.getEmail() === email)) {
      actionFile.removeEditor(email);
      message = message + translate('admin.removeWrite', { file: '\'Pavora\'' }) + getAliasEmail(email) + '\n';
    }
    SpreadsheetApp.getUi().alert(message);
    /*
  } catch (e) {
    SpreadsheetApp.getUi().alert(translate('admin.errorManage') + e.toString());
  }
  */
}
// -------------------------------------------------------------------------------------

function adminExecCreateEvents() {
  var array = [["Second Event P", "Thu Jan 16 2025 08:00:00 GMT+0100 (Ora standard dell’Europa centrale)", "Thu Jan 16 2025 18:00:00 GMT+0100 (Ora standard dell’Europa centrale)", " all=sf1 feed=nd id=Jvocj5c5 typeEv=evFreeReg org=Not given refCom=ns4 refOp=ns8 open=SI ", "H6, H7, H8, A8", "P", "ns8", "Jvocj5c5"], ["Second Event A", "Fri Jan 17 2025 08:00:00 GMT+0100 (Ora standard dell’Europa centrale)", "Fri Jan 17 2025 18:00:00 GMT+0100 (Ora standard dell’Europa centrale)", " all=sf1 feed=nd id=Jvocj5c5 typeEv=evFreeReg org=Not given refCom=ns4 refOp=ns8 open=SI ", "H6, H7, H8, I, A8", "A", "ns8", "Jvocj5c5"], ["Second Event A", "Sat Jan 18 2025 08:00:00 GMT+0100 (Ora standard dell’Europa centrale)", "Sat Jan 18 2025 18:00:00 GMT+0100 (Ora standard dell’Europa centrale)", " all=sf1 feed=nd id=Jvocj5c5 typeEv=evFreeReg org=Not given refCom=ns4 refOp=ns8 open=SI ", "H6, H7, H8, I, A8", "A", "ns8", "Jvocj5c5"], ["Second Event A", "Sun Jan 19 2025 08:00:00 GMT+0100 (Ora standard dell’Europa centrale)", "Sun Jan 19 2025 18:00:00 GMT+0100 (Ora standard dell’Europa centrale)", " all=sf1 feed=nd id=Jvocj5c5 typeEv=evFreeReg org=Not given refCom=ns4 refOp=ns8 open=SI ", "H6, H7, H8, I, A8", "A", "ns8", "Jvocj5c5"], ["Second Event E", "Mon Jan 20 2025 09:00:00 GMT+0100 (Ora standard dell’Europa centrale)", "Mon Jan 20 2025 23:00:00 GMT+0100 (Ora standard dell’Europa centrale)", " all=sf1 feed=nd id=Jvocj5c5 typeEv=evFreeReg org=Not given refCom=ns4 refOp=ns8 open=SI  vvf=SI cri=SI color=rossoChiaro", "H6, H7, H8, I, A8, R7, GP", "E", "ns8", "Jvocj5c5"], ["Second Event E", "Tue Jan 21 2025 09:00:00 GMT+0100 (Ora standard dell’Europa centrale)", "Tue Jan 21 2025 23:00:00 GMT+0100 (Ora standard dell’Europa centrale)", " all=sf1 feed=nd id=Jvocj5c5 typeEv=evFreeReg org=Not given refCom=ns4 refOp=ns8 open=SI  vvf=SI cri=SI color=rossoChiaro", "H6, H7, H8, I, A8, R7, GP", "E", "ns8", "Jvocj5c5"], ["Second Event E", "Wed Jan 22 2025 09:00:00 GMT+0100 (Ora standard dell’Europa centrale)", "Wed Jan 22 2025 13:00:00 GMT+0100 (Ora standard dell’Europa centrale)", " all=sf1 feed=nd id=Jvocj5c5 typeEv=evFreeReg org=Not given refCom=ns4 refOp=ns8 open=SI  vvf=SI cri=SI color=rossoChiaro", "H6, H7, H8, I, A8, R7, GP", "E", "ns8", "Jvocj5c5"], ["Second Event D", "Wed Jan 22 2025 13:00:00 GMT+0100 (Ora standard dell’Europa centrale)", "Wed Jan 22 2025 20:00:00 GMT+0100 (Ora standard dell’Europa centrale)", " all=sf1 feed=nd id=Jvocj5c5 typeEv=evFreeReg org=Not given refCom=ns4 refOp=ns8 open=SI ", "H6, H7, H8, I, A8", "D", "ns8", "Jvocj5c5"], ["Second Event D", "Thu Jan 23 2025 08:00:00 GMT+0100 (Ora standard dell’Europa centrale)", "Thu Jan 23 2025 18:00:00 GMT+0100 (Ora standard dell’Europa centrale)", " all=sf1 feed=nd id=Jvocj5c5 typeEv=evFreeReg org=Not given refCom=ns4 refOp=ns8 open=SI ", "H6, H7, H8, I, A8", "D", "ns8", "Jvocj5c5"], ["Second Event D", "Fri Jan 24 2025 08:00:00 GMT+0100 (Ora standard dell’Europa centrale)", "Fri Jan 24 2025 18:00:00 GMT+0100 (Ora standard dell’Europa centrale)", " all=sf1 feed=nd id=Jvocj5c5 typeEv=evFreeReg org=Not given refCom=ns4 refOp=ns8 open=SI ", "H6, H7, H8, I, A8", "D", "ns8", "Jvocj5c5"]];
  adminCreateEvents(array);
}

function testTypeEvent() {
  var testo = ' 1000 all=altro id=gYyBMrJ5 typeEv=csi org=Sistema Congressi - Università refCom= refOp= open=SI  vvf=SI cri=SI';
  var array = typeEv();
  var indice = findKey(parseEventDetails(testo).typeEv, array, 1);
}

function adminDeleteEvent() {
  //deleteEventsNoConfirm('ProvaLO', '2024-06-10', '2024-07-24') // con title
  //deleteEventsNoConfirm('neUMx4qH', '', '') // con keyword
  //deleteEvents('1bFtj8li', '', '') // con keyword
  //deleteEvents(eventId, first, last, what)
  deleteEvents('', '2024-10-01', '2024-10-31') // con keyword + 2 giorni!!!
}

function copyEvents() {
  var sourceCalendarId = myCalID()[0][0]; // Replace with Calendar A ID
  var destinationCalendarId = '58e6237414a0a05fb22cfef75ed0073ff99f38509b6b45ab3c373b215ea760f2@group.calendar.google.com'; // Replace with Calendar B ID
  var startDate = new Date('2024-11-30'); // Set your start date
  var endDate = new Date('2024-12-01'); // Set your end date

  var events = CalendarApp.getCalendarById(sourceCalendarId).getEvents(startDate, endDate);

  for (let i = 0; i < events.length; i++) {
    var event = events[i];
    CalendarApp.getCalendarById(destinationCalendarId).createEvent(event.getTitle(),
      event.getStartTime(),
      event.getEndTime(),
      {
        description: event.getDescription(),
        location: event.getLocation(),
        guests: event.getGuestList().map(guest => guest.getEmail()).join(),
        sendInvites: false
      }
    );
  }
}

function adminCreateEvents(array) {
  createUserSheet();
  updateTimeUser();
  /*
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log('first è ' + typeof (first) + ' ' + first);
  Logger.log('last è ' + typeof (last) + ' ' + last);
  sh.getRange(1, 1).setValue('Data iniziale');
  sh.getRange(2, 1).setValue(first).setNumberFormat('dd-MM-yyyy');
  sh.getRange(3, 1).setValue('Data finale');
  sh.getRange(4, 1).setValue(last).setNumberFormat('dd-MM-yyyy');
  sh.getRange(5, 1).setValue(array);
  sh.getRange(6, 1).setValue(typeof (array));
  var matrix = stringToMatrix(array, 8)
  sh.getRange(7, 1).setValue(matrix[0]);
  sh.getRange(8, 1).setValue(typeof (matrix));
  if (findKey('E', matrix, 5) >= 0) {
    showMonths(first, last, matrix[findKey('E', matrix, 5)][4]);
  } else {
    showMonths(first, last, matrix[0][4]);
  }
  */
  oggi = new Date();
  utenteEmail = Session.getEffectiveUser().getEmail();
  var eventID = (parseEventDetails(matrix[0][3]).id != '') ? parseEventDetails(matrix[0][3]).id + ' |-> ' + parseEventString(matrix[0][0]).nome : parseEventString(matrix[0][0]).nome;
  addLogRevision(oggi, translate('admin.recoverAdmin'), eventID, utenteEmail, matrix);
  //viewCalendar(); // to refresh dialog window
}
