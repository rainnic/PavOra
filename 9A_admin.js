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
// Sistema di Gestione Permessi Utenti - Versione Migliorata
// -------------------------------------------------------------------------------------
/**
 * Configurazione dei ruoli e permessi
 */
const ROLES = {
  ADMIN: 'admin',
  WRITER: 'writer',
  READER: 'reader',
  DELETED: 'deleted'
};

const PERMISSIONS = {
  CALENDAR: 'reader',
  DRIVE_EDITOR: 'editor',
  DRIVE_VIEWER: 'viewer'
};

/**
 * Classe principale per la gestione dei permessi
 */
class UserPermissionsManager {
  constructor() {
    this.ui = SpreadsheetApp.getUi();
    this.currentUser = Session.getEffectiveUser().getEmail();
    this.users = users();
    this.resources = this.initializeResources();
  }

  /**
   * Inizializza le risorse (IDs di file e calendari)
   */
  initializeResources() {
    try {
      return {
        calendarId: myCalID()[0][0],
        actionFileId: driveIDFiles()[0][0],
        variabiliFileId: driveIDFiles()[1][0],
        slideQuartiere: templateSlides()[0][0],
        slideCC: templateSlides()[1][0]
      };
    } catch (error) {
      this.showError('Errore nell\'inizializzazione delle risorse', error);
      throw error;
    }
  }

  /**
   * Funzione principale per gestire tutti i permessi
   */
  manageAccess() {
    try {
      // Verifica se l'utente corrente è admin
      if (!this.isCurrentUserAdmin()) {
        this.ui.alert(translate('alert.userPermission', { user: this.currentUser }));
        return;
      }

      const results = this.processAllUsers();
      this.showResults(results);

    } catch (error) {
      this.showError('Errore generale nella gestione dei permessi', error);
    }
  }

  /**
   * Verifica se l'utente corrente è admin
   */
  isCurrentUserAdmin() {
    const userIndex = findKey(this.currentUser, this.users, 0);
    return userIndex >= 0 && this.users[userIndex][1] === ROLES.ADMIN;
  }

  /**
   * Processa tutti gli utenti
   */
  processAllUsers() {
    const results = {
      success: [],
      errors: [],
      newWriters: []
    };

    this.users.forEach(user => {
      try {
        const [email, role] = user;
        const userResult = this.processUser(email, role);
        
        if (userResult.success) {
          results.success.push(userResult);
          if (userResult.newWriter) {
            results.newWriters.push(email);
          }
        } else {
          results.errors.push(userResult);
        }
      } catch (error) {
        results.errors.push({
          email: user[0],
          error: error.toString()
        });
      }
    });

    return results;
  }

  /**
   * Processa un singolo utente
   */
  processUser(email, role) {
    const result = {
      email: email,
      role: role,
      success: false,
      actions: [],
      newWriter: false
    };

    try {
      if (role === ROLES.DELETED) {
        result.actions = this.removeUserAccess(email);
      } else {
        result.actions = this.grantUserAccess(email, role);
        result.newWriter = this.handleNewWriter(email, role);
      }
      
      result.success = true;
    } catch (error) {
      result.error = error.toString();
    }

    return result;
  }

  /**
   * Concede accesso a un utente
   */
  grantUserAccess(email, role) {
    const actions = [];
    
    // Gestione accesso calendario
    const calendarResult = this.manageCalendarAccess(email);
    if (calendarResult) actions.push(calendarResult);

    // Gestione accesso file
    const fileResults = this.manageFileAccess(email, role);
    actions.push(...fileResults);

    return actions;
  }

  /**
   * Rimuove accesso a un utente
   */
  removeUserAccess(email) {
    const actions = [];
    
    // Rimuove accesso calendario
    const calendarResult = this.removeCalendarAccess(email);
    if (calendarResult) actions.push(calendarResult);

    // Rimuove accesso file
    const fileResults = this.removeFileAccess(email);
    actions.push(...fileResults);

    return actions;
  }

  /**
   * Gestisce l'accesso al calendario
   */
  manageCalendarAccess(email) {
    try {
      const acl = Calendar.Acl.list(this.resources.calendarId);
      const alreadyShared = acl.items.some(entry => 
        entry.scope.type === 'user' && entry.scope.value === email
      );

      if (!alreadyShared) {
        Calendar.Acl.insert({
          'scope': {
            'type': 'user',
            'value': email
          },
          'role': PERMISSIONS.CALENDAR
        }, this.resources.calendarId);
        
        return `Accesso calendario concesso a ${email}`;
      } else {
        return `${email} ha già accesso al calendario`;
      }
    } catch (error) {
      throw new Error(`Errore gestione calendario per ${email}: ${error.toString()}`);
    }
  }

  /**
   * Rimuove l'accesso al calendario
   */
  removeCalendarAccess(email) {
    try {
      const acl = Calendar.Acl.list(this.resources.calendarId);
      const userEntry = acl.items.find(entry => 
        entry.scope.type === 'user' && entry.scope.value === email
      );

      if (userEntry) {
        Calendar.Acl.remove(this.resources.calendarId, userEntry.id);
        return `Accesso calendario rimosso per ${email}`;
      }
      return null;
    } catch (error) {
      throw new Error(`Errore rimozione calendario per ${email}: ${error.toString()}`);
    }
  }

  /**
   * Gestisce l'accesso ai file
   */
  manageFileAccess(email, role) {
    const actions = [];
    const fileConfigs = this.getFileConfigs(role);

    fileConfigs.forEach(config => {
      try {
        const file = DriveApp.getFileById(config.fileId);
        const action = this.setFilePermission(file, email, config.permission, config.name);
        if (action) actions.push(action);
      } catch (error) {
        actions.push(`Errore gestione file ${config.name} per ${email}: ${error.toString()}`);
      }
    });

    return actions;
  }

  /**
   * Rimuove l'accesso ai file
   */
  removeFileAccess(email) {
    const actions = [];
    const allFiles = [
      { id: this.resources.actionFileId, name: 'Pavora' },
      { id: this.resources.variabiliFileId, name: 'PavoraCustomSettings' },
      { id: this.resources.slideQuartiere, name: 'slideQ' },
      { id: this.resources.slideCC, name: 'slideCC' }
    ];

    allFiles.forEach(fileInfo => {
      try {
        const file = DriveApp.getFileById(fileInfo.id);
        const action = this.removeFilePermission(file, email, fileInfo.name);
        if (action) actions.push(action);
      } catch (error) {
        actions.push(`Errore rimozione file ${fileInfo.name} per ${email}: ${error.toString()}`);
      }
    });

    return actions;
  }

  /**
   * Configura i file in base al ruolo
   */
  getFileConfigs(role) {
    const configs = [];

    // File action - accesso in scrittura per tutti
    configs.push({
      fileId: this.resources.actionFileId,
      permission: PERMISSIONS.DRIVE_EDITOR,
      name: 'Pavora'
    });

    // File variabili e slide - permessi in base al ruolo
    const variabiliPermission = role === ROLES.ADMIN ? PERMISSIONS.DRIVE_EDITOR : PERMISSIONS.DRIVE_VIEWER;
    
    configs.push(
      {
        fileId: this.resources.variabiliFileId,
        permission: variabiliPermission,
        name: 'PavoraCustomSettings'
      },
      {
        fileId: this.resources.slideQuartiere,
        permission: variabiliPermission,
        name: 'slideQ'
      },
      {
        fileId: this.resources.slideCC,
        permission: variabiliPermission,
        name: 'slideCC'
      }
    );

    return configs;
  }

  /**
   * Imposta i permessi per un file
   */
  setFilePermission(file, email, permission, fileName) {
    const editors = file.getEditors();
    const viewers = file.getViewers();
    
    const isEditor = editors.some(user => user.getEmail() === email);
    const isViewer = viewers.some(user => user.getEmail() === email);

    if (permission === PERMISSIONS.DRIVE_EDITOR) {
      if (!isEditor) {
        if (isViewer) file.removeViewer(email);
        file.addEditor(email);
        return `Accesso scrittura concesso a ${email} per ${fileName}`;
      } else {
        return `${email} ha già accesso in scrittura a ${fileName}`;
      }
    } else if (permission === PERMISSIONS.DRIVE_VIEWER) {
      if (!isViewer && !isEditor) {
        file.addViewer(email);
        return `Accesso lettura concesso a ${email} per ${fileName}`;
      } else {
        return `${email} ha già accesso a ${fileName}`;
      }
    }

    return null;
  }

  /**
   * Rimuove i permessi per un file
   */
  removeFilePermission(file, email, fileName) {
    const editors = file.getEditors();
    const viewers = file.getViewers();
    
    const isEditor = editors.some(user => user.getEmail() === email);
    const isViewer = viewers.some(user => user.getEmail() === email);

    const actions = [];

    if (isEditor) {
      file.removeEditor(email);
      actions.push(`Accesso scrittura rimosso per ${email} su ${fileName}`);
    }
    
    if (isViewer) {
      file.removeViewer(email);
      actions.push(`Accesso lettura rimosso per ${email} su ${fileName}`);
    }

    return actions.length > 0 ? actions.join('; ') : null;
  }

  /**
   * Gestisce l'aggiunta di nuovi writer
   */
  handleNewWriter(email, role) {
    if (role !== ROLES.WRITER) return false;

    const usersOnlineList = usersOnline();
    if (findKey(email, usersOnlineList, 0) >= 0) return false;

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(sheetsList()[1][0]);
      const lr = sh.getLastRow() + 1;
      const currentTime = new Date();

      sh.getRange(lr, 1).setValue(email);
      sh.getRange(lr, 2).setValue(currentTime).setNumberFormat('dd/MM/yy - HH:mm');
      
      return true;
    } catch (error) {
      throw new Error(`Errore nell'aggiunta del writer ${email}: ${error.toString()}`);
    }
  }

  /**
   * Mostra i risultati delle operazioni
   */
  showResults(results) {
    let message = '';
    
    if (results.success.length > 0) {
      message += `✅ Utenti processati con successo: ${results.success.length}\n`;
      
      if (results.newWriters.length > 0) {
        message += `📝 Nuovi writer aggiunti: ${results.newWriters.join(', ')}\n`;
      }
    }

    if (results.errors.length > 0) {
      message += `❌ Errori: ${results.errors.length}\n`;
      results.errors.forEach(error => {
        message += `- ${error.email}: ${error.error}\n`;
      });
    }

    if (message) {
      this.ui.alert(message);
    }
  }

  /**
   * Mostra errori
   */
  showError(context, error) {
    const message = `${context}: ${error.toString()}`;
    Logger.log(message);
    this.ui.alert(message);
  }
}

// -------------------------------------------------------------------------------------
// Funzioni di interfaccia pubblica (per mantenere compatibilità)
// -------------------------------------------------------------------------------------

/**
 * Funzione principale per la gestione dei permessi
 */
function manageAccess() {
  const manager = new UserPermissionsManager();
  manager.manageAccess();
}

/**
 * Funzione per gestire l'accesso al calendario (compatibilità)
 */
function manageCalendarAccess(email, calendarId) {
  try {
    const manager = new UserPermissionsManager();
    return manager.manageCalendarAccess(email);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Errore gestione calendario: ${error.toString()}`);
  }
}

/**
 * Funzione per gestire l'accesso ai file (compatibilità)
 */
function manageFileAccess(email, role, variabiliFileId, actionFileId, slideQuartiereId, slideCCId) {
  try {
    const manager = new UserPermissionsManager();
    return manager.manageFileAccess(email, role);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Errore gestione file: ${error.toString()}`);
  }
}

/**
 * Funzione per rimuovere l'accesso (compatibilità)
 */
function removeAccess(email, calendarId, variabiliFileId, actionFileId, slideQuartiereId, slideCCId) {
  try {
    const manager = new UserPermissionsManager();
    return manager.removeUserAccess(email);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Errore rimozione accesso: ${error.toString()}`);
  }
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