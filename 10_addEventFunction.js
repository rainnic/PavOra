/*
* Project Name: Pavora
* Copyright (c) 2025 Nicola Rainiero
*
* This software is released under the MIT License.
* Please refer to the LICENSE file for the full license text.
*/
function testShowFreeStruct() {
  //showFreeStruct('2024-09-11', '2024-09-16', 'SGP, SGG, SGB, S01, S02, S03, S04, S05, SMP, SM1, SM2, GALL, foyerSG, foyerSMU, foyerSM, foyerMe1, foyerMe2, bistro, loggia, lobbyQ4, S11, catering, foyerBar, ristorante, lbar');
  var start = Date.now();
  showFreeStruct('2025-08-01', '2025-08-31', '');
  var end = Date.now();
  Logger.log("Tempo di esecuzione funzione: " + (end - start) / 1000 + " secondi");
}

function togliVet(matrix, vet) {
  // Normalizzare i valori del vettore rimuovendo spazi bianchi
  const normalizedVet = vet.map(val => val.trim());

  return matrix.filter(row => !normalizedVet.includes(row[0]));
}

function stringToMatrix(str, cols) {
  // Converti la stringa in un array di testo
  //const arr = str.split(',').map(Number);
  const arr = str.split(',');

  // Crea una matrice vuota
  const matrix = [];

  // Calcola il numero di righe necessario
  const rows = Math.ceil(arr.length / cols);

  // Riempie la matrice
  for (let i = 0; i < rows; i++) {
    const row = arr.slice(i * cols, (i + 1) * cols);
    // perché la location è fatta così 1|2|3| ecc. --> 1,2,3
    row[4] = row[4].replace(/\|/g, ",");
    matrix.push(row);
  }

  return matrix;
}

function onlyUnique(value, index, self) {
  return self.indexOf(value) === index;
}

function showFreeStruct(first, last, selectedStruct) {
  try {
    firstDate = text2monthDays(first);
    lastDate = text2monthDays(last);
    var eventi = events2Array(convertDateBar(firstDate[1]), convertDateBar(lastDate[1]), categories()[0][0], "");
    if (eventi.length != 0) {
      var usedLocations = []
      for (let i = 0; i < eventi.length; i += 1) {
        if (!excludeAll()) {
          const isOptionated = optionated().indexOf(eventi[i][8].substring(0, 4)) < 0;
          if (eventi[i][10].length != 0 && (includeOptionated() || isOptionated)) {
            loc2array = eventi[i][10]; // .split(",");
            for (let j = 0; j < loc2array.length; j += 1) {
              usedLocations.push(String(loc2array[j])); // .map(s => s.trim()) prime 
              if (findKey(String(loc2array[j]), strutture(), 0) >= 0) {
                var relationship = strutture()[findKey(String(loc2array[j]), strutture(), 0)][10].split(","); // strutture con un grado di parentela da string a array
              }
              for (let k = 0; k < relationship.length; k += 1) {
                usedLocations.push(String(relationship[k]));
              }

            }

          }
        }
      }
      var usedUnique = usedLocations.filter(onlyUnique);
      var freeStructures = strutture();
      for (let i = 0; i < usedUnique.length; i += 1) {
        if (findKey(String(usedUnique[i]), freeStructures, 0) >= 0) {
          freeStructures.splice(findKey(String(usedUnique[i]), freeStructures, 0), 1);
        }
      }
      var structures = onlyStrcturesSelect(freeStructures);
      var structures = onlyStrcturesSelect(togliVet(strutture(), usedUnique));
    } else {
      var structures = onlyStrcturesSelect(strutture());
    }

    createUserSheet();
    structures.push(first);
    structures.push(last);
    structures.push(catering());
    structures.push(allestitore());
    structures.push(refCom());
    structures.push(refOp());
    structures.push(typeEv());
    var structSpecial = selectedStruct || '';
    structures.push(structSpecial);

    /*
    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '1B_addEventPageFinish', translate('addEventPage.addPageTitle')));
    */
    var htmlOutput = doGet(structures, '1B_addEventPageFinish', translate('addEventPage.addPageTitle'));

    // 2. Imposto le dimensioni desiderate per il dialogo.
    // I valori 800 e 600 sono presi dal tuo esempio.
    htmlOutput
      .setWidth(800)
      .setHeight(600);

    // 3. Estraggo il titolo dal risultato di doGet (che è il terzo parametro)
    // e lo uso come titolo della finestra di dialogo modale.
    var dialogTitle = translate('addEventPage.addPageTitle');

    // 4. Mostro il dialogo modale.
    SpreadsheetApp.getUi().showModelessDialog(htmlOutput, dialogTitle); //showModelessDialog oppure showModalDialog

  } catch (error) {
    SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
  }
}

function array2Events(array, calendarID) {
  var calendar = CalendarApp.getCalendarById(calendarID);

  if (!calendar) {
    //Logger.log('Calendar not found. Check your calendar ID.');
    return;
  }

  for (let i = 0; i < array.length; i++) {
    var title = array[i][0];
    var startTime = new Date(array[i][1]);
    var endTime = new Date(array[i][2]);
    var description = array[i][3];
    var author = parseEventDetails(description).refCom;
    var authorCongressi = groupName(refCom(), 2); // 2 is the group for Congress
    var location = array[i][4];
    var color = parseEventDetails(array[i][3]).color;
    var options = {
      description: description,
      location: location
    };

    event = calendar.createEvent(title, startTime, endTime, options);
    // ref: https://developers.google.com/apps-script/reference/calendar/event-color
    if (array[i][5] == 'E') {
      switch (color) {
        case 'viola':
          //var color = 3;
          event.setColor(CalendarApp.EventColor.MAUVE);
          break;
        case 'verdeChiaro':
          //var color = 3;
          event.setColor(CalendarApp.EventColor.PALE_GREEN);
          break;
        case 'rossoChiaro':
          //var color = 3;
          event.setColor(CalendarApp.EventColor.PALE_RED);
          break;
        case 'verde':
          //var color = 10;
          event.setColor(CalendarApp.EventColor.GREEN);
          break;
        case 'rosso':
          //var color = 10;
          event.setColor(CalendarApp.EventColor.RED);
          break;
        default:
        // code block
      }
    } else if (array[i][5] == 'L') {
      //var color = 11;
      event.setColor(CalendarApp.EventColor.RED);
    } else if (((array[i][5] == 'A') || (array[i][5] == 'D')) && (authorCongressi.includes(author))) {
      event.setColor(CalendarApp.EventColor.GRAY);
    }
    //Logger.log('Event created: ' + title);
  }
}

function getColorId(colorName) {
  var colors = {
    'Lavender': 1,
    'Sage': 2,
    'Grape': 3,
    'Flamingo': 4,
    'Banana': 5,
    'Tangerine': 6,
    'Peacock': 7,
    'Graphite': 8,
    'Blueberry': 9,
    'Basil': 10,
    'Tomato': 11
  };

  return colors[colorName] || null;
}

function tryAddlogRev() {
  addLogRevision('2025-07-04', 'Test', 'dfre', Session.getEffectiveUser().getEmail(), [])
}

function addLogRevision(oggi, tipoAggiunta, eventID, utenteEmail, array) {
  // Ottieni il foglio di lavoro attivo
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Ottieni il foglio di nome "registro"
  var sheet = spreadsheet.getSheetByName(sheetsList()[0][0]);
  //var sheetBackup = spreadsheetBackup.getSheetByName(sheetsList()[0][0]);

  // Se il foglio non esiste, crealo
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetsList()[0][0]);
    // Aggiungi intestazioni se il foglio è stato appena creato
    //sheet.appendRow(['Data', 'Azione', 'Event ID', 'Email Utente', 'Dettagli']);
    sheet.appendRow(translate('addEvent.appendLogRow').split(','));
  }
  /*
  if (!sheetBackup) {
    sheetBackup = spreadsheetBackup.insertSheet(sheetsList()[0][0]);
    // Aggiungi intestazioni se il foglio è stato appena creato
    sheetBackup.appendRow(['Data', 'Azione', 'Event ID', 'Email Utente', 'Dettagli']);
  }
  */

  // Prepara i dati da inserire
  var dataDaInserire = [
    oggi,				// Colonna 1: Data di oggi
    tipoAggiunta,			// Colonna 2: Tipo di aggiunta
    eventID,				// Colonna 3: ID evento
    getAliasEmail(utenteEmail),		// Colonna 4: Email dell'utente
    JSON.stringify(array) 		// Colonna 5: Array convertito in stringa
  ];

  // Aggiungi una nuova riga con i dati
  sheet.appendRow(dataDaInserire);
  var subject = translate('addEvent.logSubject') + convertDateBar(oggi) + '_' + convertHour(oggi) + ' ' + eventID;
  mandaEmail(oggi, getRealEmail(emailTarget()[0][0]), utenteEmail, emailTarget()[0][1], eventID, subject, tipoAggiunta, JSON.stringify(array));
}

function createEvents(first, last, array, what) {
  try {
    createUserSheet();
    updateTimeUser();
    if (checkUserWritePermission(myCalID()[0][0]) == true) {
      var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var matrix = stringToMatrix(array, 8)
      var mycalID = myCalID()[0][0];             // ID of the first calendar

      array2Events(matrix, mycalID);
      oggi = new Date();
      utenteEmail = Session.getEffectiveUser().getEmail();
      var eventID = (parseEventDetails(matrix[0][3]).id != '') ? parseEventDetails(matrix[0][3]).id + ' |-> ' + parseEventString(matrix[0][0]).nome : parseEventString(matrix[0][0]).nome;
      if (typeof what !== 'undefined') {
        if (what == 'hall') {
          //manageSmallRoom();
          createDailyScheduleFromCalendar(first, 30, '');
          addLogRevision(oggi, translate('addEvent.logNewOne'), eventID, utenteEmail, matrix);
return {
  success: true
};          
        }
      } else {
        //specialEvent();
        if (findKey('E', matrix, 5) >= 0) {
          const prefs = getUserBrowserSettings();
          showMonths(prefs.first, prefs.last, prefs.selectedStruct, prefs.keyword);
        } else {
          const prefs = getUserBrowserSettings();
          showMonths(prefs.first, prefs.last, prefs.selectedStruct, prefs.keyword);
        }
        addLogRevision(oggi, translate('addEvent.logNew'), eventID, utenteEmail, matrix);
return {
  success: true
};        
      }
    } else {
      SpreadsheetApp.getUi().alert(translate('modifyEvent.waitSomeTime'));
  return {
    success: false,
    reason: 'WAIT'
  };      
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
  return {
    success: false,
    reason: 'error'
  };    
  }
}