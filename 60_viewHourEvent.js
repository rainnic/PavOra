/*
* Project Name: Pavora
* Copyright (c) 2025 Nicola Rainiero
*
* This software is released under the MIT License.
* Please refer to the LICENSE file for the full license text.
*/
// Stringa di data e ora in formato locale
function convertUTCDate(localDateTime) {

  // Creazione di un oggetto Date a partire dalla stringa
  const localDate = new Date(localDateTime);

  // Estrazione dei componenti della data e ora in UTC
  const utcYear = localDate.getUTCFullYear();
  const utcMonth = String(localDate.getUTCMonth() + 1).padStart(2, '0'); // getUTCMonth() ritorna 0-11, quindi aggiungiamo 1
  const utcDay = String(localDate.getUTCDate()).padStart(2, '0');
  const utcHours = String(localDate.getUTCHours()).padStart(2, '0');
  const utcMinutes = String(localDate.getUTCMinutes()).padStart(2, '0');

  // Costruzione della stringa nel formato ISO 8601
  const utcDateTime = `${utcYear}-${utcMonth}-${utcDay}T${utcHours}:${utcMinutes}:00.000Z`;

  return utcDateTime
}

function testDailyEvents() {
  giorno = '2026-01-22';
  minuti = 60;
  createDailyScheduleFromCalendar(giorno, minuti, '', '', '');
}

function testFreeHallStruct() {
  changeTimeHall('2024-09-12T11:00:00.000Z', '2024-09-12T14:00:00.000Z', 'S5ySA9oz', '2024-09-12T13:00', '2024-09-12T17:00', 'S11,SM1,SM2,B,BPdx,Bdx')
}

function testFreeHall() {
  //start = '2025-06-21T10:00:00.000Z';
  //finish = '2025-06-21T23:00.00.000Z';
  start = '2025-06-21';
  finish = '2025-06-22';
  showFreeHall(start, finish);
}
function testFreeStructModifyEventDaily() {
  var first = '2025-06-22T07:00:00.000Z';
  var last = '2025-06-22T12:00:00.000Z';
  var finalList = [["Sweden and Martina ", "NO", "NO", [], ["SGP", "SGG", "SGB", "S05", "S11", "SMP", "SM1", "SM2", "GALL", "foyerSG", "foyerSM", "foyerMe1", "foyerMe2", "bistro", "loggia", "lobbyQ4", "catering", "foyerBar", "ristorante"], ["Cdx", "M", "L"], [], [], "immaginazione", "", "", "guzzonato", "750 | 2250 | 1480", "AC0124", "D", "2025-06-22", "09:00", "14:00", "2025-06-22T07:00:00.000Z", "2025-06-22T12:00:00.000Z", null, "levorato", "SWEDEN AND MARTINA", "cf", "", "cb", "25ISM", ""]];
  var note = 'AC0124';
  showFreeStructHallModifyEvent(first, last, finalList, note);
}

function createDailyScheduleFromCalendar(giorno, minuti, keyword, struttureScelte, mode) {
  mode = mode || ""; // oppure: mode = mode || "compact";
  resetFoglioConNuovo();

  // Ottenere gli eventi dal calendario
  var calendarId = myCalID()[0][0]; // Sostituire con l'ID del calendario desiderato
  var today = new Date(giorno);
  today.setHours(0, 0, 0, 0); // Imposta l'ora a mezzanotte per prendere tutto il giorno
  var tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1); // Data di domani
  var events = events2Array(formatDateMaster(today).giorno, formatDateMaster(today).giorno, categories()[0][0], keyword);

  // Create a new array with only one event with the same name to prevent extra work
  var filteredEvents = filterEvents(events);

  // Set the table
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);
  sheet.getRange(100, 100).setValue('');
  // Columns

  // Rows
  // Converte le date e gli orari in oggetti Date

  // Calcola il numero totale di ore
  var differenzaInMillisecondi = tomorrow.getTime() - today.getTime();
  var numeroColonne = differenzaInMillisecondi / (1000 * minuti * 60);

  // Metto le righe
  sheet.getRange(2, 1).setValue(translate('hourMenuPage.space')).setFontSize(14).setBorder(null, null, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  var sale = [];
  var j = 2;
  if ((struttureScelte == undefined) || (struttureScelte == '')) {
    for (let i = 0; i < strutture().length; i++) {
      if (strutture()[i][17] == 1) {
        j += 1;
        sheet.getRange(j, 1).setValue(strutture()[i][6]).setFontSize(12).setBorder(null, null, null, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
        sale.push(strutture()[i][0]);
      }
    }
  } else {
    sale = struttureScelte.split(",").map(s => s.trim());
    for (let i = 0; i < sale.length; i++) {
      j += 1;
      sheet.getRange(j, 1).setValue(strutture()[findKey(sale[i], strutture(), 0)][6]).setFontSize(12).setBorder(null, null, null, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
  }
  //sale.push('END', 'foyerSG', 'foyerSMU', 'foyerSM', 'foyerMe1','foyerMe2');
  //Logger.log("Le sale sono " + sale);

  // Scrive l'orario in ogni riga, incrementando di un'ora alla volta
  // Metto le colonne con le ore e trovo la corrispondenza con la sala
  var orari = [];
  var firstEventTiming = numeroColonne;
  var lastEventTiming = 0;
  var listaEventi = [];
  // Load the colors used to highlight the occupation of the structures
  var colori = methodMcolors();
  for (let i = 0; i <= numeroColonne; i++) {
    var dataCorrente = new Date(today.getTime() + i * minuti * 60 * 1000);
    sheet.getRange(2, i + 2).setValue(dataCorrente).setNumberFormat('HH:mm').setBorder(null, null, true, null, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
    for (let k = 0; k < events.length; k++) {
      if (events[k][10].length > 0) {
        for (let j = 0; j < events[k][10].length; j++) {
          if (events[k][8].substring(0, 4) === optionated()) {
            colore = colori[selectHigh(events[k][7], events[k][1])[0]][10];
          } else {
            colore = colori[selectHigh(events[k][7], 'D')[0]][selectHigh(events[k][7], 'D')[1]]; // events[k][1] --> categoria
          }
          if ((sale.indexOf(events[k][10][j]) > -1) && (events[k][0].getTime() <= dataCorrente.getTime()) && (dataCorrente.getTime() < events[k][9].getTime())) {
            if (events[k][1] == categories()[0][1]) {
              sheet.getRange(sale.indexOf(events[k][10][j]) + 3, i + 2).setValue(events[k][7] + 1).setNumberFormat('#');
              sheet.getRange(sale.indexOf(events[k][10][j]) + 3, i + 2).setBackground(colore).setHorizontalAlignment("center").setNote(events[k][8]);
            } else {
              sheet.getRange(sale.indexOf(events[k][10][j]) + 3, i + 2).setBackground(colore).setHorizontalAlignment("center").setNote(events[k][8]);
            }
            if ((i + 2) <= firstEventTiming) { firstEventTiming = i + 2 };
            if ((i + 2) >= lastEventTiming) { lastEventTiming = i + 2 };
            if (findKey(events[k][2], listaEventi, 0) < 0) {
              listaEventi.push([events[k][2], events[k][7]]);
            }
          }
        }
      }
    }
    orari.push(dataCorrente);
  }
  // Tolgo gli orari che non sono occupati da eventi
  if (mode !== 'H24') {
    if ((lastEventTiming != 0) && (lastEventTiming <= 24 * (60 / minuti))) {
      sheet.deleteColumns(lastEventTiming + 2, numeroColonne * (60 / minuti) - lastEventTiming + 2);
    }
    if ((lastEventTiming != 0) && (lastEventTiming > 24 * (60 / minuti))) {
      sheet.deleteColumns(24 * (60 / minuti) + 2, numeroColonne * (60 / minuti) - 24 * (60 / minuti) + 2);
    }
    if ((firstEventTiming != numeroColonne) && (firstEventTiming > 4 * (60 / minuti))) {
      sheet.deleteColumns(2, firstEventTiming - 2);
    }
  }

  // Legenda
  var lastRow = sheet.getLastRow() + 2;
  for (let i = 0; i < listaEventi.length; i++) {
    if (listaEventi[i][0].substring(0, 4) === optionated()) {
      colore = colori[selectHigh(listaEventi[i][1], categories()[2][1])[0]][10]; // categories()[0][1] --> 'E' categories()[4][1] --> 'P'
    } else {
      colore = colori[selectHigh(listaEventi[i][1], categories()[2][1])[0]][selectHigh(listaEventi[i][1], categories()[0][1])[1]];  // categories()[0][1] --> 'E'
    }
    var numeroEvento = listaEventi[i][1] + 1;
    sheet.getRange(lastRow + i, 1).setValue(numeroEvento + ') ' + listaEventi[i][0]).setBackground(colore).setHorizontalAlignment("left").setFontSize(14);
  }
  // Title and description
  var sheetTitle = translate('viewCalendar.mainCell'); // Table title
  var adesso = new Date();
  var adesso = formatDateMaster(adesso).giorno + ' ' + formatDateMaster(adesso).ora;
  var sottotitolo = translate('hourMenuPage.eventAt') + formatDateMaster(giorno).giorno + translate('hourMenuPage.updateAt', { adesso: adesso });
  sheet.getRange(1, 1).setValue(sheetTitle).setNumberFormat('0').setHorizontalAlignment("left").setFontSize(16).setNote(translate('hourMenuPage.eventsScheduled') + filteredEvents.map(row => row[2]).join('\n'));
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  sheet.autoResizeRows(1, sheet.getLastRow());
  sheet.getRange(1, sheet.getLastColumn()).setValue(sottotitolo).setHorizontalAlignment("right").setFontSize(10);
  // remove extra rows and columns
  var totalRows = sheet.getLastRow();
  var lc = sheet.getLastColumn();
  var mc = sheet.getMaxColumns();
  if (mc - lc != 0) {

    sheet.deleteColumns(lc + 1, mc - lc);
  } else {
    sheet.deleteColumns(7, 1);
  }
  var lr = sheet.getLastRow();
  var mr = sheet.getMaxRows();
  if (lr - 2 != 0) {
    sheet.deleteRows(lr + 1, mr - lr);
  }
  // fix +20 the size of first column
  sheet.setColumnWidth(1, sheet.getColumnWidth(1) + 20);
}

// --------------------------------------------
// Codice per inserire e modificare gli eventi
// --------------------------------------------
function showFreeHall(first, last) {
  try {
    createUserSheet();
    ClearAll();
    //updateTimeUser(); // lo metto alla fine
    var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    //Logger.log('first è ' + typeof (first));
    /*
    sh.getRange(1, 1).setValue('Data iniziale');
    sh.getRange(1, 2).setValue(first).setNumberFormat("@");
    sh.getRange(1, 10).setValue(typeof (first)).setNumberFormat("@");
    sh.getRange(2, 1).setValue('Data finale');
    sh.getRange(2, 2).setValue(last).setNumberFormat("@");
    sh.getRange(2, 10).setValue(typeof (last)).setNumberFormat("@");
    sh.getRange(2, 15).setValue(new Date(last));
    */
    //Logger.log(first + ' ' + typeof (first));
    var lastDate = new Date(last);
    //sh.getRange(4, 1).setValue(lastDate);

    var eventi = events2Array(first, last, categories()[0][0], "");
    //Logger.log(eventi);
    if (eventi.length != 0) {
      //Logger.log(eventi[0][10]);
      var usedLocations = []
      for (let i = 0; i < eventi.length; i += 1) {
        if (!excludeAll()) {
          const isOptionated = optionated().indexOf(eventi[i][8].substring(0, 4)) < 0;
          if (eventi[i][10].length != 0 && (includeOptionated() || isOptionated)) {
            loc2array = eventi[i][10]; // .split(",");
            for (let j = 0; j < loc2array.length; j += 1) {
              usedLocations.push(String(loc2array[j]));
              if (findKey(String(loc2array[j]), strutture(), 0) >= 0) {
                var relationship = strutture()[findKey(String(loc2array[j]), strutture(), 0)][10].split(","); // strutture con un grado di parentela da string a array
                //Logger.log(loc2array[j] + '------------->' + relationship);
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
        //Logger.log('string(usedUnique[i]) è ' + String(usedUnique[i]) + ' | ' + 'findKey(string(usedUnique[i]), freStrutture(), 0) è' + findKey(String(usedUnique[i]), freeStructures, 0));
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
    // Add start date and finish date at the end of structures
    structures.push(first);
    structures.push(last);
    structures.push(catering());
    structures.push(allestitore());
    structures.push(refCom());
    structures.push(refOp());
    structures.push(typeEv());

    var giorno = formatDateMaster(new Date(first)).dataXweb;
    createDailyScheduleFromCalendar(giorno, 30, '');
    if (eventi.length != 0) {
    } else {
    }
    updateTimeUser();
    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '6C2_addEditMSRPage', translate('menu.modifyEvent')));
  } catch (error) {
    SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
  }
}

function showFreeHallEdit(first) {
  try {
    resetFoglioConNuovo();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    firstDate = text2monthDays(first);
    var today = new Date(first);
    today.setHours(0, 0, 0, 0); // Imposta l'ora a mezzanotte per prendere tutto il giorno
    var tomorrow = new Date(today);
    tomorrow.setDate(today.getDate() + 1); // Data di domani
    var eventi = events2Array(first, tomorrow, categories()[0][0], "");
    if (eventi.length != 0) {
      //Logger.log(eventi[0][10]);
      var usedLocations = [];
      var allStructures = strutture();
      for (let i = 0; i < eventi.length; i += 1) {
        if (!excludeAll()) {
          const isOptionated = optionated().indexOf(eventi[i][8].substring(0, 4)) < 0;
          if (eventi[i][10].length != 0 && (includeOptionated() || isOptionated)) {
            loc2array = eventi[i][10]; // .split(",");
            for (let j = 0; j < loc2array.length; j += 1) {
              usedLocations.push(loc2array[j]);
              if (findKey(String(loc2array[j]), strutture(), 0) >= 0) {
                var relationship = allStructures[findKey(loc2array[j], allStructures, 0)][10].split(","); // strutture con un grado di parentela da string a array
              }
              for (let k = 0; k < relationship.length; k += 1) {
                usedLocations.push(relationship[k]);
              }

            }

          }
        }
      }
      //Logger.log('Le location presenti sono queste ' + usedLocations + ' Quelle uniche invece sono queste: ' + usedLocations.filter(onlyUnique));
      //sh.getRange(3, 1).setValue('Strutture in uso:');
      //sh.getRange(4, 1).setValue(usedLocations.filter(onlyUnique).toString()).setNumberFormat("@");
      var usedUnique = usedLocations.filter(onlyUnique);
      var freeStructures = strutture();
      for (let i = 0; i < usedUnique.length; i += 1) {
        if (findKey(usedUnique[i], freeStructures, 0) >= 0) {
          freeStructures.splice(findKey(usedUnique[i], freeStructures, 0), 1);
        }
      }
      var structures = onlyStrcturesSelect(freeStructures);
      var structures = onlyStrcturesSelect(togliVet(strutture(), usedUnique));
    } else {
      var freeStructures = strutture();
      var structures = onlyStrcturesSelect(freeStructures);
    }

    createUserSheet();
    structures.push(first);
    structures.push(first);

    var giorno = formatDateMaster(new Date(first)).dataXweb;
    createDailyScheduleFromCalendar(giorno, 30, '');
    if (eventi.length != 0) {
      sh.getRange(2, 1).setValue(translate('modifyEvent.usedStruct')).setFontSize(12);
      sh.getRange(2, 1).setNote(usedLocations.filter(onlyUnique).toString()).setNumberFormat("@").setFontSize(10).setWrap(true);
    } else {
      sh.getRange(2, 1).setNote(translate('modifyEvent.noUsedStruct')).setFontSize(10).setWrap(true);
    }

    updateTimeUser();
    // Temporaneamente commentato
    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '6D1_editAskMSRPage', translate('modifyEvent.editDelEvent')));
    //ClearAll();

  } catch (error) {
    SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
  }
}

function getCellHallNote(first, what) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var cell = sh.getActiveCell(); // from the google sheet
    var noteComplete = cell.getNote();
    if (noteComplete.length != 0) {
      var note = ((extractRegex(regexId, noteComplete) != 0) ? extractRegex(regexId, noteComplete) : noteComplete.split("  ")[0]);
    }
    if (note) {
      sh.getRange(3, 1).setValue(note).setFontSize(10);
      var finalList = logMatchingEvents(myCalID()[0][0], note, first, first, ' ', what);
      first = finalList[0][18].toISOString();
      last = finalList[finalList.length - 1][19].toISOString();
      showFreeStructHallModifyEvent(first, last, finalList, note);
    }
    return note || translate('modifyEvent.emptyCell')
  } catch (error) {
    SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
    return translate('modifyEvent.noEventDate')
  }
}

function changeTimeHall(firstOriginal, lastOriginal, eventNameId, firstNew, lastNew, locationOriginal) {
  var finalList = logMatchingEvents(myCalID()[0][0], eventNameId, firstOriginal, lastOriginal, locationOriginal);
  first = convertDateInputHtml(finalList[0][18]);
  first = firstNew;
  last = lastNew;
  showFreeStructHallModifyEvent(first, last, finalList, eventNameId, locationOriginal);
}

// Function to showfreeStrruct and event to modify or delete
function showFreeStructHallModifyEvent(first, last, array, eventNameId, locationOriginal) {
  try {
    createUserSheet();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    firstDate = text2monthDays(first);
    lastDate = text2monthDays(last);
    first = convertUTCDate(first);
    last = convertUTCDate(last);
    if (eventNameId != undefined) {
      var keyword = '-' + eventNameId;
    } else {
      var keyword = '';
    }
    var eventi = events2Array(first, last, categories()[0][0], keyword); // formatDateMaster(today).giorno
    if (eventi.length != 0) {
      var usedLocations = []
      for (let i = 0; i < eventi.length; i += 1) {
        if (!excludeAll()) {
          const isOptionated = optionated().indexOf(eventi[i][8].substring(0, 4)) < 0;
          if (eventi[i][10].length != 0 && (includeOptionated() || isOptionated)) {
            loc2array = eventi[i][10]; // .split(",");
            for (let j = 0; j < loc2array.length; j += 1) {
              usedLocations.push(loc2array[j]);
              if (findKey(String(loc2array[j]), strutture(), 0) >= 0) {
                var relationship = strutture()[findKey(loc2array[j], strutture(), 0)][10].split(","); // strutture con un grado di parentela da string a array
              }
              for (let k = 0; k < relationship.length; k += 1) {
                usedLocations.push(relationship[k]);
              }

            }

          }
        }
      }
      var usedUnique = usedLocations.filter(onlyUnique);
      var usedUnique = usedUnique.filter(element => !array[0][2].includes(element)); // con array[0][20] dava errore!
      var freeStructures = strutture();
      for (let i = 0; i < usedUnique.length; i += 1) {
        if (findKey(usedUnique[i], freeStructures, 0) >= 0) {
          freeStructures.splice(findKey(usedUnique[i], freeStructures, 0), 1);
        }
      }
      var structures = onlyStrcturesSelect(freeStructures);
    } else {
      var freeStructures = strutture();
      var structures = onlyStrcturesSelect(freeStructures);
    }

    createUserSheet();
    structures.unshift(array);
    structures.unshift(last);
    structures.unshift(first);
    structures.unshift(catering());
    structures.unshift(allestitore());
    structures.unshift(refCom());
    structures.unshift(refOp());
    structures.unshift(typeEv());

    createDailyScheduleFromCalendar(first, 60, '');
    if (eventi.length != 0) {
      sh.getRange(2, 1).setValue(translate('modifyEvent.usedStruct')).setFontSize(12);
      sh.getRange(2, 1).setNote(usedUnique.toString()).setNumberFormat("@").setFontSize(10).setWrap(true);
    } else {
      sh.getRange(2, 1).setNote(translate('modifyEvent.noUsedStruct')).setFontSize(10).setWrap(true);
    }

    updateTimeUser();
    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '6D2_editAddMSRPage', translate('sidebar.editEvent')));

  } catch (error) {
    SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
  }
}