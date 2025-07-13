// New method for showing months in fast mode
function testShowNewMonths() {
  //showMonths('2024-09-16', '2024-09-20', 'SGP,SGG,SGB,S01,S02,S03,S04,S05,S11,SMP,SM1,SM2,GALL,foyerSG,foyerSMU,foyerSM,foyerMe1,foyerMe2,bistro,loggia,lobbyQ4,ufficiQ8,catering,foyerBar,ristorante,lbar', '', '1');
  var start = Date.now();
  showMonths('2025-01-01', '2025-12-31', '', '', '');
  var end = Date.now();
  Logger.log("Tempo di esecuzione funzione: " + (end - start) / 1000 + " secondi");
}

// --------------------------------------------------------------------------------------
// SPECIAL DAILY EVENT
// --------------------------------------------------------------------------------------
function showFreeDailyHall(first, last, selectionData) {
  try {
    //createUserSheet();
    //ClearAll();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastDate = new Date(last);

    var eventi = events2Array(first, last, categories()[0][0], "");
    var structures;

    if (eventi.length != 0) {
      var usedLocations = [];
      for (let i = 0; i < eventi.length; i++) {
        if (!excludeAll()) {
          const isOptionated = optionated().indexOf(eventi[i][8].substring(0, 4)) < 0;
          if (eventi[i][10].length != 0 && (includeOptionated() || isOptionated)) {
            //if ((eventi[i][10].length != 0) && (optionated().indexOf(eventi[i][8].substring(0, 4)) < 0)) {
            let loc2array = eventi[i][10];
            for (let j = 0; j < loc2array.length; j++) {
              usedLocations.push(String(loc2array[j]));
              if (findKey(String(loc2array[j]), strutture(), 0) >= 0) {
                let relationship = strutture()[findKey(String(loc2array[j]), strutture(), 0)][10].split(",");
                for (let k = 0; k < relationship.length; k++) {
                  usedLocations.push(String(relationship[k]));
                }
              }
            }
          }
        }
      }

      let usedUnique = usedLocations.filter(onlyUnique);
      let freeStructures = togliVet(strutture(), usedUnique);

      // Se è presente selectionData, filtra ulteriormente
      if (selectionData && Array.isArray(selectionData) && selectionData.length > 0) {
        freeStructures = freeStructures.filter(s => selectionData.includes(s[0])); // s[0] = nome struttura
      }

      structures = onlyStrcturesSelect(freeStructures);
    } else {
      structures = onlyStrcturesSelect(strutture());
      if (selectionData && Array.isArray(selectionData) && selectionData.length > 0) {
        structures = structures.filter(s => selectionData.includes(s[0]));
      }
    }

    // Aggiunte informazioni extra
    structures.push(first, last, catering(), allestitore(), refCom(), refOp(), typeEv());
    var structSpecial = selectionData || '';
    structures.push(structSpecial);

    // Crea pianificazione giornaliera
    let giorno = formatDateMaster(new Date(first)).dataXweb;
    //createDailyScheduleFromCalendar(giorno, 30, '');

    updateTimeUser();

    SpreadsheetApp.getUi().showSidebar(doGet(structures, '6C2_addEditMSRPage', translate('menu.modifyEvent')));

  } catch (error) {
    SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
  }
}

function specialNewDailyEvent() {
  const selectionResult = getSelectionDailyData();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (selectionResult) {
    updateTimeUser();
    /*
    Logger.log(selectionResult.firstDate);
    Logger.log(selectionResult.lastDate);
    Logger.log(selectionResult.firstColumnData);
    */
    showFreeDailyHall(selectionResult.firstDate, selectionResult.lastDate, selectionResult.firstColumnData);

    const range = sheet.getRange(selectionResult.range);
    range.setValue("N");
    range.setBackground("#d5a6bd");
    range.setFontColor("#000000");
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
    range.setNote(translate('specialEvent.newEvent'));
  }
}

function getSelectionDailyData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();

  if (!range) {
    Browser.msgBox(translate('specialEvent.selectOneFirst'));
    return null;
  }

  const startRow = range.getRow();
  const endRow = range.getLastRow();
  const startCol = range.getColumn();
  const endCol = range.getLastColumn();

  const dateStart = sheet.getRange(2, startCol).getValue();
  const dateEnd = sheet.getRange(2, endCol).getValue();

  if (!dateStart || !dateEnd || isNaN(new Date(dateStart)) || isNaN(new Date(dateEnd))) {
    Browser.msgBox(translate('specialEvent.selectCellRowSec'));
    return null;
  }

  const firstColumnValues = sheet.getRange(startRow, 1, endRow - startRow + 1, 1).getValues().flat();
  const convertedValues = [];

  for (let i = 0; i < firstColumnValues.length; i++) {
    if (findKey(firstColumnValues[i], strutture(), 6) > -1) {
      convertedValues.push(strutture()[findKey(firstColumnValues[i], strutture(), 6)][0]);
    }
  }

  const startCell = sheet.getRange(startRow, startCol).getA1Notation();
  const endCell = sheet.getRange(endRow, endCol).getA1Notation();
  const selectionRange = startCell + ":" + endCell;

  return {
    firstDate: dateStart.toISOString(),
    lastDate: dateEnd.toISOString(),
    firstColumnData: convertedValues.join(","),
    range: selectionRange
  };
}



function specialDeleteDailyEvent(first) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveCell();
  var note = range.getNote(); // Prende la nota della cella selezionata

  if (!note) {
    SpreadsheetApp.getUi().alert(translate('specialEvent.noNoteCell'));
    return;
  }

  var regexId = /id=([^\s">]+)/;
  var match = note.match(regexId);
  if (!match) {
    SpreadsheetApp.getUi().alert(translate('specialEvent.noIdFound'));
    return;
  }

  var id = match[1]; // Estratto ID dalla nota
  var what = "hall"; // Se serve, altrimenti può essere lasciato vuoto
  //var activeRow = range.getRow(); // Prende la riga attuale

  // Chiamata alla funzione deleteEvents con i parametri richiesti
  Logger.log('First è ' + first);
  deleteEvents(id, first, first, what);

  // Evidenzia tutte le celle con lo stesso ID nella nota
  //highlightCellsWithId(id);
}

function specialUpdateDailyEvent(first) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveCell();
  var note = range.getNote(); // Prende la nota della cella selezionata

  if (!note) {
    SpreadsheetApp.getUi().alert(translate('specialEvent.noNoteCell'));
    return;
  }

  var regexId = /id=([^\s">]+)/;
  var match = note.match(regexId);
  if (!match) {
    SpreadsheetApp.getUi().alert(translate('specialEvent.noIdFound'));
    return;
  }

  var id = match[1]; // Estratto ID dalla nota
  var what = ""; // Se serve, altrimenti può essere lasciato vuoto
  var activeRow = range.getRow(); // Prende la riga attuale

  // Chiamata alla funzione getCellNote con i parametri richiesti
  getCellHallNote(first, 'hall');
  //getCellNote(first, last, id);

  // Evidenzia tutte le celle con lo stesso ID nella nota
  highlightCellsWithIdM(id);
}

// --------------------------------------------------------------------------------------
// SPECIAL EVENT
// --------------------------------------------------------------------------------------
function specialNewEvent() {
  const selectionResult = getSelectionData();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (selectionResult) {
    //Logger.log(selectionResult);
    updateTimeUser();
    showFreeStruct(selectionResult.firstDate, selectionResult.lastDate, selectionResult.firstColumnData)
    //showFreeStruct(selectionResult.firstDate, selectionResult.lastDate, '')

    var range = sheet.getRange(selectionResult.range); // Ottieni il range in formato A1:B2
    range.setValue("N"); // Scrive "N" in tutte le celle
    range.setBackground("	#d5a6bd"); // Magenta chiaro due (Hex)
    range.setFontColor("#000000"); // Testo nero per contrasto
    range.setHorizontalAlignment("center"); // Centra il testo
    range.setVerticalAlignment("middle");
    //range.setFontWeight("bold"); // Grassetto
    range.setNote(translate('specialEvent.newEvent')); // Aggiunge la nota

    //SpreadsheetApp.getUi().alert("Il range " + selectionData.range + " è stato evidenziato.");


  }
}

function getSelectionData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange(); // Ottieni la selezione attiva

  if (!range) {
    Browser.msgBox(translate('specialEvent.selectOneFirst'));
    return null;
  }

  var startRow = range.getRow();  // Prima riga selezionata
  var endRow = range.getLastRow(); // Ultima riga selezionata
  var startCol = range.getColumn(); // Prima colonna selezionata
  var endCol = range.getLastColumn(); // Ultima colonna selezionata

  // Controllo: la selezione deve includere almeno una colonna valida
  if (startCol > (sheet.getLastColumn() - 1)) {
    Browser.msgBox(translate('specialEvent.selectValidColumn'));
    return null;
  }

  // Estrai la prima e l'ultima data dalla quinta riga della selezione
  var dateStart = sheet.getRange(5, startCol).getValue();
  var dateEnd = sheet.getRange(5, endCol).getValue();

  if (!dateStart || !dateEnd || isNaN(new Date(dateStart)) || isNaN(new Date(dateEnd))) {
    Browser.msgBox(translate('specialEvent.selectCellRowfive'));
    return null;
  }

  // Estrarre il contenuto della prima colonna per ogni riga selezionata
  var firstColumnValues = sheet.getRange(startRow, 1, endRow - startRow + 1, 1).getValues().flat();
  var convertedValues = [];
  for (var i = 0; i < firstColumnValues.length; i++) {
    //Logger.log(firstColumnValues[i]);
    if (findKey(firstColumnValues[i], strutture(), 6) > -1) {
      convertedValues.push(strutture()[findKey(firstColumnValues[i], strutture(), 6)][0])
    }
  }

  // Creazione del range in formato A1:B2
  var startCell = sheet.getRange(startRow, startCol).getA1Notation();
  var endCell = sheet.getRange(endRow, endCol).getA1Notation();
  var selectionRange = startCell + ":" + endCell;

  // Output finale
  var result = {
    firstDate: convertDateInputHtml(dateStart),
    lastDate: convertDateInputHtml(dateEnd),
    //firstColumnData: firstColumnValues
    firstColumnData: convertedValues.join(","),
    range: selectionRange // Aggiunto il range selezionato in formato A1:B2    
  };

  Logger.log(result); // Per debug
  //showFreeStruct(convertDateInputHtml(dateStart), convertDateInputHtml(dateEnd), convertedValues)
  return result
}



function specialDeleteEvent(first, last) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveCell();
  var note = range.getNote(); // Prende la nota della cella selezionata

  if (!note) {
    SpreadsheetApp.getUi().alert(translate('specialEvent.noNoteCell'));
    return;
  }

  var regexId = /id=([^\s">]+)/;
  var match = note.match(regexId);
  if (!match) {
    SpreadsheetApp.getUi().alert(translate('specialEvent.noIdFound'));
    return;
  }

  var id = match[1]; // Estratto ID dalla nota
  //var what = ""; // Se serve, altrimenti può essere lasciato vuoto
  //var activeRow = range.getRow(); // Prende la riga attuale

  // Chiamata alla funzione deleteEvents con i parametri richiesti
  deleteEvents(id, first, last);

  // Evidenzia tutte le celle con lo stesso ID nella nota
  highlightCellsWithId(id);
}

// Funzione per evidenziare tutte le celle contenenti lo stesso ID nella nota
function highlightCellsWithId(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange(); // Prende tutto il range dei dati
  var values = range.getNotes(); // Ottiene tutte le note

  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (values[r][c] && values[r][c].includes("id=" + id)) {
        sheet.getRange(r + 1, c + 1)
          .setBackground("#d5a6bd") // Colore magenta chiaro 2
          .setHorizontalAlignment("center") // Centra il testo
          .setVerticalAlignment("middle")
          //.setFontWeight("bold") // Grassetto
          .setValue("C"); // Inserisce la lettera "C"
      }
    }
  }
}

function specialUpdateEvent(first, last) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveCell();
  var note = range.getNote(); // Prende la nota della cella selezionata

  if (!note) {
    SpreadsheetApp.getUi().alert(translate('specialEvent.noNoteCell'));
    return;
  }

  var regexId = /id=([^\s">]+)/;
  var match = note.match(regexId);
  if (!match) {
    SpreadsheetApp.getUi().alert(translate('specialEvent.noIdFound'));
    return;
  }

  var id = match[1]; // Estratto ID dalla nota
  var what = ""; // Se serve, altrimenti può essere lasciato vuoto
  var activeRow = range.getRow(); // Prende la riga attuale

  // Chiamata alla funzione getCellNote con i parametri richiesti
  getCellNote(first, last, id);

  // Evidenzia tutte le celle con lo stesso ID nella nota
  highlightCellsWithIdM(id);
}

// Funzione per evidenziare tutte le celle contenenti lo stesso ID nella nota
function highlightCellsWithIdM(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange(); // Prende tutto il range dei dati
  var values = range.getNotes(); // Ottiene tutte le note

  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (values[r][c] && values[r][c].includes("id=" + id)) {
        sheet.getRange(r + 1, c + 1)
          .setBackground("#d5a6bd") // Colore magenta chiaro 2
          .setHorizontalAlignment("center") // Centra il testo
          .setVerticalAlignment("middle")
          //.setFontWeight("bold") // Grassetto
          .setValue("M"); // Inserisce la lettera "C"
      }
    }
  }
}


// what = "newEvent" | updateDetailsEvent | deleteEvent;
function getCellSpecialNote(what, first, last) {
  //try {
  //Logger.log(what + ' | ' + first + ' | '+last);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  //createUserSheet();
  updateTimeUser();
  var cell = sh.getActiveCell(); // from the google sheet
  var activeRow = sh.getActiveRange().getRow();
  /*
  var noteComplete = cell.getNote();
  var name = cell.getValue();
  var lastColumn = sh.getLastColumn();
  var activeRow = sh.getActiveRange().getRow();
  */
  // Ottieni la riga attualmente selezionata
  //const selectedRow = range.getRow();

  // Ottieni la nota dalla seconda colonna (colonna B)
  const nameNote = cell.getNote();
  ui.alert('L\'evento è: ' + nameNote);

  if (nameNote.length != 0) {
    var note = ((extractRegex(regexId, nameNote) != 0) ? extractRegex(regexId, nameNote) : parseEventString(extractRegex(regexName, nameNote)).nome);
    var nameEvent = parseEventString(extractRegex(regexName, nameNote)).nome;
    var nameFinal = nameEvent + ' (' + note + ')';
    //ui.alert('L\'evento è: ' + nameFinal);
    var today = new Date();
    //Logger.log(today.getFullYear());
    var startDate = new Date(today.getFullYear() - 2, 0, 1); // 1 gennaio di 1 anno fa
    //Logger.log(startDate);

    // Inizializza una data finale molto futura
    var endDate = new Date(today.getFullYear() + 6, 0, 1); // 1 gennaio tra 6 anni
    var first = formatDateMaster(startDate).dataXweb;
    var last = formatDateMaster(endDate).dataXweb;
    var startList = logMatchingEvents(myCalID()[0][0], note, first, last, what);
    //ui.alert(startList[0]);
    var finalList = [];
    for (let j = 0; j < startList.length; j++) {
      //var idEvent = (startList[j][13].length == 0) ? randomID(8) : startList[j][13];
      finalList.push(startList[j][16], startList[j][17]);
    }
  }
  //var titoli = translate('viewList.titleChanges').split(',');
  //ui.alert('Pronto per verificare le variabili: '+note + ' | '+ what + ' | '+ activeRow);
  if (note && (what === 'deleteEvent')) {
    first = convertDateInputHtml(finalList[0][0]);
    last = convertDateInputHtml(finalList[finalList.length - 1][1]);
    ui.alert('Pronto a cancellare con questi dati: ' + note + ' | ' + first + ' | ' + last + ' | ' + what + ' | ' + activeRow);
    deleteEvents(note, first, last, what, activeRow);
  } else if (note && (what === 'updateSpecificEvent')) {
    if (note != '') {
      var response = ui.alert(translate('viewList.okEditSpecific'), translate('viewList.yesEditSpecific', { name: note }), ui.ButtonSet.YES_NO);
      if (response == ui.Button.YES) {
        first = convertDateInputHtml(finalList[0][18]);
        last = convertDateInputHtml(finalList[finalList.length - 1][19]);
        var listaFinale = logMatchingEvents(myCalID()[0][0], note, first, last);
        showFreeStructModifyEvent(first, last, listaFinale, note);
      }
    } else { ui.alert(translate('viewList.alertOldEdit', { name: note })); }

  } else {
    ui.alert(translate('viewList.alertEdit'));
  }
  return note || translate('modifyEvent.emptyCell')
  /*
} catch (error) {
  SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
  return translate('modifyEvent.noEventDate')
}
*/
}

// -------------------------------- BEGIN NEW FAST SHOWMONTH ---------------------------------
// New function showNewMonths(first, last, structures, keyword, period);
function showMonths(first, last, structures, keyword, period) {
  try {
    //createUserSheet();
    resetFoglioConNuovo();
    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Se il periodo non è definito, regola automaticamente il range
    if (period === undefined) {
      first = incrDay(first, -incrementDay());
      last = incrDay(last, incrementDay());
    }

    // Converte le date in intervalli mensili
    const firstDate = text2monthDays(first);
    const lastDate = text2monthDays(last);

    // Se structures non è definito o è una stringa vuota, usa le strutture di default
    if (!structures) {
      structures = struttureBigKey(readVariables('struttureScelteSup', DataStructures)).join(',');
    }

    const eventi = events2Array(convertDateBar(firstDate[0]), convertDateBar(lastDate[2]), categories()[0][0], keyword);

    var matrici = generaMatriceEventiConColori(eventi, formatDateMaster(firstDate[0]).dataXweb, formatDateMaster(lastDate[2]).dataXweb, structures);
    popolaFoglioSheet(matrici.matriceOutput, matrici.matriceNoteOutput, matrici.matriceColorOutput, matrici.matriceDomenicheOutput);
    checkToday();

    // Nasconde le colonne fuori dal range
    const lc = sh.getLastColumn();
    const rangeColStart = datediff(firstDate[0], firstDate[1]);
    const colFinishStart = lc + 1 - datediff(lastDate[1], lastDate[2]);
    const rangeColFinish = datediff(lastDate[1], lastDate[2]);

    if (rangeColStart > 0) sh.hideColumns(2, rangeColStart);
    if (rangeColFinish > 0) sh.hideColumns(colFinishStart, rangeColFinish);
  } catch (error) {
    const errorMessage = translate('alert.errorMessage') + ' (' + error.message + ')';
    SpreadsheetApp.getUi().alert(errorMessage);
  }
}

function resetFoglioConNuovo() {
  createUserSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const foglioVecchio = ss.getActiveSheet();
  const nome = foglioVecchio.getName();
  const pos = ss.getSheetByName(nome).getIndex();
  ss.deleteSheet(foglioVecchio);

  // Crea nuovo foglio con lo stesso nome nella stessa posizione
  const foglioNuovo = ss.insertSheet(nome, pos - 1);

  // Applica impostazioni di base (facoltative)
  foglioNuovo.setHiddenGridlines(false);
  foglioNuovo.setColumnWidths(1, 20, 100);
  foglioNuovo.setRowHeights(1, 100, 30);
  foglioNuovo.getRange("A1").setFontSize(10).setFontWeight("bold");

  // Blocca intestazioni
  foglioNuovo.setFrozenRows(6);
  foglioNuovo.setFrozenColumns(1);

  // (Opzionale) segnaposto
  const today = new Date();
  foglioNuovo.getRange(1, 1).setValue(translate('specialEvent.newSheet') + '\n' + formatDateMaster(today).ora + ' ' + formatDateMaster(today).giorno);
}

function generaMatriceEventiConColori(eventi, intervalloInizio, intervalloFine, luoghi) {
  // Inizializzazioni efficienti con cache
  var inizio = new Date(intervalloInizio);
  var fine = new Date(intervalloFine);
  var luoghiArray = luoghi.split(",").map(s => s.trim());

  //const nomiMesi = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
  const nomiMesi = translate('viewCalendar.months').split(', ');
  //const lettereGiorniSettimana = ['D', 'L', 'M', 'M', 'G', 'V', 'S'];
  const lettereGiorniSettimana = translate('viewCalendar.days').split(', '); // Converte in array

  // Pre-calcola la mappa dei luoghi una sola volta
  const mappaLuoghi = new Map();
  const struttureCache = strutture();

  for (const struttura of struttureCache) {
    const genitore = struttura[0];
    const figli = struttura[18] ? struttura[18].split(",").map(s => s.trim()) : [];
    mappaLuoghi.set(genitore, figli);
  }

  // Funzione ottimizzata per la ricerca dei figli con memoization
  const memoFigli = new Map();
  function getTuttiFigli(luogo) {
    if (memoFigli.has(luogo)) return memoFigli.get(luogo);

    const visitati = new Set();
    const result = new Set();

    function ricorsivo(nodo) {
      if (visitati.has(nodo)) return;
      visitati.add(nodo);

      const figliDiretti = mappaLuoghi.get(nodo) || [];
      figliDiretti.forEach(figlio => {
        result.add(figlio);
        ricorsivo(figlio);
      });
    }

    ricorsivo(luogo);
    const figli = Array.from(result);
    memoFigli.set(luogo, figli);
    return figli;
  }

  // Ottimizza la mappa luogo con figli usando Map
  const mappaLuogoConFigli = new Map();
  for (const luogo of luoghiArray) {
    const figli = getTuttiFigli(luogo);
    mappaLuogoConFigli.set(luogo, new Set([luogo, ...figli]));
  }

  // Calcola una volta sola la lista di date
  const millisecondsPerDay = 86400000;
  const dateList = [];
  const datesToIndex = new Map();
  let dayCount = 0;

  for (let t = inizio.getTime(); t <= fine.getTime(); t += millisecondsPerDay) {
    const giorno = new Date(t);
    const giornoStr = giorno.toISOString().split('T')[0];
    dateList.push(giorno);
    datesToIndex.set(giornoStr, dayCount++);
  }

  // Prepopola la mappa eventi per tutti i luoghi e date - struttura ottimizzata
  const mappaEventi = new Map();
  for (const luogo of luoghiArray) {
    const luogoMap = new Map();
    mappaEventi.set(luogo, luogoMap);

    for (const giorno of dateList) {
      const giornoStr = giorno.toISOString().split('T')[0];
      luogoMap.set(giornoStr, {
        eventi: [],
        numero: '',
        note: '',
        numeriUnici: new Set()
      });
    }
  }

  // Elabora gli eventi una sola volta per luogo
  for (const evento of eventi) {
    const dataInizio = new Date(evento[0]);
    const dataFine = new Date(evento[9]);
    const nomeEvento = evento[2];
    const numeroEvento = evento[7] + 1;
    const descrizione = evento[8];
    const luoghiEvento = evento[10].map(s => s.trim());

    // Per ogni luogo richiesto, verifica se è coinvolto
    for (const luogoRichiesto of luoghiArray) {
      const figliDelLuogo = mappaLuogoConFigli.get(luogoRichiesto);
      const coinvolto = luoghiEvento.some(l => figliDelLuogo.has(l));

      if (coinvolto) {
        // Usa una singola iterazione sulle date dell'evento
        for (let t = dataInizio.getTime(); t <= dataFine.getTime(); t += millisecondsPerDay) {
          const giorno = new Date(t);
          const giornoStr = giorno.toISOString().split('T')[0];

          // Verifica se la data è nel range richiesto
          if (t >= inizio.getTime() && t <= fine.getTime()) {
            const luogoMap = mappaEventi.get(luogoRichiesto);
            const datoGiorno = luogoMap.get(giornoStr);

            if (datoGiorno) {
              datoGiorno.eventi.push(nomeEvento);
              datoGiorno.numeriUnici.add(numeroEvento);
              datoGiorno.numero = Array.from(datoGiorno.numeriUnici).join('\n'); // | or \n
              datoGiorno.note += '[' + numeroEvento + '] ' + descrizione + '\n\n';
            }
          }
        }
      }
    }
  }

  // Prepara le matrici di output con dimensioni prealloccate
  const numGiorni = dateList.length;
  const numLuoghi = luoghiArray.length;

  // Preallocazione delle matrici per migliori prestazioni
  const matriceOutput = [];
  const matriceNoteOutput = [];
  const matriceColorOutput = [];

  // Funzione per creare una riga di array con lunghezza specificata
  function createEmptyRow(length, defaultValue = '') {
    return Array(length).fill(defaultValue);
  }

  // Intestazioni
  const intestazioneMesi = [inizio.getFullYear()];
  const intestazioneGiorno = [translate('specialEvent.date')];
  const intestazioneGiornoSett = [translate('specialEvent.structure')];

  for (const giorno of dateList) {
    //intestazioneMesi.push(giorno.getDate() === 1 ? nomiMesi[giorno.getMonth()] : '');
    intestazioneMesi.push(giorno.getDate() === 1 ? nomiMesi[giorno.getMonth()] + '\n' + giorno.getFullYear() : '');
    intestazioneGiorno.push(giorno.toISOString().split('T')[0]);
    intestazioneGiornoSett.push(lettereGiorniSettimana[giorno.getDay()]);
  }

  matriceOutput.push(intestazioneMesi);
  matriceNoteOutput.push(createEmptyRow(intestazioneMesi.length));
  matriceColorOutput.push(createEmptyRow(intestazioneMesi.length));

  matriceOutput.push(intestazioneGiorno);
  matriceNoteOutput.push(createEmptyRow(intestazioneGiorno.length));
  matriceColorOutput.push(createEmptyRow(intestazioneGiorno.length));

  matriceOutput.push(intestazioneGiornoSett);
  matriceNoteOutput.push(createEmptyRow(intestazioneGiornoSett.length));
  matriceColorOutput.push(createEmptyRow(intestazioneGiornoSett.length));

  // Cache per le ricerche di strutture
  const struttureLookup = new Map();
  for (let i = 0; i < struttureCache.length; i++) {
    struttureLookup.set(struttureCache[i][0], { index: i, nome: struttureCache[i][6] });
  }

  // Ottimizza getColoreCella con regole di regexp compilate una sola volta
  const regexE = /\(E\)/;
  const regexA = /\(A\)/;
  const regexD = /\(D\)/;
  const regexP = /\(P\)/;
  const regexL = /\(L\)/;
  const regexOpz = /Opz\./;
  const regexOff = /Off\./;

  function getColoreCella(nota) {
    const matches = nota.match(/\[(\d+)\]/g);
    if (!matches || matches.length === 0) return '';

    const unici = new Set(matches);
    if (unici.size > 1) return "#e06666"; // più eventi

    if (regexL.test(nota)) return "#c27ba0";

    let tipo = '';
    if (regexE.test(nota)) tipo = 'E';
    else if (regexA.test(nota) || regexD.test(nota)) tipo = 'AD';
    else if (regexP.test(nota)) tipo = 'P';

    if (regexOpz.test(nota)) {
      if (tipo === 'E') return "#ffd966";
      if (tipo === 'AD') return "#ffe599";
      if (tipo === 'P') return "#fff2cc";
    } else if (regexOff.test(nota)) {
      if (tipo === 'E') return "#93c47d";
      if (tipo === 'AD') return "#b6d7a8";
      if (tipo === 'P') return "#d9ead3";
    } else {
      if (tipo === 'E') return "#6fa8dc";
      if (tipo === 'AD') return "#9fc5e8";
      if (tipo === 'P') return "#cfe2f3";
    }

    return '';
  }

  // Popola le righe per luoghi in una sola passata
  for (const luogo of luoghiArray) {
    const luogoInfo = struttureLookup.get(luogo) || { nome: luogo };
    const nomeLuogo = luogoInfo.nome;

    const rigaOutput = [nomeLuogo];
    const rigaNote = [luogo];
    const rigaColori = [''];

    const luogoMap = mappaEventi.get(luogo);

    for (const giorno of dateList) {
      const giornoStr = giorno.toISOString().split('T')[0];
      const evento = luogoMap.get(giornoStr);

      if (evento && evento.eventi.length > 0) {
        rigaOutput.push(evento.numero);
        const notaPulita = evento.note.trim();
        //const notaPulita = evento.note;
        rigaNote.push(notaPulita);
        rigaColori.push(getColoreCella(notaPulita));
      } else {
        rigaOutput.push('');
        rigaNote.push('');
        rigaColori.push('');
      }
    }

    matriceOutput.push(rigaOutput);
    matriceNoteOutput.push(rigaNote);
    matriceColorOutput.push(rigaColori);
  }

  // Riga del mese nei primi giorni
  const rigaMeseSoloPrimi = ['Mese'];
  for (const giorno of dateList) {
    rigaMeseSoloPrimi.push(giorno.getDate() === 1 ? nomiMesi[giorno.getMonth()] : '');
  }
  // matriceOutput.push(rigaMeseSoloPrimi); // intestazioneMesi
  matriceOutput.push(intestazioneMesi);
  matriceNoteOutput.push(createEmptyRow(rigaMeseSoloPrimi.length));
  matriceColorOutput.push(createEmptyRow(rigaMeseSoloPrimi.length));

  const eventiEL = new Map();
  for (const giorno of dateList) {
    const giornoStr = giorno.toISOString().split('T')[0];

    for (const luogo of luoghiArray) {
      const luogoMap = mappaEventi.get(luogo);
      const evento = luogoMap.get(giornoStr);

      if (evento && evento.numeriUnici.size > 0 && evento.note) {
        if (regexE.test(evento.note) || regexL.test(evento.note)) {
          // Dividi le note per evento
          const noteParts = evento.note.split('\n\n');

          for (const numero of evento.numeriUnici) {
            if (!eventiEL.has(numero) || new Date(giornoStr) < new Date(eventiEL.get(numero).primaOccorrenza)) {
              // Trova la nota specifica per questo numero
              let notaSpecifica = '';
              for (const notaPart of noteParts) {
                if (notaPart.includes(`[${numero}]`)) {
                  notaSpecifica = notaPart.trim();
                  break;
                }
              }

              eventiEL.set(numero, {
                primaOccorrenza: giornoStr,
                nota: notaSpecifica
              });
            }
          }
        }
      }
    }
  }

  // Raggruppa per giorno gli eventi
  const eventiPerGiorno = new Map();
  for (const [numero, dati] of eventiEL.entries()) {
    const giornoStr = dati.primaOccorrenza;

    if (!eventiPerGiorno.has(giornoStr)) {
      eventiPerGiorno.set(giornoStr, []);
    }

    eventiPerGiorno.get(giornoStr).push({
      numero: numero,
      nota: dati.nota
    });
  }

  // Massimo numero di eventi in un giorno
  let maxEventiPerGiorno = 0;
  for (const eventi of eventiPerGiorno.values()) {
    maxEventiPerGiorno = Math.max(maxEventiPerGiorno, eventi.length);
  }

  // Crea le righe per eventi nel menu in basso
  for (let i = 0; i < maxEventiPerGiorno; i++) {
    const rigaEventi = [i === 0 ? translate('specialEvent.events') : ''];
    const rigaNote = [''];
    const rigaColori = [''];

    for (const giorno of dateList) {
      const giornoStr = giorno.toISOString().split('T')[0];
      const eventiDelGiorno = eventiPerGiorno.get(giornoStr) || [];

      if (i < eventiDelGiorno.length) {
        const evento = eventiDelGiorno[i];
        rigaEventi.push(evento.numero);
        rigaNote.push(evento.nota);
        rigaColori.push(getColoreCella(evento.nota));
      } else {
        rigaEventi.push('');
        rigaNote.push('');
        rigaColori.push('');
      }
    }

    matriceOutput.push(rigaEventi);
    matriceNoteOutput.push(rigaNote);
    matriceColorOutput.push(rigaColori);
  }

  // Matrice domeniche ottimizzata
  const matriceDomenicheOutput = [];
  const rigaDomeniche = ['Domenica'];

  for (const giorno of dateList) {
    rigaDomeniche.push(giorno.getDay() === 0); // true se domenica
  }

  matriceDomenicheOutput.push(rigaDomeniche);
  return { matriceOutput, matriceNoteOutput, matriceColorOutput, matriceDomenicheOutput };
}

// function popolaFoglioSheet(matriceOutput, matriceNoteOutput, matriceColorOutput, matriceDomenicheOutput, nomeFoglio) {
function popolaFoglioSheet(matriceOutput, matriceNoteOutput, matriceColorOutput, matriceDomenicheOutput) {
  // Cache dell'app
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const colori = readVariables('method2colors', DataSettings);

  // Ottieni o crea il foglio in modo ottimizzato
  var sheet = ss.getActiveSheet();

  // Ottimizza le operazioni di scrittura facendole in batch
  const righe = matriceOutput.length;
  const colonne = matriceOutput[0].length;

  // Batch operations per matrici
  if (righe > 0 && colonne > 0) {
    // Valori
    sheet.getRange(4, 1, righe, colonne)
      .setValues(matriceOutput)
      .setNumberFormat('00')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Note
    sheet.getRange(4, 1, righe, colonne).setNotes(matriceNoteOutput);

    // Colori di sfondo
    if (matriceColorOutput && matriceColorOutput.length > 0) {
      sheet.getRange(4, 1, matriceColorOutput.length, matriceColorOutput[0].length)
        .setBackgrounds(matriceColorOutput);
    }
  }

  // Operazioni di formattazione in batch
  // Intestazione
  sheet.getRange(4, 1, 1, colonne)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet.getRange(5, 1, 1, colonne)
    .setBackground('#cccccc')
    .setFontWeight('bold')
    .setNumberFormat("DD");

  // Prima colonna
  sheet.getRange(4, 1, righe, 1)
    .setBackground('#eeeeee')
    .setFontWeight('bold')
    .setNote('')
    .setHorizontalAlignment('right');

  // Auto-ridimensionamento per la prima colonna
  sheet.autoResizeColumn(1);

  // Imposta larghezza fissa per le altre colonne in un'unica operazione
  for (let i = 2; i <= colonne; i++) {
    sheet.setColumnWidth(i, 40);
  }

  // Imposta larghezza fissa per le altre righe in un'unica operazione
  for (let i = 6; i <= righe; i++) {
    sheet.setRowHeight(i, 5);
  }

  // Applica bordi alle domeniche
  if (matriceDomenicheOutput && matriceDomenicheOutput.length > 0) {
    const rigaDomeniche = matriceDomenicheOutput[0];
    const rangesDomenica = [];

    // Raggruppa le colonne da formattare
    for (let col = 1; col < rigaDomeniche.length; col++) {
      if (rigaDomeniche[col] === true) {
        rangesDomenica.push(sheet.getRange(4, col + 1, matriceOutput.length, 1));
      }
    }

    // Applica lo stile in batch alle domeniche
    if (rangesDomenica.length > 0) {
      rangesDomenica.forEach(range => {
        range.setBorder(false, false, false, true, false, false,
          "black", SpreadsheetApp.BorderStyle.SOLID);
      });
    }
  }
  // Legenda
  const lr = sheet.getLastRow();
  sheet.getRange(lr + 1, 1).setValue(translate('specialEvent.confirmed')).setFontSize(8).setFontWeight("bold").setBackground(colori[1][0]).setHorizontalAlignment("center");
  sheet.getRange(lr + 2, 1).setValue(translate('specialEvent.optionated')).setFontSize(8).setFontWeight("bold").setBackground(colori[4][0]).setHorizontalAlignment("center");
  sheet.getRange(lr + 3, 1).setValue(translate('specialEvent.offer')).setFontSize(8).setFontWeight("bold").setBackground(colori[14][0]).setHorizontalAlignment("center");
  sheet.getRange(lr + 4, 1).setValue(translate('specialEvent.work')).setFontSize(8).setFontWeight("bold").setBackground(colori[15][0]).setHorizontalAlignment("center");
  sheet.getRange(lr + 5, 1).setValue(translate('specialEvent.concurrent')).setFontSize(8).setFontWeight("bold").setBackground(colori[6][0]).setHorizontalAlignment("center");
  // translate('viewCalendar.moreEvents')
}
// -------------------------------- END NEW FAST SHOWMONTH ---------------------------------
