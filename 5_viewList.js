/*
* Project Name: Pavora
* Copyright (c) 2025 Nicola Rainiero
*
* This software is released under the MIT License.
* Please refer to the LICENSE file for the full license text.
*/
function testCreateListEvent() {
  first = '2025-01-01';
  last = '2025-01-31';
  //cosa = 'A,P,L,E,D';
  cosa = 'E';
  come = 'agg'; // day o agg
  // function createListEvent(first, last, cosa, selectedStruct, keyword, come)
  createListEvent(first, last, cosa, '', '', come);
}


function addFiltersAndIntelligentFilter() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var user = getAliasEmail(Session.getEffectiveUser().getEmail());
  var sheet = ss.getSheetByName(user); // Cambia "Sheet1" con il nome del tuo foglio

  // Rimuove eventuali filtri esistenti
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  // Aggiungi filtri a tutte le colonne
  var headerRow = sheet.getRange(2, 1, 1, sheet.getLastColumn());
  headerRow.setFontWeight('bold');
  headerRow.setBackground('grey');

  var filterRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
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

  //Logger.log('Le strutture uniche sono: ' + uniqueStructures);
  // Converti il Set in Array e ordina le strutture
  var uniqueStructuresArray = Array.from(uniqueStructures).sort();

  //Logger.log('Le strutture uniche sono: ' + uniqueStructuresArray);

  // Applica il filtro personalizzato alla colonna "Strutture"
  var criteria = SpreadsheetApp.newFilterCriteria()
    //.setVisibleValues(uniqueStructuresArray)
    //.setHiddenValues(uniqueStructuresArray)
    .build();
  filter.setColumnFilterCriteria(3, criteria);
}

function calculateEventDurations(events) {
  // Crea un oggetto per memorizzare i giorni unici per ciascun titolo principale
  const eventDurations = {};

  // Itera attraverso gli eventi per raccogliere i giorni unici per ciascun titolo principale
  events.forEach(event => {
    const title = parseEventString(event[1]).nome; // Titolo principale senza la lettera finale e senza Opz.
    const startDate = new Date(event[0]).toISOString().split('T')[0]; // Data di inizio (solo parte della data)
    const endDate = new Date(event[3]).toISOString().split('T')[0]; // Data di fine (solo parte della data)

    if (!eventDurations[title]) {
      eventDurations[title] = new Set();
    }

    // Aggiungi solo il giorno di inizio
    eventDurations[title].add(startDate);
  });

  // Calcola il numero di giorni unici per ciascun titolo principale
  const eventDurationsInDays = [];
  for (const title in eventDurations) {
    const uniqueDays = eventDurations[title].size;
    eventDurationsInDays.push([title, uniqueDays]);
  }

  return eventDurationsInDays;
}

// Funzione per ottenere la durata di un evento specifico
function getEventDuration(eventDurationsInDays, title) {
  const event = eventDurationsInDays.find(event => event[0] === title);
  return event ? event[1] : 0;
}

// Per restituirmi le tipologie da non mostrare nella ricerca.
function subtractStrings(first, second) {
  // Rimuove virgolette singole e spazi iniziali e finali, poi divide in array
  let arrayFirst = first.split(',').map(item => item.trim().replace(/'/g, ''));
  let arraySecond = second.split(',').map(item => item.trim().replace(/'/g, ''));

  // Filtra gli elementi presenti in arrayFirst ma non in arraySecond
  let result = arrayFirst.filter(item => !arraySecond.includes(item));

  // Formatta ogni elemento con uno spazio iniziale e virgolette singole
  return result.map(item => `${item}`);
}

function createListEvent(first, last, cosa, selectedStruct, keyword, come) {
  // Nuovo metodo per inizializzare il foglio
  resetFoglioConNuovo();

  // Step 1: inizializzare il foglio ed eliminare le immagini presenti
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Aggiornare le celle congelate
  sheet.setFrozenRows(2);
  if (come == 'agg') {
    sheet.setFrozenColumns(2);
  } else if (come == 'day') {
    sheet.setFrozenColumns(5);
  }
  sheet.getRange(1, 1, 200, 20).setFontSize(10);

  // Rimuovere le immagini esistenti
  var images = sheet.getImages();
  for (let i = 0; i < images.length; i++) {
    images[i].remove();
  }

  // vuoto vuol dire tutte le tipologie!!
  var tutteTipologie = 'P,A,E,D,L';
  var cosaVedere = subtractStrings(tutteTipologie, cosa);

  sheet.getRange(1, 1).setValue(first);
  sheet.getRange(2, 1).setValue(last);
  sheet.getRange(3, 1).setValue(JSON.stringify(cosaVedere));
  sheet.getRange(4, 1).setValue(selectedStruct);
  sheet.getRange(6, 1).setValue('WORKING IN PROGRESS .....');

  viewListEvents(first, last, cosa, selectedStruct, keyword, come)
}

// Convertire la data pulita con / senza anno
function convertDMBar(inputFormat) {
  function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat)
  return [pad(d.getDate()), pad(d.getMonth() + 1)].join('/')
}

function readableLocations(array) {
  // ---------> NEW TESTING <-----------------
  if (typeof (array) === 'string') { var array = string2array(array) }
  var structs = [];
  var gates = [];
  var areamqQ = 0;
  var areamqCC = 0;
  var areamqSum = 0;
  var pax = 0;
  var ae = 0;
  for (let x = 0; x < array.length; x++) {
    if (findKey(array[x], strutture(), 0) >= 0) {
      index = findKey(array[x], strutture(), 0);
      if ((strutture()[index][7] == 1) || (strutture()[index][7] == 4)) { // QF
        structs.push(strutture()[index][6]);
        areamqQ = areamqQ + strutture()[index][8];
      } else if (strutture()[index][7] == 2) { // CC
        structs.push(strutture()[index][6]);
        areamqCC = areamqCC + strutture()[index][8];
        pax = pax + strutture()[index][9];
      } else if (strutture()[index][7] == 5) { // ingressi pubblico
        gates.push(strutture()[index][6]);
      }
    }
  }
  areamqSum = areamqCC + areamqQ;
  if ((areamqCC != 0) && (areamqQ != 0)) {
    areamq = areamqSum.toLocaleString('it-IT') + '\n(Q=' + areamqQ.toLocaleString('it-IT') + ' | CC=' + areamqCC.toLocaleString('it-IT') + ')';
  } else {
    //areamq = areamqSum.toLocaleString('it-IT');
    areamq = areamqSum
  }

  var structs = structs.filter(onlyUnique);
  var structs = structs.sort((a, b) => a - b);
  var gates = gates.filter(onlyUnique);
  var gates = gates.sort((a, b) => a - b);
  if ((structs.length != 0) && (gates.length != 0)) {
    var result = '' + structs + '. Accessi: ' + gates + '.';
  } else if ((structs.length == 0) && (gates.length == 0)) {
    var result = 'Occupazione non definita.';
  } else if (structs.length == 0) {
    var result = '' + gates + '.';
  } else {
    var result = '' + structs + '.';
  }
  if ((areamq != 0) && (pax != 0)) {
    result = result + '\n Occupati ' + areamq.toLocaleString('it') + ' mq. e presenze massime autorizzate ' + pax.toLocaleString('it') + '.';
  } else if ((areamq != 0) && (pax == 0)) {
    result = result + '\n Occupati ' + areamq.toLocaleString('it') + ' mq.';
  } else if ((areamq == 0) && (pax != 0)) {
    result = result + '\n Presenze massime autorizzate ' + pax.toLocaleString('it') + '.';
  }
  var paxAE = 0;
  var areamqAE = 0;
  if ((pax < 300) && (pax > 1)) { paxAE = 2 };
  if ((pax < 600) && (pax >= 300)) { paxAE = 4 };
  if ((pax < 1000) && (pax >= 600)) { paxAE = 6 };
  if ((pax < 2000) && (pax >= 1000)) { paxAE = 6 };
  if ((pax < 3000) && (pax >= 2000)) { paxAE = 6 };
  if (pax >= 3000) { paxAE = 6 };

  if ((areamqQ < 100) && (areamqQ > 0)) { areamqAE = 0 };
  if ((areamqQ < 4000) && (areamqQ >= 100)) { areamqAE = 4 };
  if ((areamqQ < 10000) && (areamqQ >= 4000)) { areamqAE = 2 };
  if ((areamqQ < 20000) && (areamqQ >= 10000)) { areamqAE = 4 };
  if ((areamqQ < 150000) && (areamqQ >= 20000)) { areamqAE = 6 };
  if (areamqQ >= 150000) { areamqAE = 6 };
  if (array.includes('14')) {
    areamqAE = 2;
  };
  //Logger.log(pax + ' paxAE=' + paxAE + ' ' + areamqQ + ' mqAE=' + areamqAE);
  var aeSum = Number(paxAE) + Number(areamqAE);
  if ((pax != 0) && (areamqQ != 0)) {
    ae = aeSum + '\n(Q=' + areamqAE + ' | CC=' + paxAE + ')';
  } else {
    ae = aeSum;
  }

  //Logger.log('La matrice finale per l\'evento '+eventi[i][8]+' è \n\n'+result); // .toLocaleString('it-IT') 
  return {
    testo: result,
    strutture: structs.join(', '),
    mq: areamq,
    pax: pax,
    ae: ae
  }
}

function setAlternatingColorsWithConditionalFormatting(sheet, startRow, numRows, come) {
  var range = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());

  // Definisci i colori alternati
  var oddColor = '#ffffff'; // Colore per le righe dispari
  var evenColor = '#d9d9d9'; // Colore per le righe pari

  // Rimuove tutte le regole di formattazione condizionale esistenti
  sheet.clearConditionalFormatRules();

  // Regola di formattazione condizionale per le righe pari
  var evenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=ISEVEN(ROW())`)
    .setBackground(evenColor)
    .setRanges([range])
    .build();

  // Regola di formattazione condizionale per le righe dispari
  var oddRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=ISODD(ROW())`)
    .setBackground(oddColor)
    .setRanges([range])
    .build();

  // Regola di formattazione condizionale per colorare di giallo le righe con "SI" nella colonna 13 se cosa -> agg se day 11
  if (come == 'agg') {
    var yellowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=OR($O3="Opzionato";$O3="Offerta")`)
      .setBackground('#ffff00')
      .setRanges([range])
      .build();
    var redRule = SpreadsheetApp.newConditionalFormatRule()
      //.whenFormulaSatisfied('=AND($O3="SI", $T3<>"", OR($T3>$A3, $T3<=TODAY()))') // Confronta le date --> ENGLISH VERSION
      .whenFormulaSatisfied('=AND(OR($O3="Opzionato";$O3="Offerta"); $T3<>""; OR($T3>$A3; $T3<=TODAY()))') // Confronta le date --> ITALIAN VERSION
      .setBackground('#e06666') // Imposta uno sfondo rosso per evidenziare
      .setRanges([range]) // Specifica l'intervallo a cui applicare la regola
      .build();
  } else if (come == 'day') {
    var yellowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=OR($K3="Opzionato";$K3="Offerta")`)
      .setBackground('#ffff00')
      .setRanges([range])
      .build();
  }

  var rules = sheet.getConditionalFormatRules();
  if (come == 'agg') {
    rules.push(redRule);
  }
  rules.push(yellowRule);
  rules.push(evenRule);
  rules.push(oddRule);

  sheet.setConditionalFormatRules(rules);
}

// -----------------------------------------------------------
// Generare un foglio Google dalle voci del proprio calendario
// -----------------------------------------------------------
//function preparaTabellaAggr(inizio --> '2024-07-24', fine, searchWord, tipo) {
function viewListEvents(fromDate, toDate, cosa, multiBuilding, keyword, come) {

  // Converto le date in formato datetime
  var dateParts = fromDate.split("-");
  var fromDate = new Date(+dateParts[0], dateParts[1] - 1, +dateParts[2]);
  var primo = [dateParts[2], dateParts[1], dateParts[0]].join('/');
  var dateParts = toDate.split("-");
  var toDate = new Date(+dateParts[0], dateParts[1] - 1, +dateParts[2]);
  toDate = new Date(toDate.getTime() + 1 * 3600000 * 24); // add 24h to bring the date at the midnight
  var ultimo = [dateParts[2], dateParts[1], dateParts[0]].join('/');
  // Count the number of days between the two dates
  var numMilliSec = toDate.getTime() - fromDate.getTime();
  var numDays = numMilliSec / (1000 * 3600 * 24);
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetTitle = translate('viewCalendar.mainCell'); // titolo della tabella

  // Caricamento eventi dal calendario
  var cal = CalendarApp.getCalendarById(myCalID()[0][0]);
  var events = cal.getEvents(new Date(fromDate), new Date(toDate), { search: keyword });

  // Remove the events that start before the first day and end in that day (for example: 23.00-04.00)
  t = 0;
  while ((events.length > 1) && (events[t].getStartTime().getTime() < fromDate.getTime())) {
    events.splice(t, 1);
    t++
  }

  // Inizializzo la tabella degli eventi aggregati
  var listaEventiStart = [];

  // Per tenere solo le fasi definite in "cosa" vedere
  if (events.length == 0) { // non c'è alcun evento nell'intervallo
    var events = [];
    var finaleEventi = [];
    var listaEventiStart = [];
  }
  //example of cosa = [" A", " D", " L", " P"];
  var myarr = (cosa.length > 0) ? cosa : 'E'; // se cosa è vuota allora ricerca solo gli eventi
  // itera per lasciare solo gli eventi che contengono la tipologia di cosa e ridefinisci le locations

  for (let i = 0; i < events.length; i++) {
    locations = string2array(events[i].getLocation());
    cleanLocations = readableLocations(locations).strutture;
    var row = [events[i].getStartTime(), events[i].getTitle(), events[i].getStartTime(), events[i].getEndTime(), events[i].getLocation(), events[i].getDescription(), events[i].getCreators(), events[i].getLocation(), parseEventString(events[i].getTitle()).nome, '', events[i].getCreators()];

    // Rimuovo tutti gli eventi di tipo myarr
    var tipoEventi = parseEventString(events[i].getTitle()).type;
    if (myarr.indexOf(tipoEventi) >= 0) { // lascio solo gli eventi che non sono in myarr
      listaEventiStart.push(row);
    }
  }
  //  
  // FINE: Per rimuovere eventi che finiscono con A D e L
  //

  //  
  // INIZIO: Aggregatore
  //
  // Trovo tutti i singoli e distinti eventi nell'intervallo
  var nomeEventi = [];
  var listaEventi = []
  const durations = calculateEventDurations(listaEventiStart);
  for (let i = 0; i < listaEventiStart.length; i++) {
    var nomeEvento = parseEventString(listaEventiStart[i][1]).nome;
    if (nomeEventi.indexOf(nomeEvento) > -1) {
      listaEventi[findKey(nomeEvento, listaEventi, 8)][9] = getEventDuration(durations, nomeEvento);
    } else {
      listaEventiStart[i][9] = getEventDuration(durations, nomeEvento);
      listaEventi.push(listaEventiStart[i]);
      nomeEventi.push(nomeEvento);
    }
  }
  /*
  for (let i = 0; i < listaEventi.length; i++) {
    Logger.log(listaEventi[i][1]+' | Prima---------> '+listaEventi[i][9]);
    listaEventi[i][9] = listaEventi[i][9].split(',').filter(onlyUnique).join(',');
    Logger.log('Dopo---------> '+listaEventi[i][9]);
  }
  */
  // Se ricerco solo alcune strutture --> lascio solo gli eventi che contengono una delle strutture presenti
  // Fix Multibuilding
  if (multiBuilding === ' ') { multiBuilding = '' };
  // Load the structures
  // Transform keyword into an array
  if (multiBuilding === '') {
  } else {
    var structKeyArray = multiBuilding.split(',').map(function (value) {
      return value.trim();
    });
    //Logger.log('I building sono ' + structKeyArray);
    var listaStruct = [];
    for (let i = 0; i < listaEventi.length; i++) {
      var strucIterArray = listaEventi[i][7].split(',').map(function (value) {
        return value.trim();
      });
      if (structKeyArray.some(r => strucIterArray.includes(r))) {
        listaStruct.push(listaEventi[i]);
      }

    }
    //Logger.log('Le strutture selezionate sono ' + listaStruct);
    listaEventi = listaStruct;

    var listaStructStart = [];
    for (let i = 0; i < listaEventiStart.length; i++) {


      var strucIterArray = listaEventiStart[i][7].split(',').map(function (value) {
        return value.trim();
      });
      //Logger.log('Le strutture nella iterazione i sono ' + strucIterArray);
      // Add 5 if structKeyArray contain 5A, 5B
      if ((structKeyArray.includes('5A')) || (structKeyArray.includes('5B'))) {
        structKeyArray.push('5');
      }
      if (structKeyArray.some(r => strucIterArray.includes(r))) {
        listaStructStart.push(listaEventiStart[i]);
      }

    }
    //Logger.log('Le strutture selezionate sono ' + listaStruct);
    listaEventiStart = listaStructStart;


  }
  // Fix the number of the days.

  //var durations = calculateEventDurations(listaEventiStart);
  /*
  for (let i = 0; i < listaEventi.length; i++) {
    if (listaEventi[i][9] != '') {
      giorni = listaEventi[i][9].split(",").length + ' giorni: ';
    } else { giorni = 1 + ' giorno: ' };
    listaEventi[i][9] = giorni + convertDMBar(listaEventi[i][0]) + ',' + listaEventi[i][9];
    listaEventi[i][9] = listaEventi[i][9].slice(0, -1);
  }
  */
  //  
  // FINE: Aggregatore
  //

  // Preparo la tabella
  var lr = sheet.getLastRow();
  var mr = sheet.getMaxRows();
  if (mr - lr != 0) {
    //sheet.deleteRows(lr+1, mr-lr);
  }
  var lc = sheet.getLastColumn();
  var mc = sheet.getMaxColumns();
  if (mc - lc != 0) {
    //sheet.deleteColumns(lc+1, mc-lc);
  }
  var range = sheet.getRange(2, 1, lr, lc);
  range.clearContent();
  range.clearFormat();
  range.setBackground('#FFFFFF').setBorder(false, false, false, false, false, false, "grey", SpreadsheetApp.BorderStyle.DASHED).setNote('').setVerticalAlignment("middle").setFontSize(10);

  // Stili per la tabella
  var headerColor = "#999999"; // #EDD400 giallo per l'intestazione della tabella 002d62 blu sito
  var textHeaderColor = "#000000";
  var firstColor = "#ffffff"; // bianco per la prima riga alternata
  var secondColor = "#c9edee"; // E0E0E0 grigio per la seconda riga alternata oppure c9edee

  // Header of the sheet
  var adesso = new Date();
  var adesso = formatDateMaster(adesso).giorno + ' ' + formatDateMaster(adesso).ora;
  var sottotitolo = translate('viewList.eventsFrom') + primo + translate('viewList.eventsTo') + ultimo + translate('viewList.updateAt') + adesso + ')';
  if (come == 'agg') {
    //var header = [["Inizio", "Evento", "Codice", "Strutture", "∑mq", "∑pax ", "Note", "Durata (" + myarr + ")", "Referente", "Operation", "Organizzatore", "Allestitore", "Catering", "Tipo evento", "Opzionato?", "Pubblico?", "VVF", "CRI", "Num. AE", "Scadenza Opz."]];
    var header = [translate('viewList.headerAggr', { myarr: myarr.replace(/,/g, '|') }).split(',')];
    var range = sheet.getRange(2, 1, 1, 20);
  } else if (come == 'day') {
    //var header = [["Inizio", "Evento ", "Tipo", "Inizio", "Fine", "Strutture", "∑mq", "∑pax ", "Note", "Tipo evento", "Opzionato?", "Pubblico", "VVF", "CRI", "Num. AE"]];
    var header = [translate('viewList.headerList').split(',')];
    var range = sheet.getRange(2, 1, 1, 15);
  }

  range.setValues(header).setFontColor(textHeaderColor).setFontSize(12).setBackground(headerColor).setHorizontalAlignment("center");

  // Step 2: Centrare il contenuto delle colonne 4 e 5 e impostare il formato testo
  var numRows = 100; // Stima del numero di righe che verranno importate
  var rangeCol4 = sheet.getRange(1, 4, numRows);
  var rangeCol5 = sheet.getRange(1, 5, numRows);
  var rangeCol7 = sheet.getRange(1, 7, numRows);

  rangeCol4.setHorizontalAlignment("center");
  rangeCol5.setHorizontalAlignment("center");

  rangeCol4.setNumberFormat("@STRING@");
  rangeCol5.setNumberFormat("@STRING@");
  rangeCol7.setNumberFormat('#');

  if (come == 'agg') {
    if (listaEventi.length != 0) {
      var col2MaxLength = 0;
      var col3MaxLength = 0;
      var col4MaxLength = 0;
      //const durations = calculateEventDurations(listaEventi);
      for (let i = 0; i < listaEventi.length; i++) {
        var row = i + 3;
        var myformula_placeholder = '';
        // Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
        // NOTE: I've had problems with the getVisibility for some older events not having a value, so I've had do add in some NULL text to make sure it does not error

        // ORIGINALE
        var typeEvento = (findKey(parseEventDetails(listaEventi[i][5]).typeEv, typeEv(), 1) > 0) ? typeEv()[findKey(parseEventDetails(listaEventi[i][5]).typeEv, typeEv(), 1)][1] : 'nonSpec';
        var details = [[listaEventi[i][0], parseEventString(listaEventi[i][1]).nome, parseEventDetails(listaEventi[i][5]).code, readableLocations(listaEventi[i][4]).strutture, readableLocations(listaEventi[i][4]).mq, readableLocations(listaEventi[i][4]).pax, parseEventDetails(listaEventi[i][5]).descrizione, listaEventi[i][9], parseEventDetails(listaEventi[i][5]).refCom, parseEventDetails(listaEventi[i][5]).refOp, parseEventDetails(listaEventi[i][5]).org, parseEventDetails(listaEventi[i][5]).all, parseEventDetails(listaEventi[i][5]).feed, typeEvento, parseEventString(listaEventi[i][1]).opz, parseEventDetails(listaEventi[i][5]).open, parseEventDetails(listaEventi[i][5]).vvf, parseEventDetails(listaEventi[i][5]).cri, readableLocations(listaEventi[i][4]).ae, parseEventDetails(listaEventi[i][5]).opzExp]];

        var range = sheet.getRange(row, 1, 1, 20).setWrap(true);
        // Step 1: Calcolare la larghezza necessaria per la colonna 3
        for (let j = 0; j < details.length; j++) {
          if (details[j][1].length > col2MaxLength) {
            col2MaxLength = details[j][1].length;
          }
          if (details[j][2].length > col3MaxLength) {
            col3MaxLength = details[j][2].length;
          }
          if (details[j][4].length > col4MaxLength) {
            col4MaxLength = details[j][4].length;
          }
        }
        range.setValues(details).setVerticalAlignment("middle").setHorizontalAlignment("center");
        sheet.getRange(row, 2).setNote(listaEventi[i][5] + ' idNome=' + parseEventString(listaEventi[i][1]).nome + ' (' + formatEmail(String(listaEventi[i][10])) + ') | ' + listaEventi[i][4]);
        sheet.setColumnWidths(1, 17, 75);

      }
    } else {
      //var totalRows = sheet.getLastRow();
      sheet.getRange(4, 2).setValue('NESSUN EVENTO TROVATO').setFontSize(10).setHorizontalAlignment("left");
    }
  } else if (come == 'day') {

    if (listaEventiStart.length != 0) {
      var col2MaxLength = 0;
      var col6MaxLength = 0;
      var col7MaxLength = 0;
      for (let i = 0; i < listaEventiStart.length; i++) {
        var row = i + 3;
        var typeEvento = (findKey(parseEventDetails(listaEventiStart[i][5]).typeEv, typeEv(), 1) > 0) ? typeEv()[findKey(parseEventDetails(listaEventiStart[i][5]).typeEv, typeEv(), 1)][0] : 'nonSpec';
        var details = [[listaEventiStart[i][0], parseEventString(listaEventiStart[i][1]).nome, parseEventString(listaEventiStart[i][1]).type, formatDateMaster(listaEventiStart[i][2]).ora, formatDateMaster(listaEventiStart[i][3]).ora, readableLocations(listaEventiStart[i][4]).strutture, readableLocations(listaEventiStart[i][4]).mq, readableLocations(listaEventiStart[i][4]).pax, parseEventDetails(listaEventiStart[i][5]).descrizione, typeEvento, parseEventString(listaEventiStart[i][1]).opz, parseEventDetails(listaEventiStart[i][5]).open, parseEventDetails(listaEventiStart[i][5]).vvf, parseEventDetails(listaEventiStart[i][5]).cri, readableLocations(listaEventiStart[i][4]).ae]];

        var range = sheet.getRange(row, 1, 1, 15).setWrap(true);
        // Step 1: Calcolare la larghezza necessaria per la colonna 3
        for (let j = 0; j < details.length; j++) {
          if (details[j][1].length > col2MaxLength) {
            col2MaxLength = details[j][1].length;
          }
          if (details[j][5].length > col6MaxLength) {
            col6MaxLength = details[j][5].length;
          }
          if (details[j][6].length > col7MaxLength) {
            col7MaxLength = details[j][6].length;
          }
        }
        range.setValues(details).setVerticalAlignment("middle").setHorizontalAlignment("center");
        sheet.getRange(row, 2).setNote(listaEventiStart[i][5] + ' (' + formatEmail(String(listaEventiStart[i][10])) + ')');
        sheet.setColumnWidths(1, 15, 75);
        sheet.setColumnWidths(3, 1, 50);
      }
    } else {
      sheet.getRange(4, 2).setValue('NESSUN EVENTO TROVATO').setFontSize(10).setHorizontalAlignment("left");
    }
  }

  //sheet.setColumnWidths(6, 1, 500)
  sheet.autoResizeColumns(1, 20);
  if (come == 'agg') {
    // Durata giorni
    sheet.setColumnWidths(1, 7, 150);

    // Impostare la larghezza della colonna 2
    var column2Width = (col2MaxLength * 6) > 70 ? col2MaxLength * 6 : 75;
    sheet.setColumnWidth(2, column2Width);
    var column3Width = (col3MaxLength * 1) > 70 ? col3MaxLength * 1 : 150;
    sheet.setColumnWidth(4, column3Width);
    var column4Width = (col4MaxLength * 6) > 70 ? col4MaxLength * 6 : 75; // mq Regola questo fattore moltiplicativo in base alle necessità
    sheet.setColumnWidth(5, column4Width);
    sheet.setColumnWidth(12, column4Width);  // Stand Fitter
    sheet.setColumnWidth(13, column4Width);  // Catering
    sheet.setColumnWidth(14, column4Width);  // Type    
    sheet.setColumnWidth(19, column4Width);  // AE
    sheet.setColumnWidth(17, column4Width);  // VVF
    sheet.setColumnWidth(18, column4Width);  // CRI        
    sheet.getRange(1, sheet.getLastColumn()).setValue(sottotitolo).setHorizontalAlignment("right").setFontSize(10);
    sheet.getRange(1, 1).setValue(sheetTitle).setNumberFormat('0').setHorizontalAlignment("left").setFontSize(16);
    var numRows = (sheet.getLastRow() > 3) ? (nomeEventi.length) : 6;
  } else if (come == 'day') {
    // Durata giorni
    sheet.setColumnWidths(10, 1, 150);
    sheet.setColumnWidths(1, 15, 75);
    sheet.setColumnWidths(3, 1, 50);
    sheet.setColumnWidths(4, 2, 75);

    // Impostare la larghezza della colonna 2
    var column2Width = (col2MaxLength * 6) > 70 ? col2MaxLength * 6 : 75;
    sheet.setColumnWidth(2, column2Width);
    var column6Width = (col6MaxLength * 1) > 150 ? col6MaxLength * 2 : 250;
    sheet.setColumnWidth(6, column6Width);
    var column7Width = (col7MaxLength * 6) > 70 ? col7MaxLength * 6 : 75; // mq Regola questo fattore moltiplicativo in base alle necessità
    sheet.setColumnWidth(7, column7Width);
    sheet.setColumnWidth(15, column7Width);  // AE
    sheet.getRange(1, sheet.getLastColumn()).setValue(sottotitolo).setHorizontalAlignment("right").setFontSize(10);
    sheet.getRange(1, 1).setValue(sheetTitle).setNumberFormat('0').setHorizontalAlignment("left").setFontSize(16);
    var numRows = (sheet.getLastRow() > 3) ? (listaEventiStart.length) : 6;
  }


  if (nomeEventi.length >= 1) {
    setAlternatingColorsWithConditionalFormatting(sheet, 3, numRows, come);

    var totalRows = sheet.getLastRow();
    var lc = sheet.getLastColumn();
    var mc = sheet.getMaxColumns();
    if (mc - lc != 0) {
      //var range = sheet.getRange(3, 1, lr, lc);
      //range.clearContent();
      //range.clearFormat();
      sheet.deleteColumns(lc + 1, mc - lc);
    } else {
      sheet.deleteColumns(7, 1);
    }
    //Logger.log(lc);
    //Logger.log(mc);


    var lr = sheet.getLastRow();
    //Logger.log('Riga finale --> ' + lr);
    var mr = sheet.getMaxRows();
    //Logger.log('Max Riga finale --> ' + lr);
    //if (mr-lr != 0){
    if (lr - 2 != 0) {
      sheet.deleteRows(lr + 1, mr - lr);
    }
  } else {
    //Logger.log('Nessun evento trovato');
  }

  addFiltersAndIntelligentFilter();

  // Ottieni l'intervallo della prima colonna specificata
  //Logger.log('Last row è '+sheet.getLastRow()+' Max row è '+sheet.getMaxRows());
  if (come == 'agg') {
    sheet.getRange(1, 2).setValue('ATTIVO');
    sheet.getRange(1, 1, sheet.getLastRow()).setNumberFormat('dd/MM/yy');
    sheet.getRange(1, 4, sheet.getLastRow()).setNumberFormat("#,#");
    sheet.getRange(1, 5, sheet.getLastRow()).setNumberFormat("#,#");
    sheet.getRange(1, 6, sheet.getLastRow()).setNumberFormat('#');
    sheet.getRange(1, 8, sheet.getLastRow()).setNumberFormat('#');
    sheet.getRange(1, 19, sheet.getLastRow()).setNumberFormat('#');
    sheet.getRange(1, 20, sheet.getLastRow()).setNumberFormat('dd/MM/yy');
  } if (come == 'day') {
    sheet.getRange(1, 1, sheet.getLastRow()).setNumberFormat('dd/MM/yy');
    sheet.getRange(1, 7, sheet.getLastRow()).setNumberFormat("#,#");
    sheet.getRange(1, 8, sheet.getLastRow()).setNumberFormat("#,#");
    sheet.getRange(1, 15, sheet.getLastRow()).setNumberFormat('#');
  }

  // Per aggiornare i campi della tabella con quelli nelle variabili
  if (come == 'agg') {
    aggiornaColonna(refCom(), 9);
    aggiornaColonna(refOp(), 10);
    aggiornaColonna(allestitore(), 12);
    aggiornaColonna(catering(), 13);
    aggiornaColonna(typeEv(), 14);
    aggiornaColonna([['Opzionato', 'SI', 1], ['Confermato', 'NO', 1], ['Offerta', 'OFF', 1]], 15); // OPZ
    aggiornaColonna([['SI', 'SI', 1], ['NO', 'NO', 1]], 16);
    aggiornaColonna([['SI', 'SI', 1], ['NO', 'NO', 1], ['Richiesto', 'NI', 1], ['CPVLPS', 'CPVLPS', 1]], 17);
    aggiornaColonna([['SI', 'SI', 1], ['NO', 'NO', 1], ['Richiesto', 'NI', 1], ['CPVLPS', 'CPVLPS', 1]], 18);

    // Per impostare le formule sulle colonne VVF, CRI
    for (let i = 3; i <= sheet.getLastRow(); i++) {
      sheet.getRange(i, 17).setFormula('=IF(OR(AND(P' + i + '="SI"; E' + i + '>=4000);AND(P' + i + '="SI"; F' + i + '>=1000)); "SI"; "NO")'); //17 -> VVF
      sheet.getRange(i, 18).setFormula('=IF(OR(AND(P' + i + '="SI"; E' + i + '>=4000);AND(P' + i + '="SI"; G' + i + '>=1000)); "SI"; "NO")'); //18 -> CRI
    }
  } if (come == 'day') {
    //aggiornaColonna(refCom(), 9);
    //aggiornaColonna(refOp(), 10);
    //aggiornaColonna(allestitore(), 12);
    //aggiornaColonna(catering(), 13);
    aggiornaColonna(typeEv(), 10);
    aggiornaColonna([['Opzionato', 'SI', 1], ['Confermato', 'NO', 1], ['Offerta', 'OFF', 1]], 11); // OPZ
    aggiornaColonna([['SI', 'SI', 1], ['NO', 'NO', 1]], 12); // PUBBLICO
    aggiornaColonna([['SI', 'SI', 1], ['NO', 'NO', 1], ['Richiesto', 'Richiesto', 1], ['CPVLPS', 'CPVLPS', 1]], 13);
    aggiornaColonna([['SI', 'SI', 1], ['NO', 'NO', 1], ['Richiesto', 'Richiesto', 1], ['CPVLPS', 'CPVLPS', 1]], 14);

    // Per impostare le formule sulle colonne VVF, CRI
    for (let i = 3; i <= sheet.getLastRow(); i++) {
      sheet.getRange(i, 13).setFormula('=IF(OR(AND(L' + i + '="SI"; G' + i + '>=4000);AND(L' + i + '="SI"; H' + i + '>=1000)); "SI"; "NO")'); //13 -> VVF
      sheet.getRange(i, 14).setFormula('=IF(OR(AND(L' + i + '="SI"; G' + i + '>=4000);AND(L' + i + '="SI"; I' + i + '>=1000)); "SI"; "NO")'); //14 -> CRI
    }
  }
}

////////////////////////////////////
// ADD MENU AND CHANGE CALENDAR
//////////////////////////////////

function aggiornaColonna(matrice, colonnaTarget) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const foglio = ss.getActiveSheet(); // Foglio attivo
  const primaRiga = 3; // Inizia dalla terza riga (escludendo intestazioni)
  const ultimaRiga = foglio.getLastRow(); // Ultima riga con dati

  // Crea una mappa di sostituzione e un array per il menu a tendina
  const mappaValori = {};
  const valoriAttivi = []; // Per il menu a tendina (solo attivi)
  matrice.forEach(riga => {
    const [nomeCompleto, chiaveID, attivazione] = riga;
    mappaValori[chiaveID] = nomeCompleto;
    if (attivazione === 1) {
      valoriAttivi.push(nomeCompleto);
    }
  });

  // Ottiene i valori esistenti nella colonna
  const rangeColonna = foglio.getRange(primaRiga, colonnaTarget, ultimaRiga - primaRiga + 1);
  const valoriColonna = rangeColonna.getValues();

  // Sostituisce i valori nella colonna con i nomi completi
  const nuoviValori = valoriColonna.map(riga => {
    const valoreCorrente = riga[0];
    return [mappaValori[valoreCorrente] || valoreCorrente]; // Sostituisci solo se trovato
  });
  rangeColonna.setValues(nuoviValori);

  // Aggiunge il menu a tendina nella stessa colonna
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(valoriAttivi, true) // Solo valori attivi
    .setAllowInvalid(true) // Permetti valori non validi (nel caso siano spenti)
    .build();
  rangeColonna.setDataValidation(rule);
}

function trovaChiaveID(nomeCompleto, matrice) {
  // Cerca nella matrice il nome completo e restituisce la chiave ID corrispondente
  for (let i = 0; i < matrice.length; i++) {
    if (matrice[i][0] === nomeCompleto) { // matrice[i][0] contiene il nome completo
      return matrice[i][1]; // matrice[i][1] contiene la chiave ID
    }
  }
  return null; // Restituisce null se il nome completo non viene trovato
}

function confrontaVettoriConTitoli(titoli, iniziale, finale) {

  let differenze = []; // Array per memorizzare le differenze

  // Confronta i due vettori
  for (let i = 0; i < titoli.length; i++) {
    const titolo = titoli[i];
    const valoreIniziale = iniziale[i].toString().replace(/\s/g, "") || "N/A"; // "N/A" se non esiste un valore
    const valoreFinale = finale[i].toString().replace(/\s/g, "") || "N/A";

    if (valoreIniziale !== valoreFinale) {
      // Registra la differenza con il titolo
      differenze.push({
        titolo: titolo,
        valoreIniziale: valoreIniziale,
        valoreFinale: valoreFinale
      });
    }
  }

  // Mostra le differenze
  if (differenze.length > 0) {
    let messaggio = translate('viewList.diffFound');
    differenze.forEach(diff => {
      messaggio += `${diff.titolo}: "${diff.valoreIniziale}" → "${diff.valoreFinale}"\n`;
    });
    return messaggio;
  } else {
    return translate('viewList.noDiffFound');
  }
}

function onEditTrigger(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const cellaAttivazione = sheet.getRange("B1").getValue();
  const colonneIgnorate = [1, 4, 5, 6, 8, 19];

  // Controlla se la cella B1 contiene "ATTIVO" e se la modifica è stata effettuata a partire dalla terza riga
  if (cellaAttivazione === "ATTIVO" && range.getRow() >= 3) {
    const colonna = range.getColumn();
    updateTimeUser();

    // Controlla se la colonna modificata non è da ignorare
    if (!colonneIgnorate.includes(colonna)) {
      range.setBorder(true, true, true, true, false, false, "#980000", SpreadsheetApp.BorderStyle.SOLID_THICK);
    }
  }
}


function trygetCellListaNote() {
  //what = "updateDetailsEvent" | updateDetailsEvent | deleteEvent;
  what = "updateDetailsEvent";
  getCellListNote(what);
}

/**
 * Gestisce le operazioni sugli eventi del calendario dalla cella attiva del foglio
 * @param {string} what - Tipo di operazione: "updateDetailsEvent", "updateSpecificEvent", "deleteEvent"
 * @param {string} first - Data iniziale (opzionale, calcolata automaticamente)
 * @param {string} last - Data finale (opzionale, calcolata automaticamente)
 * @returns {string} Nome dell'evento o messaggio di errore
 */
function getCellListNote(what, first, last) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  updateTimeUser();
  
  // Validazione colonne minime
  const lastColumn = sh.getLastColumn();
  if (lastColumn <= 10) {
    ui.alert(translate('viewList.alertEdit'));
    return translate('modifyEvent.emptyCell');
  }
  
  // Recupero informazioni cella attiva
  const activeRow = sh.getActiveRange().getRow();
  const cell = sh.getActiveCell();
  const noteComplete = cell.getNote();
  const name = cell.getValue();
  
  // Recupero nota e valore dalla colonna B
  const nameNote = sh.getRange(activeRow, 2).getNote();
  const nameValue = sh.getRange(activeRow, 2).getValue();
  
  // Estrazione nome evento
  let note = '';
  if (nameNote.length !== 0) {
    const extractedId = extractRegex(regexId, nameNote);
    note = extractedId !== 0 
      ? extractedId 
      : parseEventString(extractRegex(regexName, nameNote)).nome;
  }
  
  // Recupero valori riga
  const range = sh.getRange(activeRow, 1, 1, lastColumn);
  const valuesRow = range.getValues();
  
  // Calcolo date di ricerca (se non fornite)
  if (!note) {
    ui.alert(translate('viewList.alertEdit'));
    return translate('modifyEvent.emptyCell');
  }
  
  const today = new Date();
  const startDate = new Date(today.getFullYear() - 2, 0, 1);
  const endDate = new Date(today.getFullYear() + 6, 0, 1);
  first = formatDateMaster(startDate).dataXweb;
  last = formatDateMaster(endDate).dataXweb;
  
  // Recupero eventi
  const startList = logMatchingEvents(myCalID()[0][0], note, first, last, what);
  
  // Costruzione finalList
  const finalList = buildFinalList(startList, valuesRow);
  
  // Gestione operazioni
  switch (what) {
    case "updateDetailsEvent":
      if (lastColumn > 16) {
        handleUpdateDetailsEvent(sh, ui, note, first, last, finalList, startList, activeRow);
      }
      break;
      
    case "deleteEvent":
      first = convertDateInputHtml(finalList[0][18]);
      last = convertDateInputHtml(finalList[finalList.length - 1][19]);
      deleteEvents(note, first, last, what, activeRow);
      break;
      
    case "updateSpecificEvent":
      if (lastColumn > 12 && note !== '') {
        handleUpdateSpecificEvent(ui, note, finalList);
      } else if (note === '') {
        ui.alert(translate('viewList.alertOldEdit', { name: note }));
      }
      break;
      
    default:
      ui.alert(translate('viewList.alertEdit'));
  }
  
  return note || translate('modifyEvent.emptyCell');
}

/**
 * Costruisce la lista finale degli eventi con tutti i dati necessari
 */
function buildFinalList(startList, valuesRow) {
  const finalList = [];
  
  for (let j = 0; j < startList.length; j++) {
    const idEvent = startList[j][13].length === 0 ? randomID(8) : startList[j][13];
    
    // Lookup degli ID dalle matrici di riferimento
    const all = trovaChiaveID(valuesRow[0][11], allestitore());
    const referenteCom = trovaChiaveID(valuesRow[0][8], refCom());
    const referenteOp = trovaChiaveID(valuesRow[0][9], refOp());
    const typeEvento = trovaChiaveID(valuesRow[0][13], typeEv());
    const feed = trovaChiaveID(valuesRow[0][12], catering());
    
    const opzExp = valuesRow[0][19] === '' ? '' : convertDateInputHtml(valuesRow[0][19]);
    
    finalList.push([
      valuesRow[0][1],   // nome
      valuesRow[0][14],  // opzionato
      valuesRow[0][15],  // pubblico
      startList[j][3],   // quartiere
      startList[j][4],   // congress
      startList[j][5],   // ingressi
      startList[j][6],   // aree
      startList[j][7],   // parcheggi
      all,               // allestitore
      valuesRow[0][16],  // vvf
      valuesRow[0][17],  // cri
      referenteCom,      // referente commerciale
      valuesRow[0][6],   // note
      idEvent,           // id evento
      startList[j][14],  // type
      startList[j][15],  // data
      startList[j][16],  // startTime
      startList[j][17],  // endTime
      startList[j][18],  // startTimeUTC
      startList[j][19],  // endTimeUTC
      startList[j][20],  // location
      referenteOp,       // referente operativo
      valuesRow[0][10],  // organizzatore
      typeEvento,        // tipo evento
      startList[j][24],  // color
      feed,              // catering
      valuesRow[0][2],   // codice amministrazione
      opzExp             // data scadenza opzione
    ]);
  }
  
  return finalList;
}

/**
 * Gestisce l'aggiornamento completo dei dettagli dell'evento
 */
function handleUpdateDetailsEvent(sh, ui, note, first, last, finalList, startList, activeRow) {
  const titoli = translate('viewList.titleChanges').split(',');
  
  first = convertDateInputHtml(finalList[0][18]);
  last = convertDateInputHtml(finalList[finalList.length - 1][19]);
  
  const eventi = findEventsByKeyword(myCalID()[0][0], note, first, last);
  
  // Costruzione matrice eventi e testo per conferma
  const matrice = [];
  let testo = '';
  for (let i = 0; i < eventi.length; i++) {
    testo += `${eventi[i].getTitle()}\t${convertDate(eventi[i].getStartTime())}\n`;
    matrice.push([
      eventi[i].getTitle(),
      eventi[i].getStartTime(),
      eventi[i].getEndTime(),
      eventi[i].getDescription(),
      eventi[i].getLocation()
    ]);
  }
  
  // Prima conferma: mostra differenze
  const confronto = confrontaVettoriConTitoli(titoli, startList[0], finalList[0]);
  const response1 = ui.alert(
    translate('viewList.updateOf') + finalList[0][0],
    confronto + translate('viewList.okTo'),
    ui.ButtonSet.YES_NO
  );
  
  if (response1 !== ui.Button.YES) {
    ui.alert(translate('modifyEvent.opNOMessage'));
    return;
  }
  
  if (!checkUserWritePermission(myCalID()[0][0])) {
    ui.alert(translate('modifyEvent.waitSomeTime'));
    return;
  }
  
  // Seconda conferma: cancellazione eventi
  const eventId = buildEventId(matrice[0]);
  let testoDel = translate('modifyEvent.warEventDel', { eventId: eventId });
  
  for (let i = 0; i < eventi.length; i++) {
    testoDel += `${eventi[i].getTitle()}\t | \t${convertDate(eventi[i].getStartTime())}\n`;
  }
  
  const response2 = ui.alert(
    translate('modifyEvent.confirmYN'),
    testoDel,
    ui.ButtonSet.YES_NO
  );
  
  if (response2 !== ui.Button.YES) {
    ui.alert(translate('modifyEvent.opNOMessage'));
    return;
  }
  
  // Cancellazione eventi esistenti
  const calendar = CalendarApp.getCalendarById(myCalID()[0][0]);
  deleteExistingEvents(calendar, eventi);
  
  // Log revisione
  const oggi = new Date();
  const utenteEmail = Session.getEffectiveUser().getEmail();
  const eventID = matrice.length !== 0 ? buildEventId(matrice[0]) : 'ERRORE';
  addLogRevision(oggi, translate('modifyEvent.logCancMod'), eventID, utenteEmail, matrice);
  
  // Creazione nuovi eventi
  const listFinalEvents = buildEventsList(finalList);
  modifyEvents(first, last, listFinalEvents.toString(), 'updateDetailsEvent', activeRow);
  
  // Aggiornamento foglio
  markRowAsModified(sh, activeRow);
}

/**
 * Gestisce l'aggiornamento specifico di un evento
 */
function handleUpdateSpecificEvent(ui, note, finalList) {
  const response = ui.alert(
    translate('viewList.okEditSpecific'),
    translate('viewList.yesEditSpecific', { name: note }),
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const first = convertDateInputHtml(finalList[0][18]);
    const last = convertDateInputHtml(finalList[finalList.length - 1][19]);
    const listaFinale = logMatchingEvents(myCalID()[0][0], note, first, last);
    showFreeStructModifyEvent(first, last, listaFinale, note);
  }
}

/**
 * Cancella gli eventi esistenti dal calendario
 */
function deleteExistingEvents(calendar, eventi) {
  for (let i = 0; i < eventi.length; i++) {
    const event = calendar.getEventById(eventi[i].getId());
    if (event) {
      event.deleteEvent();
    }
  }
}

/**
 * Costruisce la lista degli eventi per la creazione/modifica
 */
function buildEventsList(finalList) {
  const listFinalEvents = [];
  
  for (let i = 0; i < finalList.length; i++) {
    const item = finalList[i];
    
    // Titolo evento
    const title = (item[1] === 'SI' ? 'Opz. ' + item[0] : item[0]) + ' ' + item[14];
    
    // Descrizione evento
    const descriptionParts = [
      item[12],
      `all=${item[8]}`,
      `feed=${item[25]}`,
      `code=${item[26]}`,
      `id=${item[13]}`,
      `typeEv=${item[23]}`,
      `org=${item[22]}`,
      `refCom=${item[11]}`,
      `refOp=${item[21]}`,
      `open=${item[2]}`
    ];
    
    // Aggiungi dettagli evento se tipo E
    if (item[14] === 'E') {
      descriptionParts.push(`vvf=${item[9]}`, `cri=${item[10]}`, `color=${item[24]}`);
    }
    
    // Aggiungi scadenza opzione se presente
    if (item[27] !== '') {
      descriptionParts.push(`opzExp=${item[27]}`);
    }
    
    const description = descriptionParts.join(' ');
    
    // Location
    const locationArray = [item[3], item[4], item[5], item[6], item[7]];
    const location = locationArray.join('| ');
    
    listFinalEvents.push([
      title,
      item[18],  // startTimeUTC
      item[19],  // endTimeUTC
      description,
      location,
      item[14],  // type
      item[21],  // refOp
      item[13]   // idEvent
    ]);
  }
  
  return listFinalEvents;
}

/**
 * Costruisce l'ID evento per il log
 */
function buildEventId(matriceItem) {
  const details = parseEventDetails(matriceItem[3]);
  const nome = parseEventString(matriceItem[0]).nome;
  return details.id !== '' ? `${details.id} |-> ${nome}` : nome;
}

/**
 * Marca la riga come modificata
 */
function markRowAsModified(sh, activeRow) {
  const numColumns = sh.getLastColumn();
  const rowValues = Array(numColumns).fill(translate('modifyEvent.logEditMod'));
  sh.getRange(activeRow, 1, 1, numColumns).setValues([rowValues]);
}
// FINE