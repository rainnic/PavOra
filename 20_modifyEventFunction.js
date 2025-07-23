/*
* Project Name: Pavora
* Copyright (c) 2025 Nicola Rainiero
*
* This software is released under the MIT License.
* Please refer to the LICENSE file for the full license text.
*/
function testShowFreeStructEdit() {
  //showFreeStructEdit('2024-07-30', '2024-08-10');
  var matrice = findEventsByKeyword(myCalID()[0][0], 'EnBBJCna', '2024-09-29', '2024-10-25', '');
  //showFreeStructModifyEvent('2024-09-29', '2024-10-25', matrice, 'EnBBJCna');
}

function testChangeTime() {
  //    -- 2024-08-26 -- 2024-09-20 -- xuvmrnfP -- 2024-08-23 -- 2024-09-22 -- 14,BPdx
  //  -- 2024-09-16 -- 2024-10-01 -- pbnpJDhK -- 2024-09-15 -- 2024-10-01 -- 7,G78,M,AC,ANP,PI,PN
  // changeTime(firstOriginal, lastOriginal, eventNameId, firstNew, lastNew, locationOriginal)
  changeTime('2024-09-16', '2024-10-01', 'pbnpJDhK', '2024-09-16', '2024-10-01', '7,G78,8,M,AC,ANP,PI,PN')
}

function testCopyNew() {
  // -- 2025-07-06 -- 2025-07-06 -- 5RIBZtlA -- 2025-11-16 -- 2025-11-16 -- SMP,SM1,SM2,foyerSM,Cdx
  //    -- 2024-08-26 -- 2024-09-20 -- xuvmrnfP -- 2024-08-23 -- 2024-09-22 -- 14,BPdx
  //  -- 2024-09-16 -- 2024-10-01 -- pbnpJDhK -- 2024-09-15 -- 2024-10-01 -- 7,G78,M,AC,ANP,PI,PN
  // changeTime(firstOriginal, lastOriginal, eventNameId, firstNew, lastNew, locationOriginal)
  // copyEventNew
  copyEventNew('2025-07-06', '2025-07-06', '5RIBZtlA', '2025-11-16', '2025-11-16', 'SMP,SM1,SM2,foyerSM,Cdx')
}

function showFreeStructEdit(first, last) {
  try {
    firstDate = text2monthDays(first);
    lastDate = text2monthDays(last);
    var eventi = events2Array(convertDateBar(firstDate[1]), convertDateBar(lastDate[1]), categories()[0][0], "");
    if (eventi.length != 0) {
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

    structures.push(first);
    structures.push(last);

    createUserSheet();
    ClearAll();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    updateTimeUser();
    showMonths(first, last, '');
    if (eventi.length != 0) {
      sh.getRange(2, 1).setValue(translate('modifyEvent.usedStruct')).setFontSize(12);
      sh.getRange(2, 1).setNote(usedLocations.filter(onlyUnique).toString()).setNumberFormat("@").setFontSize(10).setWrap(true);
    } else {
      sh.getRange(2, 1).setNote(translate('modifyEvent.usedStruct')).setFontSize(10).setWrap(true);
    }

    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '2A_modifyEventPage', translate('modifyEvent.editDelEvent')));


  } catch (error) {
    SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
  }
}

function populateCategories(array, numberCategory) {
  result = [];
  for (let i = 0; i < array.length; i++) {
    if (array[i][7] == numberCategory) {
      result.push(array[i][0]);
    }
  }
  return result
}

function categorizeLocation(location, array) {
  // Definizione delle categorie
  const categories = {
    //quartiere: ['1', '5'],
    quartiere: populateCategories(array, 1),
    //congress: ['SG'],
    congress: populateCategories(array, 2),
    //cancelli: ['T'],
    ingressi: populateCategories(array, 3),
    // aree: ['AN']
    aree: populateCategories(array, 4),
    // parcheggi
    parcheggi: populateCategories(array, 5)
  };

  // Inizializzazione delle variabili di categoria
  let locationQuartiere = [];
  let locationCongress = [];
  let locationIngressi = [];
  let locationAree = [];
  let locationParcheggi = [];

  // Loop attraverso ogni elemento in location
  location.forEach(loc => {
    if (categories.quartiere.includes(loc)) {
      locationQuartiere.push(loc);
    } else if (categories.congress.includes(loc)) {
      locationCongress.push(loc);
    } else if (categories.ingressi.includes(loc)) {
      locationIngressi.push(loc);
    } else if (categories.aree.includes(loc)) {
      locationAree.push(loc);
    } else if (categories.parcheggi.includes(loc)) {
      locationParcheggi.push(loc);
    }
  });

  // Restituisci un oggetto con le nuove variabili
  return {
    locationQuartiere: locationQuartiere,
    locationCongress: locationCongress,
    locationIngressi: locationIngressi,
    locationAree: locationAree,
    locationParcheggi: locationParcheggi
  };
}
// USAGE:
// var data = strutture();
// let location = ['1', 'SG', 'T', 'B', '5', 'AN'];
// let categorizedLocations = categorizeLocation(location, data);
// let locQuartiere = categorizedLocations.locationQuartiere;

// La data in formato stringa
function extractionDateTime(timeStamp) {
  // Creazione di un oggetto Date a partire dalla stringa
  const date = new Date(timeStamp);

  // Estrazione dell'anno, mese e giorno per la prima variabile
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0'); // I mesi in JavaScript sono 0-indexed, quindi aggiungiamo 1
  const day = String(date.getDate()).padStart(2, '0');

  // Formattazione della prima variabile
  const formattedDate = `${year}-${month}-${day}`;

  // Estrazione delle ore e minuti per la seconda variabile
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');

  // Formattazione della seconda variabile
  const formattedTime = `${hours}:${minutes}`;

  // Restituisci un oggetto con le nuove variabili
  return {
    date: formattedDate,
    time: formattedTime
  };
}
//const dateString = 'Thu Jun 20 2024 08:00:00 GMT+0200 (Central European Summer Time)';
//console.log(extractionDateTime(dateString).date); // Output: "2024-06-20"
//console.log(extractionDateTime(dateString).time); // Output: "08:00"

// Search event with a keyword
function findEventsByKeyword(calendarId, keyword, first, last, what) {
  // Ottieni il calendario specificato
  var calendar = CalendarApp.getCalendarById(calendarId);
  // Crea un array per memorizzare gli eventi trovati
  var matchingEvents = [];

  // Inizializza una data iniziale molto antica
  var startDate, endDate;

  if (what === 'hall' && first) {
    startDate = new Date(first);
    startDate.setHours(0, 0, 0, 0);
    endDate = new Date(first);
    endDate.setHours(23, 59, 59, 999);

  } else if (!first) {
    const today = new Date();
    startDate = new Date(today.getFullYear() - 1, 0, 1);
    endDate = new Date(today.getFullYear() + 6, 0, 1);

  } else {
    const f = incrDay(new Date(first), -incrementDay() * 6);
    const l = incrDay(new Date(last), incrementDay() * 6);
    startDate = text2monthDays(f)[1];
    endDate = text2monthDays(l)[1];
  }

  // Ottieni tutti gli eventi in questo intervallo di tempo ampio
  var events = calendar.getEvents(startDate, endDate);
  // Loop attraverso tutti gli eventi e controlla se contengono la parola chiave
  for (let i = 0; i < events.length; i++) {
    var event = events[i];
    if ((event.getDescription().includes(keyword)) || (event.getTitle().includes(keyword))) {
      matchingEvents.push(event);
    }
  }

  // Restituisci gli eventi trovati
  return matchingEvents;
}

function logMatchingEvents(calendarId, keyword, first, last, locationOriginal, what) {
  var matchingEvents = findEventsByKeyword(calendarId, keyword.replace(/[\r\n]+/g, '').replace(/^\(.*?\)\s*/, ''), first, last, what);
  // Log gli eventi trovati
  var eventsList = [];
  for (let i = 0; i < matchingEvents.length; i++) {
    var event = matchingEvents[i];
    type = parseEventString(event.getTitle()).type;
    eventsList.push([event.getTitle(), event.getStartTime(), event.getEndTime(), event.getDescription(), fixLocationsEvents(event.getLocation()), type]);
  }
  // trovo le variabili che sono costanti per tutti gli eventi
  var data = strutture();
  if ((findKey('E', eventsList, 5) >= 0) || (findKey('L', eventsList, 5) >= 0)) {
    if (findKey('E', eventsList, 5) > -1) {
      index = findKey('E', eventsList, 5);
    } else {
      index = findKey('L', eventsList, 5);
    }
    var name = parseEventString(eventsList[index][0]).nome;
    var opz = parseEventString(eventsList[index][0]).opz;
    var public = parseEventDetails(eventsList[index][3]).open;
    var allestitore = parseEventDetails(eventsList[index][3]).all;
    var feed = parseEventDetails(eventsList[index][3]).feed;
    var code = parseEventDetails(eventsList[index][3]).code;
    var opzExp = parseEventDetails(eventsList[index][3]).opzExp;
    var vvf = parseEventDetails(eventsList[index][3]).vvf;
    var cri = parseEventDetails(eventsList[index][3]).cri;
    var color = parseEventDetails(eventsList[index][3]).color;
    var org = parseEventDetails(eventsList[index][3]).org;
    var refCom = parseEventDetails(eventsList[index][3]).refCom;
    var refOp = parseEventDetails(eventsList[index][3]).refOp;
    var typeEv = parseEventDetails(eventsList[index][3]).typeEv;
    var note = parseEventDetails(eventsList[index][3]).descrizione;
    var idEvent = parseEventDetails(eventsList[index][3]).id;
    if (locationOriginal) {
      var location = '';
    } else {
      var location = eventsList[index][4];
    }
    var categorizedLocations = categorizeLocation(eventsList[index][4].split(',').map(item => item.trim()), data);
    var quartiere = categorizedLocations.locationQuartiere;
    var congress = categorizedLocations.locationCongress;
    var ingressi = categorizedLocations.locationIngressi;
    var aree = categorizedLocations.locationAree;
    var parcheggi = categorizedLocations.locationParcheggi;
  } else {
    index = 0;
    var name = parseEventString(eventsList[index][0]).nome;
    var opz = parseEventString(eventsList[index][0]).opz;
    var public = parseEventDetails(eventsList[index][3]).open;
    var allestitore = parseEventDetails(eventsList[index][3]).all;
    var feed = parseEventDetails(eventsList[index][3]).feed;
    var code = parseEventDetails(eventsList[index][3]).code;
    var opzExp = parseEventDetails(eventsList[index][3]).opzExp;
    var vvf = parseEventDetails(eventsList[index][3]).vvf;
    var cri = parseEventDetails(eventsList[index][3]).cri;
    var color = parseEventDetails(eventsList[index][3]).color;
    var refCom = parseEventDetails(eventsList[index][3]).refCom;
    var refOp = parseEventDetails(eventsList[index][3]).refOp;
    var typeEv = parseEventDetails(eventsList[index][3]).typeEv;
    var org = parseEventDetails(eventsList[index][3]).org;
    var note = parseEventDetails(eventsList[index][3]).descrizione;
    var idEvent = parseEventDetails(eventsList[index][3]).id;
    var categorizedLocations = categorizeLocation(eventsList[index][4].split(',').map(item => item.trim()), data);
    var quartiere = categorizedLocations.locationQuartiere;
    var congress = categorizedLocations.locationCongress;
    var ingressi = categorizedLocations.locationIngressi;
    var aree = categorizedLocations.locationAree;
    var parcheggi = categorizedLocations.locationParcheggi;
  }

  // Create the final eventsList
  var finalList = [];
  for (let i = 0; i < eventsList.length; i++) {
    type = eventsList[i][5];
    data = extractionDateTime(eventsList[i][1]).date;
    startTime = extractionDateTime(eventsList[i][1]).time;
    endTime = extractionDateTime(eventsList[i][2]).time;
    finalList.push([name, opz, public, quartiere, congress, ingressi, aree, parcheggi, allestitore, vvf, cri, refCom, note, idEvent, type, data, startTime, endTime, eventsList[i][1], eventsList[i][2], location, refOp, org, typeEv, color, feed, code, opzExp]);
  }
  return finalList
}

// transform title e description in variable
function parseEventString(eventString) {
  let opz = 'NO';
  let nome = '';
  let type = '';

  // Controlla se la stringa contiene "Opz."
  if (eventString.startsWith('Opz.')) {
    opz = 'SI';
    // Rimuove "Opz. " dalla stringa
    eventString = eventString.slice(5);
  }

  if (eventString.startsWith('Off.')) {
    opz = 'OFF';
    // Rimuove "Off. " dalla stringa
    eventString = eventString.slice(5);
  }

  // Usa un'espressione regolare per trovare la lettera maiuscola finale
  let match = eventString.match(/^(.*)\s([A-Z])$/);

  if (match) {
    nome = match[1]; // Estrai il nome dell'evento
    type = match[2]; // Estrai la lettera maiuscola finale
  } else {
    // Nel caso in cui la stringa non corrisponda al formato atteso metti una 'E' per evento alla fine!
    nome = eventString;
    type = 'E';
  }

  return {
    opz: opz,
    nome: nome,
    type: type
  };
}

// Extract description
function parseEventDetails(eventString) {
  let descrizione = '';
  let all = '';
  let feed = '';
  let code = '';
  let id = '';
  let refCom = '';
  let refOp = '';
  let typeEv = '';
  let open = '';
  let vvf = '';
  let cri = '';
  let org = '';
  let color = '';
  let opzExp = '';

  // Usa un'espressione regolare per trovare i parametri e le loro rispettive valori
  let typeEvMatch = eventString.match(/\btypeEv=([^\s]+)/);
  if (typeEvMatch) {
    typeEv = typeEvMatch[1]; // Estrai 'typeEv'
  } else { typeEv = ''; }

  let descrizioneMatch = eventString.match(/^(.*?)\s(?=\w+=|$)/);
  if (descrizioneMatch) {
    descrizione = descrizioneMatch[1].trim(); // Estrai la descrizione
  } else if (eventString.match(/^(.*?\[.*?\])/)) {
    descrizione = eventString.match(/^(.*?\[.*?\])/)[1].trim(); // Estrai per le slide giornaliere
  } else { descrizione = ''; }

  let allMatch = eventString.match(/\ball=([^\s]+)/);
  if (allMatch) {
    all = allMatch[1]; // Estrai 'all'
  } else { all = ''; }

  let feedMatch = eventString.match(/\bfeed=([^\s]+)/);
  if (feedMatch) {
    feed = feedMatch[1]; // Estrai 'all'
  } else { feed = ''; }

  let codeMatch = eventString.match(/\bcode=([^\s]+)/);
  if (codeMatch) {
    code = codeMatch[1]; // Estrai 'all'
  } else { code = ''; }

  let idMatch = eventString.match(/\bid=([^\s]+)/);
  if (idMatch) {
    id = idMatch[1]; // Estrai 'id'
  } else { id = ''; }

  let orgMatch = eventString.match(/\borg=([^]+?)\srefCom=/); // /org=([^]+?)\srefCom=/;
  if (orgMatch) {
    org = orgMatch[1]; // Estrai 'org'
  } else { org = ''; }

  let refComMatch = eventString.match(/\brefCom=([^\s]+)/);
  if (refComMatch) {
    refCom = refComMatch[1]; // Estrai 'refCom'
  } else { refCom = ''; }

  let refOpMatch = eventString.match(/\brefOp=([^\s]+)/);
  if (refOpMatch) {
    refOp = refOpMatch[1]; // Estrai 'ref'
  } else { refOp = ''; }

  let openMatch = eventString.match(/\bopen=([^\s]+)/);
  if (openMatch) {
    open = openMatch[1]; // Estrai 'open'
  } else { open = ''; }

  let vvfMatch = eventString.match(/\bvvf=([^\s]+)/);
  if (vvfMatch) {
    vvf = vvfMatch[1]; // Estrai 'vvf'
  } else { vvf = ''; }

  let criMatch = eventString.match(/\bcri=([^\s]+)/);
  if (criMatch) {
    cri = criMatch[1]; // Estrai 'cri'
  } else { cri = ''; }

  let colorMatch = eventString.match(/\bcolor=([^\s]+)/);
  if (colorMatch) {
    color = colorMatch[1]; // Estrai 'color'
  } else { color = ''; }

  let opzExpMatch = eventString.match(/\bopzExp=([^\s]+)/);
  if (opzExpMatch) {
    opzExp = opzExpMatch[1]; // Estrai 'color'
  } else { opzExp = ''; }

  return {
    descrizione: descrizione,
    all: all,
    feed: feed,
    code: code,
    id: id,
    org: org,
    refCom: refCom,
    refOp: refOp,
    typeEv: typeEv,
    open: open,
    vvf: vvf,
    cri: cri,
    color: color,
    opzExp: opzExp
  };
}

// Regex migliorate
const regexId = /id=([^\s">]+)/; // Cattura l'ID dopo "id=" evitando spazi e caratteri HTML
const regexName = /idNome=(.+?)\s\(/; // Cattura il valore dopo "idNome=" e prima di " ("

// Funzione di estrazione con regex
function extractRegex(regex, string) {
  if (!string) return ''; // Evita errori su stringhe vuote o non definite

  const match = string.match(regex);
  return match ? match[1] : ''; // Ritorna il valore estratto o stringa vuota se non trova nulla
}

function getCellNote(first, last, id) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const cell = sh.getActiveCell(); // Cella selezionata
    const noteComplete = cell.getNote(); // Nota della cella selezionata
    let note = id || ""; // Usa l'ID passato oppure inizializza

    if (!id && noteComplete.length !== 0) {
      const extractedId = extractRegex(regexId, noteComplete);
      if (extractedId && extractedId.length === 8) {
        note = extractedId;
      } else {
        const match = noteComplete.match(/] (.*?) \(/);
        note = match ? match[1] : '';
      }
    }

    if (note) {
      sh.getRange(3, 1).setValue(note).setFontSize(10);

      let finalList;

      if (id || (note.length === 8 && /^[A-Za-z0-9]{8}$/.test(note))) {
        finalList = logMatchingEvents(myCalID()[0][0], note, first, last);
      } else {

        // Calcola intervallo date dinamico (-5 anni, +5 anni)
        const today = new Date();
        const start = new Date(today);
        start.setFullYear(today.getFullYear() - 5);
        const end = new Date(today);
        end.setFullYear(today.getFullYear() + 5);

        const startFormatted = Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM-dd");
        const endFormatted = Utilities.formatDate(end, Session.getScriptTimeZone(), "yyyy-MM-dd");

        finalList = logMatchingEvents(myCalID()[0][0], note, startFormatted, endFormatted);
      }

      // Converti date da finalList
      if (finalList && finalList.length > 0) {
        first = convertDateInputHtml(finalList[0][18]);
        last = convertDateInputHtml(finalList[finalList.length - 1][19]);
        showFreeStructModifyEvent(first, last, finalList, note);
      } else {
        SpreadsheetApp.getUi().alert(translate('modifyEvent.noEventFound', { note: note }));
      }
    }

    return note || translate('modifyEvent.emptyCell');
  } catch (error) {
    SpreadsheetApp.getUi().alert(translate('modifyEvent.noEventDate') + ' (' + first + '-' + last + '). \n' + error);
    return translate('modifyEvent.noEventDate');
  }
}

function changeTime(firstOriginal, lastOriginal, eventNameId, firstNew, lastNew, locationOriginal) {
  var finalList = logMatchingEvents(myCalID()[0][0], eventNameId, firstOriginal, lastOriginal, locationOriginal);
  first = convertDateInputHtml(finalList[0][18]);
  first = firstNew;
  last = lastNew;
  showFreeStructModifyEvent(first, last, finalList, eventNameId);
}

function copyEventNew(firstOriginal, lastOriginal, eventNameId, firstNew, lastNew, locationOriginal) {
  var finalList = logMatchingEvents(myCalID()[0][0], eventNameId, firstOriginal, lastOriginal, locationOriginal);
  first = convertDateInputHtml(finalList[0][18]);
  first = firstNew;
  last = lastNew;
  // Create a new idEvent in the finalList array
  var eventID = randomID(8);
  for (let i = 0; i < finalList.length; i += 1) {
    finalList[i][13] = eventID;
  }
  showFreeStructModifyEvent(first, last, finalList, eventNameId);
}

// Function to showfreeStrruct and event to modify or delete
function showFreeStructModifyEvent(first, last, array, eventNameId) {
  try {
    firstDate = text2monthDays(first);
    lastDate = text2monthDays(last);
    if (eventNameId != undefined) {
      var keyword = '-' + eventNameId;
    } else {
      var keyword = '';
    }
    var eventi = events2Array(convertDateBar(firstDate[1]), convertDateBar(lastDate[1]), categories()[0][0], keyword);
    if (eventi.length != 0) {
      //Logger.log(eventi[0][10]);
      var usedLocations = []
      for (let i = 0; i < eventi.length; i += 1) {
        if (!excludeAll()) {
          const isOptionated = optionated().indexOf(eventi[i][8].substring(0, 4)) < 0;
          if (eventi[i][10].length != 0 && (includeOptionated() || isOptionated)) {
            loc2array = eventi[i][10]; // .split(",");
            for (let j = 0; j < loc2array.length; j += 1) {
              usedLocations.push(loc2array[j]);
              if (findKey(String(loc2array[j]), strutture(), 0) >= 0) {
                var relationship = strutture()[findKey(loc2array[j], strutture(), 0)][10].split(',').map(s => s.trim()); // strutture con un grado di parentela da string a array
              }
              if (loc2array != 0) {
                for (let k = 0; k < relationship.length; k += 1) {
                  usedLocations.push(relationship[k]);
                }
              }

            }

          }
        }
      }
      var usedUnique = usedLocations.filter(onlyUnique);
      //var usedUnique = usedUnique.filter(element => !array[0][20].includes(element));
      var colValue = array[0][20];
      // Se colValue è null o vuoto, usa l'array alla posizione 4
      if (!colValue && Array.isArray(array[0][4])) {
        colValue = array[0][4].join(",");
      }
      // Se colValue ora è una stringa valida, procedi con il filtro
      if (typeof colValue === "string") {
        var colArray = colValue.split(",").map(s => s.trim());
        usedUnique = usedUnique.filter(element => !colArray.includes(element));
      } else {
        Logger.log("Nessun valore valido per il filtro: colValue = " + colValue);
      }      
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

    SpreadsheetApp.getUi()
      .showSidebar(doGet(structures, '2B_modifyEventPage', translate('modifyEvent.editDelEvent')));

    updateTimeUser();

  } catch (error) {
    // Mostra un messaggio tramite ui.alert
    SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
  }

}

// Function to delete Events!!!
function deleteEvents(eventId, first, last, what, activeRow) {
  try {
    createUserSheet();
    updateTimeUser();
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var eventi = findEventsByKeyword(myCalID()[0][0], eventId, first, last, what);
    //Logger.log(eventi);
    var testo = translate('modifyEvent.warEventDel', { eventId: eventId });
    var matrice = [];
    for (let i = 0; i < eventi.length; i += 1) {
      testo = testo + eventi[i].getTitle() + '\t | \t' + convertDate(eventi[i].getStartTime()) + '\n';
      matrice.push([eventi[i].getTitle(), eventi[i].getStartTime(), eventi[i].getEndTime(), eventi[i].getDescription(), eventi[i].getLocation()]);
    }
    //sh.setActiveRange(sh.getRange('A1'));
    var calendar = CalendarApp.getCalendarById(myCalID()[0][0]);
    var response = ui.alert(translate('modifyEvent.confirmYN'), testo, ui.ButtonSet.YES_NO);

    // Gestisci la risposta dell'utente
    if ((response == ui.Button.YES) && (checkUserWritePermission(myCalID()[0][0]) == true)) {
      //cancellaEventi(eventi);
      // Itera su ciascun ID evento e prova a cancellarlo
      for (let i = 0; i < eventi.length; i += 1) {
        var event = calendar.getEventById(eventi[i].getId());
        if (event) {
          event.deleteEvent();
          //Logger.log('Evento con ID ' + eventId + ' cancellato.');
        } else {
          //Logger.log('Evento con ID ' + eventId + ' non trovato.');
        }
      }
      oggi = new Date();
      utenteEmail = Session.getEffectiveUser().getEmail();
      var eventID = (parseEventDetails(matrice[0][3]).id != '') ? parseEventDetails(matrice[0][3]).id + ' |-> ' + parseEventString(matrice[0][0]).nome : parseEventString(matrice[0][0]).nome;
      if (typeof what !== 'undefined') {
        if (what == 'hall') {
          addLogRevision(oggi, translate('modifyEvent.delRoomsOK'), eventID, utenteEmail, matrice);
          createDailyScheduleFromCalendar(first, 60, '');
          specialDailyEvent();
        } else if (typeof activeRow != 'undefined') {
          addLogRevision(oggi, translate('modifyEvent.delOK'), eventID, utenteEmail, matrice);

          // Ottieni il numero di colonne del foglio
          var numColumns = sh.getLastColumn();

          // Crea un array con il testo "cancellato" per ogni cella nella riga
          var rowValues = Array(numColumns).fill(translate('modifyEvent.delOK'));

          // Scrivi il testo "cancellato" su tutte le celle della riga
          sh.getRange(activeRow, 1, 1, numColumns).setValues([rowValues]);
        }
      } else {
        addLogRevision(oggi, translate('modifyEvent.delOK'), eventID, utenteEmail, matrice);
        specialEvent();
        const prefs = getUserBrowserSettings();
        showMonths(prefs.first, prefs.last, prefs.selectedStruct, prefs.keyword);
      }
      ui.alert(translate('modifyEvent.delOKMessage'));
    } else if ((response == ui.Button.YES) && (checkUserWritePermission(myCalID()[0][0]) == false)) {
      ui.alert(translate('modifyEvent.waitSomeTime'));
    } else {
      if (what != null && what === 'hall') {
        specialDailyEvent();
        createDailyScheduleFromCalendar(first, 60, '');
        ui.alert(translate('modifyEvent.delNOMessage'));
      } else {
        specialEvent();
        const prefs = getUserBrowserSettings();
        showMonths(prefs.first, prefs.last, prefs.selectedStruct, prefs.keyword);
        ui.alert(translate('modifyEvent.delNOMessage'));
      }
    }
  } catch (error) {
    // Mostra un messaggio tramite ui.alert
    SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
  }
}

// Delete events before modify them
function deleteEventsNoConfirm(eventId, first, last, what) {
  createUserSheet();
  updateTimeUser();
  var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var eventi = findEventsByKeyword(myCalID()[0][0], eventId, first, last, what);
  var testo = '';
  var matrice = [];
  for (let i = 0; i < eventi.length; i += 1) {
    testo = testo + eventi[i].getTitle() + '\t' + convertDate(eventi[i].getStartTime()) + '\n';
    matrice.push([eventi[i].getTitle(), eventi[i].getStartTime(), eventi[i].getEndTime(), eventi[i].getLocation(), eventi[i].getDescription()]);
  }
  var calendar = CalendarApp.getCalendarById(myCalID()[0][0]);
  var response = ui.alert(translate('modifyEvent.confModTitle'), testo, ui.ButtonSet.YES_NO);

  // Gestisci la risposta dell'utente
  if ((response == ui.Button.YES) && (checkUserWritePermission(myCalID()[0][0]) == true)) {
    for (let i = 0; i < eventi.length; i += 1) {
      var event = calendar.getEventById(eventi[i].getId());
      if (event) {
        event.deleteEvent();
        //Logger.log('Evento con ID ' + eventId + ' cancellato.');
      } else {
        //Logger.log('Evento con ID ' + eventId + ' non trovato.');
      }
    }
    oggi = new Date();
    utenteEmail = Session.getEffectiveUser().getEmail();
    var eventID = (parseEventDetails(matrice[0][3]).id != '') ? parseEventDetails(matrice[0][3]).id + ' |-> ' + parseEventString(matrice[0][0]).nome : parseEventString(matrice[0][0]).nome;
    //addLogRevision(oggi, "ADMIN NUOVO", eventID, utenteEmail, matrix);
    addLogRevision(oggi, translate('modifyEvent.logCancMod'), eventID, utenteEmail, matrice);
  } else if ((response == ui.Button.YES) && (checkUserWritePermission(myCalID()[0][0]) == false)) {
    ui.alert(translate('modifyEvent.waitSomeTime'));
  } else {
    ui.alert(translate('modifyEvent.opNOMessage'));
  }
  //showMonths(first, last, eventi[0].getLocation());
  //viewCalendar(); // to refresh dialog window
  //ui.alert('Eventi cancellati con successo.');
  //} 
}


// Delete events and then modify them
function deleteThenModifyEvents(eventID, first, last, outputEvents, what) {
  try {
    createUserSheet();
    updateTimeUser();
    var ui = SpreadsheetApp.getUi(); // Se utilizzi Documenti Google, usa DocumentApp.getUi()
    var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var eventi = findEventsByKeyword(myCalID()[0][0], eventID, first, last, what);
    var matrice = [];
    var testo = '';
    for (let i = 0; i < eventi.length; i += 1) {
      testo = testo + eventi[i].getTitle() + '\t' + convertDate(eventi[i].getStartTime()) + '\n';
      matrice.push([eventi[i].getTitle(), eventi[i].getStartTime(), eventi[i].getEndTime(), eventi[i].getDescription(), eventi[i].getLocation()]);
    }
    //sh.setActiveRange(sh.getRange('A1'));
    var calendar = CalendarApp.getCalendarById(myCalID()[0][0]);
    var response = ui.alert(translate('modifyEvent.confModTitle'), testo, ui.ButtonSet.YES_NO);
    //var response = ui.alert('Conferma Cancellazione', testo, ui.ButtonSet.YES_NO);

    // Gestisci la risposta dell'utente
    if ((response == ui.Button.YES) && (checkUserWritePermission(myCalID()[0][0]) == true)) {
      //cancellaEventi(eventi);
      // Itera su ciascun ID evento e prova a cancellarlo
      if (eventi.length > 0) {
        for (let i = 0; i < eventi.length; i += 1) {
          var event = calendar.getEventById(eventi[i].getId());
          if (event) {
            event.deleteEvent();
            //Logger.log('Evento con ID ' + eventID + ' cancellato.');
          } else {
            //Logger.log('Evento con ID ' + eventID + ' non trovato.');
          }
        }
      }
      oggi = new Date();
      utenteEmail = Session.getEffectiveUser().getEmail();
      if (matrice.length != 0) {
        var eventID = (parseEventDetails(matrice[0][3]).id != '') ? parseEventDetails(matrice[0][3]).id + ' |-> ' + parseEventString(matrice[0][0]).nome : parseEventString(matrice[0][0]).nome;
        //addLogRevision(oggi, "ADMIN NUOVO", eventID, utenteEmail, matrix);
      } else {
        var eventID = 'ERROR';
      }

      if (typeof what !== 'undefined') {
        if (what == 'hall') {
          addLogRevision(oggi, translate('modifyEvent.logCancModOne'), eventID, utenteEmail, matrice);
          modifyEvents(first, last, outputEvents, what);
        }
      } else {
        addLogRevision(oggi, translate('modifyEvent.logCancMod'), eventID, utenteEmail, matrice);
        modifyEvents(first, last, outputEvents);
      }
    } else if ((response == ui.Button.YES) && (checkUserWritePermission(myCalID()[0][0]) == false)) {
      ui.alert(translate('modifyEvent.waitSomeTime'));
    } else {
      ui.alert(translate('modifyEvent.opNOMessage'));
    }


  } catch (error) {
    // Mostra un messaggio tramite ui.alert
    SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
  }
}


// Modify events!
function modifyEvents(first, last, array, what, activeRow) {
  createUserSheet();
  if (typeof activeRow == 'undefined') {
  }
  updateTimeUser();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var matrix = stringToMatrix(array, 8)
  var mycalID = myCalID()[0][0];             // ID of the first calendar
  array2Events(matrix, mycalID);
  if (typeof activeRow == 'undefined') {
    if (findKey('E', matrix, 5) >= 0) {
      //showMonths(first, last, matrix[findKey('E', matrix, 5)][4]);
      const prefs = getUserBrowserSettings();
      showMonths(prefs.first, prefs.last, prefs.selectedStruct, prefs.keyword);
    } else {
      //showMonths(first, last, matrix[0][4]);
      const prefs = getUserBrowserSettings();
      showMonths(prefs.first, prefs.last, prefs.selectedStruct, prefs.keyword);
    }
  }
  oggi = new Date();
  utenteEmail = Session.getEffectiveUser().getEmail();
  var eventID = (parseEventDetails(matrix[0][3]).id != '') ? parseEventDetails(matrix[0][3]).id + ' |-> ' + parseEventString(matrix[0][0]).nome : parseEventString(matrix[0][0]).nome;
  //addLogRevision(oggi, "ADMIN NUOVO", eventID, utenteEmail, matrix);  
  addLogRevision(oggi, translate('modifyEvent.logEditMod'), eventID, utenteEmail, matrix);
  if (typeof what !== 'undefined') {
    if (what == 'hall') {
      manageSmallRoom();
      createDailyScheduleFromCalendar(first, 60, '', '', 'H24');
    }
  } else if (what == 'updateDetailsEvent') {
    // do nothing
  } else {
    //viewCalendar(); // to refresh dialog window
    //completeMenu(); // to reload the menu
    specialEvent();
  }
}
