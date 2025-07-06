/*
* Project Name: Pavora
* Copyright (c) 2025 Nicola Rainiero
*
* This software is released under the MIT License.
* Please refer to the LICENSE file for the full license text.
*/

function testEvents2arrays() {
  var start = Date.now();
  events = events2Array('01/01/2025', '31/01/2025', categories()[0][0], '');
  /*
  for (let i = 0; i < eventiFinali.length; i++) {
    //Logger.log('Per l\'evento '+eventiFinali[i][2]+ ' || L\'ultima data è '+eventiFinali[i][11]);
  }
  */
  var end = Date.now();
  Logger.log("Tempo di esecuzione funzione: " + (end - start) / 1000 + " secondi");
  Logger.log(JSON.stringify(events));
}

function addLastStartingTime(matrix) {
  const lastStartingTime = {};

  // Prima scansione: Trova l'ultimo starting time per ogni numero progressivo
  matrix.forEach(row => {
    const progressiveNumber = row[7];
    lastStartingTime[progressiveNumber] = row[0];
  });

  // Seconda scansione: Aggiungi lo starting time dell'ultima riga per ogni riga con lo stesso numero progressivo
  matrix.forEach(row => {
    const progressiveNumber = row[7];
    row.push(lastStartingTime[progressiveNumber]);
  });

  return matrix;
}

function removeDuplicatesAndExpand(array) {
  const expandedArray = array.flatMap(item => item.split(',').map(s => s.trim()));
  return [...new Set(expandedArray)];
}

function fixLocationsEvents(stringLocation) {
  const location2change = ['SM', 'SG', 'SMU', 'CC'];
  const array = stringLocation.split(',').map(s => s.trim());
  const fixedLocationArray = array.map(location => {
    const index = findKey(location, strutture(), 0);
    if (index > -1 && location2change.includes(location) && strutture()[index][10].length !== 0) {
      return strutture()[index][10];
    }
    return location;
  });

  return removeDuplicatesAndExpand(fixedLocationArray).join(',');
}

// Retrieve calendar events in a defined period, the category is associated to a various calendar ID
// In my case I have two different calendars:
// one for the first category of events
// another for the last one
function events2Array(startDate, finishDate, category, keyword) {
  nColore = 0;

  // Funzione helper per ottenere gli eventi del calendario
  function getCalendarEvents(start, end, keyword) {
    const mycal = myCalID()[0][0]; // Sempre il primo calendario
    const cal = CalendarApp.getCalendarById(mycal);
    return keyword != null ?
      cal.getEvents(new Date(start), new Date(end), { search: keyword }) :
      cal.getEvents(new Date(start), new Date(end));
  }

  // Gestione date
  let fromDate, toDate;
  if (startDate.indexOf('-') > -1) {
    fromDate = startDate;
    toDate = finishDate;
  } else {
    const startParts = startDate.split("/");
    const endParts = finishDate.split("/");
    fromDate = new Date(+startParts[2], startParts[1] - 1, +startParts[0]);
    toDate = new Date(+endParts[2], endParts[1] - 1, +endParts[0]);
    toDate = new Date(toDate.getTime() + 24 * 3600000); // Aggiungi un giorno
  }

  // Ottieni eventi
  const events = getCalendarEvents(fromDate, toDate, keyword);

  // Se non ci sono eventi, ritorna array vuoto
  if (events.length === 0) {
    return addLastStartingTime([]);
  }

  // Processa eventi iniziali
  const listaEventi = [];
  for (const event of events) {
    const row = [
      event.getStartTime(),
      event.getTitle().replace(/\s*$/, ''),
      event.getStartTime(),
      event.getEndTime(),
      event.getDescription(),
      fixLocationsEvents(event.getLocation()),
      '',
      '',
      formatEmail(String(event.getCreators()))
    ];

    if (category === categories()[0][0]) {
      if (row[1].substring(row[1].length - 2, row[1].length - 1) !== ' ') {
        row[1] = row[1] + ' ' + categories()[0][1];
      }
      listaEventi.push(row);
    } else if (category === categories()[3][0]) {
      if (parseEventString(row[1]).type === categories()[3][1]) {
        listaEventi.push(row);
      }
    }
  }

  // Processa acronimi
  for (let i = 0; i < listaEventi.length; i++) {
    listaEventi[i][7] = listaEventi[i][1].substring(0, listaEventi[i][1].length - 1).replace(/\s+/g, '');
  }

  // Assegna colori unici
  const holderEvents = [];
  let color = 0;

  if (listaEventi.length > 0) {
    holderEvents[0] = listaEventi[0][7];
    listaEventi[0][6] = color;

    for (let i = 0; i < listaEventi.length; i++) {
      const eventIndex = holderEvents.indexOf(listaEventi[i][7]);
      if (eventIndex > -1) {
        listaEventi[i][6] = nColore + eventIndex;
      } else {
        color++;
        holderEvents[color] = listaEventi[i][7];
        listaEventi[i][6] = nColore + color;
      }
    }
  }

  // Crea array finale
  const finaleEventi = makeArray(listaEventi.length, 10);
  for (let i = 0; i < listaEventi.length; i++) {
    const parsedEvent = parseEventString(listaEventi[i][1]);

    finaleEventi[i][0] = listaEventi[i][0];                  // starting time
    finaleEventi[i][1] = parsedEvent.type;                   // category
    if (parsedEvent.opz == 'NO') {
      finaleEventi[i][2] = parsedEvent.nome;
    } else if (parsedEvent.opz == 'OFF') {
      finaleEventi[i][2] = 'Off. ' + parsedEvent.nome;                   // title without category      
    } else {                   // title without category  
      finaleEventi[i][2] = 'Opz. ' + parsedEvent.nome;                   // title without category
    }
    finaleEventi[i][3] = RemoveAccents(listaEventi[i][1])    // title without spaces
      .replace(/[^\w\s]/gi, '')
      .replace(/\s+/g, '');
    finaleEventi[i][4] = selectFill(finaleEventi[i][1]);     // hatch type
    finaleEventi[i][5] = convertHour(listaEventi[i][2]);     // start time
    finaleEventi[i][6] = convertHour(listaEventi[i][3]);     // finish time
    finaleEventi[i][7] = listaEventi[i][6];                  // unique number

    // full title
    const description = listaEventi[i][4] || ' ';
    finaleEventi[i][8] = `${finaleEventi[i][2]} (${selectType(finaleEventi[i][1]).substring(0, 1)}) [${finaleEventi[i][5]}-${finaleEventi[i][6]}] ${description} (${listaEventi[i][8]})`;

    finaleEventi[i][9] = listaEventi[i][3];                  // small title

    // locations array
    let locations = listaEventi[i][5] || '0';
    locations = locations.replaceAll(/[;.:]/g, ',');
    finaleEventi[i][10] = string2array(locations);
  }

  return addLastStartingTime(finaleEventi);
}