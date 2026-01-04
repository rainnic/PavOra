/*
* Project Name: Pavora
* Copyright (c) 2025 Nicola Rainiero
*
* This software is released under the MIT License.
* Please refer to the LICENSE file for the full license text.
*/
function testShowMonths() {
  //showMonths('2024-09-25', '2024-09-28', '');
  //showMonths('2024-09-16', '2024-09-20', 'SGP,SGG,SGB,S01,S02,S03,S04,S05,S11,SMP,SM1,SM2,GALL,foyerSG,foyerSMU,foyerSM,foyerMe1,foyerMe2,bistro,loggia,lobbyQ4,ufficiQ8,catering,foyerBar,ristorante,lbar', '', '1');
  // showMonths(first, last, structures, keyword, period)

  var start = Date.now();

  showOldMonths('2025-06-25', '2025-07-25', '', '', '');
  //ClearAll();

  var end = Date.now();
  Logger.log("Tempo di esecuzione funzione: " + (end - start) / 1000 + " secondi");

}

function testMethod2Key() {
  //showMonths('2024-09-25', '2024-09-28', '');
  //showMonths('2024-09-16', '2024-09-20', 'SGP,SGG,SGB,S01,S02,S03,S04,S05,S11,SMP,SM1,SM2,GALL,foyerSG,foyerSMU,foyerSM,foyerMe1,foyerMe2,bistro,loggia,lobbyQ4,ufficiQ8,catering,foyerBar,ristorante,lbar', '', '1');
  //methodL(4, convertDateBar(fromObject), convertDateBar(toObject), colore);
  //showMonths('2024-09-15', '2024-10-10', '1,3,4,5A,5B,7,G78,8,11,14,15,CC', '', '1');
  var struttureScelte = '';
  if ((struttureScelte == '') || (struttureScelte == undefined)) {
    structures = struttureBigKey(readVariables('struttureScelteSup', DataStructures)).join(','); // --> ORIGINAL WITH ALL THE STRUCTURES TRANSFORMED INTO A COMMA STRING
  } else {
    structures = keyArray;
  }
  first = '2024-09-15';
  last = '2024-10-10';
  first = incrDay(first, - 0);
  last = incrDay(last, 0);
  firstDate = text2monthDays(first);
  lastDate = text2monthDays(last);
  method2cKey(4, convertDateBar(firstDate[0]), convertDateBar(lastDate[0]), structures, '');
}

function verificaColoriCostanti(testo) {
  // Usa una regex per trovare tutte le occorrenze di "Colore=" seguite da un valore numerico
  const colori = [...testo.matchAll(/Colore=(\d+)/g)].map(match => match[1]);

  // Controlla se tutti i valori sono uguali
  return colori.every(colore => colore === colori[0]);
}

function text2monthDays(string) {
  const dateParts = string.split("-");
  const year = parseInt(dateParts[0], 10); // Parsing esplicito a intero
  const month = parseInt(dateParts[1], 10) - 1; // Parsing esplicito e aggiustamento mese (0-indicizzato)
  const day = parseInt(dateParts[2], 10); // Parsing esplicito

  // Usa array letterale per una creazione più veloce
  return [
    new Date(year, month, 1),
    new Date(year, month, day),
    new Date(year, month + 1, 0), // Ottieni l'ultimo giorno del mese corrente in modo più diretto
    new Date(year, 0, 1),      // 1 gennaio
    new Date(year, 11, 31)     // 31 dicembre
  ];
}

// From string '2024-04-05' to a new string --> 2024-04-incr
// Esempio di utilizzo
// console.log(incrDay('2024-06-01', 2)); // '2024-06-03'
function incrDay(dateString, daysToAdd) {
  // Converti la stringa di input in un oggetto Date
  const date = new Date(dateString);

  // Aggiungi i giorni specificati
  if (daysToAdd > 0) {
    date.setDate(date.getDate() + daysToAdd);
  } else {
    date.setDate(date.getDate() - Math.abs(daysToAdd));
  }

  // Estrai l'anno, il mese e il giorno dalla data risultante
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  // Costruisci la stringa della data risultante nel formato 'YYYY-MM-DD'
  return `${year}-${month}-${day}`;
}

// return integer between two dates
function datediff(first, second) {
  return Math.round((second - first) / (1000 * 60 * 60 * 24));
}

// Old Functions showMonths(first, last, structures, keyword, period)
function showOldMonths(first, last, structures, keyword, period) {
  try {
    //createUserSheet();
    resetFoglioConNuovo();
    // const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    // Inserisce un punto per mantenere la tabella strutturata
    // sh.getRange(90, 1500).setValue('.');

    // Se il periodo non è definito, regola automaticamente il range
    if (period === undefined) {
      first = incrDay(first, -incrementDay());
      last = incrDay(last, incrementDay());
    }

    // Converte le date in intervalli mensili
    const firstDate = text2monthDays(first);
    const lastDate = text2monthDays(last);

    // Pulisce il foglio prima di aggiungere i dati
    //ClearAll();

    // Imposta la formattazione del foglio

    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sh.getRange(90, 800).setValue('');
    sh.setFrozenRows(6);
    sh.setFrozenColumns(1);
    sh.setColumnWidths(1, 1, 250);   // Prima colonna larga 250
    sh.setColumnWidths(2, 368 - 1, 40); // Altre colonne larghe 40
    sh.setRowHeight(3, 40);  // Terza riga più alta
    sh.setRowHeights(1, 60, 21);

    // Se structures non è definito o è una stringa vuota, usa le strutture di default
    if (!structures) {
      structures = struttureBigKey(readVariables('struttureScelteSup', DataStructures)).join(',');
    }

    // Popola il foglio con i dati principali
    method2cKey(4, convertDateBar(firstDate[0]), convertDateBar(lastDate[0]), structures, keyword);
    checkToday(sh);

    // Nasconde le colonne fuori dal range
    const lc = sh.getLastColumn();
    const rangeColStart = datediff(firstDate[0], firstDate[1]);
    const colFinishStart = lc + 1 - datediff(lastDate[1], lastDate[2]);
    const rangeColFinish = datediff(lastDate[1], lastDate[2]);

    if (rangeColStart > 0) sh.hideColumns(2, rangeColStart);
    if (rangeColFinish > 0) sh.hideColumns(colFinishStart, rangeColFinish);
  } catch (error) {
    // Mostra un messaggio di errore e logga l'errore nella console
    const errorMessage = translate('alert.errorMessage') + ' (' + error.message + ')';
    SpreadsheetApp.getUi().alert(errorMessage);
    console.error(errorMessage);
  }
}

function struttureBigKey(array) {
  var structures = [];
  var container = array
  //Logger.log('container è '+container);
  //var container = struttureScelteSupKey();
  j = 0;
  for (let i = 0; i < container.length; i += 1) {
    //if (container[i][1] > 500) {
    if (container[i][1] >= 0) {
      structures[j] = container[i][0];
      j = j + 1;
    }
  }
  return structures
}

//
// method2cKey
//
// Display a chart with 2 base color: blue for approved events and orange for optionated
function method2cKey(startingRow, fromDate, toDate, structures, keyword) {
  // Load the sheet
  var sheet = SpreadsheetApp.getActiveSheet();

  var keyArray = structures.split(',').map(function (value) {
    return value.trim();
  });


  //var strucSelectSup = readVariables('struttureScelteSup', DataStructures); // NON FUNZIONA SE METTO DELLE STRUTTURE PERSONALIZZATE
  var strucSelectSup = [];
  for (let i = 0; i < keyArray.length; i += 1) {
    var index = findKey(keyArray[i], strutture(), 0);
    //Logger.log('Per keyArray[i] = ' + keyArray[i] + ' index è= ' + index);
    if (index >= 0) {
      strucSelectSup.push([strutture()[index][0], strutture()[index][8], strutture()[index][10], strutture()[index][11]]);
      //Logger.log(strutture()[index][0] + ' , ' + strutture()[index][8] + ' , ' + strutture()[index][10] + ' , ' + strutture()[index][11]);
    }
  };
  //Logger.log('All\'inizio strctselectSup è \n' + strucSelectSup);
  if ((keyArray.length == 1) || (keyArray.length == 2)) {
    /*
    keyArray.unshift('END');
    keyArray.push('END');
    */
  }
  //Logger.log(keyArray);
  //keyArray.unshift('END');
  //keyArray.push('END');
  // Full array  
  //var structures = strutture();
  var structures = readVariables('structures', DataStructures);
  // Array with the structures used in the rows
  //var strucSelect = struttureScelte(); // --> ORIGINAL WITH ALL THE STRUCTURES
  var strucSelect = keyArray; // --> NEW FOR PRINTING ONLY THE KEYWORD
  //Logger.log(strucSelect + ' ' + strucSelect[1] + ' è ' + typeof (strucSelect[1]));

  var strucBigSelect = struttureBigKey(readVariables('struttureScelteSup', DataStructures));
  // Array used for stats (strucSelect + relative square metres)
  //Logger.log('struSelectSup è '+strucSelectSup[19][0]+ ' '+strucSelectSup[19][3]);

  // External area
  const mappaSaleToGruppi = {};
  const struttureData = strutture();  // salva una volta per efficienza

  for (let i = 0; i < struttureData.length; i++) {
    const gruppo = struttureData[i][0];
    const elencoSale = struttureData[i][18];
    if (!gruppo || !elencoSale) continue;

    elencoSale.split(',').forEach(sala => {
      const salaTrimmed = sala.trim();
      if (salaTrimmed) {
        if (!mappaSaleToGruppi[salaTrimmed]) {
          mappaSaleToGruppi[salaTrimmed] = [];
        }
        mappaSaleToGruppi[salaTrimmed].push(gruppo);
      }
    });
  }

  // Load the colors used to highlight the occupation of the structures
  var colori2d = [[0, 0], [0, 0], [0, 0]];    // initialize array for approved
  var colori2dOpz = [[0, 0], [0, 0], [0, 0]]; // initialize array for optionated
  var colori2dOff = [[0, 0], [0, 0], [0, 0]]; // initialize array for optionated
  var colori = readVariables('method2colors', DataSettings);
  for (let i = 0; i < 3; i += 1) {
    colori2d[i][0] = colori[i][0];       // fill array for approved
    colori2dOpz[i][0] = colori[i + 3][0];  // fill array for optionated
    colori2dOff[i][0] = colori[i + 12][0];  // fill array for Offer
  }

  var row = startingRow;
  var date = new Date();
  var fDate = new Date(fromDate);
  var tDate = new Date(toDate);

  // Split in parts the date object
  var dateParts = fromDate.split("/");
  var fromObject = new Date(+dateParts[2], dateParts[1] - 1, 1);
  // Split in parts the date object
  var dateParts = toDate.split("/");
  var toObject = LastDayOfMonth(dateParts[2], dateParts[1])

  // Load into an array the data collected from Calendar in the prefixed period,
  // for this method the Calendar ID associated is linked to the first category
  //
  // function events2Array(startDate, finishDate, nColore, category)
  //var eventi = events2Array(convertDateBar(fromObject), convertDateBar(toObject), categories()[0][0], keyword); // To add the keyword
  var eventi = events2Array(convertDateBar(fromObject), convertDateBar(toObject), categories()[0][0], keyword); // To add the keyword
  //Logger.log(convertDateBar(fromObject) + ' | ' + convertDateBar(toObject) + ' | ' + categories()[0][0]);
  //Logger.log('Gli eventi sono ------------>' + JSON.stringify(eventi));

  // NEW array with unique numbers
  var arrayLastDay = [];
  for (let i = 0; i < eventi.length; i += 1) {
    arrayLastDay[i] = eventi[i][7];
  }
  // END array

  // Remove the events that start before the first day and end in that day (for example: 23.00-04.00)
  t = 0;
  while ((eventi.length > 1) && (eventi[t][0].getTime() < fromObject.getTime())) {
    eventi.splice(t, 1);
    t++
  }

  // -------------------
  // HEADER OF THE SHEET
  // -------------------
  var firstRowDate = row;
  sheet.getRange(row - 3, 1).setValue(translate('viewCalendar.firstCell')).setFontColor("#0b5394").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(row - 2, 1).setValue(translate('viewCalendar.secondCell')).setFontColor("#0b5394").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(row - 1, 1).setValue(translate('viewCalendar.mainCell')).setFontColor("#0b5394").setFontSize(18).setFontWeight("bold").setHorizontalAlignment("center");
  var startingYear = new Date(fromObject.getTime());
  //Logger.log('Starting year è ' + startingYear.getFullYear());
  if (startingYear.getFullYear() != toObject.getFullYear()) {
    sheet.getRange(row, 1).setValue(startingYear.getFullYear() + '⇾' + toObject.getFullYear()).setNumberFormat('@').setFontColor("#0b5394").setFontWeight("bold").setHorizontalAlignment("right");
  } else {
    sheet.getRange(row, 1).setValue(startingYear.getFullYear()).setNumberFormat('0').setFontColor("#0b5394").setFontWeight("bold").setHorizontalAlignment("right");
  }
  // Key for describing the row that contains dates and the first column that contains the selected structures
  sheet.getRange(firstRowDate + 1, 1).setValue(translate('viewCalendar.date')).setFontWeight("bold").setBorder(true, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(firstRowDate + 2, 1).setValue(translate('viewCalendar.structures')).setFontWeight("bold").setBorder(false, false, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);

  // ----------------------------------------------------------------------------------
  // Define the row that contains the days from the start to the end of the time period
  // ----------------------------------------------------------------------------------
  // Create a new array that contains for every month:
  // dispCal = [name of month, number of days in that month, year, number of month]
  // How many months are in the time period?
  var numMesi = monthDiff(fromObject, toObject);
  // Initialise an array with 4 columns and with the number of months as rows
  var dispCal = makeArray(numMesi, 4);
  for (let i = 0; i < numMesi; i += 1) {
    if (i == 0) {
      var nuovaData = new Date(fromObject.setMonth(fromObject.getMonth() + 0));
    } else {
      var nuovaData = new Date(fromObject.setMonth(fromObject.getMonth() + 1));
    }
    const monthsArray = translate('viewCalendar.months').split(', '); // Converte in array
    dispCal[i][0] = monthsArray[(nuovaData.getMonth() % 12 + 0)];
    dispCal[i][1] = getDaysInMonth(nuovaData.getMonth(), nuovaData.getFullYear());
    dispCal[i][2] = nuovaData.getFullYear();
    dispCal[i][3] = (nuovaData.getMonth() % 12 + 0);
  }

  // Count the total number of days in the period using dispCal
  var j = 0;
  var lastDay = 0;
  while (j < dispCal.length) {
    for (let i = 0; i < dispCal[j][1]; i += 1) {
    }
    lastDay = lastDay + dispCal[j][1];
    j++;
  }

  // Color Alternate Rows for the areas of the structures (NOW DISABLED, LOOK BELOW TO ENABLE)
  var totalRows = strucSelect.length;
  var totalColumns = 1 + lastDay;
  var startRow = firstRowDate + 3;
  var startColumn = 1;
  var row = startRow;
  while (row < totalRows + startRow) {
    var column = startColumn
    while (column < totalColumns + startColumn) {
      if (row % 2 == 0) {
        //sheet.getRange(row, column).setBackground('#d9d9d9'); // REMOVE THE COMMENT TO ENABLE IT
      }
      column++;
    }
    row++;
  }

  // Print the selected structures in the first column
  // REDUCE the structures reading the keyword
  for (let i = 0; i < strucSelect.length; i += 1) {
    if (findKey(strucSelect[i], structures, 0) > -1) { // --> CORRECT TO DISPLAY ALL THE ARRAY
      var zona = structures[findKey(strucSelect[i], structures, 0)][6]; // The correct name to output is in the 6 column
    } else {
      var zona = strucSelect[i] + ' non esiste!'
    }
    sheet.getRange(firstRowDate + 3 + i, 1).setValue(zona).setHorizontalAlignment("left").setFontSize(structures[findKey(strucSelect[i], structures, 0)][11]); // Wrote on the left strucSelectSup[i][3] the ROW HEIGHT COLUMN IS IN THE NUMBER 11
  }

  // First day from the period
  var pageDate = new Date(dispCal[0][2], dispCal[0][3], 1);
  // Set counters and variables
  var j = 0; // counter for the timeline of the structures per day  
  var y = 0; // counter for the timeline of the menu at the bottom
  var w = 0; // per l'avanzamento delle righe della matrice eventi nel menu
  var eventoRicorrente = [];
  var rowLib1 = 0;
  var rowKey = 0;
  var numDays = lastDay;

  //
  // For loop to print every day of the period in the columns
  //
  for (let i = 0; i < numDays; i += 1) {
    //
    // HEADER
    //
    // First: Month
    // If the day is the first of the month: write the name of the month and make the border
    if (pageDate.getDate() === 1) {
      const monthsArray = translate('viewCalendar.months').split(', '); // Converte in array
      var mese = monthsArray[(pageDate.getMonth() % 12 + 0)];
      //var mese = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"][(pageDate.getMonth() % 12 + 0)];
      sheet.getRange(firstRowDate - 1, 2 + i).setBorder(false, true, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      if (startingYear.getFullYear() != pageDate.getFullYear()) {
        sheet.getRange(firstRowDate - 1, 2 + i).setValue(pageDate.getFullYear()).setFontStyle("bold").setTextRotation(45);
      }
      sheet.getRange(firstRowDate, 2 + i).setValue(mese).setFontStyle("italic").setBorder(false, true, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(startRow + totalRows, 2 + i).setValue(mese).setFontStyle("italic").setBorder(true, true, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
    // Second: the number of the day      
    // If the day is a holiday, write it red
    const daysArray = translate('viewCalendar.days').split(', '); // Converte in array
    festivo = ((["D", "L", "M", "M", "G", "V", "S"][pageDate.getDay()] == 'D') || (specialHolidays().indexOf(convertDateBar(pageDate)) > -1) || (holidays().indexOf(convertDayMonthBar(pageDate)) > -1));
    if (festivo) {
      sheet.getRange(firstRowDate + 1, 2 + i, 2).setFontColor("red").setBorder(false, true, false, true, false, false, "grey", SpreadsheetApp.BorderStyle.DASHED);
    }
    sheet.getRange(firstRowDate + 1, 2 + i).setValue(convertDateUSA(pageDate.toDateString())).setFontSize(12).setNumberFormat("DD").setHorizontalAlignment("center").setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
    // Third: first letter of the day
    sheet.getRange(firstRowDate + 2, 2 + i).setValue(daysArray[pageDate.getDay()]).setFontSize(9).setNumberFormat("DD").setHorizontalAlignment("center").setBorder(false, false, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);

    //
    // Fill the rows that contain the selected structures
    //
    // While the events is at this day
    // Initialize an array to check if there are more events in the same structure per day
    var repetita = [];
    var r = 0;
    while ((j < eventi.length) && (convertDate(eventi[j][0]) == convertDate(pageDate))) {
      var comNonTrov = '';
      var nonSiTrova = true;

      // EXTERNAL AREA AND SALETTE Loop the structures inside an event
      for (let k = 0; k < eventi[j][10].length; k++) {
        const sala = eventi[j][10][k];
        const gruppi = mappaSaleToGruppi[sala];
        if (gruppi) {
          gruppi.forEach(gruppo => {
            if (eventi[j][10].indexOf(gruppo) < 0) {
              eventi[j][10].push(gruppo);
            }
          });
        }
      }

      // NEW Change the number at the j position
      arrayLastDay[j] = -1;
      // NEW END

      // Loop the structures inside an event
      for (let k = 0; k < eventi[j][10].length; k += 1) {
        var comNonTrov = eventi[j][8] + '\n' + comNonTrov;
        //Logger.log(eventi[j][10][k] + ' è ' + typeof (eventi[j][10][k]));


        // if a structure is in the first column (i.e. the structures selected in the settings)
        if ((strucSelect.indexOf(eventi[j][10][k]) >= 0)) {
          nonSiTrova = false;
          var rowCercata = findKey(eventi[j][10][k], strucSelectSup, 0);
          // Assign the color if the event is optionated or approved
          //if (eventi[j][8].substring(0,4) === optionated()) { // vecchio sistema
          if (optionated().indexOf(eventi[j][8].substring(0, 4)) >= 0) {
            if (parseEventString(eventi[j][8]).opz == 'OFF') {
              colore = colori2dOff[selectHigh(eventi[j][7], eventi[j][1])[0]][0];
            } else {
              colore = colori2dOpz[selectHigh(eventi[j][7], eventi[j][1])[0]][0];
            }
          } else {
            colore = colori2d[selectHigh(eventi[j][7], eventi[j][1])[0]][0];
          }
          // In the corresponding row, put the color and number if it is an event or just color in the other cases
          // Create an array to check if there are more events in the same structure per day
          repetita[r] = [rowCercata, colore, eventi[j][8], strucSelectSup[rowCercata][2], eventi[j][7]];

          // To color the event is in a child structure (salette)
          if (strucSelectSup[rowCercata][2] != 0) {
            nota = "";
            evento = false;
            for (o = 0; o < findAllKey(strucSelect.indexOf(strucSelectSup[rowCercata][2]), repetita).length; o += 1) {

              if ((findKey(strucSelect.indexOf(strucSelectSup[rowCercata][2]), repetita) > -1)) {
                nota = nota + '\n' + repetita[findAllKey(strucSelect.indexOf(strucSelectSup[rowCercata][2]), repetita)[o]][2] + ' Colore=' + repetita[findAllKey(strucSelect.indexOf(strucSelectSup[rowCercata][2]), repetita)[o]][4] + '\n';
              }
            }

            var valore = eventi[j][7];
            if (findAllKey(strucSelect.indexOf(strucSelectSup[rowCercata][2]), repetita).length > 1) {
              for (g = 0; g < findAllKey(strucSelect.indexOf(strucSelectSup[rowCercata][2]), repetita).length; g += 1) {
                if (repetita[findAllKey(strucSelect.indexOf(strucSelectSup[rowCercata][2]), repetita)[g]][3] === 0) {
                  salette = true;
                  var valore = repetita[findAllKey(strucSelect.indexOf(strucSelectSup[rowCercata][2]), repetita)[g]][4];
                }
              }
            }

            r = r + 1;
            repetita[r] = [strucSelect.indexOf(strucSelectSup[rowCercata][2]), method2colors()[8][0], eventi[j][8] + '(in ' + strutture()[findKey(eventi[j][10][k], strutture(), 0)][6] + ')', strucSelectSup[rowCercata][2], eventi[j][7]];
            var salette = false;
            if (findAllKey(strucSelect.indexOf(strucSelectSup[rowCercata][2]), repetita).length > 1) {
              for (g = 0; g < findAllKey(strucSelect.indexOf(strucSelectSup[rowCercata][2]), repetita).length; g += 1) {
                if (repetita[findAllKey(strucSelect.indexOf(strucSelectSup[rowCercata][2]), repetita)[g]][3] === 0) {
                  salette = true;
                  var valore = repetita[findAllKey(strucSelect.indexOf(strucSelectSup[rowCercata][2]), repetita)[g]][4];
                }
              }
            }
          }


          if ((findAllKey(rowCercata, repetita).length == 1)) { // default one only event

            // NEW CODE FOR PRE ALLESTIMENTI
            if (eventi[j][1] === categories()[4][1]) {
              sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setBackground(method2colors()[9][0]).setNote(eventi[j][8] + '\n' + eventi[j][10][k] + '\n' + translate('viewCalendar.date') + ' ' + convertDateBar(eventi[j][0]));
              if (optionated().indexOf(eventi[j][8].substring(0, 4)) >= 0) {
                if (parseEventString(eventi[j][8]).opz == 'OFF') {
                  sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setBackground(method2colors()[12][0]).setNote(eventi[j][8] + '\n' + eventi[j][10][k] + '\n' + translate('viewCalendar.date') + ' ' + convertDateBar(eventi[j][0]));
                } else {
                  sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setBackground(method2colors()[10][0]).setNote(eventi[j][8] + '\n' + eventi[j][10][k] + '\n' + translate('viewCalendar.date') + ' ' + convertDateBar(eventi[j][0]));
                }
              }
            }

            if ((eventi[j][1] === categories()[1][1]) || (eventi[j][1] === categories()[2][1])) {
              sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setBackground(colore).setNote(eventi[j][8] + '\n' + eventi[j][10][k] + '\n' + translate('viewCalendar.date') + ' ' + convertDateBar(eventi[j][0]));
              //} else if ((eventi[j][1] != categories()[4][1])&&((eventi[j][1] != categories()[3][1]))) { // P prenotato o L lavori
            } else if (eventi[j][1] === categories()[0][1]) { // mostra solo E eventi con il numero progressivo
              sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setValue(eventi[j][7] + 1).setNumberFormat("00").setBackground(colore).setNote(eventi[j][8] + '\n' + eventi[j][10][k] + '\n' + translate('viewCalendar.date') + ' ' + convertDateBar(eventi[j][0]));
            } else if (eventi[j][1] === categories()[3][1]) { // mostra solo L eventi con lettera progressiva e nota
              sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setValue(printToLetter(eventi[j][7] + 1)).setNumberFormat("00").setFontColor('red').setHorizontalAlignment("center").setNote(eventi[j][8] + '\n' + eventi[j][10][k] + '\n' + translate('viewCalendar.date') + ' ' + convertDateBar(eventi[j][0]));
            }
          } else if ((findAllKey(rowCercata, repetita).length > 1) && ((repetita[findAllKey(rowCercata, repetita)[0]][1] == method2colors()[0][0]) || (repetita[findAllKey(rowCercata, repetita)[0]][1] == method2colors()[3][0]))) { // A + E or A + D or D + A
            nota = "";
            evento = false;
            for (u = 0; u < findAllKey(rowCercata, repetita).length; u += 1) {
              if (repetita[findAllKey(rowCercata, repetita)[u]][1] == method2colors()[1][0]) {
                evento = true;
              }
              nota = nota + '\n' + repetita[findAllKey(rowCercata, repetita)[u]][2] + ' Colore=' + repetita[findAllKey(rowCercata, repetita)[u]][4] + '\n';
            }
            if (evento) {
              if (verificaColoriCostanti(nota)) {
                sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setValue(eventi[j][7] + 1).setNumberFormat("00").setNote(nota + '\n' + eventi[j][10][k] + '\n' + translate('viewCalendar.date') + ' ' + convertDateBar(eventi[j][0]));
              } else {
                sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setValue(eventi[j][7] + 1).setNumberFormat("00").setBackground(method2colors()[6][0]).setNote(nota + '\n' + eventi[j][10][k] + '\n' + translate('viewCalendar.date') + ' ' + convertDateBar(eventi[j][0]));
              }
            } else {
              if (verificaColoriCostanti(nota)) {
                sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setNote(nota + '\n' + eventi[j][10][k] + '\n' + translate('viewCalendar.date') + ' ' + convertDateBar(eventi[j][0]));
              } else {
                sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setBackground(method2colors()[6][0]).setNote(nota + '\n' + eventi[j][10][k] + '\n' + translate('viewCalendar.date') + ' ' + convertDateBar(eventi[j][0]));
              }
            }
          } else if ((findAllKey(rowCercata, repetita).length > 1) && ((repetita[findAllKey(rowCercata, repetita)[0]][1] == method2colors()[1][0]) || (repetita[findAllKey(rowCercata, repetita)[0]][1] == method2colors()[4][0]))) { // E + D
            nota = "";
            for (u = 0; u < findAllKey(rowCercata, repetita).length; u += 1) {
              nota = nota + '\n' + repetita[findAllKey(rowCercata, repetita)[u]][2] + ' Colore=' + repetita[findAllKey(rowCercata, repetita)[u]][4] + '\n';
            }
            //sheet.getRange(firstRowDate+3+rowCercata, 2+i).setValue(eventi[j][7]+1).setNumberFormat("00").setBackground(method2colors()[7][0]).setNote(nota +'\n'+eventi[j][10][k]+'\n'+translate('viewCalendar.date')+' '+convertDateBar(eventi[j][0]));
            if (verificaColoriCostanti(nota)) {
              sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setValue(eventi[j][7] + 1).setNumberFormat("00").setNote(nota + '\n' + eventi[j][10][k] + '\n' + translate('viewCalendar.date') + ' ' + convertDateBar(eventi[j][0]));
            } else {
              sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setValue(eventi[j][7] + 1).setNumberFormat("00").setBackground(method2colors()[6][0]).setNote(nota + '\n' + eventi[j][10][k] + '\n' + translate('viewCalendar.date') + ' ' + convertDateBar(eventi[j][0]));
            }
          } /* SALETTEE!!!! else if ((findAllKey(rowCercata, repetita).length > 1) && (repetita[findAllKey(rowCercata, repetita)[0]][1] == method2colors()[8][0])) { // SALETTE 
            nota = "";
            for (u = 0; u < findAllKey(rowCercata, repetita).length; u += 1) {
              nota = nota + '\n' + repetita[findAllKey(rowCercata, repetita)[u]][2] + '\n';
            }
            sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setValue(eventi[j][7] + 1).setNumberFormat("00").setFontSize(14).setBackground(method2colors()[8][0]).setNote(nota + '\n' + eventi[j][10][k] + '\n'+translate('viewCalendar.date')+' ' + convertDateBar(eventi[j][0]));
          }  FINE SALETTE */
          // NEW ADD A BORDER AT THE RIGHT IF THIS IS THE LAST DAY ['joe', 'jane', 'mary'].indexOf('jane') >= 0
          if (convertDate(eventi[j][11]) == convertDate(pageDate)) {
            sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setBorder(null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.DOUBLE); //SOLID_THICK .DOUBLE
          }
          r = r + 1;
        } else if (nonSiTrova) {
          nonSiTrova = true;
        }
        /*
        if (arrayLastDay.indexOf(eventi[j][7]) <= 0) {
          //Logger.log('rowcercata è '+rowCercata+ ' firstRowDate è '+firstRowDate);
          if (rowCercata == undefined) {
          } else {
            sheet.getRange(firstRowDate + 3 + rowCercata, 2 + i).setBorder(null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM); //SOLID_THICK .DOUBLE
          }
        }
        */

        // NEW END  
      }
      j++;
    }

    //
    // Fill the timeline at the bottom that contains the increasing numbers with the corresponding event
    //
    // A complicated way to assign the row spacing:
    // every 20 increments the row starts again at the top
    if (i % 10 === 0) {
    }
    if ((rowLib1 > 0) && ((i - rowKey) % 10 === 0)) {
      rowLib1 = 0;
    }
    //if (sheet.getRange(firstRowDate+1,2+i).getValue().getDate() === 1) {
    if (pageDate.getDate() === 1) {
      eventoRicorrente = [];
      rowLib1 = 0;
    }
    // While the events is at this day
    while ((y < eventi.length) && (convertDate(eventi[y][0]) == convertDate(pageDate))) {
      if (optionated().indexOf(eventi[y][8].substring(0, 4)) >= 0) {

        if (parseEventString(eventi[y][8]).opz == 'OFF') {
          colore = colori2dOff[selectHigh(eventi[y][7], eventi[y][1])[0]][0];
        } else {
          colore = colori2dOpz[selectHigh(eventi[y][7], eventi[y][1])[0]][0];
        }
      } else {
        colore = colori2d[selectHigh(eventi[y][7], eventi[y][1])[0]][0];
      }
      // Print an event only if one or more structures are in the first column
      var unoPresente = false;
      var unoBigPresente = false;
      for (let k = 0; k < eventi[y][10].length; k += 1) {
        if ((strucSelect.indexOf(eventi[y][10][k]) >= 0)) {
          unoPresente = true;
          if ((strucBigSelect.indexOf(eventi[y][10][k]) >= 0)) {
            unoBigPresente = true;
          }
        }
      }
      // Print an event only if is in the first category E =0 and L = 3 defined in main.gs
      //if ((eventoRicorrente.indexOf(eventi[y][7]) <= -1) && ((eventi[y][1] === categories()[0][1]) || (eventi[y][1] === categories()[3][1])) && (unoPresente)) {
      // Print an event only if is in the first category E =0 defined in main.gs
      if ((eventoRicorrente.indexOf(eventi[y][7]) <= -1) && ((eventi[y][1] === categories()[0][1]) || (eventi[y][1] === categories()[3][1]))) {
        // Print all events in the period selected
        if (rowLib1 === 0) {
          rowKey = i; // in which bottom row
        }
        // Print the alignment of the Event title according of the proximity of the month-end
        // left < day 25
        // right > day 25

        // Cicle for inserting the real name fo the structures in the note
        realNameStruct = '';
        for (let q = 0; q < eventi[y][10].length; q += 1) {
          if (findKey(eventi[y][10][q], strutture()) > -1) {
            if (strutture()[findKey(eventi[y][10][q], strutture())][6] != ' ') {
              realNameStruct = realNameStruct + strutture()[findKey(eventi[y][10][q], strutture())][6] + '\n';
            }
          }
        }

        if (unoPresente) {
          if (pageDate.getDate() < 25) {
            rowLib1++;
            if (unoBigPresente) {
              if (eventi[y][1] === categories()[0][1]) {
                sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 2 + i).setBackground(colore).setValue(eventi[y][7] + 1).setNumberFormat("00").setFontSize(14).setNote('(' + convertDateBar(eventi[y][0]) + ') ' + eventi[y][8] + translate('viewCalendar.where') + realNameStruct + '\n\n[' + eventi[y][10] + ']\nColore: ' + eventi[y][7]); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);
              } else {
                sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 2 + i).setValue(printToLetter(eventi[y][7] + 1)).setNumberFormat("00").setFontSize(14).setFontColor('red').setFontWeight("bold").setHorizontalAlignment("center").setNote('(' + convertDateBar(eventi[y][0]) + ') ' + eventi[y][8] + translate('viewCalendar.where') + realNameStruct + '\n\n[' + eventi[y][10] + ']\nColore: ' + eventi[y][7]); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);                
              }
              sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 3 + i).setValue(eventi[y][2].replace(/\s+/g, '').substring(0, 18)).setFontColor("black"); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);
            } else { // salette
              //method2colors()[8][0]
              if (eventi[y][1] === categories()[0][1]) {
                sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 2 + i).setBackground(colore).setValue(eventi[y][7] + 1).setNumberFormat("00").setFontSize(14).setNote('(' + convertDateBar(eventi[y][0]) + ') ' + eventi[y][8] + translate('viewCalendar.where') + realNameStruct + '\n\n[' + eventi[y][10] + ']\nColore: ' + eventi[y][7]); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);
              } else {
                sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 2 + i).setValue(printToLetter(eventi[y][7] + 1)).setNumberFormat("00").setFontSize(14).setFontColor('red').setFontWeight("bold").setHorizontalAlignment("center").setNote('(' + convertDateBar(eventi[y][0]) + ') ' + eventi[y][8] + translate('viewCalendar.where') + realNameStruct + '\n\n[' + eventi[y][10] + ']\nColore: ' + eventi[y][7]); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);                
              }
              sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 3 + i).setValue(eventi[y][2].replace(/\s+/g, '').substring(0, 18)).setFontColor("black"); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);              
            }
          } else {
            rowLib1++;
            if (unoBigPresente) {
              if (eventi[y][1] === categories()[0][1]) {
                sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 2 + i).setBackground(colore).setValue(eventi[y][7] + 1).setNumberFormat("00").setFontSize(14).setNote('(' + convertDateBar(eventi[y][0]) + ') ' + eventi[y][8] + translate('viewCalendar.where') + realNameStruct + '\n\n[' + eventi[y][10] + ']\nColore: ' + eventi[y][7]); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);
              } else {
                sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 2 + i).setValue(printToLetter(eventi[y][7] + 1)).setNumberFormat("00").setFontSize(14).setFontColor('red').setFontWeight("bold").setHorizontalAlignment("center").setNote('(' + convertDateBar(eventi[y][0]) + ') ' + eventi[y][8] + translate('viewCalendar.where') + realNameStruct + '\n\n[' + eventi[y][10] + ']\nColore: ' + eventi[y][7]); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);
              }
              sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 1 + i).setValue(eventi[y][2].replace(/\s+/g, '').substring(0, 18)).setFontColor("black").setHorizontalAlignment("right"); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);            
            } else { //salette
              sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 1 + i).setValue(eventi[y][2].replace(/\s+/g, '').substring(0, 18)).setFontColor("black").setHorizontalAlignment("right"); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);              
            }
          }
        } else { // se la struttura non è in elenco
          if (strucBigSelect.length == keyArray.length) {
            if (pageDate.getDate() < 25) {
              rowLib1++;
              sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 2 + i).setBackground(method2colors()[11][0]).setValue(eventi[y][7] + 1).setNumberFormat("00").setFontSize(14).setNote('(' + convertDateBar(eventi[y][0]) + ') ' + eventi[y][8] + translate('viewCalendar.where') + realNameStruct + '\n\n[' + eventi[y][10] + ']\nColore: ' + eventi[y][7]); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);
              sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 3 + i).setValue(eventi[y][2].replace(/\s+/g, '').substring(0, 18)).setFontColor("black"); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);
            } else {
              rowLib1++;
              sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 2 + i).setBackground(method2colors()[11][0]).setValue(eventi[y][7] + 1).setNumberFormat("00").setFontSize(14).setNote('(' + convertDateBar(eventi[y][0]) + ') ' + eventi[y][8] + translate('viewCalendar.where') + realNameStruct + '\n\n[' + eventi[y][10] + ']\nColore: ' + eventi[y][7]); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);
              sheet.getRange(firstRowDate + 3 + rowLib1 + strucSelect.length, 1 + i).setValue(eventi[y][2].replace(/\s+/g, '').substring(0, 18)).setFontColor("black").setHorizontalAlignment("right"); //.setNote(eventi[j][8]+'\n'+eventi[j][10][k]);
            }
          }
        }



        eventoRicorrente.push(eventi[y][7]);
      }
      y++;
    }
    //
    // Add a day to the counter
    //
    var pageDateZero = new Date(convertDateUSA(pageDate) + " 02:00:00 UTC +2"); //GMT UTC
    pageDate = new Date(pageDateZero.getTime() + 1 * 3600000 * 24);
  }

  //
  // Print border and texts at the bottom of the first column
  //
  sheet.getRange(startRow + totalRows - 1, 1, 1, totalColumns + startColumn - 1).setBorder(false, false, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(startRow + totalRows + 1, 1, 1, totalColumns + startColumn - 1).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(startRow + totalRows, 60);
  var today = new Date();
  sheet.getRange(startRow + totalRows, 1).setValue(translate('viewCalendar.eventsList') + '\n(' + today.toLocaleDateString() + ' ' + today.toLocaleTimeString() + ')').setFontSize(12).setHorizontalAlignment("center");
  sheet.getRange(startRow + totalRows + 1, 1).setValue(translate('viewCalendar.numberEvent')).setFontSize(10).setFontWeight("bold").setBackground(colori[1][0]).setHorizontalAlignment("center");
  sheet.getRange(startRow + totalRows + 2, 1).setValue(translate('viewCalendar.assOrDisass')).setFontSize(8).setFontWeight("bold").setBackground(colori[1][0]).setHorizontalAlignment("center");
  sheet.getRange(startRow + totalRows + 3, 1).setValue(translate('viewCalendar.preAss')).setFontSize(8).setFontWeight("bold").setBackground(colori[9][0]).setHorizontalAlignment("center");
  sheet.getRange(startRow + totalRows + 4, 1).setValue(translate('viewCalendar.numberOpz')).setFontSize(8).setFontWeight("bold").setBackground(colori[4][0]).setHorizontalAlignment("center");
  sheet.getRange(startRow + totalRows + 5, 1).setValue(translate('viewCalendar.assOrDisass')).setFontSize(8).setFontWeight("bold").setBackground(colori[4][0]).setHorizontalAlignment("center");
  sheet.getRange(startRow + totalRows + 6, 1).setValue(translate('viewCalendar.preAss')).setFontSize(8).setFontWeight("bold").setBackground(colori[10][0]).setHorizontalAlignment("center");
  sheet.getRange(startRow + totalRows + 7, 1).setValue(translate('viewCalendar.numberOff')).setFontSize(8).setFontWeight("bold").setBackground(colori[12][0]).setHorizontalAlignment("center");
  sheet.getRange(startRow + totalRows + 8, 1).setValue(translate('viewCalendar.moreEvents')).setFontSize(10).setFontWeight("bold").setBackground(colori[6][0]).setHorizontalAlignment("center");
  sheet.getRange(startRow + totalRows + 9, 1).setValue(translate('viewCalendar.letterWork')).setFontColor('red').setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center");

  var lr = sheet.getLastRow();
  //sheet.getRange(startRow + totalRows + 3, 1).setValue(today).setNumberFormat('dd/MM/yy - HH:mm)').setFontSize(14).setHorizontalAlignment("center");
  sheet.getRange(lr + 1, totalColumns + startColumn - 1).setValue('.');
  sheet.getRange(lr + 1, 1, 1, totalColumns + startColumn - 1).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  //
  // Recheck the timeline and fix the border and style if it is Sunday
  //
  var pageDate = new Date(dispCal[0][2], dispCal[0][3], 1);
  for (let i = 0; i < numDays; i += 1) {
    // If Sunday
    domenica = (["D", "L", "M", "M", "G", "V", "S"][pageDate.getDay()] == 'D');
    if (domenica) {
      sheet.getRange(firstRowDate + 1, 2 + i, 1).setBorder(true, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(firstRowDate + 2, 2 + i, 1).setBorder(false, false, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(firstRowDate + 3, 2 + i, 1).setBorder(true, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(firstRowDate + 4, 2 + i, strucSelect.length).setBorder(false, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(firstRowDate + 2 + strucSelect.length, 2 + i, 1).setBorder(false, false, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      // At the bottom
      sheet.getRange(firstRowDate + 4 + strucSelect.length, 2 + i, 1).setBorder(true, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(firstRowDate + 5 + strucSelect.length, 2 + i, lr - (firstRowDate + 5 + strucSelect.length)).setBorder(false, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(lr, 2 + i, 1).setBorder(false, false, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
    //
    // Add a day
    //
    var pageDateZero = new Date(convertDateUSA(pageDate) + " 02:00:00 UTC +2");
    pageDate = new Date(pageDateZero.getTime() + 1 * 3600000 * 24);
  }

  //
  // Remove the remaining empty rows and columns
  //
  var lr = sheet.getLastRow();
  var mr = sheet.getMaxRows();
  if (mr - lr != 0) {
    sheet.deleteRows(lr + 1, mr - lr);
  }
  var lc = sheet.getLastColumn();
  var mc = sheet.getMaxColumns();
  if (mc - lc != 0) {
    sheet.deleteColumns(lc + 1, mc - lc - 1);
  }
}

// ---------------------------------------------
// START STATS CODE
// ---------------------------------------------

function getStruttureInfo(sheet, strutture) {
  const startRow = 7;
  const col = 1;
  const maxRows = sheet.getLastRow() - startRow + 1;
  const range = sheet.getRange(startRow, col, maxRows, 1);
  const values = range.getValues().flat();

  // Ottieni l'elenco dei nomi delle strutture valide dalla colonna 7 della matrice 'strutture'
  const struttureValide = strutture.map(r => r[6]).filter(x => x && x.toString().trim() !== '');

  const strutturePresenti = []; // Vettore delle strutture consecutive valide
  const strutturePresentiSup = []; // Vettore delle strutture consecutive valide
  let lastValidStructureIndex = -1; // Indice dell'ultima struttura valida trovata nell'intervallo

  // Itera sui valori per trovare sia le strutture consecutive che l'ultima struttura valida
  for (let i = 0; i < values.length; i++) {
    const val = values[i];
    const trimmedVal = val ? val.toString().trim() : '';

    if (trimmedVal === '') {
      // Se è una riga vuota, non aggiungere a 'strutturePresenti' ma continua a cercare 'lastValidStructureIndex'
    } else if (struttureValide.includes(trimmedVal)) {
      //Logger.log(trimmedVal);
      //Logger.log(strutture[findKey(trimmedVal, strutture, 6)][8]);
      strutturePresenti.push(trimmedVal);
      strutturePresentiSup.push([trimmedVal, strutture[findKey(trimmedVal, strutture, 6)][8]]);
      lastValidStructureIndex = i; // Aggiorna l'indice dell'ultima struttura valida
    } else {
      break;
    }
  }

  // Calcola la lunghezza reale comprensiva delle righe vuote
  const lunghezzaReale = (lastValidStructureIndex !== -1) ? (lastValidStructureIndex + 1) : 0;

  return {
    struttureVettore: strutturePresenti,
    struttureArray: strutturePresentiSup,
    lunghezzaReale: lunghezzaReale
  };
}

function getVisibleColumnIndices(sheet, startCol, endCol) {
  const visibleIndices = [];
  for (let col = startCol; col <= endCol; col++) {
    if (!sheet.isColumnHiddenByUser(col)) {
      visibleIndices.push(col - startCol); // Indice relativo per l'array colors/texts
    }
  }
  return visibleIndices;
}

function classificaCella(color, text) {
  const testoPresente = text?.toString().trim() !== '';
  switch (color) {
    case "#e06666": return "multi"; // più eventi
    case "#9fc5e8": return testoPresente ? "ev_conf" : "ad_conf"; // confermati
    case "#ffe599": return testoPresente ? "ev_opz" : "ad_opz"; // opzionati
    case "#b6d7a8": return testoPresente ? "ev_off" : "ad_off"; // offerta
    default: return "none";
  }
}

function isDateCell(cellValue) {
  if (!cellValue) return false;
  return cellValue instanceof Date ||
    (typeof cellValue === 'string' && !isNaN(Date.parse(cellValue)));
}

function countRowsBackgroundNuova() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lc = sheet.getLastColumn();
  const struttureInfo = getStruttureInfo(sheet, strutture());
  const strutturePresenti = struttureInfo.struttureVettore;
  const lr = struttureInfo.lunghezzaReale;

  if (lr === 0) return; // Nessuna struttura da processare

  //Logger.log(sheet.getRange(5, 2).getValue());
  //Logger.log(typeof (sheet.getRange(5, 2).getValue()));
  if (sheet.getRange(5, 2).getValue() instanceof Date) {

    Logger.log('lr=' + lr + '  lc=' + lc);

    // Ottieni indici delle colonne visibili (relativi all'array)
    const visibleIndices = getVisibleColumnIndices(sheet, 2, lc);
    const numVisibleColumns = visibleIndices.length;

    Logger.log('Colonne visibili: ' + numVisibleColumns);

    // SINGOLA LETTURA di tutti i dati necessari
    const dataRange = sheet.getRange(7, 2, lr, lc - 1);
    const colors = dataRange.getBackgrounds();
    const texts = dataRange.getValues();

    // Leggi date per le note (una sola volta)
    const firstDay = sheet.getRange(5, 2).getValue();
    const lastDay = sheet.getRange(5, lc).getValue();

    // Definisci colori una volta sola
    const colorMap = {
      "#e06666": "multi",
      "#9fc5e8": "conf",
      "#ffe599": "opz",
      "#b6d7a8": "off"
    };

    // Prepara TUTTI gli array per batch operations
    const headerData = [
      ['Occ.', '', 'Conf.', '', 'Conf', '', 'Conf', '', 'Opz', '', 'Opz', '', 'Opz', '', 'Off', '', 'Off', '', 'Off', ''],
      ['Tot.', '', 'Tot', '', 'Ev', '', 'A+D', '', 'Tot', '', 'Ev', '', 'A+D', '', 'Tot', '', 'Ev', '', 'A+D', ''],
      ['Occ.', '%.', '', '%', '', '%', '', '%', '', '%', '', '%', '', '%', '', '%', '', '%', '', '%']
    ];

    const outputValues = [];
    const outputFormats = [];
    const outputNotes = [];

    // CICLO PRINCIPALE ottimizzato
    for (let r = 0; r < lr; r++) {
      const counts = {
        multi: 0,
        conf: 0, confEv: 0, confAd: 0,
        opz: 0, opzEv: 0, opzAd: 0,
        off: 0, offEv: 0, offAd: 0
      };

      const rowColors = colors[r];
      const rowTexts = texts[r];

      // Itera solo sulle colonne visibili
      for (const colIndex of visibleIndices) {
        const color = rowColors[colIndex];
        const hasText = rowTexts[colIndex]?.toString().trim() !== '';
        const colorType = colorMap[color];

        if (!colorType) continue;

        if (colorType === "multi") {
          counts.multi++;
        } else if (colorType === "conf") {
          counts.conf++;
          if (hasText) counts.confEv++;
          else counts.confAd++;
        } else if (colorType === "opz") {
          counts.opz++;
          if (hasText) counts.opzEv++;
          else counts.opzAd++;
        } else if (colorType === "off") {
          counts.off++;
          if (hasText) counts.offEv++;
          else counts.offAd++;
        }
      }

      const totalOcc = counts.multi + counts.conf + counts.opz + counts.off;

      if (totalOcc > 0) {
        const struttura = strutturePresenti[r];

        // Crea nota completa
        let info = `Occupazione Totale ${struttura}\n(${convertDateBar(firstDay)} - ${convertDateBar(lastDay)})`;
        info += `\n\n=== RIEPILOGO COMPLETO ===`;
        info += `\nOCCUPAZIONE TOTALE =\t${totalOcc} giorni\t(${(totalOcc / numVisibleColumns * 100).toFixed(0)}%)`;
        info += `\n\n--- CONCOMITANTI ---`;
        info += `\nTotale ConCOMITANTI =\t${counts.multi} giorni\t(${(counts.multi / numVisibleColumns * 100).toFixed(0)}%)`;
        info += `\n\n--- CONFERMATI ---`;
        info += `\nTotale Confermati =\t${counts.conf} giorni\t(${(counts.conf / numVisibleColumns * 100).toFixed(0)}%)`;
        info += `\n  Solo Eventi =\t${counts.confEv} giorni\t(${(counts.confEv / numVisibleColumns * 100).toFixed(0)}%)`;
        info += `\n  Solo All. + Dis. =\t${counts.confAd} giorni\t(${(counts.confAd / numVisibleColumns * 100).toFixed(0)}%)`;
        info += `\n\n--- OPZIONATI ---`;
        info += `\nTotale Opzionati =\t${counts.opz} giorni\t(${(counts.opz / numVisibleColumns * 100).toFixed(0)}%)`;
        info += `\n  Opz Eventi =\t${counts.opzEv} giorni\t(${(counts.opzEv / numVisibleColumns * 100).toFixed(0)}%)`;
        info += `\n  Opz All. + Dis. =\t${counts.opzAd} giorni\t(${(counts.opzAd / numVisibleColumns * 100).toFixed(0)}%)`;
        info += `\n\n--- OFFERTA ---`;
        info += `\nTotale Offerta =\t${counts.off} giorni\t(${(counts.off / numVisibleColumns * 100).toFixed(0)}%)`;
        info += `\n  Off Eventi =\t${counts.offEv} giorni\t(${(counts.offEv / numVisibleColumns * 100).toFixed(0)}%)`;
        info += `\n  Off All. + Dis. =\t${counts.offAd} giorni\t(${(counts.offAd / numVisibleColumns * 100).toFixed(0)}%)`;

        // Prepara riga completa
        outputValues.push([
          totalOcc, totalOcc / numVisibleColumns,
          counts.conf, counts.conf / numVisibleColumns,
          counts.confEv, counts.confEv / numVisibleColumns,
          counts.confAd, counts.confAd / numVisibleColumns,
          counts.opz, counts.opz / numVisibleColumns,
          counts.opzEv, counts.opzEv / numVisibleColumns,
          counts.opzAd, counts.opzAd / numVisibleColumns,
          counts.off, counts.off / numVisibleColumns,
          counts.offEv, counts.offEv / numVisibleColumns,
          counts.offAd, counts.offAd / numVisibleColumns
        ]);

        outputFormats.push(['00', '0%', '00', '0%', '00', '0%', '00', '0%', '00', '0%', '00', '0%', '00', '0%', '00', '0%', '00', '0%', '00', '0%']);

        // Solo prima colonna ha la nota
        const noteRow = new Array(20).fill('');
        noteRow[0] = info;
        outputNotes.push(noteRow);
      } else {
        // Riga vuota
        outputValues.push(new Array(20).fill(''));
        outputFormats.push(new Array(20).fill(''));
        outputNotes.push(new Array(20).fill(''));
      }
    }

    // === BATCH OPERATIONS ULTRA-OTTIMIZZATE ===

    // 1. Scrivi intestazioni (UNA SOLA CHIAMATA)
    sheet.getRange(4, lc + 1, 3, headerData[0].length).setValues(headerData);

    // 2. Scrivi TUTTI i valori (UNA SOLA CHIAMATA)
    const mainRange = sheet.getRange(7, lc + 1, outputValues.length, outputValues[0].length);
    mainRange.setValues(outputValues);

    // 3. Applica TUTTI i formati (UNA SOLA CHIAMATA per tipo)
    const formatRange = sheet.getRange(7, lc + 1, outputFormats.length, outputFormats[0].length);

    // Applica formati numerici per colonne specifiche
    const evenCols = [1, 3, 5, 7, 9, 11, 13, 15, 17, 19]; // Colonne numeri
    const oddCols = [2, 4, 6, 8, 10, 12, 14, 16, 18, 20]; // Colonne percentuali

    evenCols.forEach(col => {
      sheet.getRange(7, lc + col, lr, 1).setNumberFormat('00');
    });

    oddCols.forEach(col => {
      sheet.getRange(7, lc + col, lr, 1).setNumberFormat('0%');
    });

    // 4. Applica stili (UNA SOLA CHIAMATA)
    mainRange.setFontSize(14).setHorizontalAlignment("right");

    // 5. Applica note (ottimizzato)
    for (let r = 0; r < outputNotes.length; r++) {
      if (outputNotes[r][0]) {
        sheet.getRange(7 + r, lc + 1).setNote(outputNotes[r][0]);
      }
    }

    // 6. Larghezza colonne (UNA SOLA CHIAMATA)
    //sheet.setColumnWidths(lc + 2, 10, 50); // Tutte le colonne percentuali insieme
    //const lc = sheet.getLastColumn(); // L'ultima colonna, che ora include i tuoi nuovi dati

    // Definisci un intervallo che si trova nella parte destra, ad esempio la prima cella della prima riga di output
    // Qui assumiamo che i tuoi dati inizino dalla riga 7 e dalla colonna lc + 1.
    const targetRange = sheet.getRange(7, lc + 20, 1, 1); // Riga 7, Colonna lc+1, 1 riga, 1 colonna

    // Imposta l'intervallo come attivo. Questo sposterà la vista.
    sheet.setActiveRange(targetRange);

    // Resize columns
    //sheet.autoResizeColumns(lc, lc +20);
    Logger.log('inzio lc=' + lc);
    const lcFinal = sheet.getLastColumn();
    Logger.log('fine lc=' + lcFinal);
    sheet.setColumnWidths(lc + 1, lcFinal - lc, 50);
  } else {
    SpreadsheetApp.getUi().alert(translate('viewCalendarPage.wrongSheet'));
  }
}

function countColumnsBackgroundNuova() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lc = sheet.getLastColumn();
  const struttureArray = getStruttureInfo(sheet, strutture()).struttureArray;
  const lr = struttureArray.length - 1; // tutte tranne ultima

  if (lr === 0) return; // Nessuna struttura da processare

  //Logger.log(sheet.getRange(5, 2).getValue());
  //Logger.log(typeof (sheet.getRange(5, 2).getValue()));
  if (sheet.getRange(5, 2).getValue() instanceof Date) {

    const mr = sheet.getLastRow();
    const supTot = struttureArray.slice(0, lr).reduce((sum, s) => sum + s[1], 0);

    // Ottieni tutte le date dalla riga 5 in una sola operazione
    const dateRange = sheet.getRange(5, 2, 1, lc - 1);
    const dateValues = dateRange.getValues()[0];

    // Identifica le colonne valide (con date) e visibili
    const validColumns = [];
    for (let c = 0; c < lc - 1; c++) {
      const colIndex = c + 2; // +2 perché iniziamo dalla colonna B
      const isHidden = sheet.isColumnHiddenByUser(colIndex);
      const hasDate = isDateCell(dateValues[c]);

      if (!isHidden && hasDate) {
        validColumns.push(c);
      }
    }

    if (validColumns.length === 0) {
      Logger.log("Nessuna colonna valida trovata");
      return;
    }

    // Leggi tutti i dati necessari in batch
    const dataRange = sheet.getRange(7, 2, lr, lc - 1);
    const colors = dataRange.getBackgrounds();
    const texts = dataRange.getValues();

    // Inizializza contatori per le colonne valide
    const countersPerDay = {};
    validColumns.forEach(colIndex => {
      countersPerDay[colIndex] = {
        countAll: 0,
        countMulti: 0,
        countEvConf: 0,
        countAdConf: 0,
        countEvOpz: 0,
        countAdOpz: 0,
        countEvOff: 0,
        countAdOff: 0,
        supAll: 0,
        supMulti: 0,
        supEvConf: 0,
        supAdConf: 0,
        supEvOpz: 0,
        supAdOpz: 0,
        supEvOff: 0,
        supAdOff: 0
      };
    });

    let countSup0 = 0, countSup020 = 0, countSup2040 = 0, countSup4060 = 0, countSup6080 = 0, countSup80100 = 0;

    // Processa solo le colonne valide
    validColumns.forEach(c => {
      for (let r = 0; r < lr; r++) {
        const color = colors[r][c];
        const text = texts[r][c];
        const superficie = struttureArray[r][1];
        const tipo = classificaCella(color, text);

        if (tipo !== "none") {
          const counter = countersPerDay[c];
          counter.countAll++;
          counter.supAll += superficie;

          switch (tipo) {
            case "multi":
              counter.countMulti++;
              counter.supMulti += superficie;
              break;
            case "ev_conf":
              counter.countEvConf++;
              counter.supEvConf += superficie;
              break;
            case "ad_conf":
              counter.countAdConf++;
              counter.supAdConf += superficie;
              break;
            case "ev_opz":
              counter.countEvOpz++;
              counter.supEvOpz += superficie;
              break;
            case "ad_opz":
              counter.countAdOpz++;
              counter.supAdOpz += superficie;
              break;
            case "ev_off":
              counter.countEvOff++;
              counter.supEvOff += superficie;
              break;
            case "ad_off":
              counter.countAdOff++;
              counter.supAdOff += superficie;
              break;
          }
        }
      }

      // Calcolo riepilogo superficie occupata per intervallo
      const percSup = countersPerDay[c].supAll / supTot * 100;
      if (percSup < 1) countSup0++;
      else if (percSup < 21) countSup020++;
      else if (percSup < 41) countSup2040++;
      else if (percSup < 61) countSup4060++;
      else if (percSup < 81) countSup6080++;
      else countSup80100++;
    });

    // Scrivi intestazioni
    sheet.getRange(mr, 1).setValue('Stats Giornaliere:');
    const labels = [
      ["Unità totali", 1],
      ["% di tutti gli eventi", 2],
      ["Per multi eventi", 3],
      ["% per multi eventi", 4],
      ["Per eventi confermati", 5],
      ["% per eventi confermati", 6],
      ["Per A+D eventi confermati", 7],
      ["% per eventi confermati", 8],
      ["Per eventi opzionati", 9],
      ["% per eventi opzionati", 10],
      ["Per A+D eventi opzionati", 11],
      ["% per eventi opzionati", 12],
      ["Per eventi offerta", 13],
      ["% per eventi offerta", 14],
      ["Per A+D eventi offerta", 15],
      ["% per eventi offerta", 16],
      ["% Sup. Tot. Occ.:", 17],
      ["(su " + supTot + " mq)", 18]
    ];

    labels.forEach(([txt, offset]) => {
      sheet.getRange(mr + offset, 1).setValue(txt).setFontSize(14).setHorizontalAlignment("right");
    });

    // Scrivi dati solo per le colonne valide
    validColumns.forEach(c => {
      const d = countersPerDay[c];
      const dayCol = c + 2;
      const baseRow = mr + 1;

      // Crea info dettagliata per la nota
      const info =
        `${convertDate(dateValues[c])}\n\n` +
        `Unità occupate:\n` +
        `Totale = ${d.countAll} (${(d.countAll / lr * 100).toFixed(0)}%)\n` +
        `Multi eventi = ${d.countMulti} (${(d.countMulti / lr * 100).toFixed(0)}%)\n` +
        `Eventi confermati = ${d.countEvConf} (${(d.countEvConf / lr * 100).toFixed(0)}%)\n` +
        `A+D confermati = ${d.countAdConf} (${(d.countAdConf / lr * 100).toFixed(0)}%)\n` +
        `Eventi opzionati = ${d.countEvOpz} (${(d.countEvOpz / lr * 100).toFixed(0)}%)\n` +
        `A+D opzionati = ${d.countAdOpz} (${(d.countAdOpz / lr * 100).toFixed(0)}%)\n` +
        `Eventi offerta = ${d.countEvOff} (${(d.countEvOff / lr * 100).toFixed(0)}%)\n` +
        `A+D offerta = ${d.countAdOff} (${(d.countAdOff / lr * 100).toFixed(0)}%)\n\n` +
        `Superficie occupata:\n` +
        `Totale = ${d.supAll} mq (${(d.supAll / supTot * 100).toFixed(0)}%)`;

      // Scrivi i valori solo se > 0
      const values = [
        [d.countAll, '00'],
        [d.countAll / lr, '0%'],
        [d.countMulti, '00'],
        [d.countMulti / lr, '0%'],
        [d.countEvConf, '00'],
        [d.countEvConf / lr, '0%'],
        [d.countAdConf, '00'],
        [d.countAdConf / lr, '0%'],
        [d.countEvOpz, '00'],
        [d.countEvOpz / lr, '0%'],
        [d.countAdOpz, '00'],
        [d.countAdOpz / lr, '0%'],
        [d.countEvOff, '00'],
        [d.countEvOff / lr, '0%'],
        [d.countAdOff, '00'],
        [d.countAdOff / lr, '0%'],
        [d.supAll / supTot, '0%']
      ];

      values.forEach(([value, format], i) => {
        if (value > 0 || i === 0 || i === 1 || i === 16) { // Sempre mostra totali e superficie
          const cell = sheet.getRange(baseRow + i, dayCol);
          cell.setValue(value)
            .setNumberFormat(format)
            .setFontSize(i === 16 ? 12 : 14)
            .setHorizontalAlignment("right");

          if (i === 0) { // Aggiungi nota solo alla prima cella
            cell.setNote(info);
          }
        }
      });
    });

    // Riepilogo finale (solo se ci sono colonne valide)
    if (validColumns.length > 0) {
      const firstValidCol = Math.min(...validColumns);
      const lastValidCol = Math.max(...validColumns);
      const firstDay = dateValues[firstValidCol];
      const lastDay = dateValues[lastValidCol];

      // Usa la prima colonna visibile con data come punto di partenza per il riepilogo
      const firstVisibleDateCol = firstValidCol + 2; // +2 perché validColumns è 0-based ma le colonne del foglio partono da 2

      sheet.getRange(mr + 19, firstVisibleDateCol).setValue("RIEPILOGO DELLA SUPERFICIE GIORNALIERA OCCUPATA").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("left");
      sheet.getRange(mr + 20, firstVisibleDateCol).setValue("CON RELATIVA PERCENTUALE NEL PERIODO VISUALIZZATO:").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("left");
      sheet.getRange(mr + 19, 1).setValue(firstDay).setNumberFormat('\\DAL: dd/MM/yyyy').setFontSize(16).setHorizontalAlignment("left");
      sheet.getRange(mr + 20, 1).setValue(lastDay).setNumberFormat('AL: dd/MM/yyyy').setFontSize(16).setHorizontalAlignment("left");

      const summaryData = [
        ["Fino al 20%", countSup020],
        ["Fino al 40%", countSup2040],
        ["Fino al 60%", countSup4060],
        ["Fino al 80%", countSup6080],
        ["Fino al 100%", countSup80100],
        ["Non occupato per", countSup0]
      ];

      summaryData.forEach(([label, value], i) => {
        const row = mr + 21 + i;
        const percentage = value / validColumns.length;

        sheet.getRange(row, firstVisibleDateCol)
          .setValue(`${label} (${value} giorni)`)
          .setFontSize(12)
          .setHorizontalAlignment("left");

        sheet.getRange(row, firstVisibleDateCol + 8)
          .setValue(percentage)
          .setNumberFormat('0%')
          .setFontSize(12)
          .setFontWeight("bold")
          .setHorizontalAlignment("right");
      });

      sheet.setColumnWidths(firstVisibleDateCol, 1, 70.0);
      sheet.setColumnWidths(firstVisibleDateCol + 3, 1, 70.0);
    }
    const lrFinal = sheet.getLastRow(); // L'ultima riga, che ora include i tuoi nuovi dati

    // Definisci un intervallo che si trova nella parte destra, ad esempio la prima cella della prima riga di output
    // Qui assumiamo che i tuoi dati inizino dalla riga 7 e dalla colonna lc + 1.
    const targetRange = sheet.getRange(lrFinal, 1, 1, 1); // Riga lr, Colonna 1, 1 riga, 1 colonna

    // Imposta l'intervallo come attivo. Questo sposterà la vista.
    sheet.setActiveRange(targetRange);
  } else {
    SpreadsheetApp.getUi().alert(translate('viewCalendarPage.wrongSheet'));
  }
}
// END STATS CODE