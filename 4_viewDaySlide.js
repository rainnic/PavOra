/*
* Project Name: Pavora
* Copyright (c) 2025 Nicola Rainiero
*
* This software is released under the MIT License.
* Please refer to the LICENSE file for the full license text.
*/
function testInsertExSlide() {
  //createSlideAndExportToSheet('2025-09-10', '2025-09-11', 'cc', '', 'SI'); // cc o quartiere
  var start = Date.now();  
  createSlideAndExportToSheetImproved('2025-09-10', '2025-09-11', 'cc', '', 'SI');
  /*
        .createSlideAndExportToSheetImproved(
            startDate, 
            finishDate, 
            cosa, 
            keyword, 
            editable,
            colorAssignments  // 👈 PASSA LA MAPPA COLORI
        );
  */
  var end = Date.now();
  Logger.log("Tempo di esecuzione funzione: " + (end - start) / 1000 + " secondi");  
}

function copyAndRenamePresentation(presentationId, newPresentationName) {
  // Ottieni il file della presentazione
  var file = DriveApp.getFileById(presentationId);

  // Copia il file
  var newFile = file.makeCopy(newPresentationName);

  // Ottieni il nuovo ID della presentazione copiata
  var newPresentationId = newFile.getId();

  // Rinomina la nuova presentazione
  var newPresentation = SlidesApp.openById(newPresentationId);
  //newPresentation.rename(newPresentationName);

  return newPresentationId;
}

// ----------------------------------------------------------------------
// Funzione per copiare le slide nel foglio sheet
// ----------------------------------------------------------------------
function createSlideAndExportToSheet(first, last, cosa, keyword, editable) {
  // Inizializza il foglio
  // Nuovo metodo
  resetFoglioConNuovo();

  // Step 1: inizializzare il foglio ed eliminare le immagini presenti
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // togliere le celle congelate
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  // Inserisce un punto per mantenere la tabella strutturata
  sheet.getRange(600, 50).setValue('.');

  // Rimuovere l'immagine esistente
  var images = sheet.getImages();
  for (let i = 0; i < images.length; i++) {
    images[i].remove();
  }

  // ID della presentazione esistente
  // quartiere = 1OB-N_Msm_ModXYzQdi5LDrLSMXerQK6ZaPJI5ooKREA
  // CC = 1kclQIKUAMhk1kFvcgxFbfB6U9k4rBwm0Pu097ksK_Yc

  // Step 1: Creare una nuova presentazione Google Slides
  var today = new Date;
  var presentationName = formatDateMaster(today).dataXfile + translate('planPage.slideFile') + cosa;
  if (cosa === 'cc') {
    var presentationId = copyAndRenamePresentation(templateSlides()[1][0], presentationName);
  } else {
    var presentationId = copyAndRenamePresentation(templateSlides()[0][0], presentationName);
  }


  // Step 2: Aprire la presentazione esistente


  // Per rinizializzare la Slide

  var dateParts = first.split("-");
  var first = new Date(+dateParts[0], dateParts[1] - 1, +dateParts[2]);

  var dateParts = last.split("-");
  var last = new Date(+dateParts[0], dateParts[1] - 1, +dateParts[2]);
  var last = new Date(last.getTime() + 1 * 3600000 * 24);

  // Contare i giorni tra le due date
  var numMilliSec = last.getTime() - first.getTime();
  var numDays = (numMilliSec / (1000 * 3600 * 24) == 1) ? 0 : numMilliSec / (1000 * 3600 * 24);

  viewEvents(1, first, numDays, selectedMode(), presentationId, cosa, keyword);
  if (numDays == 0) { numDays += 1 }

  // Step 3: Riapertura per leggere le nuove slide!!!
  presentation = SlidesApp.openById(presentationId);

  // Contare le slide presenti
  var slides = presentation.getSlides();
  var numSlides = slides.length;

  for (let i = 0; i < numDays; i++) {

    if (numSlides > numDays) {
      if ((numSlides - i - numDays) >= slides.length) {
        continue;
      }

      var slide = slides[(numSlides - i - numDays)];

      if (!slide) {
        continue;
      }
    } else {
      var slide = slides[(i)];
    }

    // Step 3: Esportare la presentazione come immagine
    var exportUrl = 'https://docs.google.com/presentation/d/' + presentationId + '/export/png?pageid=' + slide.getObjectId();

    var params = {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
      }
    };
    var response = UrlFetchApp.fetch(exportUrl, params);
    var blob = response.getBlob();

    // Step 4: Controlla e ridimensiona il blob se necessario
    var imageBlob = Utilities.newBlob(blob.getBytes(), blob.getContentType(), 'slide.png');
    if (blob.getBytes().length > 2000000) {
      imageBlob = imageBlob.getAs('image/png').resize(595, 842); // Ridimensiona l'immagine
    }

    // Step 5: Inserire l'immagine nel foglio Google Sheets

    // Trova l'ultima riga con un'immagine
    var lastRow = 2; // Iniziare dalla riga 1
    var images = sheet.getImages();
    if (images.length > 0) {
      lastRow = images[images.length - 1].getAnchorCell().getRow() + 53; // Inserire 20 righe sotto l'ultima immagine
    }

    sheet.insertImage(imageBlob, 1, lastRow); // Inserisce l'immagine nella nuova riga
  }

  // Nasconde la griglia
  sheet.setHiddenGridlines(true);

  var lc = 25; //sheet.getLastColumn();
  var mc = sheet.getMaxColumns();
  if (mc - lc != 0) {
    sheet.deleteColumns(lc + 1, mc - lc);
  } else {
    sheet.deleteColumns(7, 1);
  }


  var lr = 300; //sheet.getLastRow();
  var mr = sheet.getMaxRows();
  if (lr - 2 != 0) {
    sheet.deleteRows(lr + 1, mr - lr);
  }

  // ATTENZIONE: Per cancellare la presentazione!
  if (editable === 'NO') {
    DriveApp.getFileById(presentationId).setTrashed(true);
  } else {
    // Ottieni il link diretto al file
    var fileLink = "https://drive.google.com/file/d/" + presentationId + "/view";

    // Scrivi il link nella cella A1 del foglio attivo
    sheet.getRange('A1').setValue(translate('planPage.slideFileAlert')).setFontSize(10).setHorizontalAlignment('right');
    sheet.getRange('B1').setValue(fileLink).setFontSize(10);
  }
}

function filterEvents(events) {
  // Crea un oggetto per raggruppare gli eventi per data
  const eventsByDate = {};

  // Itera attraverso gli eventi per raggrupparli per data
  events.forEach(event => {
    const date = new Date(event[0]).toISOString().split('T')[0]; // Estrai solo la parte della data (senza tempo)
    if (!eventsByDate[date]) {
      eventsByDate[date] = [];
    }
    eventsByDate[date].push(event);
  });

  const filteredEvents = [];

  // Itera attraverso i giorni nell'oggetto
  for (const date in eventsByDate) {
    const dayEvents = eventsByDate[date];
    const titles = {};

    // Raggruppa gli eventi per titolo, dando priorità agli eventi con la lettera "E"
    dayEvents.forEach(event => {
      const title = event[2]; // Titolo
      const letter = event[1]; // Lettera

      if (!titles[title]) {
        titles[title] = event;
      } else if (letter === 'E') {
        titles[title] = event;
      }
    });

    // Aggiungi gli eventi filtrati alla matrice dei risultati
    for (const title in titles) {
      filteredEvents.push(titles[title]);
    }
  }

  return filteredEvents;
}


// ----------------------------------------------------------------------
//
// View events per day
//
// ----------------------------------------------------------------------
//
//    --->   function viewEvents(user, start date, number of day to show after start date, method);
//
// what: (define how the events are displayed on the sheet)
//  "0" = by URL;        (a new sheet where events are defined by increasing numbers and by two colors: one for approved and one for optionated)
//  "1" = by Google Drive;   (a new sheet where events are defined by increasing numbers and by two colors: one for approved and one for optionated)
//  "2"  = by letters/symbol;        (over the existing sheet, the events are defined only by increasing numbers)
//
// In symbol the only way to add a background is manually inside Google Slide:
// right click on the page --> Change background --> Add to theme
function selectedMode() {
  return 2
}

function viewEvents(user, start, numberDays, method, presentationId, cosa, keyword) {

  // ID della presentazione esistente
  var presentationId = presentationId;

  // Step 1: Aprire la presentazione esistente
  var presentation = SlidesApp.openById(presentationId);

  // Load active presentation by ID
  //var presentationId = SlidesApp.getActivePresentation().getId();
  //var presentation = SlidesApp.openById(presentationId);
  // Collect the initial number of slides in order to delete them at the end
  var slides = presentation.getSlides();
  var numSlides = slides.length;
  // Set min transparency for object
  var transparency = setTransparency()[0][0];
  // Set a variable for missing structures
  var vuoto = 0;

  // Load the colors array
  var colori2d = methodMcolors();

  // Load the structures
  // Full array 
  // ['identification code', larghezza, altezza, x, y, immagine associata]
  if (cosa === 'cc') {
    var structures = centroCongressi();
  } else {
    var structures = strutture();
  }
  //Logger.log('Le strutture sono ' + JSON.stringify(structures.length));
  var date = start;
  var today = convertDateBar(date.setDate(date.getDate()));
  var nDay = (numberDays == 0) ? 1 : numberDays; // numero di giorni dopo oggi
  var nextDay = convertDateBar(date.setDate(date.getDate() + nDay));

  var fromDate = today; // example 16/10/2020
  var toDate = nextDay; // example 18/10/2020

  // Load into an array the data collected from Calendar in the prefixed period,
  // for this method the Calendar ID associated is linked to the first category
  var eventi = events2Array(fromDate, toDate, categories()[0][0], keyword);

  var fDate = new Date(fromDate);
  var tDate = new Date(toDate);

  var dateParts = fromDate.split("/");
  var fromDate = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);

  var dateParts = toDate.split("/");
  var toDate = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
  var toDate = new Date(toDate.getTime() + 1 * 3600000 * 24);

  // Count the number of days between the two dates
  var numDays = nDay;

  // Fix the start date
  var pageDate = fromDate;

  // Remove the events that start before the first day and end in that day (for example: 23.00-04.00)
  t = 0;
  while ((eventi.length > 1) && (eventi[t][0].getTime() < fromDate.getTime())) {
    eventi.splice(t, 1);
    t++
  }

  // Set counters and variables
  var y = 0; // counter for the items that form the menu on the top left
  var j = 0; // counter for the structures that cover plan
  count = 0; // counter to rispect the usage limits of 100 Write Requests per user per 100 seconds (https://developers.google.com/slides/limits)

  // Create a new array with only one event with the same name to prevent extra work
  var filteredEvents = filterEvents(eventi);

  //
  // For loop to print a slide per day
  //
  var n = 1;
  for (let i = 0; i < numDays; i += 1) {

    // random is created to generate objects with various IDs
    var random = randomID(2);

    // Add a page with header and date
    //--> CreateSlide(presentationId, pageNum, data, Width, Height, mode, cosa) cosa = quartiere o cc
    CreateSlide(presentationId, i + numSlides, random, pageDate, 134, 9, method, cosa);

    // Add the menu on the top left
    var w = 0; // counter for the spacing among the items that form the menu on the top left
    while ((y < eventi.length) && (convertDate(eventi[y][0]) == convertDate(pageDate))) {
      if (cosa === 'cc') {
        CreateBoxText(presentationId, i + numSlides, random, pageDate, 'MENU', eventi[y][3], 55, 15.0, 152, 15 + w, colori2d[selectHigh(eventi[y][7], eventi[y][1])[0]][selectHigh(eventi[y][7], eventi[y][1])[1]], '0', eventi[y][4], parseEventDetails(eventi[y][8]).descrizione, transparency, method, cosa);
        w = w + 16.25; // 25 --> 26.25 provo 15 --> 16.25
      } else if (cosa === 'quartiereVecchio') {
        CreateBoxText(presentationId, i + numSlides, random, pageDate, 'MENU', eventi[y][3], 200, 6.0, 5, 15 + w, colori2d[selectHigh(eventi[y][7], eventi[y][1])[0]][selectHigh(eventi[y][7], eventi[y][1])[1]], '0', eventi[y][4], parseEventDetails(eventi[y][8]).descrizione, transparency, method, cosa);
        w = w + 7.25;
      } else {
        CreateBoxText(presentationId, i + numSlides, random, pageDate, 'MENU', eventi[y][3], 60, 13.5, 1, 2.5 + w, colori2d[selectHigh(eventi[y][7], eventi[y][1])[0]][selectHigh(eventi[y][7], eventi[y][1])[1]], '0', eventi[y][4], parseEventDetails(eventi[y][8]).descrizione, transparency, method, cosa);
        w = w + 14.25;
      }
      count++;
      y++;
    }

    // Add the structures that cover plan (use filteredEvents!!)
    var presence = [];
    const ccNotInMappa = ['GALL', 'EMPTY', 'foyerSG', 'foyerSMU', 'foyerSM', 'foyerMe1', 'foyerMe2', 'bistro', 'loggia', 'lobbyQ4', 'ufficiQ8', 'lbar'];
    while ((j < filteredEvents.length) && (convertDate(filteredEvents[j][0]) == convertDate(pageDate))) {
      if (filteredEvents[j][10].length > 0) {
        for (let k = 0; k < filteredEvents[j][10].length; k += 1) {

          if (findKey(filteredEvents[j][10][k], structures, 0) >= 0) {
            var riga = findKey(filteredEvents[j][10][k], structures, 0);
            // Fix the position if the structure is already present
            if (findKey(filteredEvents[j][10][k], presence, 0) < 0) {
              presence.push([filteredEvents[j][10][k], 0]);
              posX = (Number(structures[riga][3]));
              posY = (Number(structures[riga][4]));
              //Logger.log(presence);
            } else {
              presence[findKey(filteredEvents[j][10][k], presence, 0)][1] = presence[findKey(filteredEvents[j][10][k], presence, 0)][1] - 2;
              pos = presence[findKey(filteredEvents[j][10][k], presence, 0)][1];
              posX = (Number(structures[riga][3]) + pos);
              posY = (Number(structures[riga][4]) + pos);
            }
            width = (Number(structures[riga][1]) * 1.0);
            height = (Number(structures[riga][2]) * 1.0);
            if (((cosa == 'quartiere') && (!ccNotInMappa.includes(filteredEvents[j][10][k]))) || (cosa == 'cc')) {
              if (width != 0) {
                CreateBoxText(presentationId, i + numSlides, random, pageDate, filteredEvents[j][10][k], filteredEvents[j][3], width, height, posX, posY, colori2d[selectHigh(filteredEvents[j][7], filteredEvents[j][1])[0]][selectHigh(filteredEvents[j][7], filteredEvents[j][1])[1]], findImage(structures[riga][5], imageFinder(), selectedMode()), filteredEvents[j][4], filteredEvents[j][9], transparency, method, cosa);
                count++;
              };
            }
          }
        }
      } //If length of structures is > of 0
      j++;
    }
    // Add a day and begin another slide or stop
    pageDate = new Date(pageDate.getTime() + 1 * 3600000 * 24);

    // Pause if necessary
    if (count > 75) { // 50
      //if (n < (numDays - 1)) {
      Utilities.sleep(100100); // in  milliseconds 100100
      count = 0; // 100 per user per 100 seconds (https://developers.google.com/slides/limits)
    };
    n += 1;
  }
  //deleteSlidesUntil(numSlides, presentationId); // per cancellare le slide pre-esistenti
  Utilities.sleep(1000);
  presentation.saveAndClose();
  Utilities.sleep(1000);
}
// ----------------------------------------------------------------------
//
// FINE View events per day
//
// ----------------------------------------------------------------------

// ---------------------------------------------------------
//
// Crea una slide vuota con l'immagine di background scelta
//
// ----------------------------------------------------------
function CreateSlide(presentationId, pageNum, random, data, Width, Height, mode) {
  var pageId = random + 'pageNum' + pageNum.toString();
  var dataScelta = random + convertDateClean(data) + '_' + pageNum.toString();
  var W = {
    magnitude: Width * 2.835,
    unit: 'PT'
  };
  var H = {
    magnitude: Height * 2.835,
    unit: 'PT'
  };
  if (mode == 0) {
    var imageUrl = findImage('PLAN', imageFinder(), 0);
  } else if (mode == 1) {
    var sourceFolderId = driveIDFolder()[2][0];
    var folder = DriveApp.getFolderById(sourceFolderId);
    var background = folder.getFilesByName(imageFinder()[6][1]).next();
    background.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var imageUrl = "https://drive.google.com/uc?export=download&id=" + background.getId();
  }
  var update_request = [{
    createSlide: {
      objectId: pageId,
      insertionIndex: pageNum,
      slideLayoutReference: {
        predefinedLayout: "BLANK"
      }
    }
  },
  {
    createShape: {
      objectId: dataScelta,
      shapeType: 'TEXT-BOX',

      elementProperties: {
        pageObjectId: pageId,
        size: {
          height: H,
          width: W
        },
        transform: {
          scaleX: 1,
          scaleY: 1,
          translateX: 75 * 2.835, // 66
          translateY: 0.25 * 2.835, // 5
          unit: 'PT'
        }
      }
    }
  },
  // Insert text into the box, using the supplied element ID.
  {
    insertText: {
      objectId: dataScelta,
      insertionIndex: 0,
      text: translate('planPage.dailySpace') + convertDate(data) + ')'
    }
  },
  {
    updateTextStyle: {
      objectId: dataScelta,
      style: {
        backgroundColor: {
          opaqueColor: {
            rgbColor: {
              red: hexToRgb("#ffffff").r,
              green: hexToRgb("#ffffff").g,
              blue: hexToRgb("#ffffff").b
            }
          }
        },
        foregroundColor: {
          opaqueColor: {
            rgbColor: {
              red: hexToRgb("#f31f4a").r,
              green: hexToRgb("#f31f4a").g,
              blue: hexToRgb("#f31f4a").b
            }
          }
        },
        bold: true,
        fontSize: {
          magnitude: 16,
          unit: 'PT'
        }
      },
      fields: 'backgroundColor.opaqueColor.rgbColor, foregroundColor.opaqueColor.rgbColor, bold, fontSize'
    }
  },
  {
    updateShapeProperties: {
      objectId: dataScelta,
      shapeProperties: {
        contentAlignment: 'MIDDLE'
      },
      fields: 'contentAlignment'
    }
  },
  {
    updateParagraphStyle: {
      objectId: dataScelta,
      style: {
        alignment: 'END'
      },
      fields: 'alignment'
    }
  }
  ];
  // Execute the request.
  var response = Slides.Presentations.batchUpdate({
    requests: update_request
  }, presentationId);

  if ((mode == 0) || (mode == 1)) {
    var update_request1 = [
      {
        updatePageProperties: {
          objectId: pageId,
          pageProperties: {
            pageBackgroundFill: {
              stretchedPictureFill: {
                contentUrl: imageUrl
              }
            }
          },
          fields: "pageBackgroundFill"
        }
      }
    ];
  }
  if ((mode == 0) || (mode == 1)) {
    // Execute the request.
    var response = Slides.Presentations.batchUpdate({
      requests: update_request1
    }, presentationId);
  }
  if (mode == 1) {
    background.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  }
}

// ----------------------------------------------------------
// Crea un box con del testo all'interno di una immagine
// ----------------------------------------------------------
function CreateBoxText(presentationId, pageNum, random, data, structure, event, Width, Height, posX, posY, color, imageText, stile, textContent, trasparenza, mode, cosa) {
  var ui = SpreadsheetApp.getUi();

  try {


    var pageId = random + 'pageNum' + pageNum.toString();
    var elementId = random + convertDateClean(data) + '_' + pageNum.toString() + structure + event.substring(0, 16) + event.substring(event.length - 1, event.length) + randomID(2);
    var elementIdText = random + convertDateClean(data) + '_' + pageNum.toString() + structure + event.substring(0, 16) + event.substring(event.length - 1, event.length) + randomID(2) + '_t';
    //var textContent = 'PROVA';
    if ((structure == 'MENU')) { //||(imageText !== 0)
      var allinea = 'START';
      var postest = 'MIDDLE';
      var taglia = 1;
    } else {
      var allinea = 'CENTER';
      var postest = selectPos(stile);
      var taglia = selectSize(stile);
    }
    if (stile === 'DOT') { var trasparenza = setTransparency()[1][0] };

    var W = {
      magnitude: Width * 2.835 * taglia,
      unit: 'PT'
    };
    var H = {
      magnitude: Height * 2.835 * taglia,
      unit: 'PT'
    };

    var imageUrl = imageText;
    if (cosa === 'cc') {
      var altTesto = 10;
    } else if (cosa === 'quartiereVecchio') {
      var altTesto = min(Width, Height) * 2;
    } else {
      var altTesto = 10;
    }

    if ((mode == 2) && (structure !== "MENU")) {
      var textContent = imageText;
    }
    if ((stile === "DOT") && (structure != "MENU")) {
      if (mode == 1) {
        var imageText = findImage('wip', imageFinder(), mode);
      } else if (mode == 0) {
        imageText = findImage('wip', imageFinder(), mode);
        var imageUrl = imageText;
      } else if (mode == 2) {
        var textContent = findImage('wip', imageFinder(), mode);
      }
    }

    var update_request = [
      {
        createShape: {
          objectId: elementIdText,
          shapeType: 'RECTANGLE',

          elementProperties: {
            pageObjectId: pageId,
            size: {
              height: H,
              width: W
            },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: posX * 2.835,
              translateY: posY * 2.835,
              unit: 'PT'
            }
          }
        }
      },
      // Insert text into the box, using the supplied element ID.
      {
        insertText: {
          objectId: elementIdText,
          insertionIndex: 0,
          text: textContent
        }
      },
      {
        updateTextStyle: {
          objectId: elementIdText,
          style: {
            foregroundColor: {
              opaqueColor: {
                rgbColor: {
                  red: hexToRgb("#000000").r,
                  green: hexToRgb("#000000").g,
                  blue: hexToRgb("#000000").b
                }
              }
            },
            bold: true,
            fontSize: {
              magnitude: altTesto,
              unit: 'PT'
            }
          },
          fields: 'backgroundColor.opaqueColor.rgbColor, foregroundColor.opaqueColor.rgbColor, bold, fontSize'
        }
      },
      {
        updateShapeProperties: {
          objectId: elementIdText,
          shapeProperties: {
            contentAlignment: postest
          },
          fields: 'contentAlignment'
        }
      },
      {
        updateParagraphStyle: {
          objectId: elementIdText,
          style: {
            alignment: allinea
          },
          fields: 'alignment'
        }
      },

      {
        updateShapeProperties: {
          objectId: elementIdText,
          shapeProperties: {
            shapeBackgroundFill: {
              solidFill: {
                color: {
                  rgbColor: {
                    red: hexToRgb(color).r,
                    green: hexToRgb(color).g,
                    blue: hexToRgb(color).b
                  }
                },
                alpha: trasparenza
              },
            },
            outline: {
              outlineFill: {
                solidFill: {
                  color: {
                    rgbColor: {
                      red: hexToRgb('#000000').r,
                      green: hexToRgb('#000000').g,
                      blue: hexToRgb('#000000').b
                    }
                  },
                  alpha: trasparenza
                }
              },
              weight: {
                magnitude: 1,
                unit: 'PT'
              },
              dashStyle: stile
            },
          },
          fields: 'shapeBackgroundFill.solidFill.color, shapeBackgroundFill.solidFill.alpha, outline.outlineFill.solidFill.color, outline.weight, outline.dashStyle'
        }
      }
    ];

    // Per caricare un immagine

    if ((imageText !== "0") && ((mode == 0) || (mode == 1))) {
      if (mode == 1) {
        var sourceFolderId = driveIDFolder()[2][0];
        var folder = DriveApp.getFolderById(sourceFolderId);
        var image = folder.getFilesByName(imageText).next();
        image.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        var imageUrl = "https://drive.google.com/uc?export=download&id=" + image.getId();
      } else {
        var imageUrl = imageText;
      }
    };

    if ((imageText !== "0") && ((mode == 0) || (mode == 1))) {
      // if (!imageText) {
      var update_request1 = [
        {
          createImage: {
            objectId: elementId,
            url: imageUrl,
            elementProperties: {
              pageObjectId: pageId,
              size: {
                height: H,
                width: W
              },
              transform: {
                scaleX: 1,
                scaleY: 1,
                translateX: posX * 2.835,
                translateY: posY * 2.835,
                unit: 'PT'
              }
            }
          }
        }
      ];
    }

    // Execute the request.
    var response = Slides.Presentations.batchUpdate({
      requests: update_request
    }, presentationId);


    // Per caricare un immagine
    if ((imageText !== "0") && ((mode == 0) || (mode == 1))) {
      var response = Slides.Presentations.batchUpdate({
        requests: update_request1
      }, presentationId);
      if (mode == 1) {
        image.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); // Added
      }
    }


  } catch (error) {
    // Mostra un messaggio tramite ui.alert
    SpreadsheetApp.getUi().alert(translate('alert.errorMessage') + ' (' + error.message + ')');
  }
}

// ------------------------------------ New ------------------------------------------------
function mergePresentationsClean(sourcePresentationIds, templateId, finalName) {

  // Crea presentazione finale dal template
  var finalPresentationId = copyAndRenamePresentation(templateId, finalName);
  Logger.log('✓ Creata presentazione finale: ' + finalPresentationId);

  var finalPresentation = SlidesApp.openById(finalPresentationId);

  // Rimuove TUTTE le slide del template
  finalPresentation.getSlides().forEach(slide => slide.remove());
  finalPresentation.saveAndClose();

  Utilities.sleep(1000);

  // Copia slide 1:1 con API ufficiale
  sourcePresentationIds.forEach((sourceId, i) => {

    Logger.log('--- Copia da presentazione ' + (i + 1) + '/' + sourcePresentationIds.length);

    var sourcePresentation = SlidesApp.openById(sourceId);
    var sourceSlides = sourcePresentation.getSlides();

    sourceSlides.forEach((slide, j) => {

      Slides.Presentations.Pages.copy(
        { presentationId: finalPresentationId },
        sourceId,
        slide.getObjectId()
      );

      Logger.log('  ✓ Slide ' + (j + 1) + ' copiata');
      Utilities.sleep(300);
    });

    sourcePresentation.saveAndClose();
  });

  Logger.log('✓ Unione completata');
  return finalPresentationId;
}

function olDCreateSlideAndExportToSheetImproved(first, last, cosa, keyword, editable) {

  resetFoglioConNuovo();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // togliere le celle congelate
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);

  // Nasconde la griglia
  sheet.setHiddenGridlines(true);

  // -------------------------
  // DATE
  // -------------------------
  var f = first.split('-');
  var l = last.split('-');

  var firstDate = new Date(+f[0], f[1] - 1, +f[2]);
  var lastDate  = new Date(+l[0], l[1] - 1, +l[2] + 1);

  var numDays = (lastDate - firstDate) / (1000 * 3600 * 24);

  Logger.log('INIZIO PROCESSO');
  Logger.log('Numero giorni: ' + numDays);
  Logger.log('Editable: ' + editable);

  // -------------------------
  // FASE 1 — SLIDE GIORNALIERE
  // -------------------------
  var presentationIds = [];
  var currentDate = new Date(firstDate);

  for (let d = 0; d < numDays; d++) {

    var dateString = formatDateForFilename(currentDate);
    var name = dateString + ' - TEMP - ' + cosa;

    var templateId = (cosa === 'cc')
      ? templateSlides()[1][0]
      : templateSlides()[0][0];

    var presId = copyAndRenamePresentation(templateId, name);

    viewEvents(
      1,
      new Date(currentDate),
      1,
      selectedMode(),
      presId,
      cosa,
      keyword
    );

    presentationIds.push(presId);
    Utilities.sleep(1500);

    currentDate.setDate(currentDate.getDate() + 1);
  }

  Logger.log('Slide giornaliere create: ' + presentationIds.length);

  // -------------------------
  // FASE 2 — UNIONE (PULITA)
  // -------------------------
// FASE 2: Unione slide (Versione Anti-Quota Exceeded)
var finalPresentationId = null;

Logger.log('==========================================');
Logger.log('Controllo editable: "' + editable + '" === "YES" o "SI" ? ' + (editable === 'YES' || editable === 'SI'));
Logger.log('==========================================');

if (editable === 'YES' || editable === 'SI') {
  Logger.log('=== FASE 2: INIZIO Unione slide (Metodo Nativo) ===');
  
  try {
    var today = new Date();
    // Assicurati che le variabili 'cosa' e le funzioni 'formatDateMaster'/'translate' siano accessibili qui
    var finalName = formatDateMaster(today).dataXfile + translate('planPage.slideFile') + cosa;
    Logger.log('Nome presentazione finale: ' + finalName);
    
    // 1. Crea la presentazione finale
    if (cosa === 'cc') {
      finalPresentationId = copyAndRenamePresentation(templateSlides()[1][0], finalName);
    } else {
      finalPresentationId = copyAndRenamePresentation(templateSlides()[0][0], finalName);
    }
    
    Logger.log('✓ Presentazione finale creata: ' + finalPresentationId);
    
    // 2. Apri la presentazione finale
    var finalPresentation = SlidesApp.openById(finalPresentationId);
    
    // 3. Rimuovi le slide template (se necessario)
    var initialSlides = finalPresentation.getSlides();
    if (initialSlides.length > 0) {
      Logger.log('Rimozione ' + initialSlides.length + ' slide template...');
      for (let i = initialSlides.length - 1; i >= 0; i--) {
        initialSlides[i].remove();
      }
    }
    
    // Salva per consolidare la rimozione
    finalPresentation.saveAndClose();
    Utilities.sleep(2000); // Pausa di sicurezza 2 secondi
    
    // 4. Ciclo sulle presentazioni giornaliere
    Logger.log('Inizio unione da ' + presentationIds.length + ' file...');
    
    for (let i = 0; i < presentationIds.length; i++) {
      try {
        // Riapri la finale ad ogni ciclo principale per evitare disconnessioni
        finalPresentation = SlidesApp.openById(finalPresentationId);
        
        var sourceId = presentationIds[i];
        Logger.log('--- Processando file ' + (i+1) + '/' + presentationIds.length + ' ---');
        
        var sourcePresentation = SlidesApp.openById(sourceId);
        var sourceSlides = sourcePresentation.getSlides();
        
        if (sourceSlides.length === 0) {
          Logger.log('  Attenzione: File sorgente vuoto, salto.');
          sourcePresentation.saveAndClose();
          continue;
        }

        // Copia le slide
        for (let j = 0; j < sourceSlides.length; j++) {
          Logger.log('  Appendendo slide ' + (j+1));
          
          // --- IL CORE DELLA SOLUZIONE ---
          // appendSlide copia TUTTO (formattazione, font, immagini) in 1 colpo solo.
          // Non usa quota API complessa.
          finalPresentation.appendSlide(sourceSlides[j]);
          
          // Pausa breve (2 secondi) tra una slide e l'altra è sufficiente
          Utilities.sleep(2000); 
        }
        
        // Chiudi il file sorgente corrente
        sourcePresentation.saveAndClose();
        
        // Salva la finale parzialmente (utile se ci sono molti file)
        finalPresentation.saveAndClose();
        
        Logger.log('  ✓ File ' + (i+1) + ' completato. Pausa di raffreddamento...');
        // Pausa media tra un file e l'altro (3 secondi)
        Utilities.sleep(3000);
        
      } catch (e) {
        Logger.log('  ✗ Errore nel file ' + (i+1) + ': ' + e.toString());
        // Se fallisce, prova a salvare la finale comunque
        try { finalPresentation.saveAndClose(); } catch(err) {}
      }
    }
    
    // Verifica finale
    finalPresentation = SlidesApp.openById(finalPresentationId);
    var finalCount = finalPresentation.getSlides().length;
    finalPresentation.saveAndClose();
    
    Logger.log('==========================================');
    Logger.log('✓ FASE 2 COMPLETATA');
    Logger.log('Totale slide finali: ' + finalCount);
    Logger.log('ID Finale: ' + finalPresentationId);
    Logger.log('==========================================');

  } catch (e) {
    Logger.log('✗ ERRORE CRITICO FASE 2: ' + e.toString());
    Logger.log(e.stack);
  }
} else {
  Logger.log('=== FASE 2: SALTATA (editable !== YES) ===');
}

  // -------------------------
  // FASE 3 — EXPORT SU SHEET
  // -------------------------
  var exportIds = (finalPresentationId)
    ? [finalPresentationId]
    : presentationIds;

  var lastRow = 2;

  exportIds.forEach(id => {

    var pres = SlidesApp.openById(id);
    pres.getSlides().forEach(slide => {

      var url = 'https://docs.google.com/presentation/d/' +
                id + '/export/png?pageid=' +
                slide.getObjectId();

      var blob = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
      }).getBlob();

      sheet.insertImage(blob, 1, lastRow);
      lastRow += 53;
    });

    pres.saveAndClose();
  });

  // -------------------------
  // FASE 4 — PULIZIA
  // -------------------------
  presentationIds.forEach(id => DriveApp.getFileById(id).setTrashed(true));

  if (finalPresentationId) {
    sheet.getRange('A1').setValue(translate('planPage.slideFileAlert'));
    sheet.getRange('B1').setValue(
      'https://docs.google.com/presentation/d/' +
      finalPresentationId + '/edit'
    );
  }

  Logger.log('PROCESSO COMPLETATO');
}

/**
 * Funzione helper per formattare la data nel nome file
 */
function formatDateForFilename(date) {
  var year = date.getFullYear();
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var day = ('0' + date.getDate()).slice(-2);
  return year + '-' + month + '-' + day;
}

//
// PER COLORAZIONE PLAN
//
/**
 * ==================================================
 * FUNZIONI APPS SCRIPT PER GESTIONE COLORI
 * ==================================================
 */

/**
 * Restituisce lista eventi unici nel periodo selezionato
 * @param {string} startDate - Data inizio (YYYY-MM-DD)
 * @param {string} finishDate - Data fine (YYYY-MM-DD)
 * @param {string} keyword - Parola chiave filtro (opzionale)
 * @return {Array} Array di oggetti evento con acronimo
 */
/**
 * getUniqueEventsForPeriod - USA evt[3] come acronimo
 */
function getUniqueEventsForPeriod(startDate, finishDate, keyword) {
  try {
    Logger.log('========================================');
    Logger.log('DEBUG: getUniqueEventsForPeriod');
    Logger.log('StartDate: ' + startDate);
    Logger.log('FinishDate: ' + finishDate);
    
    const parts1 = startDate.split('-');
    const parts2 = finishDate.split('-');
    const fromDate = parts1[2] + '/' + parts1[1] + '/' + parts1[0];
    const toDate = parts2[2] + '/' + parts2[1] + '/' + parts2[0];
    
    const eventi = events2Array(fromDate, toDate, categories()[0][0], keyword || '');
    
    Logger.log('Eventi caricati: ' + (eventi ? eventi.length : 'NULL'));
    
    if (!eventi || eventi.length === 0) {
      return [];
    }
    
    // Debug primo evento
    if (eventi.length > 0) {
      Logger.log('Primo evento:');
      Logger.log('  [3] TitleNoSpaces: ' + eventi[0][3]);
      Logger.log('  [3] Senza ultima lettera: ' + eventi[0][3].slice(0, -1));
    }
    
    const uniqueMap = new Map();
    
    eventi.forEach(function(evt) {
      // ========================================
      // RIMUOVI ULTIMA LETTERA (E, D, A, ecc.)
      // ========================================
      const fullAcronym = evt[3]; // es: "CongressoPincoPallinoE"
      const acronym = fullAcronym.slice(0, -1); // es: "CongressoPincoPallino"
      // ========================================
      
      if (!uniqueMap.has(acronym)) {
        uniqueMap.set(acronym, {
          acronym: acronym, // Senza ultima lettera
          fullName: evt[2],
          shortName: evt[2].substring(0, 30) + (evt[2].length > 30 ? '...' : ''),
          category: evt[1],
          currentColor: evt[7]
        });
      }
    });
    
    Logger.log('Eventi unici trovati: ' + uniqueMap.size);
    
    const result = Array.from(uniqueMap.values());
    
    // Log primi 3
    result.slice(0, 3).forEach(function(evt, idx) {
      Logger.log('  ' + idx + ': ' + evt.acronym);
    });
    
    Logger.log('========================================');
    return result;
    
  } catch (error) {
    Logger.log('✗ ERRORE: ' + error.message);
    throw error;
  }
}

/**
 * Converte formato data da YYYY-MM-DD a DD/MM/YYYY
 */
function convertDateFormat(dateStr) {
  const parts = dateStr.split('-');
  return parts[2] + '/' + parts[1] + '/' + parts[0];
}

/**
 * ==================================================
 * FUNZIONE MIGLIORATA: createSlideAndExportToSheetImproved
 * ==================================================
 * Genera slide una alla volta con:
 * - Mappa colori predefinita
 * - Pause tra le chiamate API
 * - Unione finale opzionale
 */
function createSlideAndExportToSheetImproved(first, last, cosa, keyword, editable, colorMap) {
  
  // Inizializza
  resetFoglioConNuovo();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  sheet.setHiddenGridlines(true);
  
  // -------------------------
  // ELABORA DATE
  // -------------------------
  var f = first.split('-');
  var l = last.split('-');
  var firstDate = new Date(+f[0], f[1] - 1, +f[2]);
  var lastDate = new Date(+l[0], l[1] - 1, +l[2] + 1);
  var numDays = (lastDate - firstDate) / (1000 * 3600 * 24);
  
  Logger.log('========================================');
  Logger.log('INIZIO PROCESSO MIGLIORATO');
  Logger.log('Numero giorni: ' + numDays);
  Logger.log('Editable: ' + editable);
  Logger.log('Mappa colori ricevuta: ' + JSON.stringify(colorMap));
  Logger.log('========================================');
  
  // -------------------------
  // FASE 1: GENERA SLIDE GIORNALIERE
  // (una alla volta con pause)
  // -------------------------
  var presentationIds = [];
  var currentDate = new Date(firstDate);
  
  for (let d = 0; d < numDays; d++) {
    
    Logger.log('--- Giorno ' + (d + 1) + ' di ' + numDays + ' ---');
    
    var dateString = formatDateForFilename(currentDate);
    var name = dateString + ' - TEMP - ' + cosa;
    
    var templateId = (cosa === 'cc') 
      ? templateSlides()[1][0] 
      : templateSlides()[0][0];
    
    var presId = copyAndRenamePresentation(templateId, name);
    
    // GENERA SLIDE CON MAPPA COLORI
viewEventsWithColorMap(
      1,
      new Date(currentDate),
      1,
      selectedMode(),
      presId,
      cosa,
      keyword,
      colorMap  // Assicurati che questo parametro sia passato
    );
    
    presentationIds.push(presId);
    
    // PAUSA TRA GIORNI (più lunga per evitare quota)
    Logger.log('Pausa di sicurezza...');
    Utilities.sleep(5000); // 5 secondi tra un giorno e l'altro
    
    currentDate.setDate(currentDate.getDate() + 1);
  }
  
  Logger.log('✓ Slide giornaliere create: ' + presentationIds.length);
  
  // -------------------------
  // FASE 2: UNIONE (SE RICHIESTA)
  // -------------------------
  var finalPresentationId = null;
  
  if (editable === 'SI') {
    Logger.log('=== FASE 2: Unione slide ===');
    
    try {
      var today = new Date();
      var finalName = formatDateMaster(today).dataXfile + 
                      translate('planPage.slideFile') + cosa;
      
      // Crea presentazione finale
      if (cosa === 'cc') {
        finalPresentationId = copyAndRenamePresentation(templateSlides()[1][0], finalName);
      } else {
        finalPresentationId = copyAndRenamePresentation(templateSlides()[0][0], finalName);
      }
      
      Logger.log('✓ Presentazione finale creata: ' + finalPresentationId);
      
      var finalPresentation = SlidesApp.openById(finalPresentationId);
      
      // Rimuovi slide template
      var initialSlides = finalPresentation.getSlides();
      if (initialSlides.length > 0) {
        for (let i = initialSlides.length - 1; i >= 0; i--) {
          initialSlides[i].remove();
        }
      }
      
      finalPresentation.saveAndClose();
      Utilities.sleep(2000);
      
      // Copia slide da ogni presentazione giornaliera
      for (let i = 0; i < presentationIds.length; i++) {
        
        Logger.log('--- Unione file ' + (i+1) + '/' + presentationIds.length + ' ---');
        
        finalPresentation = SlidesApp.openById(finalPresentationId);
        var sourcePresentation = SlidesApp.openById(presentationIds[i]);
        var sourceSlides = sourcePresentation.getSlides();
        
        if (sourceSlides.length === 0) {
          Logger.log('  Attenzione: File sorgente vuoto');
          sourcePresentation.saveAndClose();
          continue;
        }
        
        // Copia slide
        for (let j = 0; j < sourceSlides.length; j++) {
          Logger.log('  Appendendo slide ' + (j+1));
          finalPresentation.appendSlide(sourceSlides[j]);
          Utilities.sleep(2000); // Pausa tra slide
        }
        
        sourcePresentation.saveAndClose();
        finalPresentation.saveAndClose();
        
        Logger.log('  ✓ File ' + (i+1) + ' completato');
        Utilities.sleep(3000); // Pausa tra file
      }
      
      finalPresentation = SlidesApp.openById(finalPresentationId);
      var finalCount = finalPresentation.getSlides().length;
      finalPresentation.saveAndClose();
      
      Logger.log('✓ Unione completata - Totale slide: ' + finalCount);
      
    } catch (e) {
      Logger.log('✗ ERRORE FASE 2: ' + e.toString());
    }
  } else {
    Logger.log('=== FASE 2: SALTATA (editable !== SI) ===');
  }
  
  // -------------------------
  // FASE 3: EXPORT SU SHEET
  // (un giorno alla volta)
  // -------------------------
  Logger.log('=== FASE 3: Export su Sheet ===');
  
  var exportIds = (finalPresentationId) ? [finalPresentationId] : presentationIds;
  var lastRow = 2;
  
  exportIds.forEach((id, idx) => {
    
    Logger.log('Export presentazione ' + (idx + 1) + ' di ' + exportIds.length);
    
    var pres = SlidesApp.openById(id);
    var slides = pres.getSlides();
    
    slides.forEach((slide, slideIdx) => {
      
      var url = 'https://docs.google.com/presentation/d/' +
                id + '/export/png?pageid=' +
                slide.getObjectId();
      
      try {
        var blob = UrlFetchApp.fetch(url, {
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
        }).getBlob();
        
        sheet.insertImage(blob, 1, lastRow);
        lastRow += 53;
        
        Logger.log('  ✓ Slide ' + (slideIdx + 1) + ' esportata');
        
        // Pausa tra export
        Utilities.sleep(1000);
        
      } catch (e) {
        Logger.log('  ✗ Errore export slide: ' + e.toString());
      }
    });
    
    pres.saveAndClose();
  });
  
  // -------------------------
  // FASE 4: PULIZIA
  // -------------------------
  Logger.log('=== FASE 4: Pulizia file temporanei ===');
  
  presentationIds.forEach(id => {
    try {
      DriveApp.getFileById(id).setTrashed(true);
    } catch (e) {
      Logger.log('Errore eliminazione file: ' + e.toString());
    }
  });
  
  // Link finale
  if (finalPresentationId) {
    sheet.getRange('A1').setValue(translate('planPage.slideFileAlert'));
    sheet.getRange('B1').setValue(
      'https://docs.google.com/presentation/d/' +
      finalPresentationId + '/edit'
    );
  }
  
  Logger.log('========================================');
  Logger.log('✓ PROCESSO COMPLETATO');
  Logger.log('========================================');
}

/**
 * Converte formato data da YYYY-MM-DD a DD/MM/YYYY
 */
function convertDateFormat(dateStr) {
  const parts = dateStr.split('-');
  return parts[2] + '/' + parts[1] + '/' + parts[0];
}

/**
 * ==================================================
 * FUNZIONE MIGLIORATA: createSlideAndExportToSheetImproved
 * ==================================================
 * Genera slide una alla volta con:
 * - Mappa colori predefinita
 * - Pause tra le chiamate API
 * - Unione finale opzionale
 */
function createSlideAndExportToSheetImproved(first, last, cosa, keyword, editable, colorMap) {
  
  // Inizializza
  resetFoglioConNuovo();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  sheet.setHiddenGridlines(true);
  
  // -------------------------
  // ELABORA DATE
  // -------------------------
  var f = first.split('-');
  var l = last.split('-');
  var firstDate = new Date(+f[0], f[1] - 1, +f[2]);
  var lastDate = new Date(+l[0], l[1] - 1, +l[2] + 1);
  var numDays = (lastDate - firstDate) / (1000 * 3600 * 24);
  
  Logger.log('========================================');
  Logger.log('INIZIO PROCESSO MIGLIORATO');
  Logger.log('Numero giorni: ' + numDays);
  Logger.log('Editable: ' + editable);
  Logger.log('Mappa colori ricevuta: ' + JSON.stringify(colorMap));
  Logger.log('========================================');
  
  // -------------------------
  // FASE 1: GENERA SLIDE GIORNALIERE
  // (una alla volta con pause)
  // -------------------------
  var presentationIds = [];
  var currentDate = new Date(firstDate);
  
  for (let d = 0; d < numDays; d++) {
    
    Logger.log('--- Giorno ' + (d + 1) + ' di ' + numDays + ' ---');
    
    var dateString = formatDateForFilename(currentDate);
    var name = dateString + ' - TEMP - ' + cosa;
    
    var templateId = (cosa === 'cc') 
      ? templateSlides()[1][0] 
      : templateSlides()[0][0];
    
    var presId = copyAndRenamePresentation(templateId, name);
    
    // GENERA SLIDE CON MAPPA COLORI
    viewEventsWithColorMap(
      1,
      new Date(currentDate),
      1,
      selectedMode(),
      presId,
      cosa,
      keyword,
      colorMap  // 👈 PASSA LA MAPPA
    );
    
    presentationIds.push(presId);
    
    // PAUSA TRA GIORNI (più lunga per evitare quota)
    Logger.log('Pausa di sicurezza...');
    Utilities.sleep(5000); // 5 secondi tra un giorno e l'altro
    
    currentDate.setDate(currentDate.getDate() + 1);
  }
  
  Logger.log('✓ Slide giornaliere create: ' + presentationIds.length);
  
  // -------------------------
  // FASE 2: UNIONE (SE RICHIESTA)
  // -------------------------
  var finalPresentationId = null;
  
  if (editable === 'SI') {
    Logger.log('=== FASE 2: Unione slide ===');
    
    try {
      var today = new Date();
      var finalName = formatDateMaster(today).dataXfile + 
                      translate('planPage.slideFile') + cosa;
      
      // Crea presentazione finale
      if (cosa === 'cc') {
        finalPresentationId = copyAndRenamePresentation(templateSlides()[1][0], finalName);
      } else {
        finalPresentationId = copyAndRenamePresentation(templateSlides()[0][0], finalName);
      }
      
      Logger.log('✓ Presentazione finale creata: ' + finalPresentationId);
      
      var finalPresentation = SlidesApp.openById(finalPresentationId);
      
      // Rimuovi slide template
      var initialSlides = finalPresentation.getSlides();
      if (initialSlides.length > 0) {
        for (let i = initialSlides.length - 1; i >= 0; i--) {
          initialSlides[i].remove();
        }
      }
      
      finalPresentation.saveAndClose();
      Utilities.sleep(2000);
      
      // Copia slide da ogni presentazione giornaliera
      for (let i = 0; i < presentationIds.length; i++) {
        
        Logger.log('--- Unione file ' + (i+1) + '/' + presentationIds.length + ' ---');
        
        finalPresentation = SlidesApp.openById(finalPresentationId);
        var sourcePresentation = SlidesApp.openById(presentationIds[i]);
        var sourceSlides = sourcePresentation.getSlides();
        
        if (sourceSlides.length === 0) {
          Logger.log('  Attenzione: File sorgente vuoto');
          sourcePresentation.saveAndClose();
          continue;
        }
        
        // Copia slide
        for (let j = 0; j < sourceSlides.length; j++) {
          Logger.log('  Appendendo slide ' + (j+1));
          finalPresentation.appendSlide(sourceSlides[j]);
          Utilities.sleep(2000); // Pausa tra slide
        }
        
        sourcePresentation.saveAndClose();
        finalPresentation.saveAndClose();
        
        Logger.log('  ✓ File ' + (i+1) + ' completato');
        Utilities.sleep(3000); // Pausa tra file
      }
      
      finalPresentation = SlidesApp.openById(finalPresentationId);
      var finalCount = finalPresentation.getSlides().length;
      finalPresentation.saveAndClose();
      
      Logger.log('✓ Unione completata - Totale slide: ' + finalCount);
      
    } catch (e) {
      Logger.log('✗ ERRORE FASE 2: ' + e.toString());
    }
  } else {
    Logger.log('=== FASE 2: SALTATA (editable !== SI) ===');
  }
  
  // -------------------------
  // FASE 3: EXPORT SU SHEET
  // (un giorno alla volta)
  // -------------------------
  Logger.log('=== FASE 3: Export su Sheet ===');
  
  var exportIds = (finalPresentationId) ? [finalPresentationId] : presentationIds;
  var lastRow = 2;
  
  exportIds.forEach((id, idx) => {
    
    Logger.log('Export presentazione ' + (idx + 1) + ' di ' + exportIds.length);
    
    var pres = SlidesApp.openById(id);
    var slides = pres.getSlides();
    
    slides.forEach((slide, slideIdx) => {
      
      var url = 'https://docs.google.com/presentation/d/' +
                id + '/export/png?pageid=' +
                slide.getObjectId();
      
      try {
        var blob = UrlFetchApp.fetch(url, {
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
        }).getBlob();
        
        sheet.insertImage(blob, 1, lastRow);
        lastRow += 53;
        
        Logger.log('  ✓ Slide ' + (slideIdx + 1) + ' esportata');
        
        // Pausa tra export
        Utilities.sleep(1000);
        
      } catch (e) {
        Logger.log('  ✗ Errore export slide: ' + e.toString());
      }
    });
    
    pres.saveAndClose();
  });
  
  // -------------------------
  // FASE 4: PULIZIA
  // -------------------------
  Logger.log('=== FASE 4: Pulizia file temporanei ===');
  
  presentationIds.forEach(id => {
    try {
      DriveApp.getFileById(id).setTrashed(true);
    } catch (e) {
      Logger.log('Errore eliminazione file: ' + e.toString());
    }
  });
  
  // Link finale
  if (finalPresentationId) {
    sheet.getRange('A1').setValue(translate('planPage.slideFileAlert'));
    sheet.getRange('B1').setValue(
      'https://docs.google.com/presentation/d/' +
      finalPresentationId + '/edit'
    );
  }
  
  Logger.log('========================================');
  Logger.log('✓ PROCESSO COMPLETATO');
  Logger.log('========================================');
}

/**
 * ==================================================
 * FUNZIONE CORRETTA: viewEventsWithColorMap
 * ==================================================
 */
function viewEventsWithColorMap(user, start, numberDays, method, presentationId, cosa, keyword, colorMap) {

  Logger.log('>>> viewEventsWithColorMap chiamata <<<');
  Logger.log('colorMap ricevuta: ' + JSON.stringify(colorMap));

  var presentation = SlidesApp.openById(presentationId);
  var slides = presentation.getSlides();
  var numSlides = slides.length;
  var transparency = setTransparency()[0][0];
  
  var colori2d = methodMcolors();
  var structures = (cosa === 'cc') ? centroCongressi() : strutture();

  var date = new Date(start);
  var today = convertDateBar(new Date(date.getTime()));
  
  var nextDayDate = new Date(date.getTime());
  nextDayDate.setDate(nextDayDate.getDate() + ((numberDays == 0) ? 1 : numberDays));
  var toDate = convertDateBar(nextDayDate);

  // Carica eventi
  var eventi = events2Array(today, toDate, categories()[0][0], keyword);
  
  Logger.log('Eventi caricati: ' + eventi.length);

  // ========================================
  // APPLICAZIONE MAPPA COLORI - CON RIMOZIONE ULTIMA LETTERA
  // ========================================
  if (colorMap && Object.keys(colorMap).length > 0) {
    Logger.log('🎨 Applicazione mappa colori personalizzata...');
    
    var applicati = 0;
    
    eventi.forEach(function(evt) {
      // ========================================
      // RIMUOVI ULTIMA LETTERA PRIMA DI CERCARE
      // ========================================
      const fullAcronym = evt[3]; // es: "CongressoPincoPallinoE"
      const acronym = fullAcronym.slice(0, -1); // es: "CongressoPincoPallino"
      // ========================================
      
      if (colorMap.hasOwnProperty(acronym)) {
        var vecchioColore = evt[7];
        var nuovoColore = parseInt(colorMap[acronym]);
        
        // SOVRASCRIVI evt[7] con il nuovo colore
        evt[7] = nuovoColore;
        
        Logger.log('  ✓ Evento "' + fullAcronym + '" (chiave: "' + acronym + '"): ' + vecchioColore + ' → ' + nuovoColore);
        applicati++;
      } else {
        Logger.log('  ⚠️ Evento "' + fullAcronym + '" (chiave: "' + acronym + '") non trovato nella mappa, uso colore attuale: ' + evt[7]);
      }
    });
    
    Logger.log('✓ Colori applicati: ' + applicati + ' su ' + eventi.length + ' eventi');
  } else {
    Logger.log('⚠️ NESSUNA mappa colori ricevuta o vuota!');
  }
  // ========================================

  var pageDate = new Date(start);

  // Rimuovi eventi che iniziano prima
  var checkDate = new Date(+today.split('/')[2], today.split('/')[1] - 1, +today.split('/')[0]);
  var t = 0;
  while ((eventi.length > 1) && t < eventi.length && (eventi[t][0].getTime() < checkDate.getTime())) {
    eventi.splice(t, 1);
  }

  var y = 0;
  var j = 0;
  var count = 0;

  var filteredEvents = filterEvents(eventi);
  
  Logger.log('Filtered events: ' + filteredEvents.length);

  // Loop giorni
  for (let i = 0; i < numberDays; i += 1) {

    var random = randomID(2);
    CreateSlide(presentationId, i + numSlides, random, pageDate, 134, 9, method, cosa);

    // MENU IN ALTO A SINISTRA
    var w = 0;
    while ((y < eventi.length) && (convertDate(eventi[y][0]) == convertDate(pageDate))) {
      
      // USA evt[7] che ora contiene il colore dalla mappa!
      var colorIndex = eventi[y][7];
      var statusIndex = selectHigh(eventi[y][7], eventi[y][1])[1];
      
      if (colorIndex >= colori2d.length) colorIndex = 0;
      
      var finalColor = colori2d[colorIndex][statusIndex];
      
      Logger.log('  Menu: evento[' + y + '] "' + eventi[y][3] + '" → colore ' + colorIndex + ' → ' + finalColor);

      var widthBox, heightBox, xPos, yPos, spacing;
      if (cosa === 'cc') {
        widthBox = 55; heightBox = 15.0; xPos = 152; spacing = 16.25;
      } else if (cosa === 'quartiereVecchio') {
        widthBox = 200; heightBox = 6.0; xPos = 5; spacing = 7.25;
      } else {
        widthBox = 60; heightBox = 13.5; xPos = 1; spacing = 14.25;
      }

      CreateBoxText(
        presentationId, 
        i + numSlides, 
        random, 
        pageDate, 
        'MENU', 
        eventi[y][3], 
        widthBox, heightBox, xPos, 15 + w,
        finalColor,
        '0', 
        eventi[y][4], 
        parseEventDetails(eventi[y][8]).descrizione, 
        transparency, 
        method, 
        cosa
      );

      w = w + spacing;
      count++;
      y++;
      
      if (count > 40) {
        Utilities.sleep(2000);
        count = 0;
      }
    }

    // STRUTTURE SULLA MAPPA
    var presence = [];
    const ccNotInMappa = ['GALL', 'EMPTY', 'foyerSG', 'foyerSMU', 'foyerSM', 'foyerMe1', 'foyerMe2', 'bistro', 'loggia', 'lobbyQ4', 'ufficiQ8', 'lbar'];

    while ((j < filteredEvents.length) && (convertDate(filteredEvents[j][0]) == convertDate(pageDate))) {
      if (filteredEvents[j][10].length > 0) {
        for (let k = 0; k < filteredEvents[j][10].length; k += 1) {
          
          if (findKey(filteredEvents[j][10][k], structures, 0) >= 0) {
            var riga = findKey(filteredEvents[j][10][k], structures, 0);
            
            var posX, posY;
            var structName = filteredEvents[j][10][k];
            var pIndex = findKey(structName, presence, 0);
            
            if (pIndex < 0) {
              presence.push([structName, 0]);
              posX = Number(structures[riga][3]);
              posY = Number(structures[riga][4]);
            } else {
              presence[pIndex][1] = presence[pIndex][1] - 2;
              var pos = presence[pIndex][1];
              posX = Number(structures[riga][3]) + pos;
              posY = Number(structures[riga][4]) + pos;
            }

            var width = Number(structures[riga][1]) * 1.0;
            var height = Number(structures[riga][2]) * 1.0;

            if (((cosa == 'quartiere') && (!ccNotInMappa.includes(structName))) || (cosa == 'cc')) {
              if (width != 0) {
                
                var colorIndexStruct = filteredEvents[j][7];
                var statusIndexStruct = selectHigh(filteredEvents[j][7], filteredEvents[j][1])[1];
                if (colorIndexStruct >= colori2d.length) colorIndexStruct = 0;
                var finalColorStruct = colori2d[colorIndexStruct][statusIndexStruct];

                CreateBoxText(
                  presentationId, 
                  i + numSlides, 
                  random, 
                  pageDate, 
                  structName, 
                  filteredEvents[j][3], 
                  width, height, posX, posY, 
                  finalColorStruct,
                  findImage(structures[riga][5], imageFinder(), selectedMode()), 
                  filteredEvents[j][4], 
                  filteredEvents[j][9], 
                  transparency, 
                  method, 
                  cosa
                );
                count++;
              }
            }
          }
        }
      }
      j++;
      
      if (count > 40) {
        Utilities.sleep(2000);
        count = 0;
      }
    }

    pageDate = new Date(pageDate.getTime() + 1 * 3600000 * 24);
  }

  Utilities.sleep(1000);
  presentation.saveAndClose();
  Utilities.sleep(1000);
  
  Logger.log('>>> viewEventsWithColorMap completata <<<');
}

/**
 * Helper per formattare data nel nome file
 */
function formatDateForFilename(date) {
  var year = date.getFullYear();
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var day = ('0' + date.getDate()).slice(-2);
  return year + '-' + month + '-' + day;
}