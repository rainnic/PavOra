/*
* Project Name: Pavora
* Copyright (c) 2025 Nicola Rainiero
*
* This software is released under the MIT License.
* Please refer to the LICENSE file for the full license text.
*/
function testInsertExSlide() {
  createSlideAndExportToSheet('2024-12-03', '2024-12-03', 'quartiere', '', 'NO'); // cc o quartiere
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