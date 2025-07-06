/*
* Project Name: Pavora
* Copyright (c) 2025 Nicola Rainiero
*
* This software is released under the MIT License.
* Please refer to the LICENSE file for the full license text.
*/

// From sidebar
function salvaDatiInUserProperties(first, last, keyword, selectedStruct) {
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperties({
    first4Browser: first,
    last4Browser: last,
    key4Browser: keyword,
    selStruct4Browser: selectedStruct
  });
}

function getUserBrowserSettings() {
  const userProps = PropertiesService.getUserProperties();
  const defaults = {
    first4Browser: '2025-01-01',
    last4Browser: '2025-12-31',
    key4Browser: '',
    selStruct4Browser: ''
  };

  const values = userProps.getProperties();

  return {
    first: values.first4Browser || defaults.first4Browser,
    last: values.last4Browser || defaults.last4Browser,
    keyword: values.key4Browser || defaults.key4Browser,
    selectedStruct: values.selStruct4Browser || defaults.selStruct4Browser
  };
}
// How to use it?
function usaPreferenzeUtente() {
  const prefs = getUserBrowserSettings();
  Logger.log('First= ' + prefs.first);
  Logger.log('Last= ' + prefs.last);
  Logger.log('keyword= ' + prefs.keyword);
  Logger.log('Sel structures= ' + prefs.selectedStruct);
  // Esegui logica personalizzata per quell’utente
}

// exceptOpzOff
function includeOptionated() {
  var result = readVariables('includeOptionated', DataSettings);
  
  // Controllo di sicurezza per evitare l'errore "Cannot read properties of undefined"
  if (result && result.length > 0 && result[0] && result[0].length > 0) {
    return result[0][0];
  } else {
    //console.log('Nessun dato trovato per preloadSheet, restituisco valore di default');
    return false; // o qualsiasi valore di default tu preferisca
  }

}

function excludeAll() {
  var result = readVariables('excludeAll', DataSettings);
  
  // Controllo di sicurezza per evitare l'errore "Cannot read properties of undefined"
  if (result && result.length > 0 && result[0] && result[0].length > 0) {
    return result[0][0];
  } else {
    //console.log('Nessun dato trovato per preloadSheet, restituisco valore di default');
    return false; // o qualsiasi valore di default tu preferisca
  }

}

// Alias email
function aliasEmail() {
    var result = readVariables('aliasEmail', DataSettings);

  // Controllo di sicurezza per evitare l'errore "Cannot read properties of undefined"
  if (result && result.length > 0 && result[0] && result[0].length > 0) {
    return result[0][0];
  } else {
    //console.log('Nessun dato trovato per preloadSheet, restituisco valore di default');
    return false; // o qualsiasi valore di default tu preferisca
  }

}

function tryPreload() {
  Logger.log(preloadSheet());
}

// preloadSheet
function preloadSheet() {
  var result = readVariablesExt('preloadSheet', 'DataSettings');

  // Controllo di sicurezza per evitare l'errore "Cannot read properties of undefined"
  if (result && result.length > 0 && result[0] && result[0].length > 0) {
    return result[0][0];
  } else {
    //console.log('Nessun dato trovato per preloadSheet, restituisco valore di default');
    return false; // o qualsiasi valore di default tu preferisca
  }
}

// Language selected
function languageSelected() {
  return readVariables('currentLanguage', DataSettings)[0][0]
}

// entrance and parking
function entPark() {
  return readVariables('entPark', DataEventSheet)
}

// Minutes permitted with write permission
function minutesPermitted() {
  return readVariables('minutesPermitted', DataSettings)
}

// Varibile con gli ID dei folder condivisi
function driveIDFolder() {
  return readVariables('driveIDFolder', DataSettings)
}

// Varibile con gli ID dei file condivisi
function driveIDFiles() {
  return readVariables('driveIDFiles', DataSettings)
}

// sheetsList
function sheetsList() {
  return readVariables('sheetsList', DataSettings)
}


// typeEv --> tipo evento | identificativo | on/off
function typeEv() {
  return readVariables('typeEv', DataSettings)
}

// tamplateSlides --> tipo evento | identificativo | on/off
function templateSlides() {
  return readVariables('templateSlides', DataSettings)
}

// emailTarget --> tipo evento | identificativo | on/off
function emailTarget() {
  return readVariables('emailTarget', DataSettings)
}

// allestitore --> nome ditta | identificativo
function allestitore() {
  return readVariables('allestitore', DataSettings)
}

// catering --> nome ditta | identificativo
function catering() {
  return readVariables('catering', DataSettings)
}

// refCom --> nome e cognome | identificativo
function refCom() {
  return readVariables('refCom', DataSettings)
}

// refOp --> nome e cognome | identificativo
function refOp() {
  return readVariables('refOp', DataSettings)
}

// incrementDay
function incrementDay() {
  return Number(readVariables('incrementDay', DataSettings))
}

// Users in the external table
function users() {
  return readVariables('users', DataSettings)
}

// Users online
function usersOnline() {
  return readVariables('User', UsersOnline)
}

// Word to recognise the optionated events, put at the beginning of event title in the calendar app
function optionated() {
  var optionated1 = ["Opz.", "OPZ.", "Off.", "OFF."];
  var original = readVariables('optionated', DataSettings);
  var convertVariable = [];
  for (let i = 0, len = original.length; i < len; i++) {
    convertVariable[i] = original[i];
  };
  return optionated1
} // usage: optionated()

// Type of Events - Acronyms - Descriptions
function categories() {
  return readVariables('categories', DataSettings)
}

function myCalID() {
  return readVariables('myCalID', DataSettings)
}

// Where store holidays that occur every year
function holidays() {
  return readVariables('holidays', DataSettings)[0]
}

// Where store holidays that occur in a specific date (in this case only Easter Monday)
function specialHolidays() {
  return readVariables('specialHolidays', DataSettings)[0]
}

// The complete structures' array
// ['identification code', width, height, x position, y position, '0' no image associated or 'Letter' for image, 'full name for menu and the description if necessary']
function strutture() {
  return readVariables('structures', DataStructures)
}

function centroCongressi() {
  var centroCongressi = [];
  for (let i = 0; i < strutture().length; i++) {
    if ((strutture()[i][7] == 2) || (strutture()[i][12] != '')) {
      centroCongressi.push([strutture()[i][0], strutture()[i][12], strutture()[i][13], strutture()[i][14], strutture()[i][15], strutture()[i][5]]);
    }
  }
  return centroCongressi
}

// Only the strctures for menu
function onlyStrcturesSelect(matrice) {
  var data = matrice;
  var final = []; //makeArray(1,2);
  for (let i = 0; i < data.length; i++) {
    if (data[i][7] > 0) {
      final.push([data[i][0], data[i][6], data[i][7], data[i][8], data[i][9], data[i][10], data[i][17]]);
    }
  }
  //Logger.log(final.length);
  return final
}

// Color for Mc method
function methodMcolors() {
  return readVariables('methodMcolor', DataSettings)
}

// Colors for 2c method
function method2colors() {
  return readVariables('method2colors', DataSettings)
} // usage: method2colors()[row][column]

//Letters' Color for letter method
function color4Letters() {
  var colorLetters = '#ff0000';
  return colorLetters
}

// Grade of transparency
function setTransparency() {
  return readVariables('setTransparency', DataSettings)
}