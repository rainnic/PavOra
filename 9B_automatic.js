function testCreateAlertEvent() {
  createAlertEvent()
}

function createAlertEvent() {
  const today = new Date();
  const startSearch = convertDateBar(new Date());
  const nDay = 30; // numero di giorni dopo quello iniziale
  const endSearch = convertDateBar(today.setDate(today.getDate() + nDay * 12 * 5));
  // events2Array('01/01/2025', '31/12/2025', categories()[0][0], 'keyword');
  const eventi = events2Array(startSearch, endSearch, categories()[0][0], '');
  Logger.log(startSearch + ' ' + endSearch + ' ' + eventi[0]);
  Logger.log(refCom());
  Logger.log(users());
}

function isWeekdayAndTime() {
  const now = new Date();
  const dayOfWeek = now.getDay(); // 0 = Domenica, 1 = Lunedì, ..., 6 = Sabato
  const hour = now.getHours();

  // Verifica se è un giorno feriale (Lunedì = 1, Venerdì = 5)
  const isWeekday = dayOfWeek >= 1 && dayOfWeek <= 5;

  // Verifica se l'orario è compreso tra le 9:00 e le 10:00
  const isTimeBetween9And10 = hour >= 9 && hour < 10;

  //return isWeekday && isTimeBetween9And10;
  return isWeekday;
}

function formatTitle(title, length = 35) {
  if (title.length > length) {
    return title.slice(0, length - 3) + '...';
  }
  return title.padEnd(length, ' ');
}

function checkEventsAndSendEmails() {
  if (!isWeekdayAndTime()) {
    Logger.log("Non è un giorno feriale o non è nell'orario previsto (9:00-10:00). Script non eseguito.");
    return;
  }
  const fromMail = getRealEmail(users()[0][0]);
  const sender = 'Pianificazione Eventi Pavora';
  const subject = 'Elenco eventi a Calendadio in scadenza';
  const daysRange = 30; // Intervallo di tempo di 30 giorni
  const yearsRange = 5; // Anni nel futuro da considerare
  const testMode = false; // Imposta su false per inviare realmente le email

  const today = new Date();
  const futureDate = new Date();
  futureDate.setFullYear(today.getFullYear() + yearsRange);

  const mycal = myCalID()[0][0]; // Sempre il primo calendario
  const calendar = CalendarApp.getCalendarById(mycal);
  const events = calendar.getEvents(today, futureDate);

  // Fase 0: Pulizia degli eventi (rimozione duplicati basati sul titolo)
  const uniqueEvents = [];
  const seenTitles = new Set();

  events.forEach(event => {
    const title = event.getTitle();
    // Estrae la parte del titolo prima dello spazio e della lettera finale (es. "L", "P", "A", "E", "D")
    const baseTitle = title.split(/ [LPEDA]?$/)[0];

    if (!seenTitles.has(baseTitle)) {
      seenTitles.add(baseTitle);
      uniqueEvents.push(event);
    }
  });

  const eventTable = [];

  // Fase 1: Creazione della tabella degli eventi
  uniqueEvents.forEach(event => {
    const title = event.getTitle();
    if (title.startsWith('Opz.') || title.startsWith('Off.')) {
      const startDate = event.getStartTime();
      const description = event.getDescription();
      const creatorCalendarEmail = event.getCreators()[0];

      // Trova l'email reale dell'utente dalla matrice users
      const userEntry = users().find(user => user[0] === creatorCalendarEmail);
      const creatorEmail = userEntry ? userEntry[4] : creatorCalendarEmail;

      let refComEmail = '';
      let opzExpDate = '';

      // Estrazione refCom email
      const refComMatch = description.match(/refCom=(\w+)/);
      if (refComMatch) {
        const refComWord = refComMatch[1];
        const refComEntry = refCom().find(entry => entry[1] === refComWord);
        if (refComEntry) {
          refComEmail = refComEntry[4];
        }
      }

      // Estrazione opzExp date
      const opzExpMatch = description.match(/opzExp=(\d{4}-\d{2}-\d{2})/);
      if (opzExpMatch) {
        opzExpDate = opzExpMatch[1];
      }

      // Filtra gli eventi in base alle condizioni
      const isStartDateInRange = Math.floor((startDate - today) / (1000 * 60 * 60 * 24)) <= daysRange;
      const isOpzExpDatePassed = opzExpDate && new Date(opzExpDate) <= today;

      if (isStartDateInRange || isOpzExpDatePassed) {
        eventTable.push({
          title: title,
          startDate: startDate,
          creatorEmail: creatorEmail,
          description: description,
          refComEmail: refComEmail,
          opzExpDate: opzExpDate
        });
      }
    }
  });

  // Raggruppamento degli eventi per creatorEmail
  const groupedEvents = {};
  eventTable.forEach(event => {
    const creatorEmail = event.creatorEmail;
    if (!groupedEvents[creatorEmail]) {
      groupedEvents[creatorEmail] = [];
    }
    groupedEvents[creatorEmail].push(event);
  });

  // Fase 2: Invio email raggruppate (o log in test mode)
  for (const creatorEmail in groupedEvents) {
    const eventsForCreator = groupedEvents[creatorEmail];
    let emailBody = `Buongiorno,\nsono stati trovati i seguenti eventi scaduti o in scadenza:\n\n`;

    eventsForCreator.forEach(event => {
      const startDate = event.startDate;
      const opzExpDate = event.opzExpDate;
      const formattedTitle = formatTitle(event.title); // Formatta il titolo
      emailBody += `- Evento: ${formattedTitle} | (Inizio: ${convertDateBar(startDate)}) - (Scadenza: ${opzExpDate})\n\n`;
    });

    emailBody += `\nQuesti eventi sono da confermare, o da rivedere la data di scadenza o da rimuovere.\n\nGrazie,\n${sender}`;

    // Trova l'email refCom (se presente) per il primo evento del creatore
    const refComEmail = eventsForCreator[0].refComEmail;

    if (testMode) {
      // Log delle email in modalità test
      Logger.log(`Email da inviare a: ${creatorEmail}`);
      if (refComEmail && creatorEmail !== refComEmail) {
        Logger.log(`CC: ${refComEmail}`);
      }
      Logger.log(`Oggetto: ${subject}`);
      Logger.log(`Corpo: ${emailBody}`);
      Logger.log('--------------------------');
    } else {
      // Invio effettivo delle email con GmailApp
      if (refComEmail && creatorEmail !== refComEmail) {
        GmailApp.sendEmail(creatorEmail, subject, emailBody, {
          cc: refComEmail,
          bcc: getRealEmail(users()[0][0]),
          from: fromMail,
          name: sender
        });
      } else {
        GmailApp.sendEmail(creatorEmail, subject, emailBody, {
          bcc: getRealEmail(users()[0][0]),
          from: fromMail,
          name: sender
        });
      }
    }
  }
}
