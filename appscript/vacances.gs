/**
 * ============================================================
 * Backend Apps Script — vacances.html (planning Alex & Jenna)
 * ============================================================
 * Script STANDALONE : il n'est pas lié à un Sheet, il l'ouvre par ID
 * (SPREADSHEET_ID ci-dessous). Ne jamais utiliser getActiveSpreadsheet()
 * ici : elle renvoie null dans un script standalone.
 *
 * Installation : voir appscript/README.md
 *
 * Onglets utilisés dans le Sheet :
 *  - "Vacances"     : événements saisis dans l'app
 *       A: ID | B: Date début | C: Date fin | D: Catégorie | E: Personne | F: Note
 *    Créé automatiquement avec ses en-têtes s'il n'existe pas.
 *  - "Reservations" : lecture seule. Les weekends « Réservé » y sont
 *    convertis en événements « Invités » (ou « Vacances » si le message
 *    contient le mot vacances) affichés dans le planning.
 *  - "Ponts"        : A: Date début | B: Date fin | C: Note
 *    Remplace le tableau PONTS_ALEX qui était hardcodé côté front (A7).
 *    Lance initPontsSheet() une fois pour le créer. Absent → ponts = [].
 *
 * Script Property attendue (⚙️ Paramètres du projet > Propriétés du script) :
 *   API_TOKEN : jeton partagé avec le front (constante API_TOKEN de common.js)
 */

var SPREADSHEET_ID = '1qhZMGj8knN8bS7A2t6eU3wRmZwcEM9WBap_b92gWLQ4';

var SHEET_EVENTS       = 'Vacances';
var SHEET_RESERVATIONS = 'Reservations';
var SHEET_PONTS        = 'Ponts';

var TIMEZONE = 'Europe/Paris';

// Catégories acceptées par addEvent. « Invités » n'est pas saisissable :
// elle est dérivée côté serveur depuis l'onglet Reservations.
var CATEGORIES_VALIDES = ['Libre', 'Alex', 'Jenna', 'Vacances', 'Pont'];

var MAX_LEN = { personne: 100, note: 1000 };

// ------------------------------------------------------------
// ROUTAGE
// ------------------------------------------------------------
function doGet(e) {
  try {
    if (!checkToken(e && e.parameter && e.parameter.token)) {
      return jsonResponse({ error: 'unauthorized' });
    }
    var action = e.parameter.action;
    if (action === 'getAll') {
      return jsonResponse({
        events:   getEvents(),
        imported: getImportedFromReservations(),
        ponts:    getPonts()
      });
    }
    if (action === 'getEvents') {
      return jsonResponse(getEvents());
    }
    return jsonResponse({ error: 'unknown_action' });
  } catch (err) {
    console.error('doGet error: ' + err);
    return jsonResponse({ error: 'server_error' });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!checkToken(body.token)) {
      return jsonResponse({ success: false, message: 'Non autorisé.' });
    }
    if (body.action === 'addEvent')    return jsonResponse(addEvent(body));
    if (body.action === 'deleteEvent') return jsonResponse(deleteEvent(body));
    return jsonResponse({ success: false, message: 'Action inconnue.' });
  } catch (err) {
    // Ne JAMAIS renvoyer err.toString() au client (fuite d'info interne).
    console.error('doPost error: ' + err);
    return jsonResponse({ success: false, message: 'Erreur serveur, réessaie plus tard.' });
  }
}

// ------------------------------------------------------------
// AUTH (A2)
// ------------------------------------------------------------
function checkToken(token) {
  var expected = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  return !!expected && token === expected;
}

// ------------------------------------------------------------
// ACCÈS AU CLASSEUR (script standalone : openById obligatoire)
// ------------------------------------------------------------
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getEventsSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_EVENTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EVENTS);
    var headers = ['ID', 'Date début', 'Date fin', 'Catégorie', 'Personne', 'Note'];
    var hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setValues([headers]);
    hr.setFontWeight('bold');
    hr.setBackground('#4a4a4a');
    hr.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ------------------------------------------------------------
// LECTURE — événements saisis dans l'app
// ------------------------------------------------------------
function getEvents() {
  var sheet = getEventsSheet();
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    out.push({
      id:        String(data[i][0]),
      dateStart: formatDate(data[i][1]),
      dateEnd:   formatDate(data[i][2]),
      categorie: String(data[i][3] || ''),
      personne:  String(data[i][4] || ''),
      note:      String(data[i][5] || '')
    });
  }
  return out;
}

// ------------------------------------------------------------
// LECTURE — weekends réservés (onglet Reservations) convertis en
// événements du planning. « nom » est utilisé par le front pour
// afficher « Invités · <prénom> ».
// ------------------------------------------------------------
function getImportedFromReservations() {
  var sheet = getSpreadsheet().getSheetByName(SHEET_RESERVATIONS);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var statut  = data[i][3];
    var message = (data[i][6] || '').toString().toLowerCase();
    var label   = data[i][0];
    var dStart  = data[i][1];
    var dEnd    = data[i][2];
    if (!dStart) continue;
    if (statut !== 'Réservé') continue;

    var categorie = message.indexOf('vacances') !== -1 ? 'Vacances' : 'Invités';

    result.push({
      id:        'import-' + i,
      dateStart: formatDate(dStart),
      dateEnd:   formatDate(dEnd || dStart),
      categorie: categorie,
      personne:  'Tous',
      note:      String(label || ''),
      nom:       data[i][4] ? String(data[i][4]) : '',
      imported:  true
    });
  }
  return result;
}

// ------------------------------------------------------------
// LECTURE — ponts (A7). Plus de tableau hardcodé côté front ;
// fallback = liste vide si l'onglet n'existe pas encore.
// ------------------------------------------------------------
function getPonts() {
  var sheet = getSpreadsheet().getSheetByName(SHEET_PONTS);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    out.push({
      dateStart: formatDate(data[i][0]),
      dateEnd:   formatDate(data[i][1] || data[i][0]),
      note:      String(data[i][2] || '')
    });
  }
  return out;
}

// ------------------------------------------------------------
// INITIALISATION — à lancer UNE FOIS depuis l'éditeur Apps Script.
// Crée l'onglet "Ponts" avec ses en-têtes s'il n'existe pas encore.
// ------------------------------------------------------------
function initPontsSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PONTS);

  if (sheet) {
    Logger.log('L\'onglet "' + SHEET_PONTS + '" existe déjà — rien à faire.');
    return;
  }

  sheet = ss.insertSheet(SHEET_PONTS);
  var headers = ['Date début', 'Date fin', 'Note'];
  var hr = sheet.getRange(1, 1, 1, headers.length);
  hr.setValues([headers]);
  hr.setFontWeight('bold');
  hr.setBackground('#2d4a3e');
  hr.setFontColor('#ffffff');
  sheet.getRange(2, 1, 200, 2).setNumberFormat('dd/mm/yyyy');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  Logger.log('✅ Onglet "' + SHEET_PONTS + '" créé. Ajoute une ligne par pont : ' +
             'date de début, date de fin, note (ex. "Pont Ascension").');
}

// ------------------------------------------------------------
// AJOUT (A3 validations serveur + LockService)
// ------------------------------------------------------------
function addEvent(body) {
  var categorie = String(body.categorie || '');
  var personne  = String(body.personne || '').trim();
  var dateStart = String(body.dateStart || '').trim();
  var dateEnd   = String(body.dateEnd || dateStart).trim();
  var note      = String(body.note || '').trim();

  if (CATEGORIES_VALIDES.indexOf(categorie) === -1) {
    return { success: false, message: 'Catégorie invalide.' };
  }
  if (!personne || personne.length > MAX_LEN.personne) {
    return { success: false, message: 'Nom invalide (max ' + MAX_LEN.personne + ' caractères).' };
  }
  if (!isValidDate(dateStart) || !isValidDate(dateEnd)) {
    return { success: false, message: 'Date invalide (format AAAA-MM-JJ attendu).' };
  }
  if (dateEnd < dateStart) {
    return { success: false, message: 'La date de fin doit être après la date de début.' };
  }
  if (note.length > MAX_LEN.note) {
    return { success: false, message: 'Note trop longue (max ' + MAX_LEN.note + ' caractères).' };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    console.error('LockService : verrou non obtenu — ' + err);
    return { success: false, message: 'Le service est occupé, réessaie dans quelques secondes.' };
  }

  try {
    var sheet = getEventsSheet();
    var id = 'evt-' + new Date().getTime();

    // parseDate() construit une date locale (fuseau du script = Europe/Paris),
    // ce qui évite le décalage d'un jour d'un new Date('AAAA-MM-JJ') en UTC.
    sheet.appendRow([
      id,
      parseDate(dateStart),
      parseDate(dateEnd),
      sanitizeCell(categorie),
      sanitizeCell(personne),
      sanitizeCell(note)
    ]);
    sheet.getRange(sheet.getLastRow(), 2, 1, 2).setNumberFormat('dd/mm/yyyy');
    SpreadsheetApp.flush();

    return { success: true, message: 'Événement ajouté.', id: id };
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------
// SUPPRESSION (LockService)
// ------------------------------------------------------------
function deleteEvent(body) {
  var id = String(body.id || '');
  if (!id) {
    return { success: false, message: 'Identifiant manquant.' };
  }
  // Les événements importés sont dérivés de l'onglet Reservations :
  // ils ne sont pas supprimables depuis le planning.
  if (id.indexOf('import-') === 0) {
    return { success: false, message: 'Un weekend réservé se modifie depuis la page de réservation.' };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    console.error('LockService : verrou non obtenu — ' + err);
    return { success: false, message: 'Le service est occupé, réessaie dans quelques secondes.' };
  }

  try {
    var sheet = getEventsSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true, message: 'Événement supprimé.' };
      }
    }
    return { success: false, message: 'Événement introuvable (déjà supprimé ?).' };
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------
// UTILITAIRES
// ------------------------------------------------------------
function isValidDate(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  var parts = s.split('-');
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  var d = parseInt(parts[2], 10);
  if (m < 1 || m > 12 || d < 1 || d > 31) return false;
  var probe = new Date(y, m - 1, d);
  return probe.getFullYear() === y && probe.getMonth() === m - 1 && probe.getDate() === d;
}

/** 'AAAA-MM-JJ' → Date locale (pas UTC : évite le décalage d'un jour). */
function parseDate(s) {
  var parts = String(s).split('-');
  return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
}

/** Date (ou chaîne) → 'AAAA-MM-JJ' dans le fuseau du planning. */
function formatDate(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, TIMEZONE, 'yyyy-MM-dd');
  }
  return String(value);
}

/** Protection contre l'injection de formule dans le Sheet (=, +, -, @). */
function sanitizeCell(value) {
  if (value === null || value === undefined) return '';
  var s = String(value);
  if (s === '') return '';
  if (/^[=+\-@\t\r]/.test(s)) return "'" + s;
  return s;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
