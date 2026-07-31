/**
 * ============================================================
 * Backend Apps Script — vacances.html (planning Alex & Jenna)
 * ============================================================
 *
 * Installation : voir appscript/README.md
 *
 * Feuilles Google Sheet attendues :
 *  - "Events" (en-têtes en ligne 1) :
 *       A: id | B: dateStart | C: dateEnd | D: categorie | E: personne | F: note
 *    Créée automatiquement avec ses en-têtes si elle n'existe pas encore.
 *  - "Imported" (optionnelle, lecture seule, même structure que Events) :
 *       utilisée pour des événements importés depuis un calendrier externe
 *       (ex agenda partagé, ICS…). Si l'onglet n'existe pas, imported = [].
 *  - "Ponts" (en-têtes en ligne 1) :
 *       A: dateStart | B: dateEnd | C: note
 *    Gère les jours de pont / fermeture entreprise directement dans le
 *    Sheet (remplace le tableau PONTS_ALEX qui était hardcodé côté front,
 *    cf. A7 du PRD). Si l'onglet n'existe pas, ponts = [].
 *
 * Script Properties attendues (Extensions > Propriétés du projet
 * > Propriétés du script) :
 *   API_TOKEN : jeton partagé avec le front (constante API_TOKEN
 *               dans common.js, à synchroniser manuellement)
 */

var SHEET_EVENTS   = 'Events';
var SHEET_IMPORTED = 'Imported';
var SHEET_PONTS    = 'Ponts';
var CATEGORIES_VALIDES = ['Libre', 'Alex', 'Jenna', 'Invités', 'Vacances', 'Pont'];

// ------------------------------------------------------------
// ROUTAGE
// ------------------------------------------------------------
function doGet(e) {
  try {
    if (!checkToken(e.parameter && e.parameter.token)) {
      return jsonResponse({ error: 'unauthorized' });
    }
    var action = e.parameter.action;
    if (action === 'getAll') {
      return jsonResponse({
        events:   getEvents(),
        imported: getImported(),
        ponts:    getPonts()
      });
    }
    return jsonResponse({ error: 'unknown_action' });
  } catch (err) {
    return jsonResponse({ error: 'server_error' });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    if (!checkToken(body.token)) {
      return jsonResponse({ success: false, message: 'Non autorisé.' });
    }
    if (body.action === 'addEvent')    return addEvent(body);
    if (body.action === 'deleteEvent') return deleteEvent(body);
    return jsonResponse({ success: false, message: 'Action inconnue.' });
  } catch (err) {
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
// LECTURE
// ------------------------------------------------------------
function getEventsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_EVENTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EVENTS);
    sheet.appendRow(['id', 'dateStart', 'dateEnd', 'categorie', 'personne', 'note']);
  }
  return sheet;
}

function getEvents() {
  var sheet = getEventsSheet();
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    out.push({
      id:        String(row[0]),
      dateStart: formatDate(row[1]),
      dateEnd:   formatDate(row[2]),
      categorie: String(row[3] || ''),
      personne:  String(row[4] || ''),
      note:      String(row[5] || '')
    });
  }
  return out;
}

function getImported() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IMPORTED);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[1]) continue;
    out.push({
      id:        String(row[0] || 'imported-' + i),
      dateStart: formatDate(row[1]),
      dateEnd:   formatDate(row[2]),
      categorie: String(row[3] || 'Invités'),
      personne:  String(row[4] || ''),
      note:      String(row[5] || '')
    });
  }
  return out;
}

// Clé "ponts" (A7) : gérée directement dans l'onglet "Ponts" du Sheet.
// Plus de tableau hardcodé côté front ; fallback = liste vide si
// l'onglet n'existe pas encore.
function getPonts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PONTS);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    out.push({
      dateStart: formatDate(row[0]),
      dateEnd:   formatDate(row[1] || row[0]),
      note:      String(row[2] || '')
    });
  }
  return out;
}

// ------------------------------------------------------------
// INITIALISATION — à lancer UNE FOIS depuis l'éditeur Apps Script.
// Crée l'onglet "Ponts" avec ses en-têtes s'il n'existe pas encore.
// Remplace l'ancien tableau PONTS_ALEX hardcodé dans vacances.html :
// les ponts s'éditent désormais directement dans cet onglet.
// ------------------------------------------------------------
function initPontsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PONTS);

  if (sheet) {
    SpreadsheetApp.getUi().alert(
      'L\'onglet "' + SHEET_PONTS + '" existe déjà — rien à faire. ' +
      'Ajoute tes ponts directement dedans.'
    );
    return;
  }

  sheet = ss.insertSheet(SHEET_PONTS);
  var headers = ['Date début', 'Date fin', 'Note'];
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#2d4a3e');
  headerRange.setFontColor('#ffffff');
  sheet.getRange(2, 1, 200, 2).setNumberFormat('dd/mm/yyyy');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  SpreadsheetApp.getUi().alert(
    '✅ Onglet "' + SHEET_PONTS + '" créé. Ajoute une ligne par pont : ' +
    'date de début, date de fin, note (ex. "Pont Ascension").'
  );
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
    return jsonResponse({ success: false, message: 'Catégorie invalide.' });
  }
  if (!personne || personne.length > 100) {
    return jsonResponse({ success: false, message: 'Nom invalide (max 100 caractères).' });
  }
  if (!isValidDate(dateStart) || !isValidDate(dateEnd)) {
    return jsonResponse({ success: false, message: 'Date invalide (format AAAA-MM-JJ attendu).' });
  }
  if (dateEnd < dateStart) {
    return jsonResponse({ success: false, message: 'La date de fin doit être après la date de début.' });
  }
  if (note.length > 1000) {
    return jsonResponse({ success: false, message: 'Note trop longue (max 1000 caractères).' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return jsonResponse({ success: false, message: 'Le service est occupé, réessaie dans quelques secondes.' });
  }

  try {
    var sheet = getEventsSheet();
    var id = Utilities.getUuid();
    sheet.appendRow([
      id,
      dateStart,
      dateEnd,
      sanitizeCell(categorie),
      sanitizeCell(personne),
      sanitizeCell(note)
    ]);
    return jsonResponse({ success: true, message: 'Événement ajouté.', id: id });
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
    return jsonResponse({ success: false, message: 'Identifiant manquant.' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return jsonResponse({ success: false, message: 'Le service est occupé, réessaie dans quelques secondes.' });
  }

  try {
    var sheet = getEventsSheet();
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === id) {
        sheet.deleteRow(i + 1);
        return jsonResponse({ success: true, message: 'Événement supprimé.' });
      }
    }
    return jsonResponse({ success: false, message: 'Événement introuvable (déjà supprimé ?).' });
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------
// UTILITAIRES
// ------------------------------------------------------------
function isValidDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(s);
}

function formatDate(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    var y = value.getFullYear();
    var m = String(value.getMonth() + 1).padStart(2, '0');
    var d = String(value.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return String(value);
}

// Protection contre l'injection de formule dans le Sheet (=, +, -, @ en
// tête de valeur).
function sanitizeCell(value) {
  var s = String(value);
  if (/^[=+\-@]/.test(s)) return "'" + s;
  return s;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
