/**
 * ============================================================
 * Backend Apps Script — index.html (réservation weekend "groupe"
 * à St-Georges-sur-Cher)
 * ============================================================
 *
 * Installation : voir appscript/README.md
 *
 * Feuille Google Sheet attendue, nommée "Weekends", avec en-têtes
 * en ligne 1 et colonnes :
 *   A: label          (ex "28-29 mars 2026")
 *   B: statut          ("Libre" ou "Réservé")
 *   C: nom             (rempli à la réservation)
 *   D: nbPersonnes     (rempli à la réservation)
 *   E: email           (rempli à la réservation)
 *   F: message         (rempli à la réservation)
 *   G: dateReservation (rempli à la réservation)
 *
 * Script Properties attendues (Extensions > Propriétés du projet
 * > Propriétés du script) :
 *   API_TOKEN : jeton partagé avec le front (constante API_TOKEN
 *               dans common.js, à synchroniser manuellement)
 */

var SHEET_NAME    = 'Weekends';
var STATUT_LIBRE  = 'Libre';
var STATUT_RESERVE = 'Réservé';

// ------------------------------------------------------------
// ROUTAGE
// ------------------------------------------------------------
function doGet(e) {
  try {
    if (!checkToken(e.parameter && e.parameter.token)) {
      return jsonResponse({ error: 'unauthorized' });
    }
    var action = e.parameter.action;
    if (action === 'getWeekends') {
      return jsonResponse(getWeekends());
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
    return submitReservation(body);
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
// LECTURE DES CRÉNEAUX LIBRES
// ------------------------------------------------------------
function getWeekends() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  var out = [];
  // ligne 1 = en-têtes
  for (var i = 1; i < values.length; i++) {
    var row    = values[i];
    var label  = row[0];
    var statut = row[1];
    if (label && statut === STATUT_LIBRE) {
      out.push({ rowIndex: i + 1, label: String(label) });
    }
  }
  return out;
}

// ------------------------------------------------------------
// RÉSERVATION (A3 validations serveur + LockService)
// ------------------------------------------------------------
function submitReservation(body) {
  var rowIndex    = parseInt(body.rowIndex, 10);
  var label       = String(body.label || '');
  var nom         = String(body.nom || '').trim();
  var nbPersonnes = parseInt(body.nbPersonnes, 10);
  var email       = String(body.email || '').trim();
  var message     = String(body.message || '').trim();

  // -- Validations serveur (ne jamais faire confiance au front) --
  if (!rowIndex || rowIndex < 2) {
    return jsonResponse({ success: false, code: 'INVALID_INPUT', message: 'Créneau invalide.' });
  }
  if (!nom || nom.length > 100) {
    return jsonResponse({ success: false, code: 'INVALID_INPUT', message: 'Nom invalide (max 100 caractères).' });
  }
  if (!nbPersonnes || nbPersonnes < 1 || nbPersonnes > 10) {
    return jsonResponse({ success: false, code: 'INVALID_INPUT', message: 'Le nombre de personnes doit être entre 1 et 10.' });
  }
  if (!isValidEmail(email)) {
    return jsonResponse({ success: false, code: 'INVALID_INPUT', message: 'Email invalide.' });
  }
  if (message.length > 1000) {
    return jsonResponse({ success: false, code: 'INVALID_INPUT', message: 'Message trop long (max 1000 caractères).' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return jsonResponse({ success: false, message: 'Le service est occupé, réessaie dans quelques secondes.' });
  }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse({ success: false, message: 'Configuration serveur invalide.' });
    }

    // Vérifie que la ligne correspond toujours au label envoyé par le
    // front (identifiant de créneau stable, cf. contrat inter-lots).
    var currentLabel = sheet.getRange(rowIndex, 1).getValue();
    if (String(currentLabel) !== label) {
      return jsonResponse({ success: false, code: 'STALE_SLOT', message: 'Ce créneau a changé, recharge la page.' });
    }
    var currentStatut = sheet.getRange(rowIndex, 2).getValue();
    if (currentStatut !== STATUT_LIBRE) {
      return jsonResponse({ success: false, code: 'ALREADY_BOOKED', message: "Ce weekend vient d'être réservé par quelqu'un d'autre 😅" });
    }

    sheet.getRange(rowIndex, 2).setValue(STATUT_RESERVE);
    sheet.getRange(rowIndex, 3).setValue(sanitizeCell(nom));
    sheet.getRange(rowIndex, 4).setValue(nbPersonnes);
    sheet.getRange(rowIndex, 5).setValue(sanitizeCell(email));
    sheet.getRange(rowIndex, 6).setValue(sanitizeCell(message));
    sheet.getRange(rowIndex, 7).setValue(new Date());

    var mailSent = true;
    try {
      MailApp.sendEmail({
        to: email,
        subject: 'Weekend à St-Georges-sur-Cher confirmé 🍷',
        body: 'Salut ' + nom + ',\n\nTon weekend du ' + label + ' est confirmé !\n\nÀ bientôt.'
      });
    } catch (mailErr) {
      mailSent = false;
    }

    if (!mailSent) {
      return jsonResponse({ success: true, message: "Réservation confirmée, mais l'email de confirmation n'a pas pu partir 📭" });
    }
    return jsonResponse({ success: true, message: 'Réservation confirmée ! Tu vas recevoir un email récapitulatif 🎉' });
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------
// VALIDATION EMAIL
// ------------------------------------------------------------
function isValidEmail(email) {
  if (!email || email.length > 254) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// ------------------------------------------------------------
// PROTECTION INJECTION DE FORMULE (A3)
// Préfixe d'une apostrophe toute valeur commençant par =, +, -, @
// pour empêcher l'exécution d'une formule dans le Sheet.
// ------------------------------------------------------------
function sanitizeCell(value) {
  var s = String(value);
  if (/^[=+\-@]/.test(s)) return "'" + s;
  return s;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
