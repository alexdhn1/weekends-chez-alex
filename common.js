/* ============================================================
   common.js — helpers partagés entre index.html et vacances.html
   (LOT A du PRD d'audit : A2/A3/A5/A8)
============================================================ */

// ------------------------------------------------------------
// Token partagé envoyé sur toutes les requêtes (GET et POST) vers
// les Apps Script du dossier appscript/. Remplace cette valeur par
// le token que tu as mis dans les Script Properties de CHAQUE
// Apps Script (clé API_TOKEN) — voir appscript/README.md.
// ------------------------------------------------------------
var API_TOKEN = 'uVJdFa0CV7H4uJvGv6xWsGu3I_P2ND3a';

// ------------------------------------------------------------
// Détection mobile via screen.width (pas affecté par les iframes
// dans lesquelles Apps Script embarque parfois la page, contrairement
// aux media queries CSS classiques).
// ------------------------------------------------------------
function detectMobile() {
  if (window.screen.width <= 768 || window.innerWidth <= 768) {
    document.body.classList.add('is-mobile');
  }
}

// ------------------------------------------------------------
// GET JSON avec token en query string + vérification de r.ok.
// params (optionnel) : objet de paramètres additionnels (ex action).
// ------------------------------------------------------------
function getJson(url, params) {
  var query = 'token=' + encodeURIComponent(API_TOKEN);
  if (params) {
    Object.keys(params).forEach(function(key) {
      query += '&' + encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
    });
  }
  var sep = url.indexOf('?') === -1 ? '?' : '&';
  return fetch(url + sep + query).then(function(r) {
    if (!r.ok) throw new Error('http_' + r.status);
    return r.json();
  });
}

// ------------------------------------------------------------
// POST JSON avec token dans le corps + vérification de r.ok.
// Content-Type text/plain : Apps Script exige ce type pour éviter
// le preflight CORS.
// ------------------------------------------------------------
function postJson(url, data) {
  var body = {};
  if (data) {
    Object.keys(data).forEach(function(key) { body[key] = data[key]; });
  }
  body.token = API_TOKEN;
  return fetch(url, {
    method:  'POST',
    headers: { 'Content-Type': 'text/plain' },
    body:    JSON.stringify(body)
  }).then(function(r) {
    if (!r.ok) throw new Error('http_' + r.status);
    return r.json();
  });
}

// ------------------------------------------------------------
// Affiche un message texte dans un élément, avec une classe de base
// (ex "message" sur index.html, "msg" sur vacances.html) + un type
// ("success" / "error"). Toujours via textContent (jamais innerHTML).
// ------------------------------------------------------------
function showMessage(elId, texte, type, baseClass) {
  var el = document.getElementById(elId);
  if (!el) return;
  el.textContent = texte;
  el.className = (baseClass || 'message') + ' ' + type;
}
