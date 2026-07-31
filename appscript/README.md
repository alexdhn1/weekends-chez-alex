# appscript/ — backends Google Apps Script proposés

Ce dossier contient le code **proposé** pour les deux backends Apps
Script utilisés par les pages statiques du repo (`index.html` et
`vacances.html`). Ces scripts ne font pas partie de ce repo statique :
ils vivent normalement dans deux projets Apps Script liés chacun à un
Google Sheet, et sont déployés en application web (`/exec`).

- `reservations.gs` → backend de `index.html` (réservation du weekend
  "groupe" à St-Georges-sur-Cher).
- `vacances.gs` → backend de `vacances.html` (planning Alex & Jenna).

## 1. Pourquoi ce code existe

Suite à l'audit sécurité du repo (LOT A), les deux pages front
envoient désormais un **token** sur chaque requête (GET et POST), via
la constante `API_TOKEN` définie dans `common.js`. Ce dossier fournit
le code serveur correspondant :

- vérification du token contre une Script Property,
- validations serveur (jamais confiance au front seul),
- `LockService` sur les écritures pour éviter les races (deux
  réservations simultanées sur le même créneau, etc.),
- sanitisation anti-injection de formule avant tout `setValue` de
  donnée utilisateur dans un Sheet.

## 2. Installation pas à pas (à faire par le propriétaire du Sheet)

Pour **chacun** des deux scripts :

1. Ouvre le Google Sheet correspondant (celui qui contient l'onglet
   `Weekends` pour les réservations, ou les onglets `Events` /
   `Imported` / `Ponts` pour le planning vacances).
2. Menu **Extensions > Apps Script**.
3. Si un script existe déjà : remplace son contenu par celui du
   fichier `.gs` correspondant de ce dossier (`reservations.gs` ou
   `vacances.gs`). Sinon, crée un nouveau fichier et colle le contenu.
4. Menu **Project Settings** (l'icône ⚙️ dans la barre latérale) >
   section **Script Properties** > **Add script property** :
   - Clé : `API_TOKEN`
   - Valeur : un secret long et aléatoire que tu génères toi-même
     (ex. avec un gestionnaire de mots de passe). **Ne réutilise pas
     un mot de passe existant.**
5. Reporte cette même valeur dans `common.js` du repo statique, à la
   place de `'CHANGE_ME_TOKEN'` :
   ```js
   var API_TOKEN = 'la-valeur-que-tu-as-mise-dans-Script-Properties';
   ```
   Ce fichier est chargé par les deux pages (`index.html` et
   `vacances.html`), donc un seul endroit à modifier côté front.
6. **Déployer > Gérer les déploiements > ✏️ (modifier)** sur le
   déploiement existant, puis choisis **Nouvelle version** et
   redéploie. (Si tu modifies le code d'un déploiement existant sans
   changer de version, l'URL `/exec` change de comportement immédiatement
   pour tout le monde — c'est voulu ici, mais fais-le en dehors des
   heures d'utilisation si possible.)
7. Vérifie que l'URL `/exec` affichée après déploiement correspond bien
   à celle déjà présente dans `API_URL` / `API` des fichiers HTML. Si
   Apps Script en génère une nouvelle, mets-la à jour dans
   `index.html` (`API_URL`) ou `vacances.html` (`API`).

## 3. Structure des Google Sheets attendue

### Pour `reservations.gs` (onglet `Weekends`)

| Colonne | Contenu |
|---|---|
| A | `label` (ex "28-29 mars 2026") |
| B | `statut` (`Libre` ou `Réservé`) |
| C | `nom` (rempli à la réservation) |
| D | `nbPersonnes` (rempli à la réservation) |
| E | `email` (rempli à la réservation) |
| F | `message` (rempli à la réservation) |
| G | `dateReservation` (rempli à la réservation) |

Ligne 1 = en-têtes. Ligne 2 et suivantes = un créneau par ligne.

### Pour `vacances.gs`

- Onglet **`Events`** (créé automatiquement avec ses en-têtes si absent) :
  `id | dateStart | dateEnd | categorie | personne | note`
- Onglet **`Imported`** (optionnel, lecture seule, même structure) :
  pour des événements importés d'un calendrier externe. S'il n'existe
  pas, l'API renvoie simplement une liste vide.
- Onglet **`Ponts`** (remplace l'ancien tableau `PONTS_ALEX` hardcodé
  dans `vacances.html`) : `dateStart | dateEnd | note`. Ajoute/édite
  les jours de pont directement dans cet onglet — plus besoin de
  toucher au code front. S'il n'existe pas, `ponts` renvoie `[]`.

## 4. Ce que fait le code livré

- **Auth (A2)** : `checkToken()` compare le token reçu (query string
  en GET, champ `token` du JSON en POST) à la Script Property
  `API_TOKEN`. Si absent ou différent → réponse `unauthorized` /
  `{success:false}`.
- **Validations serveur (A3)** :
  - `nbPersonnes` entre 1 et 10 (réservations),
  - `dateEnd >= dateStart` (planning),
  - email validé par regex, longueur max 254,
  - nom max 100 caractères, message/note max 1000 caractères,
  - catégorie restreinte à une énumération connue (planning).
- **LockService** : toute écriture (réservation, ajout, suppression)
  est protégée par `LockService.getScriptLock()` /
  `waitLock(10000)` / `finally releaseLock()`, pour éviter qu'une
  double soumission quasi simultanée ne corrompe les données.
- **Anti-injection de formule** : `sanitizeCell()` préfixe d'une
  apostrophe toute valeur utilisateur commençant par `=`, `+`, `-` ou
  `@` avant écriture dans le Sheet (protection contre les formules
  malveillantes ouvertes ensuite dans Sheets/Excel).
- **Identifiant de créneau stable (réservations)** : le serveur revérifie
  que `label` correspond toujours au contenu de la ligne `rowIndex`
  avant de réserver ; sinon `code: 'STALE_SLOT'`.
- **Échec d'envoi d'email non bloquant** : si `MailApp.sendEmail`
  échoue, la réservation reste actée mais le message renvoyé au client
  prévient que l'email de confirmation n'est pas parti.

## 5. Limitations connues / à surveiller

- `LockService.getScriptLock()` protège contre les écritures
  concurrentes **au sein d'un même script**. Il ne protège pas contre
  une modification manuelle du Sheet pendant l'exécution.
- Le token est un secret partagé simple (pas d'OTP, pas d'expiration).
  Il suffit pour dissuader le scraping/spam automatisé basique, mais
  ne remplace pas une authentification utilisateur réelle. Change-le
  si tu soupçonnes une fuite (ex. capture réseau côté client — le
  token est visible dans le code source et les requêtes réseau, ce
  qui est inhérent à une app 100% statique sans backend d'auth dédié).
