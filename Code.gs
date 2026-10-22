/**
 * Centr Supply Ops — back-end Google Sheets
 * ------------------------------------------------------------
 * 1. Ouvre ta Google Sheet > Extensions > Apps Script, colle ce fichier.
 * 2. Exécute la fonction setup() une fois (autorise l'accès) :
 *    elle crée les onglets et affiche ta CLÉ D'ACCÈS dans le journal.
 * 3. Déployer > Nouveau déploiement > Application Web
 *    - Exécuter en tant que : Moi
 *    - Qui a accès : Tout le monde
 *    Copie l'URL qui finit par /exec et colle-la dans Réglages de l'app.
 * Après chaque modification de ce script : Déployer > Gérer les déploiements
 * > modifier > Nouvelle version (sinon l'ancienne version reste en ligne).
 */

const SCHEMA = {
  clients:    ['id','nom','type','contact','telephone','email','prix_prep','prix_upsell','prix_retour','emballage','consignes','portail','notes','actif','created','updated'],
  produits:   ['id','client_id','sku','nom','emplacement','prix','cout','seuil','delai','couverture','fefo','actif','created','updated'],
  commandes:  ['id','numero','date','client_id','destinataire','lignes','ups','statut','date_exp','transporteur','suivi','montant','paiement','montant_paye','date_paiement','photos','notes','created','updated'],
  retours:    ['id','dossier','date','client_id','commande_id','produit_id','qte','motif','etat','decision','notes','created','updated'],
  mouvements: ['id','date','produit_id','type','qte','lot','dlc','note','created','updated'],
  frais:      ['id','date','client_id','libelle','qte','pu','montant','paiement','montant_paye','date_paiement','auto','notes','created','updated'],
  provisions: ['id','date','client_id','type','sens','montant','ref','note','created','updated'],
  parametres: ['id','valeur','created','updated'],
  journal:    ['id','date','user','sheet','ref','action','created','updated']
};
// Champs date au format AAAA-MM-JJ
const JOURNAL_MAX = 1500; // nb max de lignes d'historique renvoyées à l'app
// Colonnes laissées telles quelles (dates -> vraies dates dans la Sheet)
const RAW_FIELDS = ['date', 'date_paiement', 'date_exp', 'dlc'];

/** À lancer une fois (et à relancer sans risque si tu ajoutes des colonnes). */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function (name) {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const hdr = SCHEMA[name];
    const lastCol = Math.max(sh.getLastColumn(), 1);
    const cur = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    if (cur.join('') === '') {
      sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
    } else {
      hdr.forEach(function (h) {
        if (cur.indexOf(h) < 0) sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      });
    }
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#e2ede7');
    sh.setFrozenRows(1);
  });
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('TOKEN', token);
  }
  DriveApp.getRootFolder(); // autorise l'accès à Google Drive (photos des colis)
  Logger.log('✅ Onglets prêts. CLÉ D\'ACCÈS à coller dans l\'app : ' + token);
}

/** Pour changer de clé (si elle a fuité) : lance-la puis colle la nouvelle dans l'app. */
function nouvelleCle() {
  const token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('TOKEN', token);
  Logger.log('Nouvelle clé : ' + token);
}

function doGet() {
  return out({ ok: true, service: 'Centr Supply Ops' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return out({ ok: false, error: 'Requête invalide' }); }
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token) return out({ ok: false, error: 'Lance d\'abord setup() dans Apps Script' });
  // Lien client en lecture seule : n'expose que les données de ce client
  if (body.action === 'portal') {
    try { return out({ ok: true, data: portal(body.portail) }); }
    catch (err) { return out({ ok: false, error: String(err && err.message || err) }); }
  }
  if (body.token !== token) return out({ ok: false, error: 'Clé d\'accès invalide' });
  try {
    switch (body.action) {
      case 'ping': return out({ ok: true });
      case 'all':  return out({ ok: true, data: readAll() });
      case 'bulk': return out({ ok: true, n: bulk(body.ops || []) });
      case 'photo': return out({ ok: true, data: savePhoto(body) });
      default:     return out({ ok: false, error: 'Action inconnue' });
    }
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  }
}

/** Photo d'un colis -> Google Drive : Centr Supply — Photos colis / <client> / <AAAA-MM> / <n°>_<date>.jpg */
function savePhoto(b) {
  if (!b.data) throw new Error('Photo vide');
  const props = PropertiesService.getScriptProperties();
  let root = null;
  const rid = props.getProperty('PHOTO_DIR');
  if (rid) { try { root = DriveApp.getFolderById(rid); if (root.isTrashed()) root = null; } catch (e) { root = null; } }
  if (!root) { root = DriveApp.createFolder('Centr Supply — Photos colis'); props.setProperty('PHOTO_DIR', root.getId()); }
  const sub = function (parent, name) { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); };
  const t = String(b.t || new Date().toISOString());
  const folder = sub(sub(root, String(b.client || 'Sans client').slice(0, 80)), t.slice(0, 7));
  const name = String(b.numero || 'commande').replace(/[\\/:*?"<>|#]/g, '') + '_' + t.slice(0, 16).replace(/[:T]/g, '-') + '_' + String(b.k || '').slice(-5) + '.jpg';
  const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b.data), 'image/jpeg', name));
  f.setDescription('Commande ' + b.numero + ' · ' + b.client + ' · ' + t + (b.u ? ' · ' + b.u : ''));
  return { id: f.getId(), url: f.getUrl() };
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function readAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const res = {};
  Object.keys(SCHEMA).forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) { res[name] = []; return; }
    const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    const hdr = v[0].map(String);
    res[name] = v.slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) {
      const o = {};
      hdr.forEach(function (h, i) {
        if (!h) return;
        let x = r[i];
        if (x instanceof Date) x = h === 'date' && name === 'journal' ? x.toISOString() : Utilities.formatDate(x, tz, 'yyyy-MM-dd');
        o[h] = x;
      });
      return o;
    });
    if (name === 'journal' && res[name].length > JOURNAL_MAX) res[name] = res[name].slice(-JOURNAL_MAX);
  });
  return res;
}

function portal(tok) {
  tok = String(tok || '');
  if (tok.length < 16) throw new Error('Lien invalide');
  const all = readAll();
  const c = all.clients.filter(function (x) { return String(x.portail) === tok; })[0];
  if (!c) throw new Error('Ce lien a été désactivé ou n\'existe pas.');
  const cid = String(c.id);
  const mine = function (x) { return String(x.client_id) === cid; };
  const commandes = all.commandes.filter(mine);
  const own = all.produits.filter(mine);
  const ids = {}; own.forEach(function (p) { ids[String(p.id)] = 1; });
  const used = {};
  commandes.forEach(function (o) { try { JSON.parse(o.lignes || '[]').forEach(function (l) { used[String(l.p)] = 1; }); } catch (e) {} });
  const produits = own.concat(all.produits.filter(function (p) { return !ids[String(p.id)] && used[String(p.id)]; })
    .map(function (p) { return { id: p.id, client_id: p.client_id, nom: p.nom, sku: p.sku, actif: p.actif }; }));
  return {
    clients: [{ id: c.id, nom: c.nom, type: c.type, prix_prep: c.prix_prep, prix_upsell: c.prix_upsell, prix_retour: c.prix_retour, emballage: c.emballage, consignes: c.consignes, actif: true }],
    produits: produits,
    commandes: commandes,
    retours: all.retours.filter(mine),
    mouvements: all.mouvements.filter(function (m) { return ids[String(m.produit_id)]; }),
    frais: all.frais.filter(mine),
    provisions: (all.provisions || []).filter(mine),
    parametres: all.parametres.filter(function (x) { return x.id === 'prix'; }),
    journal: []
  };
}

/** Crée l'onglet ou les colonnes manquantes (mises à jour de l'app sans relancer setup). */
function ensureSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name]]).setFontWeight('bold').setBackground('#e2ede7');
    sh.setFrozenRows(1);
    return sh;
  }
  const cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  SCHEMA[name].forEach(function (h) {
    if (cur.indexOf(h) < 0) { sh.getRange(1, sh.getLastColumn() + 1).setValue(h).setFontWeight('bold'); cur.push(h); }
  });
  return sh;
}

function cell(h, v) {
  if (v === undefined || v === null) return '';
  // Texte qui commence par 0, +, =, -, @ : on force le texte (téléphones, SKU, anti-formule)
  if (typeof v === 'string' && RAW_FIELDS.indexOf(h) < 0 && /^[=+\-@0-9]/.test(v)) return "'" + v;
  return v;
}

function bulk(ops) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cache = {};
    function ctx(name) {
      if (cache[name]) return cache[name];
      if (!SCHEMA[name]) throw new Error('Onglet inconnu : ' + name);
      const sh = ensureSheet(ss, name);
      const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      const n = sh.getLastRow();
      const ids = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
      // base = nb de lignes déjà dans la Sheet ; pending = nouvelles lignes écrites en une fois à la fin
      return (cache[name] = { sh: sh, hdr: hdr, ids: ids, base: ids.length, pending: [] });
    }
    ops.forEach(function (op) {
      const c = ctx(op.sheet);
      if (op.action === 'upsert') {
        const row = op.row || {};
        const i = c.ids.indexOf(String(row.id));
        if (i >= 0 && i < c.base) {
          const range = c.sh.getRange(i + 2, 1, 1, c.hdr.length);
          const cur = range.getValues()[0];
          range.setValues([c.hdr.map(function (h, k) { return (h in row) ? cell(h, row[h]) : cur[k]; })]);
        } else if (i >= c.base) {
          const arr = c.pending[i - c.base];
          c.hdr.forEach(function (h, k) { if (h in row) arr[k] = cell(h, row[h]); });
        } else {
          c.pending.push(c.hdr.map(function (h) { return cell(h, row[h]); }));
          c.ids.push(String(row.id));
        }
      } else if (op.action === 'delete') {
        const j = c.ids.indexOf(String(op.id));
        if (j >= 0 && j < c.base) { c.sh.deleteRow(j + 2); c.ids.splice(j, 1); c.base--; }
        else if (j >= c.base) { c.pending.splice(j - c.base, 1); c.ids.splice(j, 1); }
      }
    });
    Object.keys(cache).forEach(function (name) {
      const c = cache[name];
      if (c.pending.length) c.sh.getRange(c.sh.getLastRow() + 1, 1, c.pending.length, c.hdr.length).setValues(c.pending);
    });
    SpreadsheetApp.flush();
    return ops.length;
  } finally {
    lock.releaseLock();
  }
}
