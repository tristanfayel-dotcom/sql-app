// ----------------------------------------------------------------------
// API backend pour l'app "Titulaires & Remplaçants SQL", en remplacement
// de l'Apps Script "Application Web" (trop lent en démarrage à froid).
// Parle directement à l'API Google Sheets v4 avec un compte de service,
// sans dépendance npm (juste le module "crypto" natif de Node + fetch,
// tous deux déjà présents dans l'environnement Netlify Functions), pour
// rester déployable même sans étape de build locale.
//
// Variables d'environnement requises (à définir dans Netlify, jamais dans
// ce fichier) : GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY,
// GOOGLE_SPREADSHEET_ID.
// ----------------------------------------------------------------------

const SHEET_TITULAIRES = 'Titulaires';
const SHEET_REMPLACANTS = 'Remplacants';
const SHEET_MANCHE = 'Manche';
const SHEET_REPONSES = 'Reponses';
const SHEET_ADMIN = 'Admin';

// ---------- Auth : JWT signé (RS256) -> jeton d'accès OAuth2 ----------
let cachedToken = null; // { token, expiresAt } — réutilisé tant que le conteneur reste "chaud"

function base64url(input){
  return Buffer.from(input)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

function signJwt(){
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const key = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: 'RS256', typ: 'JWT' };
  const claim = {
    iss: email,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600
  };
  const signingInput = base64url(JSON.stringify(header)) + '.' + base64url(JSON.stringify(claim));
  const crypto = require('crypto');
  const signer = crypto.createSign('RSA-SHA256');
  signer.update(signingInput);
  signer.end();
  const signature = signer.sign(key)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
  return signingInput + '.' + signature;
}

async function getAccessToken(){
  if(cachedToken && cachedToken.expiresAt > Date.now() + 30000){
    return cachedToken.token;
  }
  const jwt = signJwt();
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: 'grant_type=' + encodeURIComponent('urn:ietf:params:oauth:grant-type:jwt-bearer') + '&assertion=' + encodeURIComponent(jwt)
  });
  const data = await res.json();
  if(!res.ok || !data.access_token){
    throw new Error('Auth Google échouée : ' + JSON.stringify(data));
  }
  cachedToken = { token: data.access_token, expiresAt: Date.now() + (data.expires_in * 1000) };
  return data.access_token;
}

// ---------- Petits wrappers autour de l'API Sheets v4 ----------
async function sheetsFetch(path, options){
  const token = await getAccessToken();
  const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + path;
  const res = await fetch(url, Object.assign({}, options, {
    headers: Object.assign({ 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' }, (options && options.headers) || {})
  }));
  const data = await res.json();
  if(!res.ok){
    throw new Error('Erreur Sheets API : ' + JSON.stringify(data));
  }
  return data;
}

async function batchGet(ranges){
  // valueRenderOption=UNFORMATTED_VALUE : sans ça, l'API Sheets renvoie les
  // nombres au format d'affichage de la cellule (ex. "51,46" en texte, avec
  // une virgule) plutôt que le nombre brut 51.46, ce qui cassait le calcul
  // des moyennes (Number("51,46") = NaN en JS).
  const q = ranges.map(function(r){ return 'ranges=' + encodeURIComponent(r); }).join('&')
    + '&valueRenderOption=UNFORMATTED_VALUE';
  const data = await sheetsFetch('/values:batchGet?' + q, { method: 'GET' });
  return data.valueRanges.map(function(vr){ return vr.values || []; });
}

async function updateRange(range, values){
  return sheetsFetch('/values/' + encodeURIComponent(range) + '?valueInputOption=RAW', {
    method: 'PUT',
    body: JSON.stringify({ range: range, values: values })
  });
}

async function appendRange(range, values){
  return sheetsFetch('/values/' + encodeURIComponent(range) + ':append?valueInputOption=RAW', {
    method: 'POST',
    body: JSON.stringify({ values: values })
  });
}

async function clearRange(range){
  return sheetsFetch('/values/' + encodeURIComponent(range) + ':clear', { method: 'POST', body: '{}' });
}

// ---------- Logique métier (reprise de Code.gs) ----------
async function readState(){
  const [tRows, rRows, mRows, respRows] = await batchGet([
    SHEET_TITULAIRES + '!A1:N',
    SHEET_REMPLACANTS + '!A1:C',
    SHEET_MANCHE + '!A1:A2',
    SHEET_REPONSES + '!A1:I'
  ]);

  const teams = [];
  for(let i = 1; i < tRows.length; i++){
    const row = tRows[i];
    if(!row || !row[0]) continue;
    const titulaires = [];
    for(let k = 0; k < 4; k++){
      const name = row[5 + k*2], moy = row[6 + k*2];
      if(name) titulaires.push({ n: String(name), m: Number(moy) || 0 });
    }
    teams.push({
      id: row[0], gender: row[1], group: row[2], category: row[3], label: row[4],
      titulaires: titulaires,
      max: (row[13] === '' || row[13] === undefined || row[13] === null) ? null : Number(row[13])
    });
  }

  const remplacantsM = [], remplacantsF = [];
  for(let j = 1; j < rRows.length; j++){
    const rr = rRows[j];
    if(!rr || !rr[1]) continue;
    const entry = { n: String(rr[1]), m: Number(rr[2]) || 0 };
    if(rr[0] === 'M') remplacantsM.push(entry); else if(rr[0] === 'F') remplacantsF.push(entry);
  }

  const manche = mRows.length > 1 ? Number(mRows[1][0]) : 1;

  const responses = {};
  for(let x = 1; x < respRows.length; x++){
    const rw = respRows[x];
    if(!rw || Number(rw[0]) !== manche) continue;
    const needReplacement = !!rw[3];
    const mode = rw[6] ? String(rw[6]) : (needReplacement ? 'pool' : 'none');
    let borrowed = [];
    if(rw[7]){ try{ borrowed = JSON.parse(String(rw[7])) || []; }catch(e){ borrowed = []; } }
    responses[rw[1]] = {
      present: rw[2] ? String(rw[2]).split('|') : [],
      needReplacement: needReplacement,
      replacements: rw[4] ? String(rw[4]).split('|') : [],
      updatedAt: rw[5],
      mode: mode,
      borrowed: borrowed,
      validated: !!rw[8]
    };
  }

  return { manche: manche, teams: teams, remplacantsM: remplacantsM, remplacantsF: remplacantsF, responses: responses };
}

async function checkPin(pin){
  const [rows] = await batchGet([SHEET_ADMIN + '!B1']);
  const real = rows && rows[0] ? String(rows[0][0]) : '';
  return String(pin || '') === real;
}

async function verifyPin(pin){
  return { ok: await checkPin(pin) };
}

async function getCurrentManche(){
  const [rows] = await batchGet([SHEET_MANCHE + '!A1:A2']);
  return rows.length > 1 ? Number(rows[1][0]) : 1;
}

async function saveResponse(teamId, patch){
  const manche = await getCurrentManche();
  const [respRows] = await batchGet([SHEET_REPONSES + '!A1:I']);
  let rowIndex = -1; // 1-based ligne réelle dans la feuille
  for(let i = 1; i < respRows.length; i++){
    if(Number(respRows[i][0]) === manche && respRows[i][1] === teamId){ rowIndex = i + 1; break; }
  }
  const existing = rowIndex > 0 ? respRows[rowIndex - 1] : null;
  const present = patch.present !== undefined ? patch.present : (existing ? String(existing[2] || '').split('|').filter(Boolean) : []);
  const needReplacement = patch.needReplacement !== undefined ? patch.needReplacement : (existing ? !!existing[3] : false);
  const replacements = patch.replacements !== undefined ? patch.replacements : (existing ? String(existing[4] || '').split('|').filter(Boolean) : []);
  const mode = patch.mode !== undefined ? patch.mode : (existing ? String(existing[6] || 'none') : 'none');
  let borrowed = patch.borrowed !== undefined ? patch.borrowed : [];
  if(patch.borrowed === undefined && existing && existing[7]){
    try{ borrowed = JSON.parse(String(existing[7])) || []; }catch(e){ borrowed = []; }
  }
  const validated = patch.validated !== undefined ? patch.validated : (existing ? !!existing[8] : false);
  const row = [manche, teamId, present.join('|'), needReplacement ? 1 : 0, replacements.join('|'), Date.now(), mode, JSON.stringify(borrowed), validated ? 1 : 0];

  if(rowIndex > 0){
    await updateRange(SHEET_REPONSES + '!A' + rowIndex + ':I' + rowIndex, [row]);
  } else {
    await appendRange(SHEET_REPONSES + '!A1:I1', [row]);
  }
  return { ok: true };
}

async function setManche(n){
  await updateRange(SHEET_MANCHE + '!A2', [[n]]);
  return readState();
}

async function clearManche(pin){
  if(!(await checkPin(pin))) return { error: 'pin' };
  const manche = (await readState()).manche;
  const [respRows] = await batchGet([SHEET_REPONSES + '!A1:I']);
  const header = respRows[0];
  const kept = respRows.slice(1).filter(function(rw){ return Number(rw[0]) !== manche; });
  await clearRange(SHEET_REPONSES + '!A2:I100000');
  if(kept.length){
    await updateRange(SHEET_REPONSES + '!A2:I' + (kept.length + 1), kept);
  }
  return readState();
}

async function saveAdmin(pin, payload){
  if(!(await checkPin(pin))) return { error: 'pin' };

  const teamRows = payload.teams.map(function(tm){
    let row = [tm.id, tm.gender, tm.group, tm.category, tm.label];
    for(let i = 0; i < 4; i++){
      if(tm.titulaires[i]) row = row.concat([tm.titulaires[i].n, tm.titulaires[i].m]); else row = row.concat(['', '']);
    }
    row.push(tm.max === null || tm.max === undefined ? '' : tm.max);
    return row;
  });
  await clearRange(SHEET_TITULAIRES + '!A2:N100000');
  if(teamRows.length) await updateRange(SHEET_TITULAIRES + '!A2:N' + (teamRows.length + 1), teamRows);

  const rempRows = [];
  payload.remplacantsM.forEach(function(p){ rempRows.push(['M', p.n, p.m]); });
  payload.remplacantsF.forEach(function(p){ rempRows.push(['F', p.n, p.m]); });
  await clearRange(SHEET_REMPLACANTS + '!A2:C100000');
  if(rempRows.length) await updateRange(SHEET_REMPLACANTS + '!A2:C' + (rempRows.length + 1), rempRows);

  return readState();
}

// ---------- Répartiteur d'actions ----------
const ACTIONS = { readState, verifyPin, saveResponse, setManche, clearManche, saveAdmin };

exports.handler = async function(event){
  const headers = {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS'
  };
  if(event.httpMethod === 'OPTIONS'){
    return { statusCode: 204, headers: headers, body: '' };
  }
  try{
    let action, args;
    if(event.httpMethod === 'POST'){
      const body = JSON.parse(event.body || '{}');
      action = body.action; args = body.args || [];
    } else {
      const qs = event.queryStringParameters || {};
      action = qs.action;
      args = qs.args ? JSON.parse(qs.args) : [];
    }
    const fn = ACTIONS[action];
    if(!fn) return { statusCode: 400, headers: headers, body: JSON.stringify({ __error: 'Action inconnue : ' + action }) };
    const result = await fn.apply(null, args);
    return { statusCode: 200, headers: headers, body: JSON.stringify(result) };
  } catch(err){
    return { statusCode: 200, headers: headers, body: JSON.stringify({ __error: String(err && err.message || err) }) };
  }
};
