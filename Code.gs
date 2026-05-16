// ============================================================
//  BUDGET TGR — Google Apps Script Web App
//  Sheet ID : 1BaZFoPOFi6E-VXtFANxPD5ExKo7RWrdASqStAQqr-kQ
// ============================================================

const SPREADSHEET_ID = '1jqHsmn0aFQGQ5TTs0RceuIb2aFxp_eeXD5xfJNWFnrk';

// ────────────────────────────────────────────────────────────
//  SETUP  — exécuter UNE SEULE FOIS après le premier déploiement
// ────────────────────────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Supprimer la feuille vide par défaut si elle existe
  const defaultSheet = ss.getSheetByName('Feuille 1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  _initSheet(ss, 'Config',       ['key', 'value']);
  _initSheet(ss, 'Users',        ['id','username','password','role','active','createdAt']);
  _initSheet(ss, 'Transactions', ['id','date','type','amount','category','description','username','createdAt']);
  _initSheet(ss, 'Categories',   ['id','name','type','active']);

  // Config
  const cfg = ss.getSheetByName('Config');
  cfg.getRange(2,1,3,2).setValues([
    ['householdName','TGR'],
    ['currency','€'],
    ['setupDone','true']
  ]);

  // Utilisateurs
  const usr = ss.getSheetByName('Users');
  usr.getRange(2,1,2,6).setValues([
    [1,'Glory','123456abc','admin','true', new Date().toISOString()],
    [2,'Ruth', '1234',     'user', 'true', new Date().toISOString()]
  ]);

  // Catégories
  const cat = ss.getSheetByName('Categories');
  const cats = [
    [1,'Loyer/Charges','expense','true'],
    [2,'Courses alimentaires','expense','true'],
    [3,'Transport','expense','true'],
    [4,'Santé','expense','true'],
    [5,'Loisirs','expense','true'],
    [6,'Vêtements','expense','true'],
    [7,'Restaurants','expense','true'],
    [8,'Téléphone/Internet','expense','true'],
    [9,'Éducation','expense','true'],
    [10,'Autres dépenses','expense','true'],
    [11,'Salaire','income','true'],
    [12,'Prime','income','true'],
    [13,'Freelance','income','true'],
    [14,'Remboursement','income','true'],
    [15,'Autres revenus','income','true']
  ];
  cat.getRange(2,1,cats.length,4).setValues(cats);

  Logger.log('✅ Setup terminé avec succès !');
  return 'Setup OK';
}

function _initSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  else sh.clearContents();
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  return sh;
}

// ────────────────────────────────────────────────────────────
//  CORS + ROUTAGE
// ────────────────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'login')
    return _resp(handleLogin(e.parameter.username, e.parameter.password));

  const user = _verifyToken(e.parameter.token);
  if (!user) return _resp({success:false, error:'Non autorisé'});

  switch(action) {
    case 'getTransactions':   return _resp(getTransactions(e.parameter, user));
    case 'getCategories':     return _resp(getCategories());
    case 'getDashboard':      return _resp(getDashboard(e.parameter));
    case 'getUsers':          return _resp(getUsers(user));
    case 'getConfig':         return _resp(getConfig());
    case 'getMonthlySummary': return _resp(getMonthlySummary(e.parameter));
    case 'getAllCategories':   return _resp(getAllCategories(user));
    default: return _resp({success:false, error:'Action inconnue'});
  }
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch(_) { return _resp({success:false, error:'JSON invalide'}); }

  const user = _verifyToken(body.token);
  if (!user) return _resp({success:false, error:'Non autorisé'});

  switch(body.action) {
    case 'addTransaction':    return _resp(addTransaction(body, user));
    case 'updateTransaction': return _resp(updateTransaction(body, user));
    case 'deleteTransaction': return _resp(deleteTransaction(body, user));
    case 'addCategory':       return _resp(addCategory(body, user));
    case 'updateCategory':    return _resp(updateCategory(body, user));
    case 'deleteCategory':    return _resp(deleteCategory(body, user));
    case 'addUser':           return _resp(addUser(body, user));
    case 'toggleUser':        return _resp(toggleUser(body, user));
    case 'updatePassword':    return _resp(updatePassword(body, user));
    default: return _resp({success:false, error:'Action inconnue'});
  }
}

function _resp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ────────────────────────────────────────────────────────────
//  AUTH
// ────────────────────────────────────────────────────────────
function handleLogin(username, password) {
  const data = _sheet('Users').getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const [id, uname, pwd, role, active] = data[i];
    if (uname.toString().toLowerCase() === username.toString().toLowerCase()
      && pwd.toString() === password.toString()
      && active.toString() === 'true') {
      const token = _makeToken(id, uname, role);
      return {success:true, token, user:{id, username:uname, role}};
    }
  }
  return {success:false, error:'Identifiants incorrects'};
}

function _makeToken(id, username, role) {
  const payload = `${id}|${username}|${role}|${new Date().toDateString()}`;
  return Utilities.base64Encode(payload);
}

function _verifyToken(token) {
  if (!token) return null;
  try {
    const raw = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const p = raw.split('|');
    if (p.length < 4 || p[3] !== new Date().toDateString()) return null;
    return {id:p[0], username:p[1], role:p[2]};
  } catch(_) { return null; }
}

// ────────────────────────────────────────────────────────────
//  TRANSACTIONS
// ────────────────────────────────────────────────────────────
function getTransactions(params, user) {
  const data = _sheet('Transactions').getDataRange().getValues();
  let rows = [];
  for (let i = 1; i < data.length; i++) {
    const [id,date,type,amount,category,description,username,createdAt] = data[i];
    if (!id) continue;
    rows.push({id, date:_fmtDate(date), type, amount:parseFloat(amount)||0, category, description, username, createdAt:_fmtDate(createdAt)});
  }
  if (params.type)     rows = rows.filter(r => r.type === params.type);
  if (params.category) rows = rows.filter(r => r.category === params.category);
  if (params.username) rows = rows.filter(r => r.username === params.username);
  if (params.month)    rows = rows.filter(r => r.date.startsWith(params.month));
  rows.sort((a,b) => new Date(b.date) - new Date(a.date));
  return {success:true, data:rows};
}

function addTransaction(body, user) {
  const sh = _sheet('Transactions');
  const data = sh.getDataRange().getValues();
  const id = _nextId(data);
  const date = body.date || Utilities.formatDate(new Date(),'UTC','yyyy-MM-dd');
  sh.appendRow([id, date, body.type, parseFloat(body.amount)||0, body.category, body.description||'', user.username, new Date().toISOString()]);
  return {success:true, id};
}

function updateTransaction(body, user) {
  const sh = _sheet('Transactions');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() !== body.id.toString()) continue;
    if (user.role !== 'admin' && data[i][6] !== user.username)
      return {success:false, error:'Permission refusée'};
    sh.getRange(i+1,2,1,5).setValues([[body.date, body.type, parseFloat(body.amount)||0, body.category, body.description||'']]);
    return {success:true};
  }
  return {success:false, error:'Introuvable'};
}

function deleteTransaction(body, user) {
  const sh = _sheet('Transactions');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() !== body.id.toString()) continue;
    if (user.role !== 'admin' && data[i][6] !== user.username)
      return {success:false, error:'Permission refusée'};
    sh.deleteRow(i+1);
    return {success:true};
  }
  return {success:false, error:'Introuvable'};
}

// ────────────────────────────────────────────────────────────
//  DASHBOARD
// ────────────────────────────────────────────────────────────
function getDashboard(params) {
  const data = _sheet('Transactions').getDataRange().getValues();
  const now = new Date();
  const currentMonth = params.month || _monthKey(now);

  let mIncome = 0, mExpense = 0;
  const byCategory = {}, byMonth = {};
  const recent = [];

  for (let i = 1; i < data.length; i++) {
    const [id,date,type,amount,category,description,username] = data[i];
    if (!id) continue;
    const d = new Date(date);
    const mk = _monthKey(d);
    const amt = parseFloat(amount)||0;

    if (mk === currentMonth) {
      if (type === 'income') mIncome += amt; else mExpense += amt;
      if (type === 'expense') byCategory[category] = (byCategory[category]||0) + amt;
    }

    if (!byMonth[mk]) byMonth[mk] = {income:0, expense:0};
    if (type === 'income') byMonth[mk].income += amt; else byMonth[mk].expense += amt;

    recent.push({id, date:_fmtDate(date), type, amount:amt, category, description, username});
  }

  recent.sort((a,b) => new Date(b.date) - new Date(a.date));

  // Garder les 6 derniers mois pour le graphique barres
  const sortedMonths = Object.keys(byMonth).sort().slice(-12);
  const chartMonths = {};
  sortedMonths.forEach(m => chartMonths[m] = byMonth[m]);

  return {success:true, data:{currentMonth, mIncome, mExpense, balance:mIncome-mExpense, byCategory, byMonth:chartMonths, recent:recent.slice(0,10)}};
}

// ────────────────────────────────────────────────────────────
//  RÉCAP MENSUEL
// ────────────────────────────────────────────────────────────
function getMonthlySummary(params) {
  const data = _sheet('Transactions').getDataRange().getValues();
  const now = new Date();
  const month = params.month || _monthKey(now);
  const summary = {income:{}, expense:{}};
  let totalIncome = 0, totalExpense = 0;

  for (let i = 1; i < data.length; i++) {
    const [id,date,type,amount,category] = data[i];
    if (!id) continue;
    if (!_fmtDate(date).startsWith(month)) continue;
    const amt = parseFloat(amount)||0;
    if (type === 'income') { summary.income[category] = (summary.income[category]||0)+amt; totalIncome += amt; }
    else { summary.expense[category] = (summary.expense[category]||0)+amt; totalExpense += amt; }
  }
  return {success:true, data:{month, summary, totalIncome, totalExpense, balance:totalIncome-totalExpense}};
}

// ────────────────────────────────────────────────────────────
//  CATÉGORIES
// ────────────────────────────────────────────────────────────
function getCategories() {
  const data = _sheet('Categories').getDataRange().getValues();
  const cats = [];
  for (let i = 1; i < data.length; i++) {
    const [id,name,type,active] = data[i];
    if (!id || active.toString() !== 'true') continue;
    cats.push({id, name, type});
  }
  return {success:true, data:cats};
}

function getAllCategories(user) {
  if (user.role !== 'admin') return {success:false, error:'Admin uniquement'};
  const data = _sheet('Categories').getDataRange().getValues();
  const cats = [];
  for (let i = 1; i < data.length; i++) {
    const [id,name,type,active] = data[i];
    if (!id) continue;
    cats.push({id, name, type, active:active.toString()});
  }
  return {success:true, data:cats};
}

function addCategory(body, user) {
  if (user.role !== 'admin') return {success:false, error:'Admin uniquement'};
  const sh = _sheet('Categories');
  const data = sh.getDataRange().getValues();
  sh.appendRow([_nextId(data), body.name, body.type, 'true']);
  return {success:true};
}

function updateCategory(body, user) {
  if (user.role !== 'admin') return {success:false, error:'Admin uniquement'};
  const sh = _sheet('Categories');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() !== body.id.toString()) continue;
    sh.getRange(i+1,2,1,2).setValues([[body.name, body.type]]);
    return {success:true};
  }
  return {success:false, error:'Introuvable'};
}

function deleteCategory(body, user) {
  if (user.role !== 'admin') return {success:false, error:'Admin uniquement'};
  const sh = _sheet('Categories');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() !== body.id.toString()) continue;
    sh.getRange(i+1,4).setValue('false');
    return {success:true};
  }
  return {success:false, error:'Introuvable'};
}

// ────────────────────────────────────────────────────────────
//  UTILISATEURS
// ────────────────────────────────────────────────────────────
function getUsers(user) {
  if (user.role !== 'admin') return {success:false, error:'Admin uniquement'};
  const data = _sheet('Users').getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const [id,username,,role,active,createdAt] = data[i];
    if (!id) continue;
    users.push({id, username, role, active:active.toString(), createdAt:_fmtDate(createdAt)});
  }
  return {success:true, data:users};
}

function addUser(body, user) {
  if (user.role !== 'admin') return {success:false, error:'Admin uniquement'};
  const sh = _sheet('Users');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString().toLowerCase() === body.username.toString().toLowerCase())
      return {success:false, error:'Nom d\'utilisateur déjà pris'};
  }
  sh.appendRow([_nextId(data), body.username, body.password, body.role||'user', 'true', new Date().toISOString()]);
  return {success:true};
}

function toggleUser(body, user) {
  if (user.role !== 'admin') return {success:false, error:'Admin uniquement'};
  const sh = _sheet('Users');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() !== body.id.toString()) continue;
    sh.getRange(i+1,5).setValue(data[i][4].toString()==='true' ? 'false' : 'true');
    return {success:true};
  }
  return {success:false, error:'Utilisateur introuvable'};
}

function updatePassword(body, user) {
  if (user.role !== 'admin') return {success:false, error:'Admin uniquement'};
  const sh = _sheet('Users');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() !== body.id.toString()) continue;
    sh.getRange(i+1,3).setValue(body.password);
    return {success:true};
  }
  return {success:false, error:'Utilisateur introuvable'};
}

// ────────────────────────────────────────────────────────────
//  CONFIG
// ────────────────────────────────────────────────────────────
function getConfig() {
  const data = _sheet('Config').getDataRange().getValues();
  const cfg = {};
  for (let i = 1; i < data.length; i++) cfg[data[i][0]] = data[i][1];
  return {success:true, data:cfg};
}

// ────────────────────────────────────────────────────────────
//  HELPERS
// ────────────────────────────────────────────────────────────
function _sheet(name) { return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name); }
function _nextId(data) { let m=0; for(let i=1;i<data.length;i++){const n=parseInt(data[i][0])||0; if(n>m)m=n;} return m+1; }
function _monthKey(d) { return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`; }
function _fmtDate(d) {
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d,'UTC','yyyy-MM-dd');
  return d.toString().substring(0,10);
}
