const APP = Object.freeze({
  title: 'Trade Compass',
  propertyKey: 'TRADE_NOTE_SPREADSHEET_ID',
  sheets: {
    trades: { name: 'Trades', headers: ['ID','種別','日付','銘柄コード','銘柄名','数量','単価','平均取得単価','利確目標','売却元ID','実現損益','備考','作成日時','更新日時'] },
    stocks: { name: 'Stocks', headers: ['ID','銘柄コード','銘柄名','購入目標値','現在値','年初来高値','年初来安値','取得日時','備考','作成日時','更新日時','利確目標','決算情報','決算取得日時'] },
    policies: { name: 'Policies', headers: ['ID','日付','鉄のおきて','常に上位','作成日時','更新日時'] },
    favorites: { name: 'Favorites', headers: ['ID','銘柄コード','銘柄名','市場','業種','投資理由','注目材料','リスク','目標株価','メモ','作成日時','更新日時'] }
  }
});

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle(APP.title).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getInitialData() {
  const configured = Boolean(PropertiesService.getScriptProperties().getProperty(APP.propertyKey));
  if (!configured) return { configured: false, appTitle: APP.title };
  ensureSheets_();
  return Object.assign({ configured: true, appTitle: APP.title }, getAllData_());
}

function saveSettings(input) {
  const id = extractSpreadsheetId_(input && input.spreadsheetId);
  if (!id) throw new Error('スプレッドシートのURLまたはIDを入力してください。');
  const ss = SpreadsheetApp.openById(id);
  PropertiesService.getScriptProperties().setProperty(APP.propertyKey, id);
  ensureSheets_();
  return { ok: true, name: ss.getName() };
}

function createSpreadsheet() {
  const ss = SpreadsheetApp.create('Trade Compass データ');
  PropertiesService.getScriptProperties().setProperty(APP.propertyKey, ss.getId());
  ensureSheets_();
  return { ok: true, name: ss.getName(), url: ss.getUrl() };
}

function addPurchase(input) {
  validateRequired_(input, ['date','symbol','name','quantity','price']);
  const quantity = positiveNumber_(input.quantity, '購入数');
  const price = positiveNumber_(input.price, '取得単価');
  const id = Utilities.getUuid();
  const now = new Date();
  appendObject_('trades', {
    'ID': id, '種別': '購入', '日付': dateValue_(input.date), '銘柄コード': clean_(input.symbol).toUpperCase(),
    '銘柄名': clean_(input.name), '数量': quantity, '単価': price,
    '平均取得単価': numberOrBlank_(input.averagePrice) || price, '利確目標': numberOrBlank_(input.targetPrice),
    '売却元ID': '', '実現損益': '', '備考': clean_(input.notes), '作成日時': now, '更新日時': now
  });
  return { ok: true, id: id };
}

function addSale(input) {
  validateRequired_(input, ['date','purchaseId','quantity','price']);
  const quantity = positiveNumber_(input.quantity, '売却数');
  const price = positiveNumber_(input.price, '売却単価');
  const purchase = findTrade_(input.purchaseId);
  if (!purchase || purchase.type !== '購入') throw new Error('紐づける購入記録が見つかりません。');
  const remaining = getRemainingQuantity_(purchase.id);
  if (quantity > remaining) throw new Error('売却数が残数量（' + remaining + '）を超えています。');
  const avg = Number(purchase.averagePrice || purchase.price);
  const profit = Math.round((price - avg) * quantity * 100) / 100;
  const now = new Date();
  appendObject_('trades', {
    'ID': Utilities.getUuid(), '種別': '売却', '日付': dateValue_(input.date), '銘柄コード': purchase.symbol,
    '銘柄名': purchase.name, '数量': quantity, '単価': price, '平均取得単価': avg, '利確目標': '',
    '売却元ID': purchase.id, '実現損益': profit, '備考': clean_(input.notes), '作成日時': now, '更新日時': now
  });
  return { ok: true, profit: profit };
}

function deleteTrade(id) {
  const trade = findTrade_(id);
  if (!trade) throw new Error('記録が見つかりません。');
  if (trade.type === '購入' && getSalesForPurchase_(id).length) throw new Error('売却が紐づいている購入記録は削除できません。先に売却記録を削除してください。');
  deleteRowById_('trades', id);
  return { ok: true };
}

function saveStock(input) {
  validateRequired_(input, ['symbol','name']);
  upsert_('stocks', input.id, {
    '銘柄コード': clean_(input.symbol).toUpperCase(), '銘柄名': clean_(input.name),
    '購入目標値': numberListForCell_(input.buyTargets), '現在値': numberOrBlank_(input.currentPrice),
    '年初来高値': numberOrBlank_(input.yearHigh), '年初来安値': numberOrBlank_(input.yearLow),
    '取得日時': input.quotedAt ? new Date(input.quotedAt) : '', '備考': clean_(input.notes),
    '利確目標': numberListForCell_(input.takeProfitTargets), '決算情報': clean_(input.earningsData),
    '決算取得日時': input.earningsFetchedAt ? new Date(input.earningsFetchedAt) : ''
  });
  return { ok: true };
}

function fetchMarketData(symbol) {
  const ticker = normalizeTicker_(symbol);
  if (!ticker) throw new Error('銘柄コードを入力してください。');
  const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(ticker) + '?range=1y&interval=1d&events=history';
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (response.getResponseCode() !== 200) throw new Error('株価を取得できませんでした。銘柄コードを確認してください。');
  const json = JSON.parse(response.getContentText());
  const result = json && json.chart && json.chart.result && json.chart.result[0];
  if (!result) throw new Error('株価データが見つかりません。');
  const quote = result.indicators && result.indicators.quote && result.indicators.quote[0];
  const highs = (quote && quote.high || []).filter(isFiniteNumber_);
  const lows = (quote && quote.low || []).filter(isFiniteNumber_);
  const meta = result.meta || {};
  if (!highs.length || !lows.length) throw new Error('1年分の高値・安値を取得できませんでした。');
  return {
    symbol: meta.symbol || ticker, name: meta.longName || meta.shortName || '', currency: meta.currency || '',
    currentPrice: meta.regularMarketPrice || lastFinite_(quote.close), yearHigh: Math.max.apply(null, highs),
    yearLow: Math.min.apply(null, lows), quotedAt: new Date().toISOString()
  };
}

function fetchRecentEarnings(symbol) {
  const ticker = normalizeTicker_(symbol);
  if (!ticker) throw new Error('銘柄コードを入力してください。');
  const end = Math.floor(Date.now() / 1000);
  const start = end - 60 * 60 * 24 * 365 * 3;
  const types = ['quarterlyDilutedEPS','quarterlyTotalRevenue','quarterlyNetIncome'];
  const url = 'https://query1.finance.yahoo.com/ws/fundamentals-timeseries/v1/finance/timeseries/' + encodeURIComponent(ticker) +
    '?symbol=' + encodeURIComponent(ticker) + '&type=' + encodeURIComponent(types.join(',')) +
    '&period1=' + start + '&period2=' + end + '&merge=false';
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions:true, headers:{'User-Agent':'Mozilla/5.0'} });
  if (response.getResponseCode() !== 200) throw new Error('決算情報を取得できませんでした。時間をおいて再度お試しください。');
  const json = JSON.parse(response.getContentText());
  const results = json && json.timeseries && json.timeseries.result || [];
  function latest_(type) {
    const series = results.find(r => r.meta && r.meta.type && r.meta.type.indexOf(type) >= 0);
    const rows = series && series[type] || [];
    return rows.length ? rows[rows.length - 1] : null;
  }
  const eps = latest_('quarterlyDilutedEPS');
  const revenue = latest_('quarterlyTotalRevenue');
  const income = latest_('quarterlyNetIncome');
  const latest = [eps,revenue,income].filter(Boolean).sort((a,b) => String(b.asOfDate||'').localeCompare(String(a.asOfDate||'')))[0];
  if (!latest) throw new Error('直近の決算数値が見つかりません。');
  return {
    symbol:ticker, period:latest.asOfDate || '', epsActual:reportedValue_(eps),
    epsEstimate:null, epsSurprisePercent:null, revenue:reportedValue_(revenue),
    netIncome:reportedValue_(income), currency:(latest.currencyCode || ''),
    nextEarningsDate:'', fetchedAt:new Date().toISOString()
  };
}

function reportedValue_(row) {
  return row && row.reportedValue ? rawValue_(row.reportedValue) : null;
}

function savePolicy(input) {
  validateRequired_(input, ['date','rule']);
  upsert_('policies', input.id, { '日付': dateValue_(input.date), '鉄のおきて': clean_(input.rule), '常に上位': Boolean(input.pinned) });
  return { ok: true };
}

function saveFavorite(input) {
  validateRequired_(input, ['symbol','name']);
  upsert_('favorites', input.id, {
    '銘柄コード': clean_(input.symbol).toUpperCase(), '銘柄名': clean_(input.name), '市場': clean_(input.market),
    '業種': clean_(input.sector), '投資理由': clean_(input.thesis), '注目材料': clean_(input.catalyst),
    'リスク': clean_(input.risk), '目標株価': numberOrBlank_(input.targetPrice), 'メモ': clean_(input.notes)
  });
  return { ok: true };
}

function deleteRecord(type, id) {
  if (!['stocks','policies','favorites'].includes(type)) throw new Error('削除対象が不正です。');
  deleteRowById_(type, id);
  return { ok: true };
}

function getAllData_() {
  const trades = readObjects_('trades').map(mapTrade_);
  const purchases = trades.filter(t => t.type === '購入').map(p => {
    const sold = trades.filter(s => s.type === '売却' && s.purchaseId === p.id).reduce((sum, s) => sum + Number(s.quantity), 0);
    return Object.assign({}, p, { soldQuantity: sold, remainingQuantity: Number(p.quantity) - sold });
  });
  const sales = trades.filter(t => t.type === '売却');
  const realized = sales.reduce((sum, t) => sum + Number(t.profit || 0), 0);
  const invested = purchases.reduce((sum, p) => sum + p.remainingQuantity * Number(p.averagePrice || p.price), 0);
  const wins = sales.filter(t => Number(t.profit) > 0).length;
  return {
    trades: trades.sort(sortDateDesc_), purchases: purchases.sort(sortDateDesc_), openPurchases: purchases.filter(p => p.remainingQuantity > 0),
    stocks: readObjects_('stocks').map(mapStock_).sort(sortUpdatedDesc_),
    policies: readObjects_('policies').map(mapPolicy_).sort((a,b) => Number(b.pinned)-Number(a.pinned) || sortDateDesc_(a,b)),
    favorites: readObjects_('favorites').map(mapFavorite_).sort(sortUpdatedDesc_),
    summary: { realizedProfit: realized, openCost: invested, tradeCount: trades.length, winRate: sales.length ? Math.round(wins / sales.length * 1000) / 10 : 0 }
  };
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(APP.propertyKey);
  if (!id) throw new Error('先に設定タブで保存先スプレッドシートを設定してください。');
  return SpreadsheetApp.openById(id);
}

function ensureSheets_() {
  const ss = getSpreadsheet_();
  Object.keys(APP.sheets).forEach(key => {
    const def = APP.sheets[key];
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) sheet = ss.insertSheet(def.name);
    const existing = sheet.getLastColumn() ? sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(), def.headers.length)).getValues()[0] : [];
    def.headers.forEach((header, i) => { if (existing[i] !== header) sheet.getRange(1,i+1).setValue(header); });
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,def.headers.length).setFontWeight('bold').setBackground('#12261f').setFontColor('#ffffff');
  });
}

function readObjects_(key) {
  const def = APP.sheets[key];
  const sheet = getSpreadsheet_().getSheetByName(def.name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2,1,sheet.getLastRow()-1,def.headers.length).getValues().map(row => {
    const obj = {}; def.headers.forEach((h,i) => obj[h] = row[i]); return obj;
  }).filter(obj => obj.ID);
}

function appendObject_(key, obj) {
  const def = APP.sheets[key];
  const sheet = getSpreadsheet_().getSheetByName(def.name);
  sheet.appendRow(def.headers.map(h => Object.prototype.hasOwnProperty.call(obj,h) ? obj[h] : ''));
}

function upsert_(key, id, values) {
  const def = APP.sheets[key];
  const sheet = getSpreadsheet_().getSheetByName(def.name);
  const now = new Date();
  if (id) {
    const row = findRowById_(sheet, id);
    if (!row) throw new Error('更新対象が見つかりません。');
    const existing = sheet.getRange(row,1,1,def.headers.length).getValues()[0];
    const obj = {}; def.headers.forEach((h,i) => obj[h] = existing[i]);
    Object.assign(obj, values, {'更新日時': now});
    sheet.getRange(row,1,1,def.headers.length).setValues([def.headers.map(h => obj[h])]);
  } else {
    appendObject_(key, Object.assign({'ID': Utilities.getUuid(), '作成日時': now, '更新日時': now}, values));
  }
}

function deleteRowById_(key, id) {
  const sheet = getSpreadsheet_().getSheetByName(APP.sheets[key].name);
  const row = findRowById_(sheet, id);
  if (!row) throw new Error('削除対象が見つかりません。');
  sheet.deleteRow(row);
}

function findRowById_(sheet, id) {
  if (sheet.getLastRow() < 2) return 0;
  const found = sheet.getRange(2,1,sheet.getLastRow()-1,1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return found ? found.getRow() : 0;
}

function findTrade_(id) { return readObjects_('trades').map(mapTrade_).find(t => t.id === id); }
function getSalesForPurchase_(id) { return readObjects_('trades').map(mapTrade_).filter(t => t.type === '売却' && t.purchaseId === id); }
function getRemainingQuantity_(id) { const p=findTrade_(id); return Number(p.quantity)-getSalesForPurchase_(id).reduce((s,t)=>s+Number(t.quantity),0); }

function mapTrade_(o) { return { id:String(o.ID),type:String(o['種別']),date:toDateString_(o['日付']),symbol:String(o['銘柄コード']||''),name:String(o['銘柄名']||''),quantity:Number(o['数量']||0),price:Number(o['単価']||0),averagePrice:Number(o['平均取得単価']||0),targetPrice:numberOrNull_(o['利確目標']),purchaseId:String(o['売却元ID']||''),profit:numberOrNull_(o['実現損益']),notes:String(o['備考']||''),updatedAt:toIso_(o['更新日時'])}; }
function mapStock_(o) { return { id:String(o.ID),symbol:String(o['銘柄コード']||''),name:String(o['銘柄名']||''),buyTargets:numberListFromCell_(o['購入目標値']),takeProfitTargets:numberListFromCell_(o['利確目標']),currentPrice:numberOrNull_(o['現在値']),yearHigh:numberOrNull_(o['年初来高値']),yearLow:numberOrNull_(o['年初来安値']),quotedAt:toIso_(o['取得日時']),earningsData:jsonObjectOrNull_(o['決算情報']),earningsFetchedAt:toIso_(o['決算取得日時']),notes:String(o['備考']||''),updatedAt:toIso_(o['更新日時'])}; }
function mapPolicy_(o) { return { id:String(o.ID),date:toDateString_(o['日付']),rule:String(o['鉄のおきて']||''),pinned:o['常に上位']===true||String(o['常に上位']).toLowerCase()==='true',updatedAt:toIso_(o['更新日時'])}; }
function mapFavorite_(o) { return { id:String(o.ID),symbol:String(o['銘柄コード']||''),name:String(o['銘柄名']||''),market:String(o['市場']||''),sector:String(o['業種']||''),thesis:String(o['投資理由']||''),catalyst:String(o['注目材料']||''),risk:String(o['リスク']||''),targetPrice:numberOrNull_(o['目標株価']),notes:String(o['メモ']||''),updatedAt:toIso_(o['更新日時'])}; }

function validateRequired_(obj, fields) { fields.forEach(k => { if (obj[k] === undefined || obj[k] === null || String(obj[k]).trim() === '') throw new Error('必須項目を入力してください。'); }); }
function positiveNumber_(v,label) { const n=Number(v); if(!isFinite(n)||n<=0) throw new Error(label+'は0より大きい数値を入力してください。'); return n; }
function numberOrBlank_(v) { return v === '' || v === null || v === undefined ? '' : Number(v); }
function numberOrNull_(v) { return v === '' || v === null || v === undefined ? null : Number(v); }
function numberListForCell_(v) { const values=Array.isArray(v)?v:String(v||'').split(','); const nums=values.map(Number).filter(n=>isFinite(n)&&n>0); return nums.length?JSON.stringify(nums):''; }
function numberListFromCell_(v) { if(v===''||v===null||v===undefined)return []; try{const a=JSON.parse(String(v));if(Array.isArray(a))return a.map(Number).filter(n=>isFinite(n)&&n>0)}catch(e){} const n=Number(v);return isFinite(n)&&n>0?[n]:[]; }
function jsonObjectOrNull_(v) { if(!v)return null; try{const o=JSON.parse(String(v));return o&&typeof o==='object'?o:null}catch(e){return null} }
function rawValue_(v) { return v&&typeof v==='object'&&Object.prototype.hasOwnProperty.call(v,'raw')?v.raw:(v===undefined?null:v); }
function clean_(v) { return String(v == null ? '' : v).trim(); }
function dateValue_(v) { const d=new Date(v+'T00:00:00'); if(isNaN(d)) throw new Error('日付が不正です。'); return d; }
function toDateString_(v) { if(!v)return ''; const d=v instanceof Date?v:new Date(v); return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd'); }
function toIso_(v) { if(!v)return ''; const d=v instanceof Date?v:new Date(v); return isNaN(d)?'':d.toISOString(); }
function extractSpreadsheetId_(v) { const s=clean_(v); const m=s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/); return m?m[1]:s.match(/^[a-zA-Z0-9-_]+$/)?s:''; }
function normalizeTicker_(v) { const s=clean_(v).toUpperCase(); return /^\d{4}$/.test(s)?s+'.T':s; }
function isFiniteNumber_(v) { return typeof v==='number'&&isFinite(v); }
function lastFinite_(arr) { for(let i=(arr||[]).length-1;i>=0;i--)if(isFiniteNumber_(arr[i]))return arr[i]; return null; }
function sortDateDesc_(a,b) { return String(b.date||'').localeCompare(String(a.date||'')) || String(b.updatedAt||'').localeCompare(String(a.updatedAt||'')); }
function sortUpdatedDesc_(a,b) { return String(b.updatedAt||'').localeCompare(String(a.updatedAt||'')); }
