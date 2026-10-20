/**
 * 名刺登録アプリ（Eightの簡易版）
 * スマホで撮る → Drive の文字読み取り → 確認画面で直して登録 → スプシに1行 → Gemini で提案余地商材と課題
 *
 * スクリプトプロパティ：
 *   SHEET_ID        登録先のスプレッドシートID
 *   FOLDER_ID       名刺画像を置くフォルダID
 *   GEMINI_API_KEY  Gemini API のキー（コードには書かない）
 *   GEMINI_MODEL    省略時 gemini-flash-lite-latest（最新の軽量版を自動で使う）
 * 「サービス」で Drive API（v3）を追加しておく。
 */

const SHEET_NAME = '名刺';
const HEADERS = ['登録日時', '登録者', '会社名', '部署・役職', '氏名', 'メール', '電話', '携帯', '住所', 'URL',
  '提案余地商材', '提案の理由', '想定される課題', '判定状態', 'メモ', '名刺画像', '読み取った文字'];
const NONE = '対象外・不明';
// 商材の強み・弱みは「商材」シートで現場が直す（AI はそれを読んで選ぶ）。下は最初に入れる中身
const PRODUCT_SHEET = '商材';
const PRODUCT_HEADERS = ['商材名', '強み', '弱み・向かない相手', '向いている会社'];
const PRODUCT_SEED = [
  ['LINE構築', 'LINE公式アカウント・Lステップの初期設計から構築まで任せられる。予約・問い合わせ・顧客管理をLINEにまとめられる',
    '構築後に社内で回す人がいないと効果が続かない', 'LINEをまだ使っていない／使っているが配信だけの店舗・来店型・予約型の会社'],
  ['LINE運用', '配信・分析・改善を代行し、社内の手間をかけずに成果を伸ばせる', '月額の費用が続く。LINEを持っていない会社は先に構築が要る',
    'すでにLINE公式を持っているが更新が止まっている・手が回っていない会社'],
  ['LINE研修', '社員が自分で構築・運用できるようになる。内製化で外注費を抑えられる', '担当者の時間を確保できない会社では定着しない',
    '社員数が多い・複数店舗・内製したい意向がある会社'],
  ['AI業務改善', '議事録・資料作成・問い合わせ対応などの定型業務をAIで減らせる', 'LINEと直接関係しない。業務の棚卸しに先方の協力が要る',
    '事務作業が多い・人手不足を感じている中小企業'],
  ['LINEグループまとめるくん', 'LINEグループに散らばった連絡・情報をまとめ、Chatwork・Slack・メールに届ける。手軽に始められる',
    'LINEグループを使っていない会社（LINE WORKS など）は対象外', '社内や取引先との連絡をLINEグループで回している会社。まずは小さく試したい相手']
];
const COL = Object.fromEntries(HEADERS.map((h, i) => [h, i + 1]));

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('名刺登録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 画面を Google の外（GitHub Pages）に置いたときの受け口。
 * Google にログインしたブラウザだと /exec の画面が開けないため、画面は外に置き、ここへ POST する。
 * 本文は text/plain の JSON（ブラウザの事前確認を起こさないため）：{action:'read'|'save', ...}
 */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'read') out = { ok: true, data: readCard(req.base64, req.mimeType) };
    else if (req.action === 'save') out = { ok: true, data: saveCard(req.card || {}) };
    else throw new Error('不明な操作です');
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function prop_(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v && fallback === undefined) throw new Error('スクリプトプロパティ ' + key + ' が未設定です');
  return v || fallback;
}

function sheet_() {
  const ss = SpreadsheetApp.openById(prop_('SHEET_ID'));
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 「商材」シート（無ければ作って最初の中身を入れる） */
function products_() {
  const ss = SpreadsheetApp.openById(prop_('SHEET_ID'));
  let sh = ss.getSheetByName(PRODUCT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PRODUCT_SHEET);
    sh.getRange(1, 1, 1, PRODUCT_HEADERS.length).setValues([PRODUCT_HEADERS]).setFontWeight('bold');
    sh.getRange(2, 1, PRODUCT_SEED.length, PRODUCT_HEADERS.length).setValues(PRODUCT_SEED);
    sh.setFrozenRows(1);
  }
  return sh.getDataRange().getValues().slice(1).filter(r => r[0]);
}

/** 初回だけ手で実行：スプレッドシート・画像フォルダ・2つのシート・1時間おきのトリガーを作る */
function initialSetup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SHEET_ID')) props.setProperty('SHEET_ID', SpreadsheetApp.create('名刺一覧（名刺登録アプリ）').getId());
  if (!props.getProperty('FOLDER_ID')) props.setProperty('FOLDER_ID', DriveApp.createFolder('名刺画像（名刺登録アプリ）').getId());
  sheet_();
  products_();
  const ss = SpreadsheetApp.openById(prop_('SHEET_ID'));
  const blank = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);
  setupTrigger();
  console.log('スプレッドシート: ' + ss.getUrl());
}

/** 画面から：画像（base64）を受け取り、保存して文字を読み、欄の候補を返す */
function readCard(base64, mimeType) {
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType || 'image/jpeg',
    'meishi_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '.jpg');
  const image = DriveApp.getFolderById(prop_('FOLDER_ID')).createFile(blob);

  // Drive の文字読み取り：画像を Google ドキュメントとして作ると OCR される
  const doc = Drive.Files.create(
    { name: 'ocr_tmp_' + image.getId(), mimeType: 'application/vnd.google-apps.document' },
    blob, { ocrLanguage: 'ja' });
  let text = '';
  try {
    text = DocumentApp.openById(doc.id).getBody().getText();
  } finally {
    DriveApp.getFileById(doc.id).setTrashed(true);
  }
  const fields = parseCard_(text);
  fields.imageUrl = image.getUrl();
  fields.rawText = text;
  return fields;
}

/** 読み取った文字から欄の候補を作る（確認画面で人が直す前提） */
function parseCard_(text) {
  const lines = text.split(/\n+/).map(s => s.trim()).filter(Boolean);
  const all = lines.join('\n');
  const pick = re => { const m = all.match(re); return m ? m[1].trim() : ''; };

  const email = pick(/([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})/);
  const url = pick(/((?:https?:\/\/|www\.)[A-Za-z0-9.\-\/_%?=&#~]+)/);
  // 電話：0始まりで数字10〜11桁（郵便番号を拾わない）。FAX の行は除く
  const phones = (all.replace(/FAX[^\n]*?(?=TEL|$)/gi, ' ')
    .match(/(?<![\d-])0\d{1,4}[-‐－ ()（）]{0,2}\d{1,4}[-‐－ ()（）]{0,2}\d{3,4}(?![\d-])/g) || [])
    .filter(p => { const n = p.replace(/\D/g, '').length; return n >= 10 && n <= 11; });
  const mobile = phones.find(p => /^0[789]0/.test(p.replace(/\D/g, ''))) || '';
  const tel = phones.find(p => p !== mobile) || '';
  const company = (lines.find(l => /(株式会社|有限会社|合同会社|一般社団法人|社団法人|財団法人|医療法人|学校法人|\(株\)|（株）|Inc\.|Co\.,? ?Ltd)/.test(l)) || '');
  // 住所：都道府県から始まる部分。次の行がビル名なら足す
  const PREF = /(北海道|東京都|京都府|大阪府|.{2,3}県)/;
  const addrIdx = lines.findIndex(l => PREF.test(l) && /[市区町村郡]/.test(l));
  let address = '';
  if (addrIdx >= 0) {
    address = lines[addrIdx].slice(lines[addrIdx].search(PREF)).replace(/\s*(TEL|Tel|FAX|Mail).*$/, '');
    const next = lines[addrIdx + 1] || '';
    if (/(ビル|階|F$|号室|館)/.test(next) && !/TEL|FAX|@/i.test(next)) address += ' ' + next;
  }
  // 氏名：「姓 名」の行、または「漢字の氏名＋ローマ字」の行の漢字部分
  let name = lines.find(l => /^[一-龥々ヶ]{1,4}[ 　][一-龥々ヶ]{1,4}$/.test(l)) || '';
  if (!name) {
    const m = lines.map(l => l.match(/^([一-龥々ヶ]{1,4}[ 　]?[一-龥々ヶ]{1,4})\s+[A-Z][a-z]+/)).find(Boolean);
    name = m ? m[1] : '';
  }
  // 役職：役職名を含む行を優先し、無ければ部署の行
  const notNC = l => l !== company && l !== name && l.length <= 50;
  const title = (lines.find(l => /(代表|取締役|社長|長$|長 |マネージャー|リーダー|Manager|CEO|COO|CTO|Sales|担当)/.test(l) && notNC(l)) ||
                 lines.find(l => /(部|課|室|グループ|事業部|オフィス)/.test(l) && notNC(l)) || '');
  return { company, title, name, email, tel, mobile, address, url };
}

/** 画面から：確認済みの内容を1行登録し、続けて提案余地商材を判定する */
function saveCard(f) {
  if (!f.company && !f.name) throw new Error('会社名か氏名のどちらかは入れてください');
  const sh = sheet_();
  const row = new Array(HEADERS.length).fill('');
  row[COL['登録日時'] - 1] = new Date();
  row[COL['登録者'] - 1] = f.registrant || Session.getActiveUser().getEmail() || '';
  row[COL['会社名'] - 1] = f.company || '';
  row[COL['部署・役職'] - 1] = f.title || '';
  row[COL['氏名'] - 1] = f.name || '';
  row[COL['メール'] - 1] = f.email || '';
  row[COL['電話'] - 1] = f.tel || '';
  row[COL['携帯'] - 1] = f.mobile || '';
  row[COL['住所'] - 1] = f.address || '';
  row[COL['URL'] - 1] = f.url || '';
  row[COL['判定状態'] - 1] = '未判定';
  row[COL['メモ'] - 1] = f.memo || '';
  row[COL['名刺画像'] - 1] = f.imageUrl || '';
  row[COL['読み取った文字'] - 1] = (f.rawText || '').slice(0, 2000);
  sh.appendRow(row);
  const r = sh.getLastRow();
  // 判定（HP読み込み＋Gemini）は待たずに返す。少し後に裏で動かす
  try { scheduleClassify_(); } catch (e) { console.warn('判定の予約に失敗。1時間おきのやり直しで拾います: ' + e); }
  return { row: r };
}

/** 30秒後に1回だけ classifySoon を動かす予約を置く（既に予約があれば置かない） */
function scheduleClassify_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'classifySoon');
    if (!exists) ScriptApp.newTrigger('classifySoon').timeBased().after(30 * 1000).create();
  } finally {
    lock.releaseLock();
  }
}

/** 登録直後の予約から呼ばれる：予約を消してから「未判定」の行を判定する */
function classifySoon() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'classifySoon')
    .forEach(t => ScriptApp.deleteTrigger(t));
  retryUnclassified();
}

/** 1時間おきのトリガー：「未判定」の行をやり直す */
function retryUnclassified() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const states = sh.getRange(2, COL['判定状態'], last - 1, 1).getValues();
  let done = 0;
  for (let i = 0; i < states.length && done < 20; i++) {
    if (states[i][0] !== '未判定') continue;
    try { classifyRow_(sh, i + 2); done++; } catch (e) { console.warn('行' + (i + 2) + ': ' + e); break; }
  }
}

/** 初回だけ手で実行：1時間おきのトリガーを作る */
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'retryUnclassified')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('retryUnclassified').timeBased().everyHours(1).create();
}

/** Gemini に渡すのは会社名とホームページの文章だけ（氏名・連絡先は渡さない） */
function classifyRow_(sh, r) {
  const company = sh.getRange(r, COL['会社名']).getValue();
  const url = sh.getRange(r, COL['URL']).getValue();
  const out3 = (p, reason, issue) =>
    sh.getRange(r, COL['提案余地商材'], 1, 4).setValues([[p, reason, issue, '判定済み']]);
  if (!company && !url) { out3(NONE, '会社名が無い', ''); return; }
  const products = products_();
  const names = products.map(p => p[0]);
  const site = fetchSiteText_(url);
  const prompt = [
    'あなたは株式会社MARKELINE（LINE公式アカウント・Lステップの構築／運用代行／研修、AI業務改善、LINEグループまとめるくん）の営業アシスタントです。',
    '名刺交換した次の会社について、①ホームページから読み取れる会社の課題、②こちらの商材で提案の余地があるもの、を答えてください。',
    'こちらの商材（強み／弱み・向かない相手／向いている会社）：',
    products.map(p => '・' + p[0] + '：強み=' + p[1] + '／弱み=' + p[2] + '／向いている=' + p[3]).join('\n'),
    '提案余地商材は上の商材名から最大2つ（合う順、「、」区切り）。どれも合わない・情報が足りないときは「' + NONE + '」。',
    '課題はホームページに書いてあることから推測し、書いていないことは作らない。',
    '会社名：' + company,
    'ホームページの文章（抜粋）：' + (site || '取得できず'),
    'JSONだけで答える：{"products":"商材名","reason":"60字以内の理由","issue":"60字以内の課題（分からなければ空）"}'
  ].join('\n');

  const model = prop_('GEMINI_MODEL', 'gemini-flash-lite-latest');
  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent',
    {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': prop_('GEMINI_API_KEY') },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { responseMimeType: 'application/json', temperature: 0 }
      })
    });
  if (res.getResponseCode() !== 200) throw new Error('Gemini ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
  const out = JSON.parse(JSON.parse(res.getContentText()).candidates[0].content.parts[0].text);
  const picked = String(out.products || '').split(/[、,／\/]/).map(x => x.trim()).filter(x => names.includes(x)).slice(0, 2);
  out3(picked.length ? picked.join('、') : NONE, String(out.reason || '').slice(0, 100), String(out.issue || '').slice(0, 100));
}

function fetchSiteText_(url) {
  if (!url) return '';
  try {
    const u = /^https?:/.test(url) ? url : 'https://' + url;
    const html = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true }).getContentText();
    return html.replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<[^>]+>/g, ' ').replace(/&[a-z]+;/g, ' ').replace(/\s+/g, ' ').slice(0, 3000);
  } catch (e) {
    return '';
  }
}
