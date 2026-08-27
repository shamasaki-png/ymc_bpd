/**
 * トピック報告アプリ 用 Google Apps Script
 * -------------------------------------------------
 * このスクリプトは、既存の業務日報スプレッドシートに紐づけて設置し、
 * スマホの報告フォーム（index.html）から送信されたデータを
 * A〜H列に1行ずつ追記します。また、過去の報告の一覧取得・編集・削除にも
 * 対応し、複数人が同時に編集した場合の競合検知も行います。
 *
 * ▼ 設置手順
 * 1. マスターデータのスプレッドシートを開く
 * 2.「拡張機能」→「Apps Script」を開く
 * 3. デフォルトの Code.gs の中身を全て削除し、このファイルの内容を貼り付け
 * 4. 下記の SHEET_NAME を、実際の対象シート名に合わせて修正
 *    （タブに表示されているシート名。例のスプレッドシートでは「業務日報」）
 * 5.「デプロイ」→「新しいデプロイ」
 *      種類：ウェブアプリ
 *      説明：任意（例：トピック報告フォーム）
 *      次のユーザーとして実行：自分
 *      アクセスできるユーザー：全員
 *    → 「デプロイ」をクリックし、承認を行う
 * 6. 発行された「ウェブアプリのURL」をコピー
 * 7. index.html 内の APPS_SCRIPT_URL にそのURLを貼り付ける
 *
 * ※ 既にデプロイ済みの場合、コードを更新しただけでは反映されません。
 *   「デプロイ」→「デプロイを管理」→ 既存のデプロイの鉛筆アイコン →
 *   バージョン「新バージョン」を選択して「デプロイ」し直してください。
 *   （WebアプリのURL自体は変わりません）
 * -------------------------------------------------
 */

// 対象シート名（スプレッドシートのタブ名に合わせて変更してください）
const SHEET_NAME = "業務日報";

// A〜H列の並び（1始まり）
const COL = {
  date: 1,
  reporter: 2,
  customerType: 3,
  companyName: 4,
  officeName: 5,
  contactPerson: 6,
  method: 7,
  content: 8
};

const ROW_KEYS = ['date','reporter','customerType','companyName','officeName','contactPerson','method','content'];

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDateValue_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v ? String(v) : '';
}

// ---------------------------------------------
// GET: 一覧取得（?action=list）／ 疎通確認
// ---------------------------------------------
function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : null;
    if (action === 'list') {
      return listRows_();
    }
    return jsonResponse({ status: "ok", message: "トピック報告 API は正常に稼働しています。" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

function listRows_() {
  const sheet = getSheet_();
  if (!sheet) {
    return jsonResponse({ status: "error", message: "シートが見つかりません: " + SHEET_NAME });
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse({ status: "success", rows: [] });
  }
  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const rows = [];
  values.forEach(function (r, idx) {
    // 完全な空行はスキップ
    if (!r[0] && !r[1] && !r[7]) return;
    rows.push({
      rowNumber: idx + 2, // 実際のシート上の行番号（更新・削除時に使用）
      date: formatDateValue_(r[COL.date - 1]),
      reporter: r[COL.reporter - 1],
      customerType: r[COL.customerType - 1],
      companyName: r[COL.companyName - 1],
      officeName: r[COL.officeName - 1],
      contactPerson: r[COL.contactPerson - 1],
      method: r[COL.method - 1],
      content: r[COL.content - 1]
    });
  });
  rows.reverse(); // 新しい報告を先頭に
  return jsonResponse({ status: "success", rows: rows });
}

// ---------------------------------------------
// POST: 新規登録（create）／ 更新（update）／ 削除（delete）
// ---------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'create';
    if (action === 'update') return updateRow_(data);
    if (action === 'delete') return deleteRow_(data);
    return createRow_(data);
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

function validateRequired_(data) {
  const required = ["date", "reporter", "customerType", "companyName", "method", "content"];
  for (const key of required) {
    if (!data[key]) {
      return "必須項目が不足しています: " + key;
    }
  }
  return null;
}

function createRow_(data) {
  const errMsg = validateRequired_(data);
  if (errMsg) return jsonResponse({ status: "error", message: errMsg });

  const sheet = getSheet_();
  if (!sheet) {
    return jsonResponse({ status: "error", message: "シートが見つかりません: " + SHEET_NAME });
  }

  sheet.appendRow([
    data.date, data.reporter, data.customerType, data.companyName,
    data.officeName || "", data.contactPerson || "", data.method, data.content
  ]);

  return jsonResponse({ status: "success" });
}

// 現在シート上にある行の内容を、一覧取得時と同じ形のオブジェクトで取得
function getRowAsObject_(sheet, rowNumber) {
  const vals = sheet.getRange(rowNumber, 1, 1, 8).getValues()[0];
  return {
    date: formatDateValue_(vals[COL.date - 1]),
    reporter: vals[COL.reporter - 1] || '',
    customerType: vals[COL.customerType - 1] || '',
    companyName: vals[COL.companyName - 1] || '',
    officeName: vals[COL.officeName - 1] || '',
    contactPerson: vals[COL.contactPerson - 1] || '',
    method: vals[COL.method - 1] || '',
    content: vals[COL.content - 1] || ''
  };
}

// 2つの行データが一致するか（すべてのフィールドを文字列として比較）
function rowsEqual_(a, b) {
  for (const k of ROW_KEYS) {
    const av = (a[k] === null || a[k] === undefined) ? '' : String(a[k]).trim();
    const bv = (b[k] === null || b[k] === undefined) ? '' : String(b[k]).trim();
    if (av !== bv) return false;
  }
  return true;
}

// クライアントが編集を開始した時点のスナップショット(original)を検証し、
// その後に他の人が変更していないかを確認する。
// 戻り値: null なら問題なし。問題があれば jsonResponse をそのまま返せるオブジェクト。
function checkConflict_(sheet, rowNumber, original) {
  if (rowNumber > sheet.getLastRow() || rowNumber < 2) {
    return jsonResponse({
      status: "deleted",
      message: "この報告は既に他の人によって削除されています。一覧を更新してご確認ください。"
    });
  }
  if (original) {
    const current = getRowAsObject_(sheet, rowNumber);
    if (!rowsEqual_(current, original)) {
      return jsonResponse({
        status: "conflict",
        message: "この報告は、あなたが編集を開始した後に他の人が更新しています。最新の内容を確認してください。",
        current: current
      });
    }
  }
  return null;
}

function updateRow_(data) {
  const errMsg = validateRequired_(data);
  if (errMsg) return jsonResponse({ status: "error", message: errMsg });

  const rowNumber = Number(data.rowNumber);
  if (!rowNumber || rowNumber < 2) {
    return jsonResponse({ status: "error", message: "更新対象の行が指定されていません。" });
  }

  const sheet = getSheet_();
  if (!sheet) {
    return jsonResponse({ status: "error", message: "シートが見つかりません: " + SHEET_NAME });
  }

  const conflict = checkConflict_(sheet, rowNumber, data.original);
  if (conflict) return conflict;

  sheet.getRange(rowNumber, 1, 1, 8).setValues([[
    data.date, data.reporter, data.customerType, data.companyName,
    data.officeName || "", data.contactPerson || "", data.method, data.content
  ]]);

  return jsonResponse({ status: "success" });
}

function deleteRow_(data) {
  const rowNumber = Number(data.rowNumber);
  if (!rowNumber || rowNumber < 2) {
    return jsonResponse({ status: "error", message: "削除対象の行が指定されていません。" });
  }

  const sheet = getSheet_();
  if (!sheet) {
    return jsonResponse({ status: "error", message: "シートが見つかりません: " + SHEET_NAME });
  }

  const conflict = checkConflict_(sheet, rowNumber, data.original);
  if (conflict) return conflict;

  sheet.deleteRow(rowNumber);
  return jsonResponse({ status: "success" });
}
