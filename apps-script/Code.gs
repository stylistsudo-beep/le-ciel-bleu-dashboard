/**
 * LE CIEL BLEU 卸営業ダッシュボード連携スクリプト
 *
 * 使い方（スプレッドシートのメニューを探さなくてOKな方法）：
 * 1. ブラウザで https://script.google.com/ を開く（Googleにログインした状態で）。
 * 2. 「＋ 新しいプロジェクト」をクリック。
 * 3. 開いたエディタの中身を全部消して、このファイルの内容をまるごと貼り付けて保存
 *    （Cmd/Ctrl + S、プロジェクト名は何でもOK）。
 * 4. エディタ上部の関数選択で「seedInitialData」を選び、実行ボタン（▷）を1回だけ押す。
 *    → 初回のみ権限の許可を求められるので許可する
 *      （「Google で確認されていません」という警告が出たら「詳細」→「(プロジェクト名)に移動」で進める。
 *        自分で作成したスクリプトなので問題ありません）。
 *    → スプレッドシート「le ciel blue　営業先リスト」に「アプリ連携」という新しいシートが作られ、
 *      現在アプリに入っている28件が書き込まれる。
 * 5. 「デプロイ → 新しいデプロイ」→ 種類の歯車アイコンから「ウェブアプリ」を選択。
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 *    でデプロイし、発行された「ウェブアプリのURL」をコピーする。
 * 6. アプリ側（ダッシュボード）の「連携設定」にそのURLを貼り付ける。
 * 7. パスワード保護をかける場合：エディタ左側の「プロジェクトの設定」（歯車アイコン）→
 *    「スクリプト プロパティ」→ プロパティを追加：
 *      プロパティ: APP_TOKEN
 *      値: 好きなパスワード
 *    を保存する。これでアプリを開く際にこのパスワードの入力が必須になる
 *    （このファイルにパスワードは一切書き込まれないので、コードを公開しても安全）。
 *
 * ※コードを更新した場合、既存のデプロイには自動反映されません。
 *   「デプロイ→デプロイを管理→編集（鉛筆アイコン）→バージョン：新バージョン→デプロイ」で反映してください
 *   （URLは変わりません）。
 *
 * 以後、シートを直接編集してもアプリ起動時に反映され、
 * アプリ側で編集・追加・削除すると、このシートにも自動で反映されます。
 *
 * ※スプレッドシートの「拡張機能 → Apps Script」から開ける環境の場合は、
 *   そちらから同じコードを貼り付けても構いません（どちらでも動作します）。
 */

// 対象スプレッドシートのID（URLの https://docs.google.com/spreadsheets/d/【ここ】/edit の部分）
const SPREADSHEET_ID = "1mF99ZmjdRykW2hVt3LSIzQLB3PU2EXQSQa-FW1NTg2A";

const SHEET_NAME = "アプリ連携";
const HEADERS = ["ID", "店名", "地域", "優先度", "ステータス", "次のアクション", "次のアクション期限", "更新日時"];

// パスワード保護：スクリプト プロパティ「APP_TOKEN」が設定されている場合のみ、
// 一致するtokenを送ってきたリクエストだけを許可する（未設定なら保護なしで従来通り動作）。
function getRequiredToken_() {
  return PropertiesService.getScriptProperties().getProperty("APP_TOKEN") || "";
}

function isAuthorized_(token) {
  const required = getRequiredToken_();
  if (!required) return true;
  return token === required;
}

function getSpreadsheet_() {
  // スプレッドシートに紐づけて実行している場合はそちらを優先し、
  // standalone（script.google.comで直接作成）の場合はIDで開く。
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {
    // no-op
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  const token = (e && e.parameter && e.parameter.token) || "";
  if (!isAuthorized_(token)) {
    return jsonResponse_({ ok: false, error: "unauthorized" });
  }
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const stores = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // IDが空の行はスキップ
    stores.push({
      id: String(row[0]),
      name: row[1],
      region: row[2],
      priority: row[3],
      status: row[4],
      nextAction: row[5] || "",
      nextActionDate: formatDate_(row[6]),
    });
  }
  return jsonResponse_({ ok: true, stores: stores });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!isAuthorized_(body.token || "")) {
      return jsonResponse_({ ok: false, error: "unauthorized" });
    }
    const action = body.action;
    const sheet = getSheet_();

    if (action === "add") {
      const id = Utilities.getUuid();
      sheet.appendRow([
        id,
        body.name || "",
        body.region || "",
        body.priority || "",
        body.status || "",
        body.nextAction || "",
        body.nextActionDate || "",
        new Date(),
      ]);
      return jsonResponse_({ ok: true, id: id });
    }

    if (action === "update") {
      const rowIndex = findRowById_(sheet, body.id);
      if (rowIndex === -1) return jsonResponse_({ ok: false, error: "not_found" });
      sheet.getRange(rowIndex, 2, 1, 6).setValues([[
        body.name || "",
        body.region || "",
        body.priority || "",
        body.status || "",
        body.nextAction || "",
        body.nextActionDate || "",
      ]]);
      sheet.getRange(rowIndex, 8).setValue(new Date());
      return jsonResponse_({ ok: true });
    }

    if (action === "delete") {
      const rowIndex = findRowById_(sheet, body.id);
      if (rowIndex === -1) return jsonResponse_({ ok: false, error: "not_found" });
      sheet.deleteRow(rowIndex);
      return jsonResponse_({ ok: true });
    }

    return jsonResponse_({ ok: false, error: "unknown_action" });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function findRowById_(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-indexed + header
  }
  return -1;
}

function formatDate_(value) {
  if (!value) return null;
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value);
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 「アプリ連携」シートが存在しない場合にだけ、見出し行を作成する。
 * すでにセットアップ済みのスプレッドシートでは何もしない（実データの上書きを防止）。
 * 新しくこのシートで始める場合の参考として、サンプル行を1件だけ追加する。
 */
function seedInitialData() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) {
    return; // 既存データがある場合は何もしない
  }
  sheet = ss.insertSheet(SHEET_NAME);
  sheet.appendRow(HEADERS);
  sheet.setFrozenRows(1);

  const now = new Date();
  sheet.appendRow([Utilities.getUuid(), "サンプル店舗", "香川", "B", "未接触", "", "", now]);
}
