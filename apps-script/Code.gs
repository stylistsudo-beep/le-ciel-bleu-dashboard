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
 *
 * 【スプレッドシート整理用（任意）】
 * 関数選択で「consolidateReferenceTabs」を選んで実行すると、
 * 香川・岡山・広島の候補店舗27件の詳細情報（住所・連絡先・取り扱いブランドなど）を
 * 「候補店舗リスト（詳細）」という1つの新しいタブにまとめます。
 * 「アプリ連携」タブには一切影響しません。元の調査タブもそのまま残るので、
 * 内容を見比べて問題なければ、古いタブは手動で削除してください。
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

    if (action === "reorder") {
      const orderIds = (body.order || []).map(String);
      const data = sheet.getDataRange().getValues();
      const rowsById = {};
      const extraRows = [];
      for (let i = 1; i < data.length; i++) {
        if (!data[i][0]) continue;
        rowsById[String(data[i][0])] = data[i];
      }
      const mentioned = {};
      const newRows = [];
      orderIds.forEach(function (id) {
        if (rowsById[id]) {
          newRows.push(rowsById[id]);
          mentioned[id] = true;
        }
      });
      // 並び替えリストに含まれていない行があれば末尾に残す（安全策）
      for (let i = 1; i < data.length; i++) {
        const id = String(data[i][0]);
        if (data[i][0] && !mentioned[id]) newRows.push(data[i]);
      }
      if (newRows.length > 0) {
        sheet.getRange(2, 1, newRows.length, HEADERS.length).setValues(newRows);
      }
      return jsonResponse_({ ok: true });
    }

    if (action === "replaceAll") {
      // CSVインポートなど、店舗一覧をまるごと入れ替える操作用。
      // 既存の行はいったんクリアし、渡された内容で新しいIDを振り直して書き込む。
      const list = body.stores || [];
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
      }
      const now = new Date();
      const rows = list.map(function (s) {
        return [
          Utilities.getUuid(),
          s.name || "",
          s.region || "",
          s.priority || "",
          s.status || "",
          s.nextAction || "",
          s.nextActionDate || "",
          now,
        ];
      });
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
      }
      return jsonResponse_({ ok: true, count: rows.length });
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

/**
 * 徳島・愛媛・高知の候補店舗21件を「アプリ連携」シートに追加し、
 * 香川・岡山・広島と同じくアプリで管理できるようにする。
 * すでに同名の店舗が「アプリ連携」にある場合はスキップする（二重追加防止）ので、
 * 何度実行しても安全。1回だけ手動実行すればよい。
 */
function addTokushimaEhimeKochiToApp() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const existingNames = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) existingNames[data[i][1]] = true;
  }

  // [店名, 地域, 優先度, ステータス, 次のアクション]
  const newStores = [
    // 徳島
    ["THE DAY", "徳島", "S", "アプローチ中", ""],
    ["roly poly", "徳島", "S", "未接触", ""],
    ["Hiraoka", "徳島", "S", "未接触", ""],
    ["Reply（リプリー）", "徳島", "B", "未接触", ""],
    ["Cherie（クレメントプラザ）", "徳島", "A", "未接触", ""],
    ["MAISON ÉCRU（メゾンエクリュ）", "徳島", "A", "未接触", ""],
    // 愛媛
    ["Six", "愛媛", "B", "アプローチ中", "連絡待ち"],
    ["Cres.", "愛媛", "B", "アプローチ中", "連絡待ち"],
    ["PARIGOT松山店（パリゴ）", "愛媛", "B", "未接触", ""],
    ["Bless of Bless", "愛媛", "B", "未接触", ""],
    ["IRIE", "愛媛", "B", "未接触", ""],
    ["THINGS", "愛媛", "B", "未接触", ""],
    ["&PAVONE", "愛媛", "B", "未接触", ""],
    ["ドレム", "愛媛", "B", "未接触", ""],
    // 高知
    ["mio jeje (ミーオジェジェ)", "高知", "B", "アプローチ中", "返信待ち"],
    ["PuKu (プク)", "高知", "B", "アプローチ中", "返信待ち"],
    ["Brain Christy (クリスティ)", "高知", "B", "アプローチ中", ""],
    ["LIVING（リビング）", "高知", "B", "アプローチ中", "返信待ち"],
    ["ISEYA（イセヤ）", "高知", "A", "未接触", ""],
    ["chambre de charme（シャンブルドゥシャーム）高知店", "高知", "B", "未接触", ""],
    ["アンシャンテ Enchante", "高知", "A", "未接触", ""],
  ];

  const now = new Date();
  let added = 0;
  newStores.forEach(function (s) {
    if (existingNames[s[0]]) return; // 既に追加済みならスキップ
    sheet.appendRow([Utilities.getUuid(), s[0], s[1], s[2], s[3], s[4], "", now]);
    added++;
  });

  Logger.log(added + "件を追加しました（" + (newStores.length - added) + "件は既存のためスキップ）");
}

/**
 * 香川・岡山・広島・徳島・愛媛・高知の候補店舗48件分の詳細情報
 * （住所・連絡先・取り扱いブランドなど）を「候補店舗リスト（詳細）」タブにまとめる。
 * 優先度・ステータス・次のアクション・次のアクション期限の4列は、
 * 「アプリ連携」シートを参照する数式（VLOOKUP）にしてあるので、
 * アプリ側で更新するたびに自動的に反映される。
 * 「アプリ連携」タブ自体には一切触れない。
 * 元の調査タブ（香川／岡山／広島／徳島／愛媛／高知など）もそのまま残る。
 * 整理が済んだら手動で削除してよい。
 * 1回だけ手動実行する。すでに「候補店舗リスト（詳細）」タブがあれば中身を作り直す。
 */
function consolidateReferenceTabs() {
  const ss = getSpreadsheet_();
  const REF_SHEET_NAME = "候補店舗リスト（詳細）";
  let sheet = ss.getSheetByName(REF_SHEET_NAME);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(REF_SHEET_NAME);
  }

  const headers = ["店名", "地域", "優先度", "ステータス", "次のアクション", "次のアクション期限", "業態", "住所", "電話番号", "公式HP", "インスタグラム", "オンラインストア", "取り扱いブランド", "企業名", "営業時間", "定休日", "備考"];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);

  // [店名, 地域, 業態, 住所, 電話番号, 公式HP, インスタグラム, オンラインストア, 取り扱いブランド, 企業名, 営業時間, 定休日, 備考]
  const detailRows = [
    // 香川
    ["Tsuneya", "香川", "コンテンポラリーセレクト", "高松市丸亀町6-3", "087-821-8529", "", "Tsuneya (@_tsuneya_foyer)", "Foyer Online Store", "VALENTINO・DRIES VAN NOTEN・STELLA McCARTNEY・JIL SANDER・Maison Margiela・MARNI・BRUNELLO CUCINELLI・FABIANA FILIPPI・J&M DEVIDSON・JOHN SMEDLEY・MONCLER・HERNO・HYKE・M-fil・YVES SALOMON", "株式会社つねや／株式会社フォワイエ", "", "", ""],
    ["BLISS", "香川", "レディースセレクト", "高松市丸亀町6-4", "087-811-2382", "https://foyeronline.com/contact.html", "https://www.instagram.com/foyer_bliss/", "Foyer Online Store", "ADIEU TRISTESSE・ADIEU TRISTESSE LOISIR・conges payes・FRAPBOIS・KEI Hayama PLUS・H.A.K・OSKLEN・BEATING HEART・i BLUES・Lion GALLERY・WEEKEND Max Mara", "株式会社つねや／株式会社フォワイエ", "10:30～19:00", "毎週月曜日（祝日の場合は翌日火曜日）", ""],
    ["OUI FOYER", "香川", "セレクト", "高松市瓦町1丁目10-3 FoyerⅡビル1・2F", "087-837-3181", "https://www.foyer.co.jp/shop/oui.html", "https://www.instagram.com/oui_foyer/", "", "A.P.C.・assiette・CABINET・chigo・dansko・de bon coer・ELEY KISHIMOTO・HUIS.・Johnstons・Manna・MASTER&Co.・MNNG・mononogu・Pantherella・PRISTINE・SHRIKE・SLEEPY JONES・THOUSAND TALES・VERSEAU・くらしきぬ・大法紡績", "株式会社つねや／株式会社フォワイエ", "10:30～19:00", "毎週月曜日（祝日の場合は翌日火曜日）、毎月第二・第四火曜日", ""],
    ["REMIX store", "香川", "個人セレクト", "高松市常磐町2-8-1 corridorSF", "087-835-0533", "https://www.remixstore.jp/", "https://www.instagram.com/remixstore.jp/", "https://remixstore.official.ec/", "Chika Kisada・IIROT・BOWTE・SEA・Sea New York・WALANCE・MARILYN MOON・EUCHRONIA・A VACATION・OVERNEATH・BONEE・PERVERZE・KEYHOLDER・palomawool・Opera SPORT・REKISAMI・TORRAZZO DONNA・PRANK PROJECT・ANTHOM・JUUKIFF・PELLICO・Sapir Bachar・La Maison de Lyllis・SERGE de bleu・Uhr・determ;・LEVI'S VINTAGE CLOTHING・Levi's・MARROW・Sister Jane・Carne Bollente・GOOD SQUISH・JUSTINE CLENQUET・NINFA・Sisi Joia・Waa.・miffew・Munich・Marcomonde・BAGGU・LOVE MY NEIGHBORHOOD®", "株式会社THE REMIX/THE REMIX Co.,Ltd.", "11:00-19:00　日：11:00-18:00", "定休日：水、不定", ""],
    ["ジーンズファクトリー高松店", "香川", "セレクトショップ", "高松市伏石町２１４９−１２", "087-867-0222", "https://www.jeansfactory.jp/f/form", "https://www.instagram.com/jeansfactory_takamatsu/", "https://www.jeansfactory.jp/", "", "", "", "", ""],
    ["RADICA douce harmonie", "香川", "セレクトショップ", "高松市丸亀町7-16 丸亀町グリーン西館1F", "087-873-2327", "https://www.ndcjapan.jp/", "https://www.instagram.com/radica_takamatsu/", "", "", "株式会社エヌディシージャパン/NDC Japan co., ltd.", "11:00-20:00", "", ""],
    ["STEF", "香川", "セレクトショップ", "高松市瓦町2-12-2-3階", "087-884-3750", "", "https://www.instagram.com/stef_2025_/", "", "TELA・KURO・TODAYFUL・BANANA TIME・WHITE SAND・GREEN BUTTER・RELEVE'", "", "12:00〜19:00", "水曜、第1・3火曜", "2025年にオープンした大人の女性向けクリーンカジュアルセレクト。新興店舗として機動性が高く、ブランドの初期パートナーとして開拓ポテンシャルが高い。"],
    ["CLOAK（クローク）", "香川", "レディースセレクトショップ", "香川県高松市丸亀町4-12 B1F", "087-813-4620", "https://cloak-room.jp/", "https://www.instagram.com/cloak_2014/", "", "YURI PARK・JOHNSTONS・TELA・REMI RELIEF・CINOH・BLACK CRANE等", "", "", "", "丸亀町商店街の好立地。TELA・CINOH等大人コンテンポラリーを取扱いLE CIEL BLEUとテイスト相性良し。既存取引先の集中する高松丸亀町エリアの新規開拓先として有力。"],
    ["AMBIENT SHARE", "香川", "セレクトショップ（レディース）", "高松市多肥下町1528-1", "087-868-7466", "", "https://www.instagram.com/ambient_share/", "", "", "", "", "", "体型に合わせたスタイリング提案に定評あり。口コミ評価が高い"],
    ["MANU", "香川", "セレクトショップ", "高松市多肥下町1549-4", "087-816-7223", "", "", "", "", "", "", "", "高品質・厳選アイテム中心。丁寧なスタイリング提案に定評。公式Instagramは検索で特定できず未確認"],
    ["red earth", "香川", "セレクトショップ（レディース）", "高松町2487-11", "087-804-0922", "", "https://www.instagram.com/red_earth_takamatsu/", "", "", "", "", "", "シンプル系デザイン中心のレディースショップ"],
    ["Lal", "香川", "セレクトショップ（レディース）", "高松市兵庫町1-23", "087-813-0105", "", "", "", "", "", "", "", "レディース向けカジュアルセレクト。近隣に「ラル(Lar)」という店舗情報があるが住所が完全一致せず同一店舗か未確認。公式Instagramも特定できず未確認"],
    ["GARDEN", "香川", "セレクトショップ", "高松市川原町2-12-2", "087-833-0622", "", "https://www.instagram.com/garden_dsk/", "", "", "", "", "", "都会的なスタイル提案に定評のあるセレクトショップ"],
    // 岡山
    ["インターナショナルリレーションアール＆ゴー", "岡山", "セレクトショップ", "2-5 Togiya Kita-ku Okaya Japan　700-0826", "", "https://international-relation.jp/", "https://www.instagram.com/relation_honten", "https://shop.international-relation.jp/", "", "", "", "", ""],
    ["Due collection（デュエコレクション）", "岡山", "", "2001-3-25", "086-231-3448", "https://duecollection.stores.jp/", "https://www.instagram.com/duecollection/", "https://duecollection.stores.jp/", "Jil Sander・Sacai・JW Anderson等 インポート×国内ミックス", "", "10:30〜19:30", "火曜日（祝日は営業）", "Jil Sander・Sacai・JW Andersonなどモード系インポートに強い実力店。ドメスティックコンテンポラリーの新規導入にも意欲的な可能性があり要ヒアリング。"],
    ["Lucca（ルッカ）", "岡山", "セレクトショップ（レディース専門・ハイクラス）", "岡山県岡山市北区表町1丁目4-35", "086-234-4888", "https://lucca.beauty/", "https://www.instagram.com/luccainsta/", "https://luccaokayama.thebase.in/", "PLAN C等、国内外のハイクラスブランドを取扱うレディース専門店", "", "11:00〜19:00", "不定休", "表町エリアのレディース専門ハイクラス店。世界観のあるブランドセレクトでLE CIEL BLEUの世界観とマッチしやすい"],
    ["proof（プルーフ）", "岡山", "セレクトショップ（レディース専門）", "岡山県岡山市北区東島田町1丁目2-9 1F", "", "https://www.shop-proof.jp/", "https://www.instagram.com/shop_proof/", "", "suzuki takayuki・TOUJOURS・YAECA・GASA*等、大人女性向けブランドを取扱うレディース専門店", "", "", "", "電話番号不明のためInstagram DMでのアプローチが必要。大人女性向けブランドに強く、LE CIEL BLEUの顧客層と重なる可能性高い"],
    ["PARIGOT OKAYAMA（パリゴ 岡山一番街店）", "岡山", "セレクトショップ（レディース・ハイブランドMIX）", "岡山県岡山市北区駅元町1-1-1 岡山一番街地下2号", "086-801-4188", "https://www.parigot.co.jp/shop/okayama", "https://www.instagram.com/parigot_okayama_women/", "", "", "", "平日11:00-20:00／土日祝10:00-20:00", "無休（年末年始除く）", "Chloe、STELLA McCARTNEY、Mame Kurogouchi、CLANEなどハイブランド〜ドメスティックまで幅広く展開する大人フェミニン路線。LE CIEL BLEUと客層の相性が良さそう"],
    ["6DIRECTIONS/efu2", "岡山", "セレクトショップ（ナチュラル系・レディース有）", "岡山県岡山市北区表町3丁目18-63", "086-223-0156", "https://6directions.net/", "https://www.instagram.com/6directions.efu2/", "", "", "", "", "", "EEL Products、HARVESTY、TRAVAIL MANUELなどナチュラル系ブランドが中心。フェミニン度はやや控えめなので要検討"],
    // 広島
    ["VINCENT & MIA（ヴィンセント・アンド・ミア）", "広島", "ドメスティック中心のコンテンポラリー/モード系セレクト（メンズ・レディース）", "〒730-0037 広島県広島市中区中町5-1 長沼第二ビル1F", "082-247-6023", "https://vincent-mia.com", "https://www.instagram.com/vincent_and_mia/", "https://vincent-mia.shop/", "AURALEE・BATONER・BLAMINK・BOWTE・CIOTA・HERILL・JOHN SMEDLEY・MARNI・Scye・THE RERACS・LOEFF・POSTELEGANT・YAECA・YLEVE", "", "12:00-20:00", "水曜", "2004年開店・20周年の老舗セレクト。AURALEE・BLAMINK・YAECA・THE RERACSなど上質なドメスティックコンテンポラリーに強く、LE CIEL BLEUの客層と重なる可能性が高い要商談店。"],
    ["GALAXY（ギャラクシー）", "広島", "ラグジュアリー〜コンテンポラリー/モードのメンズ・レディースセレクト", "〒730-0036 広島県広島市中区袋町2-23", "082-247-1310", "https://latineve.com", "https://www.instagram.com/galaxy_hiroshima/", "https://latineve.com/pages/13", "08sircus・ENFOLD・MARNI・Maison Margiela・BALENCIAGA・Chloé・CECILIE BAHNSEN・JACQUEMUS・MEIMEIJ・Max Mara・VALENTINO・Christian Wijnants・Neil Barrett・PIERRE HARDY・DSQUARED2・Y-3", "株式会社ラタン・イヴ（LATIN EVE）", "11:00-20:00", "水曜", "広島有数のハイブランド・モード系実力店。ENFOLDを毎シーズン継続入荷しており、LE CIEL BLEUの世界観と親和性が高い。単価帯は高めだが要ヒアリング。"],
    ["VIA LUISA（ビア・ルイーザ）", "広島", "ウィメンズ専門コンテンポラリー/モードセレクト", "〒737-0046 広島県呉市中通3-3-29", "0823-25-8620", "https://latineve.com", "", "https://latineve.com/pages/15", "ENFOLD・MARNI・Chloé・MEIMEIJ・CECILIE BAHNSEN・Christian Wijnants（ウィメンズ中心）", "株式会社ラタン・イヴ（LATIN EVE）", "11:00-20:00", "火曜", "GALAXYの姉妹ウィメンズ専門店。所在地は広島市外（呉市）だが同一バイヤー窓口のため、GALAXYと併せて一括アプローチ可能。"],
    ["PARIGOT 広島店（パリゴ）", "広島", "アップスケールなコンテンポラリー/モード系セレクト（正規取扱・400ブランド以上）", "〒730-0034 広島県広島市中区新天地4-2 アクセ広島 2F/3F", "082-504-8411", "https://www.parigot.jp/f/hiroshima", "", "", "08sircus・ENFOLD・CLANE・kolor・HYKE・Mame Kurogouchi・TOGA・beautiful people・MM6 Maison Margiela・Lemaire・CFCL・Charlotte Chesnais・TATRAS・Chloé（国内外400ブランド以上）", "株式会社アクセ", "11:00-20:00", "無休（年末年始除く）", "全国7拠点程度を展開するアップスケールなセレクト。08sircus・ENFOLD・CLANEなどLE CIEL BLEUと直接比較可能なブランドを既に扱い好相性。取扱は店舗・シーズンで変動のため商談前に要確認。"],
    ["CREER（クレエ）", "広島", "ナチュラル〜モードのリラックス系コンテンポラリーセレクト", "〒730-0051 広島県広島市中区大手町2-5-11 ハルゼングランデリブ5F", "082-236-9878", "https://www.e-cloth.jp/", "https://www.instagram.com/creer/", "", "R&D.M.Co-・TOUJOURS・ALDIN・tao・noir kei ninomiya・PLAY COMME des GARÇONS・suzuki takayuki・Veritecoeur・chimala・FRANK LEDER", "", "店舗営業13:00-17:00（不定営業）", "不定（NEWSで告知）", "TOUJOURS・suzuki takayukiなど大人女性向けのモード〜ナチュラルに強く、LE CIEL BLEUのリラックス/クリーン寄りラインと部分的に親和性あり。テイストはやや作家性寄り、営業日限定的な点に留意。"],
    ["LUCY LUE（ルーシールー）", "広島", "レディース専門セレクトショップ", "〒730-0036 広島県広島市中区袋町2-10", "082-541-3122", "", "https://www.instagram.com/lucy_lue_hiroshima/", "https://lucylue.theshop.jp/", "THOMAS MAGPIE・HERENCHIA・anana・pomtata・Squady・YANUK・Dignite collier・manon・CLOCHE・Audrey and John wad", "グッズカンパニー", "10:45-19:30", "定休なし", "12年目を迎える働く大人の女性向けレディース専門店。価格・テイストはLE CIEL BLEUよりやや手頃・カジュアル寄りだが客層は一部重複。導入余地を要ヒアリング。"],
    ["Zalife（ザライフ）", "広島", "ドメスティック〜モードのユニセックスセレクト", "広島県広島市中区袋町", "", "https://www.zalife.jp/", "", "", "OUTIL・Maison Mihara Yasuhiro・DESCENTE ALLTERRAIN・DESCENTE PAUSE・810s・moonstar", "", "", "", "ややメンズ/機能系（アウトドア・ワーク）寄りのため、レディースコンテンポラリーとしての親和性は上位店より低い。レディース比率・客層を要ヒアリング。"],
    ["PB CLOSET", "広島", "Instagramライブ販売中心のカジュアルショップ", "広島県廿日市市（Instagramライブ／オンライン中心）", "", "", "", "", "", "", "", "", "「広島 大人セレクトショップ」として言及されるが、実体は廿日市市拠点・Instagramライブ／オンライン中心・低価格帯（ワンピース¥7,000〜1万円台）で、デザイナーズ・コンテンポラリーの取扱なし。対象外と判定。"],
    // 徳島
    ["THE DAY", "徳島", "", "徳島県徳島市山城西３丁目４１−１", "088-660-6565", "https://theday.shop/", "https://www.instagram.com/theday_jp/", "https://the-day.stores.jp/", "A.P.C.・A PUPIL・anana・AVIE・Cape Hights・CLANE・COOHEM・ELENDEEK・EN'DAY・EZUMi・Fire Service・FLUMOR・HERENCIA・HERIN.CYE・IN-PROCESS Tokyo・JOHN SMEDLEY・KELEN・Lallia Mu・le chanter・MACKAGE・MARGARET HOWELL・MARGAUX・MASION MAVERICK PRESENTS・MHL.・MidiUmi・MIDIUMISOILID・mizuiro ind・NAKAGAMI・nicholson&nichoison・RIM.ARK・Risley・Rita・STAMP & DAILRY・SUGAR ROSE・TRADITIONAL WEATHERWEAR・UN3D・UNIVERSAL TISSU・utilite・WHYTO.・YLEVE", "SCALE Co.", "11:00~19:00", "火曜日", "CLANEやYLEVEなどLE CIEL BLEUと非常に近いテイストを高い完成度で展開する徳島最重要アカウント。トータル提案力が抜群で相性完璧。"],
    ["roly poly", "徳島", "", "徳島県徳島市八百屋町２丁目１９", "088-626-8989", "", "https://www.instagram.com/rolypoly.r/?hl=ja", "https://saoriiroas002.stores.jp/", "", "", "11:00~20:00", "水曜日", "30〜50代の大人の女性に絶大な支持を誇る。TODAYFULやAMERI等のアプローチ実績が厚く、構築的モードの導入に極めて前向き。"],
    ["Hiraoka", "徳島", "", "徳島県徳島市東新町１丁目１４", "088-652-3510", "", "https://www.instagram.com/hiraoka_shouten/", "", "", "", "11:00~19:00", "火曜日", "徳島を代表する最高峰のデザイナーズブティック。Mame Kurogouchiの正規取扱店であり、本物志向の顧客を獲得できる非常に貴重なチャネル。"],
    ["Reply（リプリー）", "徳島", "セレクトショップ", "徳島県徳島市南内町3-7 高木ビル1F", "088-623-3370", "", "", "", "トレンド・ブランドに捕らわれないセレクト（多ブランド混合）", "", "12:00〜20:00", "12/31〜1/2を除く", "10年以上の老舗。トレンド・ブランドに捕らわれないセレクト。公式HPなし、Facebook中心"],
    ["Cherie（クレメントプラザ）", "徳島", "セレクトショップ（レディース・フェミニン）", "徳島県徳島市寺島本町西1-61 徳島駅クレメントプラザ2F", "088-626-8668", "https://www.clementplaza.com/cherie/", "", "", "", "", "10:00-19:00", "", "FRAY I.D、SNIDEL、CELFORD、gelato pique、LAISSE PASSE、Debut de FioreなどLE CIEL BLEUと極めて近いテイストのフェミニンブランドを多数展開。徳島駅直結でアクセスも良好。公式Instagramアカウントは未確認のため要調査"],
    ["MAISON ÉCRU（メゾンエクリュ）", "徳島", "セレクトショップ（レディース）", "徳島県徳島市山城町東浜傍示5-367", "088-678-7157", "https://www.ecru905.com/", "@ecru905", "", "", "", "11:00-19:00", "水曜日", "「友だちの家に遊びに来たような」寛げる空間で、流行に左右されない上質なレディースアイテムを厳選する独立系ブティック。LE CIEL BLEUの世界観と親和性高い"],
    // 愛媛
    ["Six", "愛媛", "セレクトショップ", "愛媛県松山市千舟町4-1-3 1F", "089-913-5525", "https://six06.jp/", "https://www.instagram.com/six06jp/", "https://shop.six06.jp/", "Maison Margiela・ANN DEMEULEMEESTER・HYKE・TOGA・QUIITO・Masnou design・gicipi", "", "12:00~20:00", "不定休", "愛媛エリア随一のハイエンドモードセレクト。HYKEやMargielaなどの抜群の実績を持ち、パリのコレクションにも出張する本格派セレクトショップ。LE CIEL BLEUの最高峰ターゲット。"],
    ["Cres.", "愛媛", "セレクトショップ", "愛媛県今治市朝倉上甲2427−2", "0898-77-6106", "", "https://www.instagram.com/cres_imabari/?hl=ja", "", "Kics Document.・YASHIKI・meanswhile・texnh・REVERBERATE・RhodolirioN・todayful", "", "11:00~19:00", "水曜 / 第２火曜", "北高下町から、2025年4月に今治市朝倉の緑豊かな土地へ移転オープン。「わざわざ行きたい店」として松山などの遠方からも顧客が集まる超優良店。"],
    ["PARIGOT松山店（パリゴ）", "愛媛", "セレクトショップ（チェーン系複合ブランド）", "愛媛県松山市大街道２丁目5番12号 AEL MATSUYAMA 1階", "089-935-8488", "https://www.parigot.jp/f/matsuyama", "", "https://www.parigot.jp/", "パリゴオリジナル、多ブランドセレクト（レディース中心）", "", "10:00〜19:00", "不定休（施設準拠）", "広島尾道店と同じ株式会社アクセ運営のチェーン店。松山中心商業施設内で集客力高い。複数店舗展開提案の可能性あり"],
    ["Bless of Bless", "愛媛", "セレクトショップ", "松山市大街道2-1-12", "089-933-6318", "https://www.bless-life.com/", "https://www.instagram.com/blessofbless.life/", "", "", "", "", "", "ユニーク＆女性らしいアイテム中心。フェミニン色が強く相性良さそう"],
    ["IRIE", "愛媛", "セレクトショップ", "松山市湊町3-7-11 朝日ビル2F", "089-907-0234", "", "https://www.instagram.com/iriemyc/", "", "", "", "", "", "体型を覚えて提案してくれると評判、高評価。インスタは未確定のため裏取り必要"],
    ["THINGS", "愛媛", "セレクトショップ", "松山市湊町4-14-3 渡部ビル1F東", "089-993-6537", "", "https://www.instagram.com/things_ehime/", "", "", "", "", "", "老舗セレクトのオーナーが独立して開いた店"],
    ["&PAVONE", "愛媛", "セレクトショップ（雑貨・アクセサリー中心）", "松山市大街道2-3-3", "089-921-8045", "https://andpavone.net/", "https://www.instagram.com/andpavone/", "", "", "", "", "", "バッグ・スカーフ・アクセサリー中心、雑貨寄りだが上質"],
    ["ドレム", "愛媛", "セレクトショップ", "松山市湊町3-5-12", "089-997-7404", "", "", "", "", "", "", "", "オーナーの目利きに定評、老舗の風格"],
    // 高知
    ["mio jeje (ミーオジェジェ)", "高知", "セレクトショップ", "高知市高そね18-10", "088-883-0821", "https://ameblo.jp/miojeje/", "https://www.instagram.com/miojeje_/", "https://miojeje.thebase.in/", "HERNO・MaxMara・ビーティングハート・セオリー・セオリーリュクス・GEYGRY・ディニテコリエ・インディマーク・ミカーレ", "", "11:00~18:00", "火曜、水曜", "高知県屈指の驚異的な実売力を誇るレディースブティック。バイヤー自身のブログ・動画発信能力が極めて高く、実売を伴う最強のパートナー。"],
    ["PuKu (プク)", "高知", "セレクトショップ", "高知市中万々170-6", "", "", "https://www.instagram.com/puku__kochi/", "https://puku2022.base.shop/", "", "", "11:00~19:00", "水曜", "高知の若年〜ミドル感度層から圧倒的支持を得る店舗。TODAYFUL等の高いアプローチ実績を持ち、LE CIEL BLEUの新規顧客開拓に最適。"],
    ["Brain Christy (クリスティ)", "高知", "セレクトショップ", "高知市本町2丁目1-37", "088-822-0090", "https://www.brain-christy.com/", "https://www.instagram.com/christy_christy2nd/", "", "", "", "11:00〜19:30 水曜日は17:00close", "火曜", "Tomorrowland系列のクリーンなトラッドに強く、仕立ての良さに定評があるセレクト。LE CIEL BLEUのクリーンなアウターや定番ワンピースの親和性大。"],
    ["LIVING（リビング）", "高知", "セレクトショップ（ライフスタイル・雑貨も取扱）", "高知県高知市廿代町14-11 カーザ・デレ・モーレ1F", "088-821-8527", "https://livingvsr.thebase.in/", "https://www.instagram.com/living__jp", "https://livingvsr.thebase.in/", "新品・ヴィンテージミックス、花・アート・フレグランス等も取扱うライフスタイル型セレクト", "", "平日：13:00〜19:00、土日祝：12:00〜19:00", "水曜日", "高知市中心部、ライフスタイル提案型セレクト。アパレ以外も幅広く提案しているため世界観フィットすれば差別化可能"],
    ["ISEYA（イセヤ）", "高知", "セレクトショップ（レディース）", "高知市帯屋町1-14-4", "088-871-0870", "https://www.kochi-iseya.com/", "https://www.instagram.com/iseya.kochi/", "", "Mila Owen・FRAYI.D", "", "11:30-19:30", "", "Mila Owen/FRAYI.D専門店。フェミニン×大人可愛いの実店舗投稿多数で、LE CIEL BLEUと客層・価格帯・テイストの親和性が非常に高い最有力候補"],
    ["chambre de charme（シャンブルドゥシャーム）高知店", "高知", "セレクトショップ（レディース、AMBIDEX系）", "高知市南久保10-39", "088-855-9498", "https://www.ambidex-store.jp/shop/shoplist/fc_kochi.aspx", "https://www.instagram.com/chambre_de_charme4241/", "", "AMBIDEX・PAR ICI・yuni・bulle de savon", "", "11:00-20:00（水曜10:00-17:00）", "", "AMBIDEX系はモード・エレガンス寄りの大人フェミニン。LE CIEL BLEUよりやや上のテイストだが親和性あり"],
    ["アンシャンテ Enchante", "高知", "セレクトショップ（個人経営）", "高知市潮新町2-12-19", "", "https://kochi-enchante.com/", "@enchante_kochi", "", "", "", "", "", "オーナー自身がパーソナルコーディネーター。体型提案も行う"],
  ];

  const rows = detailRows.map(function (r, i) {
    const rowNum = i + 2; // シート上の行番号（見出し行の次から）
    const name = r[0];
    return [
      name,
      r[1],
      '=IFERROR(VLOOKUP($A' + rowNum + ', アプリ連携!$B:$H, 3, FALSE), "")',
      '=IFERROR(VLOOKUP($A' + rowNum + ', アプリ連携!$B:$H, 4, FALSE), "")',
      '=IFERROR(VLOOKUP($A' + rowNum + ', アプリ連携!$B:$H, 5, FALSE), "")',
      '=IFERROR(VLOOKUP($A' + rowNum + ', アプリ連携!$B:$H, 6, FALSE), "")',
      r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12],
    ];
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 「次のアクション期限」の表記がおかしい問題の修正。
 * 原因：日付文字列（例:"2026-09-04"）をシートに書き込むと、
 * スプレッドシートが自動的に日付型として認識し直してしまい、
 * ロケール依存の表示（例:9/4/2026）や「候補店舗リスト（詳細）」側の
 * VLOOKUP結果でシリアル値のようにズレて見えることがある。
 * 対策：「次のアクション期限」列を「書式なしテキスト」に固定し、
 * すでに日付型になってしまっている値は yyyy-MM-dd の文字列に戻す。
 * 1回だけ手動実行すればよい（以後の追加・更新でも自動的にテキストのまま保存される）。
 */
function fixNextActionDateFormat() {
  const ss = getSpreadsheet_();

  // 「アプリ連携」の G列（次のアクション期限）
  const appSheet = ss.getSheetByName(SHEET_NAME);
  if (appSheet) {
    const lastRow = appSheet.getLastRow();
    if (lastRow > 1) {
      const range = appSheet.getRange(2, 7, lastRow - 1, 1);
      const values = range.getValues();
      const fixed = values.map(function (row) {
        const v = row[0];
        if (v instanceof Date) {
          return [Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd")];
        }
        return [v];
      });
      range.setNumberFormat("@"); // 書式なしテキストに固定
      range.setValues(fixed);
    }
  }

  // 「候補店舗リスト（詳細）」の F列（次のアクション期限、VLOOKUP式）
  const refSheet = ss.getSheetByName("候補店舗リスト（詳細）");
  if (refSheet) {
    const lastRow = refSheet.getLastRow();
    if (lastRow > 1) {
      refSheet.getRange(2, 6, lastRow - 1, 1).setNumberFormat("@");
    }
  }

  Logger.log("次のアクション期限の表記を修正しました。");
}

/**
 * 「アプリ連携」に漏れていた OUI FOYER（香川）を追加する。
 * 「候補店舗リスト（詳細）」には情報があるが、「アプリ連携」には
 * 元々登録されていなかったため、香川が12件（本来13件）になっていた。
 * すでに存在する場合は何もしない（二重追加防止）。1回だけ手動実行すればよい。
 */
function addOuiFoyerToApp() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === "OUI FOYER") {
      Logger.log("OUI FOYERは既に登録済みです。");
      return;
    }
  }
  sheet.appendRow([Utilities.getUuid(), "OUI FOYER", "香川", "B", "アプローチ中", "連絡待ち", "", new Date()]);
  Logger.log("OUI FOYERを追加しました。");
}
