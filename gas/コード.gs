/* =========================================================
   IBS体調管理システム - GAS（Google Apps Script）側コード
   =========================================================
   役割:
   - doPost : アプリから送られた記録をシートに追記する
   - doGet  : アプリ起動時に「設定マスタ」と「直近7日の集計」を返す
   - buildDashboard : 分析シート一式を作り直す（日次トリガー/メニューから実行）
   - initialSetup   : シートの初期作成（最初に1回だけ実行）

   ※ 分析はすべて「日付」をキーに結合しているため、将来
     「周期ログ」などのシートを足す場合も同じ方式で拡張できます。
   ========================================================= */

const TZ = "Asia/Tokyo";

// シート名（変えたい場合はここを直す）
const SHEET_EVENT = "イベントログ";
const SHEET_DAILY = "デイリーログ";
const SHEET_SETTINGS = "設定";
const SHEET_STATE = "日別ステート";
const SHEET_WEEKLY = "週次サマリー";
const SHEET_TRIGGER = "トリガー分析";
const SHEET_CALENDAR = "カレンダービュー";
const SHEET_CYCLE = "サイクルビュー";
const SHEET_MENTAL = "メンタル×体調";
const SHEET_RHYTHM = "排便リズム";
const SHEET_RX = "処方";           // 処方の管理（現行/過去）
const SHEET_RXLOG = "処方経過";     // 受診報告用の日別ビュー
const SHEET_PERIOD = "周期ログ";    // 生理周期の開始日
const SHEET_PERIOD_ANALYSIS = "周期×体調"; // 生理周期と症状の分析

// 満腹度の選択肢（アプリ側と合わせる）
const FULLNESS_OPTIONS = ["食べすぎて気持ち悪い", "胃もたれ", "ちょうどよく満腹", "八分目", "足りない"];

const EVENT_HEADERS = ["記録日時", "種別", "発生時刻", "ブリストルスケール", "腹痛", "残便感", "薬名", "タイミング", "メモ"];
const DAILY_HEADERS = ["日付", "睡眠時間", "食事タグ", "食事メモ", "水分量", "運動", "気分", "ストレス", "メンタルメモ", "メモ", "観察タグ",
                       "朝食タグ", "朝食メモ", "昼食タグ", "昼食メモ", "夕食タグ", "夕食メモ", "満腹度",
                       "朝食満腹度", "昼食満腹度", "夕食満腹度"];
const RX_HEADERS = ["薬名", "状態", "分類", "用法", "タイミング", "1回量", "開始日", "日数", "数量", "終了日", "服用メモ", "表示順"];

// 服薬チェックのタイミング表示順
const TIMING_ORDER = ["朝食後", "昼食後", "夕食前", "夕食後", "外用"];

const WEEKDAYS_JP = ["日", "月", "火", "水", "木", "金", "土"];

/* =========================================================
   初期セットアップ（最初に1回だけ実行する）
   ========================================================= */
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);

  // --- イベントログ ---
  let sh = getOrCreateSheet_(SHEET_EVENT);
  if (sh.getLastRow() === 0) {
    sh.appendRow(EVENT_HEADERS);
    sh.setFrozenRows(1);
  }

  // --- デイリーログ ---
  sh = getOrCreateSheet_(SHEET_DAILY);
  if (sh.getLastRow() === 0) {
    sh.appendRow(DAILY_HEADERS);
    sh.setFrozenRows(1);
  }

  // --- 設定（マスタと閾値の一元管理） ---
  sh = getOrCreateSheet_(SHEET_SETTINGS);
  if (sh.getLastRow() === 0) {
    // 薬マスタ（A〜E列。残数・1回量は将来の拡張用で今は未使用）
    sh.getRange(1, 1, 1, 5).setValues([["薬名", "表示順", "有効", "残数", "1回量"]]);
    const meds = [
      ["イリボー", 1, "有効"], ["ミヤBM", 2, "有効"], ["ブスコパン", 3, "有効"],
      ["アトモキセチン", 4, "有効"], ["チラーヂン", 5, "有効"], ["マグミット", 6, "有効"],
    ];
    sh.getRange(2, 1, meds.length, 3).setValues(meds);

    // 食事タグマスタ（G〜I列）
    sh.getRange(1, 7, 1, 3).setValues([["タグ名", "表示順", "有効"]]);
    const tags = ["辛いもの", "乳製品", "アルコール", "カフェイン", "脂っこいもの", "外食", "アイス", "炭酸ジュース", "お菓子"]
      .map((t, i) => [t, i + 1, "有効"]);
    sh.getRange(2, 7, tags.length, 3).setValues(tags);

    // 閾値など（K〜L列。次回通院日は将来のリマインダー用）
    sh.getRange(1, 11, 1, 2).setValues([["項目", "値"]]);
    sh.getRange(2, 11, 5, 2).setValues([
      ["便秘警戒日数", 3],
      ["下痢判定ブリストル", 6],
      ["下痢判定回数", 3],
      ["次回通院日", ""],
      ["合言葉", ""],
    ]);
    sh.setFrozenRows(1);
  }

  // 初期作成される空の「シート1」があれば削除
  const def = ss.getSheetByName("シート1") || ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  SpreadsheetApp.getUi().alert("初期セットアップが完了しました。\n「イベントログ」「デイリーログ」「設定」の3シートが作成されています。");
}

/* =========================================================
   カスタムメニュー（スプレッドシートを開くと自動で追加される）
   ========================================================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("IBS分析")
    .addItem("ダッシュボードを今すぐ更新", "buildDashboard")
    .addItem("日次トリガーを設定（毎朝5時に自動更新）", "setupDailyTrigger")
    .addItem("重複記録をチェックして削除", "cleanDuplicateEvents")
    .addSeparator()
    .addItem("処方アップデート(2026-07-31)を適用", "applyRxUpdate20260731")
    .addItem("アップデート(2026-08-02)を適用", "applyUpdate20260802")
    .addItem("アップデート(2026-08-10)を適用", "applyUpdate20260810")
    .addItem("処方アップデート(2026-09-02)を適用", "applyRxUpdate20260902")
    .addItem("初期セットアップ（最初に1回）", "initialSetup")
    .addToUi();
}

/* =========================================================
   処方アップデート(2026-09-02)の適用
   9/2の消化器内科受診で出た新処方を「処方」シートに反映する。
   - ポリフル: ジェネリック（ポリカルボフィルCa細粒0.6g）に変更、28日分
   - マグミット: 朝昼夕 各1錠(330mg)で再開、28日分
   - ヘモクロン: 継続、14日分
   - ボラザG軟膏: 継続、全量48g
   - グーフィス: 処方から外れたため「過去」へ
   - ミヤBM: 変更なし（継続）
   何度実行しても安全。ウェブアプリの再デプロイは不要（シートを書き換えるだけ）
   ========================================================= */
function applyRxUpdate20260902() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rx = ss.getSheetByName(SHEET_RX);
  if (!rx) {
    SpreadsheetApp.getUi().alert("「処方」シートがありません。先に「処方アップデート(2026-07-31)を適用」を実行してください。");
    return;
  }
  const values = rx.getDataRange().getValues();
  const head = values[0];
  const col = {};
  RX_HEADERS.forEach(function (h) { col[h] = head.indexOf(h); });
  const setCell = function (r, h, v) { if (col[h] >= 0) rx.getRange(r + 1, col[h] + 1).setValue(v); };
  const letter = function (h) { return String.fromCharCode(65 + col[h]); };
  const start = new Date(2026, 8, 2); // 2026-09-02
  const done = [];

  // 薬名ごとに書き換える内容
  const updates = {
    "ポリフル": {
      "状態": "現行", "分類": "IBS治療薬（便の水分調整）",
      "用法": "1日3回 朝昼夕食後 細粒0.6g（1日1.8g）", "タイミング": "朝食後,昼食後,夕食後",
      "1回量": "0.6g", "開始日": start, "日数": 28,
      "服用メモ": "9/2からジェネリック（ポリカルボフィルCa細粒）に変更。中身は同じ", "表示順": 1,
    },
    "マグミット（酸化マグネシウム）": {
      "状態": "現行", "分類": "便秘治療薬（酸化マグネシウム）",
      "用法": "1日3回 朝昼夕食後 各1錠(330mg)", "タイミング": "朝食後,昼食後,夕食後",
      "1回量": "1錠", "開始日": start, "日数": 28, "数量": "",
      "服用メモ": "2026-09-02の処方で再開", "表示順": 2,
    },
    "ヘモクロン": {
      "状態": "現行", "開始日": start, "日数": 14,
      "服用メモ": "他の薬より早く（14日で）終わる",
    },
    "ボラザG軟膏": {
      "状態": "現行", "開始日": start, "数量": "48g（2.4g×20本）",
    },
    "グーフィス": {
      "状態": "過去", "服用メモ": "2026-09-02の処方変更で終了（7/31〜）",
    },
  };

  const found = {};
  for (let r = 1; r < values.length; r++) {
    const name = String(values[r][col["薬名"]] || "").trim();
    if (!(name in updates)) continue;
    found[name] = true;
    const u = updates[name];
    Object.keys(u).forEach(function (h) { setCell(r, h, u[h]); });
    // 終了日の数式（開始日＋日数−1）と書式を整える
    if (col["終了日"] >= 0 && col["開始日"] >= 0 && col["日数"] >= 0) {
      const g = letter("開始日") + (r + 1), hh = letter("日数") + (r + 1);
      rx.getRange(r + 1, col["終了日"] + 1).setFormula('=IF(AND(' + g + '<>"",' + hh + '<>""),' + g + '+' + hh + '-1,"")')
        .setNumberFormat("yyyy-mm-dd");
      rx.getRange(r + 1, col["開始日"] + 1).setNumberFormat("yyyy-mm-dd");
      rx.getRange(r + 1, col["日数"] + 1).setNumberFormat("0");
    }
    done.push(name + " → " + (u["状態"] || "更新"));
  }

  // 行がなかった薬は追加する（マグミットの行名が違う場合など）
  Object.keys(updates).forEach(function (name) {
    if (found[name] || updates[name]["状態"] === "過去") return;
    const u = updates[name];
    const row = head.map(function (h) { return h === "薬名" ? name : (h in u ? u[h] : ""); });
    rx.appendRow(row);
    const r = rx.getLastRow();
    if (col["終了日"] >= 0) {
      const g = letter("開始日") + r, hh = letter("日数") + r;
      rx.getRange(r, col["終了日"] + 1).setFormula('=IF(AND(' + g + '<>"",' + hh + '<>""),' + g + '+' + hh + '-1,"")').setNumberFormat("yyyy-mm-dd");
      rx.getRange(r, col["開始日"] + 1).setNumberFormat("yyyy-mm-dd");
      rx.getRange(r, col["日数"] + 1).setNumberFormat("0");
    }
    done.push(name + " → 行を追加（現行）");
  });

  SpreadsheetApp.getUi().alert(
    "2026-09-02の処方を反映しました。\n\n・" + done.join("\n・") +
    "\n\nアプリを開き直すと服薬チェックが新しい処方になります。\n「ダッシュボードを今すぐ更新」で処方経過シートも新処方基準になります。"
  );
}

/* =========================================================
   アップデート(2026-08-10)の適用
   - デイリーログに「満腹度」列を追加
   - イベントログに「クライアントID」列を追加（二重送信の防止用）
   - 「周期ログ」シートを作成（生理周期。初回データとして2026-08-05を登録）
   - 処方シートの「日数」列の書式不具合を修復（1970-01-01バグの根治）
   何度実行しても安全
   ========================================================= */
function applyUpdate20260810() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const done = [];

  // 1) 列の追加
  ["満腹度", "朝食満腹度", "昼食満腹度", "夕食満腹度"].forEach(function (h) {
    if (ensureHeader_(mustSheet_(SHEET_DAILY), h)) done.push("デイリーログに「" + h + "」列を追加");
  });
  if (ensureHeader_(mustSheet_(SHEET_EVENT), "クライアントID")) done.push("イベントログに「クライアントID」列を追加（二重送信防止用）");

  // 2) 周期ログシート（生理周期の開始日を記録する）
  if (!ss.getSheetByName(SHEET_PERIOD)) {
    const sh = ss.insertSheet(SHEET_PERIOD);
    sh.getRange(1, 1, 1, 2).setValues([["開始日", "メモ"]]).setFontWeight("bold");
    sh.setFrozenRows(1);
    // デイリーログのメモ「生理開始」(2026-08-05)から初回データを登録
    sh.getRange(2, 1, 1, 2).setValues([[new Date(2026, 7, 5), "デイリーログのメモから自動登録"]]);
    sh.getRange(2, 1, 100, 1).setNumberFormat("yyyy-mm-dd");
    done.push("「周期ログ」シートを作成（8/5の「生理開始」メモを初回データとして登録）");
  }

  // 3) 処方シートの「日数」列の修復
  //    誤って日付書式になっていたセル（28→1900-01-27）を数値に戻す
  const rx = ss.getSheetByName(SHEET_RX);
  if (rx) {
    const values = rx.getDataRange().getValues();
    const head = values[0];
    const dCol = head.indexOf("日数");
    if (dCol >= 0) {
      let fixed = 0;
      for (let r = 1; r < values.length; r++) {
        const v = values[r][dCol];
        if (v instanceof Date) {
          rx.getRange(r + 1, dCol + 1).setValue(readDayCount_(v));
          fixed++;
        }
      }
      if (values.length > 1) rx.getRange(2, dCol + 1, values.length - 1, 1).setNumberFormat("0");
      if (fixed) done.push("処方シートの「日数」列を修復（" + fixed + "件を日付→数値に戻しました）");
    }
  }

  SpreadsheetApp.getUi().alert(
    done.length
      ? "アップデートを適用しました。\n\n・" + done.join("\n・") +
        "\n\nこのあと「重複記録をチェックして削除」と「ダッシュボードを今すぐ更新」も実行してください。"
      : "すでに適用済みです（変更はありません）。"
  );
}

/* =========================================================
   重複記録のクリーニング
   「種別・薬名・タイミング・ブリストル・発生時刻（分単位）」が
   すべて同じ行を重複とみなし、最初の1件だけ残して削除する
   ========================================================= */
function cleanDuplicateEvents() {
  const sh = mustSheet_(SHEET_EVENT);
  const values = sh.getDataRange().getValues();
  if (values.length < 3) {
    SpreadsheetApp.getUi().alert("記録が少ないためチェック対象がありません。");
    return;
  }
  const head = values[0];
  const iKind = head.indexOf("種別"), iAt = head.indexOf("発生時刻"),
        iBri = head.indexOf("ブリストルスケール"), iMed = head.indexOf("薬名"),
        iTim = head.indexOf("タイミング");

  const seen = {};
  const toDelete = []; // {row, label}
  for (let r = 1; r < values.length; r++) {
    const at = values[r][iAt];
    if (!(at instanceof Date)) continue;
    const key = [
      values[r][iKind], values[r][iMed] || "", iTim >= 0 ? values[r][iTim] || "" : "",
      values[r][iBri] || "", Utilities.formatDate(at, TZ, "yyyy-MM-dd HH:mm"),
    ].join("|");
    if (seen[key]) {
      toDelete.push({
        row: r + 1,
        label: Utilities.formatDate(at, TZ, "M/d HH:mm") + " " + values[r][iKind] + " " +
               (values[r][iMed] || (values[r][iBri] ? "ブリストル" + values[r][iBri] : "")),
      });
    } else {
      seen[key] = true;
    }
  }

  const ui = SpreadsheetApp.getUi();
  if (!toDelete.length) {
    ui.alert("重複記録は見つかりませんでした。");
    return;
  }
  const list = toDelete.map(function (d) { return "・" + d.label; }).join("\n");
  const res = ui.alert(
    "重複記録の削除",
    toDelete.length + "件の重複が見つかりました。最初の1件を残して削除します。\n\n" + list + "\n\n削除してよいですか？",
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;
  // 下の行から消す（行番号がずれないように）
  toDelete.reverse().forEach(function (d) { sh.deleteRow(d.row); });
  ui.alert(toDelete.length + "件を削除しました。\n「ダッシュボードを今すぐ更新」を実行すると分析に反映されます。");
}

/* =========================================================
   アップデート(2026-08-02)の適用
   食事の朝昼夕分割に対応する列をデイリーログに追加する。
   何度実行しても安全
   ========================================================= */
function applyUpdate20260802() {
  const done = [];
  const dl = mustSheet_(SHEET_DAILY);
  ["朝食タグ", "朝食メモ", "昼食タグ", "昼食メモ", "夕食タグ", "夕食メモ"].forEach(function (h) {
    if (ensureHeader_(dl, h)) done.push("デイリーログに「" + h + "」列を追加");
  });
  SpreadsheetApp.getUi().alert(
    done.length ? "アップデートを適用しました。\n\n・" + done.join("\n・") : "すでに適用済みです（変更はありません）。"
  );
}

/* =========================================================
   処方アップデート(2026-07-31)の適用
   何度実行しても安全（すでに適用済みの部分はスキップされる）
   ========================================================= */
function applyRxUpdate20260731() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const done = [];

  // 1) 既存シートに新しい列を追加（なければ末尾に足す）
  if (ensureHeader_(mustSheet_(SHEET_EVENT), "タイミング")) done.push("イベントログに「タイミング」列を追加");
  if (ensureHeader_(mustSheet_(SHEET_DAILY), "観察タグ")) done.push("デイリーログに「観察タグ」列を追加");

  // 2) 処方シートの作成（現行4剤＋旧処方を履歴として登録）
  if (!ss.getSheetByName(SHEET_RX)) {
    const rx = ss.insertSheet(SHEET_RX);
    rx.getRange(1, 1, 1, RX_HEADERS.length).setValues([RX_HEADERS]).setFontWeight("bold");
    rx.setFrozenRows(1);
    const start = new Date(2026, 6, 31); // 2026-07-31
    const rows = [
      ["ポリフル", "現行", "IBS治療薬（便の水分調整）", "1日3回 朝昼夕食後 各1錠(500mg)", "朝食後,昼食後,夕食後", "1錠", start, 28, "", "", "コップ1杯以上の十分な水で服用", 1],
      ["グーフィス", "現行", "便秘治療薬（毎日の排便維持）", "1日1回 夕食前 2錠(5mg×2)", "夕食前", "2錠", start, 28, "", "", "必ず食「前」に。食後だと効果が弱まる", 2],
      ["ヘモクロン", "現行", "痔治療薬（内服）", "1日3回 朝昼夕食後 各1カプセル(200mg)", "朝食後,昼食後,夕食後", "1カプセル", start, 14, "", "", "他の薬より早く（14日で）終わる", 3],
      ["ボラザG軟膏", "現行", "痔治療薬（外用）", "1日1〜2回 塗布(2.4g)", "外用", "適量", start, "", "20本", "", "", 4],
      ["ミヤBM", "現行", "整腸剤（酪酸菌）", "1日1回 朝食後 2錠", "朝食後", "2錠", "", "", "", "", "続けて飲む薬（終了日なし）", 5],
      ["イリボー", "頓服", "", "", "", "", "", "", "", "", "頓服（必要なときに使用）", 11],
      ["ブスコパン", "頓服", "", "", "", "", "", "", "", "", "頓服（必要なときに使用）", 12],
      ["マグミット（酸化マグネシウム）", "過去", "", "", "", "", "", "", "", "", "2026-07-31の処方変更前まで使用", 13],
    ];
    rx.getRange(2, 1, rows.length, RX_HEADERS.length).setValues(rows);
    // 終了日は「開始日＋日数−1」の数式で自動計算（開始日を変えれば追従する）
    for (let r = 2; r <= rows.length + 1; r++) {
      rx.getRange(r, 10).setFormula('=IF(AND(G' + r + '<>"",H' + r + '<>""),G' + r + '+H' + r + '-1,"")');
    }
    rx.getRange(2, 7, rows.length, 1).setNumberFormat("yyyy-mm-dd"); // 開始日
    rx.getRange(2, 8, rows.length, 1).setNumberFormat("0");          // 日数（数値のまま）
    rx.getRange(2, 10, rows.length, 1).setNumberFormat("yyyy-mm-dd"); // 終了日
    done.push("「処方」シートを作成（現行4剤＋旧処方4件）");
  }

  // 3) 設定シート: 観察タグマスタを追加（副作用チェック用）
  const st = mustSheet_(SHEET_SETTINGS);
  const stHead = st.getRange(1, 1, 1, st.getLastColumn()).getValues()[0];
  if (stHead.indexOf("観察タグ名") < 0) {
    const col = st.getLastColumn() + 2; // 1列空けて右側に追加
    st.getRange(1, col, 1, 3).setValues([["観察タグ名", "表示順", "有効"]]).setFontWeight("bold");
    const obs = [["腹痛", 1, "有効"], ["お腹ゴロゴロ", 2, "有効"], ["ガス・お腹の張り", 3, "有効"], ["吐き気", 4, "有効"], ["問題なし", 5, "有効"]];
    st.getRange(2, col, obs.length, 3).setValues(obs);
    done.push("設定シートに観察タグマスタを追加");
  }

  // 4) 薬マスタの整理
  //    - 新処方4剤を追加（有効）
  //    - イリボー・ブスコパンは頓服として有効のまま（「その他の服薬」ボタンに表示）
  //    - ミヤBMは服薬チェック側で管理するため頓服ボタンからは外す（無効）
  //    - マグミットは旧処方として無効化（行は残る＝履歴）
  //    ※アトモキセチン・チラーヂンは別疾患の継続薬の可能性があるため有効のまま
  const medColIdx = stHead.indexOf("薬名");
  if (medColIdx >= 0) {
    const values = st.getDataRange().getValues();
    const names = [];
    let maxOrder = 0, lastRow = 1;
    for (let r = 1; r < values.length; r++) {
      const n = String(values[r][medColIdx] || "").trim();
      if (!n) continue;
      names.push(n);
      lastRow = r + 1;
      maxOrder = Math.max(maxOrder, Number(values[r][medColIdx + 1]) || 0);
    }
    const newMeds = ["ポリフル", "グーフィス", "ヘモクロン", "ボラザG軟膏"];
    const toAdd = newMeds.filter(function (n) { return names.indexOf(n) < 0; });
    if (toAdd.length) {
      const rows = toAdd.map(function (n, i) { return [n, maxOrder + i + 1, "有効"]; });
      st.getRange(lastRow + 1, medColIdx + 1, rows.length, 3).setValues(rows);
      done.push("薬マスタに追加: " + toAdd.join("、"));
    }
    const desired = { "イリボー": "有効", "ブスコパン": "有効", "ミヤBM": "無効", "マグミット": "無効" };
    const changed = [];
    for (let r = 1; r < values.length; r++) {
      const n = String(values[r][medColIdx] || "").trim();
      if (n in desired && String(values[r][medColIdx + 2]) !== desired[n]) {
        st.getRange(r + 1, medColIdx + 3).setValue(desired[n]);
        changed.push(n + "→" + desired[n]);
      }
    }
    if (changed.length) done.push("薬マスタの有効/無効を更新: " + changed.join("、"));
  }

  // 5) 処方シートの調整（すでに作成済みの場合にも適用される）
  //    ミヤBMを「現行」（朝食後2錠）に、イリボー/ブスコパンを「頓服」に
  const rxSh = ss.getSheetByName(SHEET_RX);
  if (rxSh) {
    const rxVals = rxSh.getDataRange().getValues();
    const rxHead = rxVals[0];
    const col = {};
    RX_HEADERS.forEach(function (h) { col[h] = rxHead.indexOf(h); });
    const setCell = function (r, h, v) { if (col[h] >= 0) rxSh.getRange(r + 1, col[h] + 1).setValue(v); };

    let miyaRow = -1;
    for (let r = 1; r < rxVals.length; r++) {
      const n = String(rxVals[r][col["薬名"]] || "").trim();
      if (n === "ミヤBM" && miyaRow < 0) miyaRow = r;
      if ((n === "イリボー" || n === "ブスコパン") && String(rxVals[r][col["状態"]]) === "過去") {
        setCell(r, "状態", "頓服");
        setCell(r, "服用メモ", "頓服（必要なときに使用）");
        done.push(n + "を「頓服」に変更");
      }
    }
    if (miyaRow >= 0) {
      if (String(rxVals[miyaRow][col["状態"]]) !== "現行") {
        setCell(miyaRow, "状態", "現行");
        setCell(miyaRow, "分類", "整腸剤（酪酸菌）");
        setCell(miyaRow, "用法", "1日1回 朝食後 2錠");
        setCell(miyaRow, "タイミング", "朝食後");
        setCell(miyaRow, "1回量", "2錠");
        setCell(miyaRow, "服用メモ", "続けて飲む薬（終了日なし）");
        setCell(miyaRow, "表示順", 5);
        done.push("ミヤBMを服薬チェック（朝食後2錠）に追加");
      }
    } else {
      const map = { "薬名": "ミヤBM", "状態": "現行", "分類": "整腸剤（酪酸菌）", "用法": "1日1回 朝食後 2錠",
                    "タイミング": "朝食後", "1回量": "2錠", "服用メモ": "続けて飲む薬（終了日なし）", "表示順": 5 };
      rxSh.appendRow(rxHead.map(function (h) { return h in map ? map[h] : ""; }));
      done.push("ミヤBMを服薬チェック（朝食後2錠）に追加");
    }
  }

  SpreadsheetApp.getUi().alert(
    done.length
      ? "処方アップデートを適用しました。\n\n・" + done.join("\n・") +
        "\n\n※やめていない薬が無効化されていたら、設定シートで「有効」に戻してください。"
      : "すでに適用済みです（変更はありません）。"
  );
}

// 「日数」セルの値を数値として読む。
// セルが誤って日付書式になっていた場合（例: 28 → 1900-01-27）もシリアル値に戻して読む
function readDayCount_(v) {
  if (v instanceof Date) {
    return Math.round((v - new Date(1899, 11, 30)) / 86400000) || null;
  }
  return Number(v) || null;
}

// シートの1行目に指定の見出しがなければ末尾に追加する。追加したらtrue
function ensureHeader_(sh, name) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (head.indexOf(name) >= 0) return false;
  sh.getRange(1, lastCol + 1).setValue(name).setFontWeight("bold");
  return true;
}

/* =========================================================
   日次トリガー設定（毎朝5時ごろに buildDashboard を自動実行）
   ========================================================= */
function setupDailyTrigger() {
  // 二重登録を防ぐため、既存の同トリガーを消してから作る
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === "buildDashboard") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("buildDashboard").timeBased().everyDays(1).atHour(5).create();
  SpreadsheetApp.getUi().alert("設定しました。毎朝5時ごろにダッシュボードが自動更新されます。");
}

/* =========================================================
   doPost : アプリからの記録を受け取る
   ※ アプリ側は text/plain でJSON文字列を送ってくる
   ========================================================= */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 同時送信で行がぶつからないように順番待ち
  try {
    const body = JSON.parse(e.postData.contents);

    // 合言葉チェック（設定シートに合言葉が書かれている場合のみ有効）
    if (!checkAuth_(body.auth)) {
      return json_({ ok: false, authError: true, error: "合言葉が一致しません" });
    }

    if (body.type === "status") {
      // アプリ起動時の設定マスタ＋簡易集計の取得（旧doGetの役割）
      const cfg = getSettings_();
      return json_({
        ok: true,
        settings: { medicines: cfg.medicines, tags: cfg.tags, obsTags: cfg.obsTags, alertDays: cfg.alertDays },
        summary: computeSummary_(cfg),
        prescriptions: currentRx_(),
        todayMeds: todayMeds_(),
        todayEvents: todayEvents_(),
      });
    }
    if (body.type === "event") {
      appendEvent_(body.data, body.clientId);
    } else if (body.type === "daily") {
      upsertDaily_(body.data);
    } else if (body.type === "deleteMedToday") {
      // 服薬チェックの取り消し（今日の該当行を1件削除）
      const removed = deleteMedToday_(body.data || {});
      return json_({ ok: true, removed: removed });
    } else {
      throw new Error("不明なデータ種別: " + body.type);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// イベント（排便・服薬）を1行追記
// ※列の並びが変わっても壊れないよう、見出し名に合わせて書き込む
// ※二重送信対策: 同じクライアントID、または直近に同一内容（発生時刻が1分以内）の
//   行がすでにある場合は追記せずスキップする
function appendEvent_(d, clientId) {
  const sh = mustSheet_(SHEET_EVENT);
  const newAt = d.occurredAt ? new Date(d.occurredAt) : new Date();

  // 直近50行と照合して重複を弾く
  const last = sh.getLastRow();
  if (last >= 2) {
    const from = Math.max(2, last - 49);
    const values = sh.getRange(from, 1, last - from + 1, sh.getLastColumn()).getValues();
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const iKind = head.indexOf("種別"), iAt = head.indexOf("発生時刻"),
          iBri = head.indexOf("ブリストルスケール"), iMed = head.indexOf("薬名"),
          iTim = head.indexOf("タイミング"), iCid = head.indexOf("クライアントID");
    for (let r = 0; r < values.length; r++) {
      // 再送による重複（同じクライアントID）
      if (iCid >= 0 && clientId && String(values[r][iCid]) === String(clientId)) return;
      // 二度押しによる重複（同一内容かつ発生時刻が1分以内）
      const at = values[r][iAt];
      if (!(at instanceof Date)) continue;
      const same =
        String(values[r][iKind]) === String(d.kind || "") &&
        String(values[r][iMed] || "") === String(d.med || "") &&
        (iTim < 0 || String(values[r][iTim] || "") === String(d.timing || "")) &&
        String(values[r][iBri] || "") === String(d.bristol || "");
      if (same && Math.abs(at - newAt) < 60 * 1000) return;
    }
  }

  const map = {
    "記録日時": new Date(), // サーバー側で自動付与
    "種別": d.kind || "",
    "発生時刻": d.occurredAt ? new Date(d.occurredAt) : new Date(),
    "ブリストルスケール": d.bristol || "",
    "腹痛": d.pain || "",
    "残便感": d.zanben || "",
    "薬名": d.med || "",
    "タイミング": d.timing || "",
    "メモ": d.memo || "",
    "クライアントID": clientId || "",
  };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(head.map(function (h) { return h in map ? map[h] : ""; }));
}

// デイリーログを追記（同じ日付が既にあれば上書き）
function upsertDaily_(d) {
  const sh = mustSheet_(SHEET_DAILY);

  // 朝昼夕の食事（新形式）。旧形式（tags/mealMemo）で来た場合もそのまま動く
  const meals = d.meals || {};
  const meal = function (k) { return meals[k] || {}; };
  const unionTags = [];
  ["morning", "noon", "evening"].forEach(function (k) {
    (meal(k).tags || []).forEach(function (t) { if (unionTags.indexOf(t) < 0) unionTags.push(t); });
  });
  (d.tags || []).forEach(function (t) { if (unionTags.indexOf(t) < 0) unionTags.push(t); });
  const memoParts = [];
  if (meal("morning").memo) memoParts.push("朝:" + meal("morning").memo);
  if (meal("noon").memo) memoParts.push("昼:" + meal("noon").memo);
  if (meal("evening").memo) memoParts.push("夕:" + meal("evening").memo);

  const map = {
    "日付": d.date || fmtDate_(new Date()),
    "睡眠時間": d.sleep === "" || d.sleep == null ? "" : Number(d.sleep),
    "食事タグ": unionTags.join(","), // 3食の合算（トリガー分析はこの列を使う）
    "食事メモ": d.mealMemo || memoParts.join(" ／ "),
    "朝食タグ": (meal("morning").tags || []).join(","),
    "朝食メモ": meal("morning").memo || "",
    "昼食タグ": (meal("noon").tags || []).join(","),
    "昼食メモ": meal("noon").memo || "",
    "夕食タグ": (meal("evening").tags || []).join(","),
    "夕食メモ": meal("evening").memo || "",
    "水分量": d.water === "" || d.water == null ? "" : Number(d.water),
    "運動": d.exercise || "",
    "気分": d.mood || "",
    "ストレス": d.stress || "",
    "メンタルメモ": d.mentalMemo || "",
    "メモ": d.memo || "",
    "観察タグ": (d.obsTags || []).join(","),
    // 満腹度は食事ごとに記録。「満腹度」列にはまとめ（後方互換用）を書く
    "朝食満腹度": meal("morning").fullness || "",
    "昼食満腹度": meal("noon").fullness || "",
    "夕食満腹度": meal("evening").fullness || "",
    "満腹度": d.fullness || fullnessSummary_(meals),
  };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = head.map(function (h) { return h in map ? map[h] : ""; });
  const last = sh.getLastRow();
  let written = false;
  if (last >= 2) {
    const dates = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      if (fmtDate_(dates[i][0]) === String(d.date)) {
        sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
        written = true;
        break;
      }
    }
  }
  if (!written) sh.appendRow(row);

  // 生理開始のトグルを周期ログに反映（trueで登録・falseでその日の行を削除）
  if (typeof d.periodStart === "boolean") {
    updatePeriodLog_(map["日付"], d.periodStart);
  }
}

// 「満腹度」列に書くまとめ文字列（例: 朝:八分目 ／ 夕:食べすぎて気持ち悪い）
function fullnessSummary_(meals) {
  const parts = [];
  if (meals.morning && meals.morning.fullness) parts.push("朝:" + meals.morning.fullness);
  if (meals.noon && meals.noon.fullness) parts.push("昼:" + meals.noon.fullness);
  if (meals.evening && meals.evening.fullness) parts.push("夕:" + meals.evening.fullness);
  return parts.join(" ／ ");
}

// 周期ログに「開始日」を登録/解除する
function updatePeriodLog_(dateKey, isStart) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PERIOD);
  if (!sh) return; // シート未作成なら何もしない（アップデート未適用時）
  const last = sh.getLastRow();
  let foundRow = -1;
  if (last >= 2) {
    const dates = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      if (dates[i][0] instanceof Date && fmtDate_(dates[i][0]) === String(dateKey)) {
        foundRow = i + 2;
        break;
      }
    }
  }
  if (isStart && foundRow < 0) {
    sh.appendRow([parseDate_(dateKey), "アプリから登録"]);
  } else if (!isStart && foundRow >= 0) {
    sh.deleteRow(foundRow);
  }
}

// 周期ログの開始日一覧（日付キーの昇順）
function readCycleStarts_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PERIOD);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let r = 1; r < values.length; r++) {
    if (values[r][0] instanceof Date) out.push(fmtDate_(values[r][0]));
  }
  out.sort();
  return out;
}

// その日が周期何日目か（直近の開始日から数えて1日目〜。開始日がなければnull、45日超は打ち切り）
function cycleDayFor_(dateKey, starts) {
  let latest = null;
  for (let i = 0; i < starts.length; i++) {
    if (starts[i] <= dateKey) latest = starts[i];
  }
  if (!latest) return null;
  const diff = Math.round((parseDate_(dateKey) - parseDate_(latest)) / 86400000) + 1;
  return diff <= 45 ? diff : null;
}

// 今日の服薬チェックを1件取り消す（薬名とタイミングが一致する今日の行を削除）
function deleteMedToday_(d) {
  const sh = mustSheet_(SHEET_EVENT);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return false;
  const head = values[0];
  const iKind = head.indexOf("種別"), iAt = head.indexOf("発生時刻"),
        iMed = head.indexOf("薬名"), iTim = head.indexOf("タイミング");
  const today = fmtDate_(new Date());
  for (let r = values.length - 1; r >= 1; r--) { // 新しい行から探す
    if (String(values[r][iKind]) !== "服薬") continue;
    if (String(values[r][iMed]) !== String(d.med)) continue;
    if (iTim >= 0 && String(values[r][iTim] || "") !== String(d.timing || "")) continue;
    if (!(values[r][iAt] instanceof Date) || fmtDate_(values[r][iAt]) !== today) continue;
    sh.deleteRow(r + 1);
    return true;
  }
  return false;
}

/* =========================================================
   doGet : データは返さない（合言葉なしで見られるのを防ぐため、
   設定・集計の取得もすべて doPost の type:"status" に統一）
   ========================================================= */
function doGet(e) {
  return json_({ ok: true, message: "IBSログのAPIは動作中です。データの閲覧・記録はアプリから行ってください。" });
}

// 合言葉の照合。設定シートに合言葉が未記入の間は誰でも通す
// （初期セットアップ中でも動くようにするため）
function checkAuth_(given) {
  const secret = getSettings_().passphrase;
  if (!secret) return true;
  return String(given || "").trim() === secret;
}

/* =========================================================
   設定シートの読み込み
   「薬名」「タグ名」「項目」という見出しを探して読むので、
   将来列を追加しても壊れない
   ========================================================= */
function getSettings_() {
  const sh = mustSheet_(SHEET_SETTINGS);
  const values = sh.getDataRange().getValues();
  const head = values[0] || [];

  const medCol = head.indexOf("薬名");
  const tagCol = head.indexOf("タグ名");
  const keyCol = head.indexOf("項目");

  // マスタ表を読む共通処理（名前・表示順・有効の3列前提）
  function readMaster(col) {
    if (col < 0) return [];
    const rows = [];
    for (let r = 1; r < values.length; r++) {
      const name = String(values[r][col] || "").trim();
      if (!name) continue;
      const order = Number(values[r][col + 1]) || 999;
      const enabled = String(values[r][col + 2] || "有効") !== "無効";
      if (enabled) rows.push({ name: name, order: order });
    }
    rows.sort((a, b) => a.order - b.order);
    return rows.map((x) => x.name);
  }

  // 閾値などのキー・値を読む
  const map = {};
  if (keyCol >= 0) {
    for (let r = 1; r < values.length; r++) {
      const k = String(values[r][keyCol] || "").trim();
      if (k) map[k] = values[r][keyCol + 1];
    }
  }

  return {
    medicines: readMaster(medCol),
    tags: readMaster(tagCol),
    obsTags: readMaster(head.indexOf("観察タグ名")),
    alertDays: Number(map["便秘警戒日数"]) || 3,
    diarrheaBristol: Number(map["下痢判定ブリストル"]) || 6,
    diarrheaCount: Number(map["下痢判定回数"]) || 3,
    passphrase: String(map["合言葉"] == null ? "" : map["合言葉"]).trim(), // 外部には返さないこと
  };
}

/* =========================================================
   直近7日の簡易集計＋便秘連続日数（アプリ上部カード用）
   ========================================================= */
function computeSummary_(cfg) {
  const events = readEvents_();
  const today = startOfDay_(new Date());

  // 排便があった日の集合
  const stoolDays = {};
  events.forEach((ev) => {
    if (ev.kind === "排便") stoolDays[fmtDate_(ev.at)] = true;
  });

  // 便秘連続日数 = 今日からさかのぼって「排便なし」が続く日数
  let streak = 0;
  const d = new Date(today);
  while (streak < 366) {
    if (stoolDays[fmtDate_(d)]) break;
    streak++;
    d.setDate(d.getDate() - 1);
    if (Object.keys(stoolDays).length === 0) break; // データが1件もない場合
  }
  if (Object.keys(stoolDays).length === 0) streak = 0;

  // 直近7日（今日を含む）の集計
  const from = new Date(today);
  from.setDate(from.getDate() - 6);
  let stoolCount = 0, medCount = 0, bristolSum = 0, bristolN = 0;
  const painDaySet = {};
  events.forEach((ev) => {
    if (ev.at < from) return;
    if (ev.kind === "排便") {
      stoolCount++;
      if (ev.bristol) { bristolSum += ev.bristol; bristolN++; }
      if (ev.pain && ev.pain !== "なし") painDaySet[fmtDate_(ev.at)] = true;
    } else if (ev.kind === "服薬") {
      medCount++;
    }
  });

  return {
    constipationDays: streak,
    last7: {
      stoolCount: stoolCount,
      avgBristol: bristolN ? Math.round((bristolSum / bristolN) * 10) / 10 : null,
      painDays: Object.keys(painDaySet).length,
      medCount: medCount,
    },
  };
}

/* =========================================================
   処方シートの読み込みと服薬チェック関連
   ========================================================= */
function readRx_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RX);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0];
  const idx = {};
  RX_HEADERS.forEach(function (h) { idx[h] = head.indexOf(h); });
  const get = function (row, h) { return idx[h] >= 0 ? row[idx[h]] : ""; };
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const name = String(get(values[r], "薬名") || "").trim();
    if (!name) continue;
    out.push({
      name: name,
      status: String(get(values[r], "状態") || "").trim(),
      category: String(get(values[r], "分類") || ""),
      usage: String(get(values[r], "用法") || ""),
      timings: String(get(values[r], "タイミング") || "").split(",").map(function (s) { return s.trim(); }).filter(String),
      dose: String(get(values[r], "1回量") || ""),
      start: get(values[r], "開始日") instanceof Date ? get(values[r], "開始日") : null,
      days: readDayCount_(get(values[r], "日数")),
      qty: String(get(values[r], "数量") || ""),
      note: String(get(values[r], "服用メモ") || ""),
      order: Number(get(values[r], "表示順")) || 999,
    });
  }
  out.sort(function (a, b) { return a.order - b.order; });
  return out;
}

// 現行処方をアプリ用に整形（終了日・残り日数を計算して返す）
function currentRx_() {
  const today = startOfDay_(new Date());
  return readRx_().filter(function (r) { return r.status === "現行"; }).map(function (r) {
    let endDate = null, remainingDays = null;
    if (r.start && r.days) {
      const end = startOfDay_(r.start);
      end.setDate(end.getDate() + r.days - 1);
      endDate = fmtDate_(end);
      remainingDays = Math.max(0, Math.round((end - today) / 86400000) + 1);
    }
    return {
      name: r.name, category: r.category, usage: r.usage, timings: r.timings,
      dose: r.dose, startDate: r.start ? fmtDate_(r.start) : null, days: r.days,
      qty: r.qty, endDate: endDate, remainingDays: remainingDays, note: r.note,
    };
  });
}

// 今日すでに記録された服薬（薬名＋タイミング）の一覧
function todayMeds_() {
  const today = fmtDate_(new Date());
  return readEvents_()
    .filter(function (ev) { return ev.kind === "服薬" && fmtDate_(ev.at) === today; })
    .map(function (ev) { return { med: ev.med, timing: ev.timing || "" }; });
}

// 今日の全イベント（アプリの「今日の記録」履歴表示用）
function todayEvents_() {
  const today = fmtDate_(new Date());
  return readEvents_()
    .filter(function (ev) { return fmtDate_(ev.at) === today; })
    .map(function (ev) {
      return {
        time: Utilities.formatDate(ev.at, TZ, "HH:mm"),
        kind: ev.kind,
        bristol: ev.bristol,
        pain: ev.pain,
        med: ev.med,
        timing: ev.timing || "",
        memo: ev.memo || "",
      };
    });
}

/* =========================================================
   夕食前リマインダーは不要とのことで廃止。
   万一トリガーが設定済みだった場合も、実行時に自動で解除される
   ========================================================= */
function eveningReminder() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "eveningReminder") ScriptApp.deleteTrigger(t);
  });
}

/* =========================================================
   生ログの読み込み（見出し名で列を探すので列追加に強い）
   ========================================================= */
function readEvents_() {
  const sh = mustSheet_(SHEET_EVENT);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0];
  const iKind = head.indexOf("種別"), iAt = head.indexOf("発生時刻"),
        iBri = head.indexOf("ブリストルスケール"), iPain = head.indexOf("腹痛"),
        iZan = head.indexOf("残便感"), iMed = head.indexOf("薬名"),
        iTim = head.indexOf("タイミング"), iMemo = head.indexOf("メモ");
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const at = values[r][iAt];
    if (!(at instanceof Date)) continue; // 発生時刻が読めない行はスキップ
    out.push({
      kind: String(values[r][iKind] || ""),
      at: at,
      bristol: Number(values[r][iBri]) || null,
      pain: String(values[r][iPain] || ""),
      zanben: String(values[r][iZan] || ""),
      med: String(values[r][iMed] || ""),
      timing: iTim >= 0 ? String(values[r][iTim] || "") : "",
      memo: iMemo >= 0 ? String(values[r][iMemo] || "") : "",
    });
  }
  out.sort((a, b) => a.at - b.at);
  return out;
}

function readDaily_() {
  const sh = mustSheet_(SHEET_DAILY);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return {};
  const head = values[0];
  const idx = {};
  DAILY_HEADERS.forEach((h) => (idx[h] = head.indexOf(h)));
  const byDate = {}; // 日付キー（"2026-07-30"形式）で引ける形にする
  for (let r = 1; r < values.length; r++) {
    const dateVal = values[r][idx["日付"]];
    if (!dateVal) continue;
    const key = fmtDate_(dateVal);
    byDate[key] = {
      sleep: numOrNull_(values[r][idx["睡眠時間"]]),
      tags: String(values[r][idx["食事タグ"]] || "").split(",").map((s) => s.trim()).filter(String),
      water: numOrNull_(values[r][idx["水分量"]]),
      exercise: String(values[r][idx["運動"]] || ""),
      mood: numOrNull_(values[r][idx["気分"]]),
      stress: numOrNull_(values[r][idx["ストレス"]]),
      obsTags: idx["観察タグ"] >= 0
        ? String(values[r][idx["観察タグ"]] || "").split(",").map((s) => s.trim()).filter(String)
        : [],
      // その日の満腹度一覧（朝昼夕の各列＋旧形式の「満腹度」列を合わせて読む）
      fullnessList: ["朝食満腹度", "昼食満腹度", "夕食満腹度", "満腹度"]
        .map((h) => (idx[h] >= 0 ? String(values[r][idx[h]] || "").trim() : ""))
        .filter((v) => FULLNESS_OPTIONS.indexOf(v) >= 0),
    };
  }
  return byDate;
}

/* =========================================================
   ダッシュボード一括生成（日次トリガー/メニューから実行）
   ========================================================= */
function buildDashboard() {
  const cfg = getSettings_();
  const events = readEvents_();
  const daily = readDaily_();

  // すべての分析の土台になる「日別ステート」を先に計算
  const states = computeDailyStates_(events, daily, cfg);

  // 生理周期の「周期何日目か」を各日に付与（周期ログ未作成なら全てnull）
  const cycleStarts = readCycleStarts_();
  states.forEach(function (s) { s.cycleDay = cycleDayFor_(s.key, cycleStarts); });

  writeStateSheet_(states);
  buildWeekly_(states, events);
  buildTriggerAnalysis_(states, daily);
  buildCalendar_(states);
  buildCycle_(states);
  buildMental_(states);
  buildRhythm_(events);
  buildRxProgress_(states, daily, events);
  buildPeriodAnalysis_(states, cycleStarts);
  buildLegend_();
}

/* =========================================================
   周期×体調（生理周期と症状の関係）
   周期ログの開始日をもとに、周期のフェーズごとに症状を集計する。
   体調変化が「周期由来」か「食事由来」かの切り分けに使う
   ========================================================= */
function buildPeriodAnalysis_(states, cycleStarts) {
  if (!cycleStarts.length) return; // 周期ログ未作成・データなしなら何もしない
  const sh = resetSheet_(SHEET_PERIOD_ANALYSIS);

  sh.getRange(1, 1).setValue("生理周期×体調").setFontWeight("bold").setFontSize(13);
  sh.getRange(2, 1).setValue("記録済みの開始日: " + cycleStarts.join("、"));

  // 平均周期長（開始日が2回以上あれば計算できる）
  if (cycleStarts.length >= 2) {
    const gaps = [];
    for (let i = 1; i < cycleStarts.length; i++) {
      gaps.push(Math.round((parseDate_(cycleStarts[i]) - parseDate_(cycleStarts[i - 1])) / 86400000));
    }
    const avg = Math.round(gaps.reduce(function (a, b) { return a + b; }, 0) / gaps.length);
    sh.getRange(3, 1).setValue("平均周期: " + avg + "日（記録" + cycleStarts.length + "回）");
  } else {
    sh.getRange(3, 1).setValue("開始日が2回以上たまると平均周期が計算されます");
  }

  // フェーズ別の集計（日数区分は一般的な目安）
  const phases = [
    { label: "月経期（1〜5日目）", from: 1, to: 5 },
    { label: "卵胞期（6〜13日目）", from: 6, to: 13 },
    { label: "排卵期（14〜16日目）", from: 14, to: 16 },
    { label: "黄体期（17日目〜）", from: 17, to: 45 },
  ];
  const head = ["フェーズ", "日数", "排便回数/日", "平均ブリストル", "腹痛日率(%)", "下痢日率(%)", "便秘日率(%)", "平均気分", "平均ストレス"];
  sh.getRange(5, 1, 1, head.length).setValues([head]).setFontWeight("bold").setBackground("#e5e7eb");

  const rows = phases.map(function (ph) {
    const days = states.filter(function (s) { return s.cycleDay != null && s.cycleDay >= ph.from && s.cycleDay <= ph.to; });
    if (!days.length) return [ph.label, 0, "", "", "", "", "", "", ""];
    let stool = 0, bSum = 0, bN = 0, pain = 0, dia = 0, cons = 0, moodSum = 0, moodN = 0, strSum = 0, strN = 0;
    days.forEach(function (s) {
      stool += s.stoolCount;
      if (s.avgBristol != null) { bSum += s.avgBristol; bN++; }
      if (s.pain) pain++;
      if (s.state === "下痢日") dia++;
      if (s.state === "便秘日") cons++;
      if (s.mood != null) { moodSum += s.mood; moodN++; }
      if (s.stress != null) { strSum += s.stress; strN++; }
    });
    const n = days.length;
    return [
      ph.label, n,
      Math.round((stool / n) * 100) / 100,
      bN ? Math.round((bSum / bN) * 10) / 10 : "",
      Math.round((pain / n) * 100),
      Math.round((dia / n) * 100),
      Math.round((cons / n) * 100),
      moodN ? Math.round((moodSum / moodN) * 10) / 10 : "",
      strN ? Math.round((strSum / strN) * 10) / 10 : "",
    ];
  });
  sh.getRange(6, 1, rows.length, head.length).setValues(rows);
  sh.getRange(11, 1).setValue("※ フェーズの日数区分は一般的な目安です。データがたまるほど傾向が見えてきます").setFontSize(9);

  if (states.filter(function (s) { return s.cycleDay != null; }).length < 5) return;
  sh.insertChart(
    sh.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(sh.getRange(5, 1, rows.length + 1, 1))
      .addRange(sh.getRange(5, 4, rows.length + 1, 1))
      .addRange(sh.getRange(5, 5, rows.length + 1, 1))
      .setPosition(13, 1, 0, 0)
      .setOption("title", "周期フェーズ別の平均ブリストルと腹痛日率")
      .setOption("height", 300).setOption("width", 560)
      .build()
  );
}

/* =========================================================
   説明・凡例シート
   各シートの見方、ブリストルスケールの説明、色分けの凡例
   ========================================================= */
function buildLegend_() {
  const sh = resetSheet_("説明・凡例");
  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(2, 460);
  sh.setColumnWidth(3, 160);
  let r = 1;

  // --- 各シートの見方 ---
  sh.getRange(r, 1).setValue("📖 各シートの見方").setFontWeight("bold").setFontSize(13);
  r += 1;
  const guide = [
    ["シート名", "何が見られるか"],
    ["イベントログ", "排便・服薬の生記録（1行＝1回）。アプリから自動追記される"],
    ["デイリーログ", "1日1行の生活記録（睡眠・食事・気分など）。アプリから自動保存される"],
    ["設定", "薬・食事タグ・観察タグ・しきい値・合言葉。ここを編集するとアプリに反映"],
    ["処方", "現行・頓服・過去の薬。処方が変わったらここを編集"],
    ["日別ステート", "毎日を「便秘日/通常日/下痢日」に自動分類した表。すべての分析の土台"],
    ["週次サマリー", "週ごとの排便回数・便の状態・腹痛・服薬の推移（グラフ付き）"],
    ["トリガー分析", "食べた物タグ別に「翌日に悪化しやすいか」を比較。数字のセルが赤いほど悪化傾向"],
    ["カレンダービュー", "日々の状態をカレンダー形式で色分け（凡例は下記）"],
    ["サイクルビュー", "便秘→腹痛→下痢の波を1つのグラフで見る。下向き＝便秘の深さ、上向き＝下痢の強さ"],
    ["メンタル×体調", "気分・ストレス・睡眠と、同日/翌日の症状の関係"],
    ["排便リズム", "排便が多い時間帯・曜日"],
    ["処方経過", "受診報告用。薬の服用と排便の変化を日付で並べた表（青い行＝処方開始日）"],
    ["周期ログ", "生理の開始日（アプリの🩸ボタンか、このシートへの入力で記録）"],
    ["周期×体調", "生理周期のフェーズ別に症状を集計。周期由来か食事由来かの切り分け用"],
  ];
  sh.getRange(r, 1, guide.length, 2).setValues(guide);
  sh.getRange(r, 1, 1, 2).setFontWeight("bold").setBackground("#e5e7eb");
  r += guide.length + 2;

  // --- ブリストルスケールとは ---
  sh.getRange(r, 1).setValue("💩 ブリストルスケールとは").setFontWeight("bold").setFontSize(13);
  sh.getRange(r + 1, 1, 1, 3).setValues([["便の形を1〜7の数字で表す世界共通の物差しです。4が理想的な状態。1に近いほど便秘傾向、7に近いほど下痢傾向です。", "", ""]]);
  r += 3;
  const BRISTOL_LEGEND = [
    [1, "コロコロ便（硬くて木の実のよう）", "強い便秘傾向", "#92400e"],
    [2, "硬い便（ソーセージ状だが硬い）", "便秘傾向", "#b45309"],
    [3, "やや硬い便（表面にひび割れ）", "やや便秘傾向", "#d97706"],
    [4, "普通便（なめらかなバナナ状）", "理想的", "#0e9488"],
    [5, "やや軟らかい便（半固形）", "やや下痢傾向", "#f59e0b"],
    [6, "泥状便（形がくずれている）", "下痢傾向", "#f97316"],
    [7, "水様便（固形物がない液体）", "強い下痢傾向", "#dc2626"],
  ];
  sh.getRange(r, 1, 1, 3).setValues([["番号", "便の状態", "意味"]]).setFontWeight("bold").setBackground("#e5e7eb");
  BRISTOL_LEGEND.forEach(function (b, i) {
    sh.getRange(r + 1 + i, 1, 1, 3).setValues([[b[0], b[1], b[2]]]);
    sh.getRange(r + 1 + i, 1).setBackground(b[3]).setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  });
  r += BRISTOL_LEGEND.length + 3;

  // --- カレンダービューの色分け ---
  sh.getRange(r, 1).setValue("🗓 カレンダービューの色分け").setFontWeight("bold").setFontSize(13);
  r += 1;
  const AMBER = ["#fef3c7", "#fde68a", "#fcd34d", "#f59e0b", "#d97706"];
  const RED = ["#fecaca", "#fca5a5", "#f87171", "#ef4444", "#dc2626"];
  sh.getRange(r, 1).setValue("便秘日");
  sh.getRange(r, 2).setValue("排便がなかった日。連続日数が長いほど色が濃い（1日→5日以上）");
  AMBER.forEach(function (c, i) { sh.getRange(r, 3 + i).setBackground(c); });
  r += 1;
  sh.getRange(r, 1).setValue("通常日").setBackground("#e5e7eb");
  sh.getRange(r, 2).setValue("普通の排便があった日");
  r += 1;
  sh.getRange(r, 1).setValue("下痢日");
  sh.getRange(r, 2).setValue("下痢だった日。回数×硬さが強いほど色が濃い");
  RED.forEach(function (c, i) { sh.getRange(r, 3 + i).setBackground(c); });
  r += 1;
  sh.getRange(r, 1).setValue("▲マーク");
  sh.getRange(r, 2).setValue("その日に腹痛があったことを示す");
  r += 2;
  sh.getRange(r, 1).setValue("※ 判定のしきい値（下痢と判定するブリストル値・回数、便秘の警戒日数）は「設定」シートで変更できます").setFontSize(9);
}

/* =========================================================
   処方経過（受診報告用）
   排便記録と服薬記録を同じ日付で並べ、
   グーフィス開始前後の排便リズム・便の硬さの変化を比較する
   ========================================================= */
function buildRxProgress_(states, daily, events) {
  const rx = readRx_().filter(function (r) { return r.status === "現行"; });
  if (rx.length === 0) return; // 処方シート未設定なら何もしない
  const sh = resetSheet_(SHEET_RXLOG);

  // 基準日 = グーフィスの開始日（なければ現行処方の最も早い開始日）
  const gf = rx.filter(function (r) { return r.name.indexOf("グーフィス") >= 0; })[0];
  const baseStart = (gf && gf.start) ||
    rx.map(function (r) { return r.start; }).filter(Boolean).sort(function (a, b) { return a - b; })[0];
  if (!baseStart) { sh.getRange(1, 1).setValue("処方シートに開始日が入っていません"); return; }
  const startKey = fmtDate_(baseStart);

  const stateByKey = {};
  states.forEach(function (s) { stateByKey[s.key] = s; });

  // 開始前7日と開始後の比較サマリー
  function metrics(list) {
    if (!list.length) return ["", "", "", ""];
    let stool = 0, bSum = 0, bN = 0, cons = 0;
    list.forEach(function (s) {
      stool += s.stoolCount;
      if (s.avgBristol != null) { bSum += s.avgBristol; bN++; }
      if (s.state === "便秘日") cons++;
    });
    return [
      list.length,
      Math.round((stool / list.length) * 100) / 100,
      bN ? Math.round((bSum / bN) * 10) / 10 : "",
      Math.round((cons / list.length) * 100),
    ];
  }
  const before = states.filter(function (s) {
    const d = parseDate_(s.key), b = parseDate_(startKey);
    const diff = (b - d) / 86400000;
    return diff >= 1 && diff <= 7;
  });
  const after = states.filter(function (s) { return s.key >= startKey; });

  sh.getRange(1, 1).setValue("処方経過レポート（" + startKey + " 開始）").setFontWeight("bold").setFontSize(12);
  sh.getRange(2, 1, 1, 5).setValues([["期間", "日数", "排便回数/日", "平均ブリストル", "便秘日の割合(%)"]]).setFontWeight("bold");
  sh.getRange(3, 1, 2, 5).setValues([
    ["開始前7日"].concat(metrics(before)),
    ["開始後"].concat(metrics(after)),
  ]);

  // 次回受診の目安（28日分の薬が終わる日）と各薬の終了日
  const cur = currentRx_();
  const endInfo = cur.filter(function (r) { return r.endDate; })
    .map(function (r) { return r.name + ": " + r.endDate + "まで（あと" + r.remainingDays + "日）"; });
  sh.getRange(6, 1).setValue("薬の終了日: " + (endInfo.join(" ／ ") || "－")).setFontSize(9);
  const maxEnd = cur.map(function (r) { return r.endDate; }).filter(Boolean).sort().slice(-1)[0];
  if (maxEnd) sh.getRange(7, 1).setValue("次回受診の目安: " + maxEnd + " ごろ（28日分の薬が終わる頃）").setFontSize(9).setFontWeight("bold");

  // 日別テーブル（開始7日前〜今日）: 排便と服薬を同じ日付で横に並べる
  const medNames = rx.map(function (r) { return r.name; });
  const expected = {};
  rx.forEach(function (r) { expected[r.name] = r.timings.indexOf("外用") >= 0 ? null : r.timings.length; });

  // 日付ごとの薬別服用回数
  const medByDay = {};
  events.forEach(function (ev) {
    if (ev.kind !== "服薬") return;
    const key = fmtDate_(ev.at);
    if (!medByDay[key]) medByDay[key] = {};
    medByDay[key][ev.med] = (medByDay[key][ev.med] || 0) + 1;
  });

  const head = ["日付", "曜日", "ステート", "排便回数", "平均ブリストル", "腹痛"]
    .concat(medNames.map(function (n) { return n; }))
    .concat(["観察タグ"]);
  const startRow = 9;
  sh.getRange(startRow, 1, 1, head.length).setValues([head]).setFontWeight("bold");

  const from = parseDate_(startKey);
  from.setDate(from.getDate() - 7);
  const today = startOfDay_(new Date());
  const rows = [];
  for (let d = new Date(from); d <= today; d.setDate(d.getDate() + 1)) {
    const key = fmtDate_(d);
    const s = stateByKey[key];
    const dl = daily[key] || {};
    const medCells = medNames.map(function (n) {
      const c = (medByDay[key] && medByDay[key][n]) || 0;
      if (!c) return "";
      return expected[n] ? c + "/" + expected[n] : "○" + (c > 1 ? "×" + c : "");
    });
    rows.push(
      [new Date(d), WEEKDAYS_JP[d.getDay()],
       s ? s.state : "", s ? s.stoolCount : "", s ? orBlank_(s.avgBristol) : "", s && s.pain ? "あり" : ""]
      .concat(medCells)
      .concat([(dl.obsTags || []).join("、")])
    );
  }
  if (rows.length) {
    sh.getRange(startRow + 1, 1, rows.length, head.length).setValues(rows);
    sh.getRange(startRow + 1, 1, rows.length, 1).setNumberFormat("M/d");
    // 開始日の行に目印の背景色
    for (let i = 0; i < rows.length; i++) {
      if (fmtDate_(rows[i][0]) === startKey) {
        sh.getRange(startRow + 1 + i, 1, 1, head.length).setBackground("#dbeafe");
      }
    }
  }
  sh.setFrozenRows(startRow);

  if (rows.length < 2) return;
  sh.insertChart(
    sh.newChart().setChartType(Charts.ChartType.LINE)
      .addRange(sh.getRange(startRow, 1, rows.length + 1, 1))
      .addRange(sh.getRange(startRow, 4, rows.length + 1, 2))
      .setPosition(2, 8, 0, 0)
      .setOption("title", "排便回数と平均ブリストル値の推移（" + startKey + " 処方開始）")
      .setOption("height", 320).setOption("width", 640)
      .build()
  );
}

/* =========================================================
   日別ステート判定
   便秘日=排便なし / 下痢日=ブリストル閾値以上 or 回数閾値以上 / 通常日
   便秘連続日数・サイクル番号・腹痛までの便秘日数もここで計算
   ========================================================= */
function computeDailyStates_(events, daily, cfg) {
  if (events.length === 0 && Object.keys(daily).length === 0) return [];

  // データの開始日 〜 今日 の全日を対象にする
  let minDate = startOfDay_(new Date());
  events.forEach((ev) => { const d = startOfDay_(ev.at); if (d < minDate) minDate = d; });
  Object.keys(daily).forEach((k) => { const d = parseDate_(k); if (d < minDate) minDate = d; });
  const today = startOfDay_(new Date());

  // 日ごとにイベントをまとめる
  const byDay = {};
  events.forEach((ev) => {
    const key = fmtDate_(ev.at);
    if (!byDay[key]) byDay[key] = { stools: [], meds: 0, pain: false };
    if (ev.kind === "排便") byDay[key].stools.push(ev);
    if (ev.kind === "服薬") byDay[key].meds++;
    if (ev.pain && ev.pain !== "なし") byDay[key].pain = true;
  });

  const states = [];
  let streak = 0;      // 便秘連続日数
  let cycleNo = 0;     // サイクル番号
  let prevState = null;

  for (let d = new Date(minDate); d <= today; d.setDate(d.getDate() + 1)) {
    const key = fmtDate_(d);
    const ev = byDay[key] || { stools: [], meds: 0, pain: false };
    const n = ev.stools.length;
    const bristols = ev.stools.map((s) => s.bristol).filter(Boolean);
    const avgB = bristols.length ? Math.round((bristols.reduce((a, b) => a + b, 0) / bristols.length) * 10) / 10 : null;
    const maxB = bristols.length ? Math.max.apply(null, bristols) : null;

    let state;
    if (n === 0) state = "便秘日";
    else if ((maxB != null && maxB >= cfg.diarrheaBristol) || n >= cfg.diarrheaCount) state = "下痢日";
    else state = "通常日";

    streak = state === "便秘日" ? streak + 1 : 0;
    if (state === "便秘日" && prevState !== "便秘日") cycleNo++; // 便秘の始まり＝新しいサイクル

    const dl = daily[key] || {};
    states.push({
      date: new Date(d), key: key, weekday: WEEKDAYS_JP[d.getDay()],
      stoolCount: n, avgBristol: avgB, maxBristol: maxB,
      pain: ev.pain, medCount: ev.meds,
      state: state, streak: streak, cycle: cycleNo || null,
      sleep: dl.sleep != null ? dl.sleep : null,
      mood: dl.mood != null ? dl.mood : null,
      stress: dl.stress != null ? dl.stress : null,
      tags: dl.tags || [],
    });
    prevState = state;
  }

  // 各サイクルの「腹痛発生までの便秘日数」= サイクル内で最初に腹痛が出た日の便秘連続日数
  const cycleFirstPain = {};
  states.forEach((s) => {
    if (s.cycle && s.pain && !(s.cycle in cycleFirstPain)) {
      cycleFirstPain[s.cycle] = s.streak;
    }
  });
  states.forEach((s) => {
    s.painAfterDays = s.cycle && s.cycle in cycleFirstPain ? cycleFirstPain[s.cycle] : null;
  });

  return states;
}

// 日別ステートをシートに書き出す
function writeStateSheet_(states) {
  const sh = resetSheet_(SHEET_STATE);
  const head = ["日付", "曜日", "排便回数", "平均ブリストル", "最大ブリストル", "腹痛あり", "服薬回数",
                "ステート", "便秘連続日数", "サイクル番号", "サイクル内腹痛までの便秘日数", "気分", "ストレス", "睡眠時間", "周期日数"];
  const rows = states.map((s) => [
    s.date, s.weekday, s.stoolCount, orBlank_(s.avgBristol), orBlank_(s.maxBristol),
    s.pain ? "あり" : "", s.medCount, s.state, s.streak, orBlank_(s.cycle), orBlank_(s.painAfterDays),
    orBlank_(s.mood), orBlank_(s.stress), orBlank_(s.sleep), orBlank_(s.cycleDay),
  ]);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
  if (rows.length) {
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    sh.getRange(2, 1, rows.length, 1).setNumberFormat("yyyy-mm-dd");
  }
  sh.setFrozenRows(1);
}

/* =========================================================
   週次サマリー（月曜始まりの週ごとに集計＋グラフ）
   ========================================================= */
function buildWeekly_(states, events) {
  const sh = resetSheet_(SHEET_WEEKLY);
  const head = ["週開始日", "日数", "排便回数/日", "平均ブリストル", "腹痛日率(%)", "服薬回数"];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");

  // 週開始日（月曜）ごとにまとめる
  const weeks = {};
  states.forEach((s) => {
    const key = fmtDate_(mondayOf_(s.date));
    if (!weeks[key]) weeks[key] = { days: 0, stool: 0, bSum: 0, bN: 0, painDays: 0, meds: 0 };
    const w = weeks[key];
    w.days++;
    w.stool += s.stoolCount;
    if (s.avgBristol != null) { w.bSum += s.avgBristol; w.bN++; }
    if (s.pain) w.painDays++;
    w.meds += s.medCount;
  });

  const keys = Object.keys(weeks).sort();
  const rows = keys.map((k) => {
    const w = weeks[k];
    return [
      parseDate_(k), w.days,
      Math.round((w.stool / w.days) * 100) / 100,
      w.bN ? Math.round((w.bSum / w.bN) * 10) / 10 : "",
      Math.round((w.painDays / w.days) * 100),
      w.meds,
    ];
  });
  if (rows.length) {
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    sh.getRange(2, 1, rows.length, 1).setNumberFormat("yyyy-mm-dd");
  }
  sh.setFrozenRows(1);
  if (rows.length < 2) return; // データが少ないうちはグラフを作らない

  // 折れ線: 排便回数/日と平均ブリストルの推移
  sh.insertChart(
    sh.newChart().setChartType(Charts.ChartType.LINE)
      .addRange(sh.getRange(1, 1, rows.length + 1, 1))
      .addRange(sh.getRange(1, 3, rows.length + 1, 2))
      .setPosition(2, 8, 0, 0)
      .setOption("title", "排便回数/日と平均ブリストル値の推移")
      .setOption("height", 300).setOption("width", 560)
      .build()
  );
  // 棒: 腹痛日率と服薬回数
  sh.insertChart(
    sh.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(sh.getRange(1, 1, rows.length + 1, 1))
      .addRange(sh.getRange(1, 5, rows.length + 1, 2))
      .setPosition(18, 8, 0, 0)
      .setOption("title", "腹痛日率(%)と服薬回数")
      .setOption("height", 300).setOption("width", 560)
      .build()
  );
}

/* =========================================================
   トリガー分析（食事タグ別に「翌日」の症状を比較）
   ========================================================= */
function buildTriggerAnalysis_(states, daily) {
  const sh = resetSheet_(SHEET_TRIGGER);
  const head = ["食事タグ", "該当日数", "翌日の平均ブリストル", "翌日の腹痛発生率(%)", "翌日の下痢日率(%)"];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");

  // 日付キー→翌日のステートを引けるようにしておく
  const stateByKey = {};
  states.forEach((s) => (stateByKey[s.key] = s));

  function nextDayOf(key) {
    const d = parseDate_(key);
    d.setDate(d.getDate() + 1);
    return stateByKey[fmtDate_(d)];
  }

  // タグごとに翌日の症状を集計
  const tagSet = {};
  Object.keys(daily).forEach((k) => daily[k].tags.forEach((t) => (tagSet[t] = true)));
  const rows = [];

  // 比較の基準になる「全日平均」を最初の行に置く
  const all = aggregateNextDays_(Object.keys(daily), nextDayOf);
  rows.push(["（全体平均）", all.n, all.avgB, all.painPct, all.diaPct]);

  Object.keys(tagSet).sort().forEach((tag) => {
    const dates = Object.keys(daily).filter((k) => daily[k].tags.indexOf(tag) >= 0);
    const a = aggregateNextDays_(dates, nextDayOf);
    rows.push([tag, a.n, a.avgB, a.painPct, a.diaPct]);
  });

  sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  sh.setFrozenRows(1);

  // 悪化しているセルほど赤くなる色付け（条件付き書式）
  const rules = [];
  [3, 4, 5].forEach((col) => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpoint("#ffffff")
        .setGradientMaxpoint("#f87171")
        .setRanges([sh.getRange(3, col, Math.max(rows.length - 1, 1), 1)])
        .build()
    );
  });
  sh.setConditionalFormatRules(rules);
  sh.getRange("A" + (rows.length + 3)).setValue("※ 該当日数が少ないタグは参考程度に見てください（偶然の影響が大きいため）");

  // --- 満腹度別の症状比較（いずれかの食事で該当した日を集計） ---
  let fr = rows.length + 5;
  sh.getRange(fr, 1).setValue("満腹度別の症状（いずれかの食事で該当した日）").setFontWeight("bold");
  fr += 1;
  const fHead = ["満腹度", "該当日数", "同日の平均ブリストル", "同日の腹痛率(%)", "翌日の平均ブリストル", "翌日の腹痛率(%)"];
  sh.getRange(fr, 1, 1, fHead.length).setValues([fHead]).setFontWeight("bold");
  const fRows = FULLNESS_OPTIONS.map(function (opt) {
    const keys = Object.keys(daily).filter(function (k) { return (daily[k].fullnessList || []).indexOf(opt) >= 0; });
    let n = 0, sB = 0, sBn = 0, sPain = 0, nB = 0, nBn = 0, nPain = 0, nN = 0;
    keys.forEach(function (k) {
      const s = stateByKey[k];
      if (s) {
        n++;
        if (s.avgBristol != null) { sB += s.avgBristol; sBn++; }
        if (s.pain) sPain++;
      }
      const nx = nextDayOf(k);
      if (nx) {
        nN++;
        if (nx.avgBristol != null) { nB += nx.avgBristol; nBn++; }
        if (nx.pain) nPain++;
      }
    });
    return [
      opt, keys.length,
      sBn ? Math.round((sB / sBn) * 10) / 10 : "",
      n ? Math.round((sPain / n) * 100) : "",
      nBn ? Math.round((nB / nBn) * 10) / 10 : "",
      nN ? Math.round((nPain / nN) * 100) : "",
    ];
  });
  sh.getRange(fr + 1, 1, fRows.length, fHead.length).setValues(fRows);
}

// 指定した日付リストの「翌日」の症状をまとめて集計する共通処理
function aggregateNextDays_(dateKeys, nextDayOf) {
  let n = 0, bSum = 0, bN = 0, pain = 0, dia = 0;
  dateKeys.forEach((k) => {
    const next = nextDayOf(k);
    if (!next) return;
    n++;
    if (next.avgBristol != null) { bSum += next.avgBristol; bN++; }
    if (next.pain) pain++;
    if (next.state === "下痢日") dia++;
  });
  return {
    n: n,
    avgB: bN ? Math.round((bSum / bN) * 10) / 10 : "",
    painPct: n ? Math.round((pain / n) * 100) : "",
    diaPct: n ? Math.round((dia / n) * 100) : "",
  };
}

/* =========================================================
   カレンダービュー（日別ステートを色で表示）
   便秘日=アンバー系（連続日数が長いほど濃い）
   通常日=グレー / 下痢日=赤系（回数×硬さで濃い） / 腹痛=▲
   ========================================================= */
function buildCalendar_(states) {
  const sh = resetSheet_(SHEET_CALENDAR);
  sh.getRange(1, 1, 1, 8).setValues([["週", "月", "火", "水", "木", "金", "土", "日"]]).setFontWeight("bold");

  if (states.length === 0) return;

  const AMBER = ["#fef3c7", "#fde68a", "#fcd34d", "#f59e0b", "#d97706"]; // 連続1日→5日以上
  const RED = ["#fecaca", "#fca5a5", "#f87171", "#ef4444", "#dc2626"];   // 下痢の強さ1→5以上
  const GRAY = "#e5e7eb";

  const stateByKey = {};
  states.forEach((s) => (stateByKey[s.key] = s));

  const first = mondayOf_(states[0].date);
  const last = states[states.length - 1].date;

  const labels = [], colors = [], notes = [];
  for (let wk = new Date(first); wk <= last; wk.setDate(wk.getDate() + 7)) {
    const rowLabel = [Utilities.formatDate(wk, TZ, "M/d週")];
    const rowColor = ["#ffffff"];
    for (let i = 0; i < 7; i++) {
      const d = new Date(wk);
      d.setDate(d.getDate() + i);
      const s = stateByKey[fmtDate_(d)];
      if (!s) { rowLabel.push(""); rowColor.push("#ffffff"); continue; }
      const dayText = (d.getDate() === 1 ? Utilities.formatDate(d, TZ, "M/d") : String(d.getDate())) + (s.pain ? " ▲" : "");
      rowLabel.push(dayText);
      if (s.state === "便秘日") rowColor.push(AMBER[Math.min(s.streak, 5) - 1]);
      else if (s.state === "下痢日") rowColor.push(RED[Math.min(diarrheaIntensity_(s), 5) - 1]);
      else rowColor.push(GRAY);
    }
    labels.push(rowLabel);
    colors.push(rowColor);
  }

  sh.getRange(2, 1, labels.length, 8).setValues(labels).setBackgrounds(colors).setHorizontalAlignment("center");
  for (let c = 2; c <= 8; c++) sh.setColumnWidth(c, 52);

  // 凡例
  const legendRow = labels.length + 4;
  sh.getRange(legendRow, 1).setValue("凡例:");
  sh.getRange(legendRow, 2).setValue("便秘日").setBackground(AMBER[2]);
  sh.getRange(legendRow, 3).setValue("通常日").setBackground(GRAY);
  sh.getRange(legendRow, 4).setValue("下痢日").setBackground(RED[2]);
  sh.getRange(legendRow, 5).setValue("▲=腹痛あり");
  sh.getRange(legendRow + 1, 1, 1, 5).setValues([["", "色が濃いほど連続日数が長い", "", "色が濃いほど回数×硬さが強い", ""]]).setFontSize(8);
}

// 下痢の強さ（回数×硬さから1〜5段階）
function diarrheaIntensity_(s) {
  const hardness = s.maxBristol != null ? Math.max(s.maxBristol - 4, 1) : 1; // 5→1, 6→2, 7→3
  return Math.max(1, Math.min(5, Math.round(s.stoolCount * hardness / 2)));
}

/* =========================================================
   サイクルビュー
   便秘の深さ（連続日数）を下向き、下痢の強さを上向きのバーで表示
   腹痛発生日は▲マーカー。腹痛までの平均便秘日数も算出
   ========================================================= */
function buildCycle_(states) {
  const sh = resetSheet_(SHEET_CYCLE);
  const head = ["日付", "便秘の深さ(日)", "下痢の強さ", "腹痛マーカー"];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");

  const rows = states.map((s) => [
    s.date,
    s.state === "便秘日" ? -s.streak : 0,
    s.state === "下痢日" ? diarrheaIntensity_(s) : 0,
    s.pain ? 0.3 : "", // 軸の近くに▲を出すための値
  ]);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    sh.getRange(2, 1, rows.length, 1).setNumberFormat("M/d");
  }
  sh.setFrozenRows(1);

  // 腹痛までの平均便秘日数（= アプリの警戒ラインの根拠になる数字）
  const perCycle = {};
  states.forEach((s) => {
    if (s.cycle && s.painAfterDays != null) perCycle[s.cycle] = s.painAfterDays;
  });
  const painDays = Object.keys(perCycle).map((k) => perCycle[k]);
  const avgPain = painDays.length
    ? Math.round((painDays.reduce((a, b) => a + b, 0) / painDays.length) * 10) / 10
    : null;

  sh.getRange(1, 6, 3, 2).setValues([
    ["サイクル数", states.length ? (states[states.length - 1].cycle || 0) : 0],
    ["腹痛が発生したサイクル数", painDays.length],
    ["腹痛までの平均便秘日数", avgPain != null ? avgPain : "まだ計算できません"],
  ]);
  sh.getRange(1, 6, 3, 1).setFontWeight("bold");
  sh.getRange(4, 6).setValue("※この日数がアプリの警戒ライン（設定シートの便秘警戒日数）の目安になります").setFontSize(8);

  if (rows.length < 2) return;

  sh.insertChart(
    sh.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(sh.getRange(1, 1, rows.length + 1, 4))
      .setPosition(6, 6, 0, 0)
      .setOption("title", "便秘⇄下痢のサイクル（下=便秘の深さ / 上=下痢の強さ / ▲=腹痛）")
      .setOption("isStacked", true)
      .setOption("height", 340).setOption("width", 700)
      .setOption("series", {
        0: { color: "#d97706" },
        1: { color: "#dc2626" },
        2: { type: "line", lineWidth: 0, pointSize: 8, pointShape: "triangle", color: "#7c3aed" },
      })
      .build()
  );
}

/* =========================================================
   メンタル×体調（気分・ストレス・睡眠と同日/翌日の症状）
   ========================================================= */
function buildMental_(states) {
  const sh = resetSheet_(SHEET_MENTAL);
  const stateByKey = {};
  states.forEach((s) => (stateByKey[s.key] = s));

  function nextOf(s) {
    const d = new Date(s.date);
    d.setDate(d.getDate() + 1);
    return stateByKey[fmtDate_(d)];
  }

  // 指定の切り口（気分1〜5など）ごとに同日・翌日の症状を平均する共通処理
  function makeTable(title, groups, getGroup) {
    const head = [title, "日数", "同日 平均ブリストル", "同日 腹痛率(%)", "翌日 平均ブリストル", "翌日 腹痛率(%)"];
    const rows = groups.map((g) => {
      const days = states.filter((s) => getGroup(s) === g.value);
      let n = days.length, sB = 0, sBn = 0, sPain = 0, nB = 0, nBn = 0, nPain = 0, nN = 0;
      days.forEach((s) => {
        if (s.avgBristol != null) { sB += s.avgBristol; sBn++; }
        if (s.pain) sPain++;
        const nx = nextOf(s);
        if (nx) {
          nN++;
          if (nx.avgBristol != null) { nB += nx.avgBristol; nBn++; }
          if (nx.pain) nPain++;
        }
      });
      return [
        g.label, n,
        sBn ? Math.round((sB / sBn) * 10) / 10 : "",
        n ? Math.round((sPain / n) * 100) : "",
        nBn ? Math.round((nB / nBn) * 10) / 10 : "",
        nN ? Math.round((nPain / nN) * 100) : "",
      ];
    });
    return [head].concat(rows);
  }

  const scale5 = [1, 2, 3, 4, 5].map((v) => ({ value: v, label: v }));
  const sleepBins = [
    { value: "~5", label: "5時間未満" }, { value: "5-6", label: "5〜6時間" },
    { value: "6-7", label: "6〜7時間" }, { value: "7-8", label: "7〜8時間" },
    { value: "8~", label: "8時間以上" },
  ];
  function sleepBin(s) {
    if (s.sleep == null) return null;
    if (s.sleep < 5) return "~5";
    if (s.sleep < 6) return "5-6";
    if (s.sleep < 7) return "6-7";
    if (s.sleep < 8) return "7-8";
    return "8~";
  }

  const tables = [
    makeTable("気分", scale5, (s) => s.mood),
    makeTable("ストレス", scale5, (s) => s.stress),
    makeTable("睡眠時間", sleepBins, sleepBin),
  ];

  let row = 1;
  const chartAnchors = [];
  tables.forEach((t) => {
    sh.getRange(row, 1, 1, t[0].length).setValues([t[0]]).setFontWeight("bold");
    sh.getRange(row + 1, 1, t.length - 1, t[0].length).setValues(t.slice(1));
    chartAnchors.push({ row: row, len: t.length });
    row += t.length + 2;
  });

  // ストレス段階別の「翌日」症状グラフ
  const a = chartAnchors[1];
  sh.insertChart(
    sh.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(sh.getRange(a.row, 1, a.len, 1))
      .addRange(sh.getRange(a.row, 5, a.len, 2))
      .setPosition(2, 8, 0, 0)
      .setOption("title", "ストレス段階別 翌日の症状")
      .setOption("height", 300).setOption("width", 520)
      .build()
  );
}

/* =========================================================
   排便リズム（時間帯別ヒストグラム＋曜日別傾向）
   ========================================================= */
function buildRhythm_(events) {
  const sh = resetSheet_(SHEET_RHYTHM);
  const stools = events.filter((ev) => ev.kind === "排便");

  // 時間帯別（0〜23時）
  const hourly = new Array(24).fill(0);
  stools.forEach((ev) => hourly[ev.at.getHours()]++);
  sh.getRange(1, 1, 1, 2).setValues([["時間帯", "排便回数"]]).setFontWeight("bold");
  const hourRows = hourly.map((n, h) => [h + "時", n]);
  sh.getRange(2, 1, 24, 2).setValues(hourRows);

  // 曜日×時間帯（ざっくり4区分）
  const zones = ["深夜(0-6)", "午前(6-12)", "午後(12-18)", "夜(18-24)"];
  const grid = WEEKDAYS_JP.map(() => [0, 0, 0, 0]);
  stools.forEach((ev) => {
    grid[ev.at.getDay()][Math.floor(ev.at.getHours() / 6)]++;
  });
  sh.getRange(1, 4, 1, 5).setValues([["曜日"].concat(zones)]).setFontWeight("bold");
  const order = [1, 2, 3, 4, 5, 6, 0]; // 月曜始まりで表示
  const wdRows = order.map((wd) => [WEEKDAYS_JP[wd]].concat(grid[wd]));
  sh.getRange(2, 4, 7, 5).setValues(wdRows);

  if (stools.length === 0) return;

  sh.insertChart(
    sh.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(sh.getRange(1, 1, 25, 2))
      .setPosition(11, 4, 0, 0)
      .setOption("title", "排便の発生時刻（時間帯別）")
      .setOption("legend", { position: "none" })
      .setOption("height", 300).setOption("width", 640)
      .build()
  );
}

/* =========================================================
   共通の小道具
   ========================================================= */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function mustSheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error("シート「" + name + "」がありません。初期セットアップを実行してください。");
  return sh;
}
// 分析用シートを空にして作り直す（グラフ・条件付き書式も消す）
function resetSheet_(name) {
  const sh = getOrCreateSheet_(name);
  sh.clear();
  sh.setConditionalFormatRules([]);
  sh.getCharts().forEach((c) => sh.removeChart(c));
  return sh;
}
function fmtDate_(d) {
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), TZ, "yyyy-MM-dd");
}
function parseDate_(key) {
  const p = String(key).split("-");
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}
function startOfDay_(d) {
  return parseDate_(fmtDate_(d));
}
function mondayOf_(d) {
  const x = startOfDay_(d);
  const diff = (x.getDay() + 6) % 7; // 月曜=0になるように
  x.setDate(x.getDate() - diff);
  return x;
}
function numOrNull_(v) {
  return v === "" || v == null ? null : Number(v);
}
function orBlank_(v) {
  return v == null ? "" : v;
}
