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

const EVENT_HEADERS = ["記録日時", "種別", "発生時刻", "ブリストルスケール", "腹痛", "残便感", "薬名", "メモ"];
const DAILY_HEADERS = ["日付", "睡眠時間", "食事タグ", "食事メモ", "水分量", "運動", "気分", "ストレス", "メンタルメモ", "メモ"];

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
    sh.getRange(2, 11, 4, 2).setValues([
      ["便秘警戒日数", 3],
      ["下痢判定ブリストル", 6],
      ["下痢判定回数", 3],
      ["次回通院日", ""],
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
    .addSeparator()
    .addItem("初期セットアップ（最初に1回）", "initialSetup")
    .addToUi();
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
    if (body.type === "event") {
      appendEvent_(body.data);
    } else if (body.type === "daily") {
      upsertDaily_(body.data);
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
function appendEvent_(d) {
  const sh = mustSheet_(SHEET_EVENT);
  sh.appendRow([
    new Date(),                       // 記録日時（サーバー側で自動付与）
    d.kind || "",
    d.occurredAt ? new Date(d.occurredAt) : new Date(),
    d.bristol || "",
    d.pain || "",
    d.zanben || "",
    d.med || "",
    d.memo || "",
  ]);
}

// デイリーログを追記（同じ日付が既にあれば上書き）
function upsertDaily_(d) {
  const sh = mustSheet_(SHEET_DAILY);
  const row = [
    d.date || fmtDate_(new Date()),
    d.sleep === "" || d.sleep == null ? "" : Number(d.sleep),
    (d.tags || []).join(","),
    d.mealMemo || "",
    d.water === "" || d.water == null ? "" : Number(d.water),
    d.exercise || "",
    d.mood || "",
    d.stress || "",
    d.mentalMemo || "",
    d.memo || "",
  ];
  const last = sh.getLastRow();
  if (last >= 2) {
    const dates = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      if (fmtDate_(dates[i][0]) === String(d.date)) {
        sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
        return;
      }
    }
  }
  sh.appendRow(row);
}

/* =========================================================
   doGet : 設定マスタ＋直近7日の簡易集計を返す
   ========================================================= */
function doGet(e) {
  try {
    const cfg = getSettings_();
    return json_({
      ok: true,
      settings: {
        medicines: cfg.medicines,
        tags: cfg.tags,
        alertDays: cfg.alertDays,
      },
      summary: computeSummary_(cfg),
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
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
    alertDays: Number(map["便秘警戒日数"]) || 3,
    diarrheaBristol: Number(map["下痢判定ブリストル"]) || 6,
    diarrheaCount: Number(map["下痢判定回数"]) || 3,
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
   生ログの読み込み（見出し名で列を探すので列追加に強い）
   ========================================================= */
function readEvents_() {
  const sh = mustSheet_(SHEET_EVENT);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0];
  const iKind = head.indexOf("種別"), iAt = head.indexOf("発生時刻"),
        iBri = head.indexOf("ブリストルスケール"), iPain = head.indexOf("腹痛"),
        iZan = head.indexOf("残便感"), iMed = head.indexOf("薬名");
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

  writeStateSheet_(states);
  buildWeekly_(states, events);
  buildTriggerAnalysis_(states, daily);
  buildCalendar_(states);
  buildCycle_(states);
  buildMental_(states);
  buildRhythm_(events);
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
                "ステート", "便秘連続日数", "サイクル番号", "サイクル内腹痛までの便秘日数", "気分", "ストレス", "睡眠時間"];
  const rows = states.map((s) => [
    s.date, s.weekday, s.stoolCount, orBlank_(s.avgBristol), orBlank_(s.maxBristol),
    s.pain ? "あり" : "", s.medCount, s.state, s.streak, orBlank_(s.cycle), orBlank_(s.painAfterDays),
    orBlank_(s.mood), orBlank_(s.stress), orBlank_(s.sleep),
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
