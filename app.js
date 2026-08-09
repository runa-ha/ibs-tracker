/* =========================================================
   IBSログ アプリ本体
   - 起動時にGAS（doGet）から設定マスタと簡易集計を取得
   - 送信はキュー（localStorage）に積んでから順に送る
     → 電波がなくても記録でき、復帰時に自動再送される
   ========================================================= */

"use strict";

// ---------- 定数 ----------
const LS_QUEUE = "ibs_queue";       // 未送信データの保存キー
const LS_SETTINGS = "ibs_settings"; // 設定マスタのキャッシュキー
const LS_AUTH = "ibs_auth";         // 合言葉の保存キー（この端末の中にだけ保存される）
const LS_DAILY = "ibs_daily_draft"; // 今日タブの下書き（日付が変わると自動リセット）

const BRISTOL = [
  { n: 1, name: "コロコロ便", desc: "硬くて木の実のような便" },
  { n: 2, name: "硬い便", desc: "ソーセージ状だが硬い" },
  { n: 3, name: "やや硬い便", desc: "表面にひび割れがある" },
  { n: 4, name: "普通便", desc: "なめらかなバナナ状" },
  { n: 5, name: "やや軟らかい便", desc: "水分が多く軟らかい半固形" },
  { n: 6, name: "泥状便", desc: "形がくずれた泥のような便" },
  { n: 7, name: "水様便", desc: "固形物がない水のような便" },
];
const BRISTOL_COLORS = ["#92400e", "#b45309", "#d97706", "#0e9488", "#f59e0b", "#f97316", "#dc2626"];

const PAIN_OPTIONS = ["なし", "軽い", "中", "強い"];
const ZANBEN_OPTIONS = ["なし", "あり"];
const EXERCISE_OPTIONS = ["なし", "軽い", "しっかり"];
const MOOD_EMOJI = ["😞", "😕", "😐", "🙂", "😄"];
const FULLNESS_OPTIONS = ["食べすぎて気持ち悪い", "胃もたれ", "ちょうどよく満腹", "八分目", "足りない"];

// 服薬チェックのタイミング表示順（GAS側のTIMING_ORDERと合わせる）
const TIMING_ORDER = ["朝食後", "昼食後", "夕食前", "夕食後", "外用"];

// ---------- 状態 ----------
let settings = { medicines: [], tags: [], obsTags: [], alertDays: 3 }; // GASから取得（取得前はデフォルト）
let prescriptions = []; // 現行処方（「処方」シート由来）
let todayMeds = [];     // 今日すでに記録した服薬 [{med, timing}]
let todayEvents = [];   // 今日の記録一覧（履歴表示用。サーバーから取得）
let selectedBristol = null;
let selectedPain = "なし";
let selectedZanben = "なし";
let selectedMed = null;
let sleepHours = 7.0;
let waterMl = 0;
let selectedExercise = null;
let selectedMood = null;
let selectedStress = null;
let selectedFullness = null;
let periodStart = false; // 今日が生理開始日か
const medTapGuard = {};  // 服薬チェックの二度押しガード
// 食事タグは朝・昼・夕で別々に持つ
let mealTags = { morning: new Set(), noon: new Set(), evening: new Set() };
let selectedObsTags = new Set();
let flushing = false;
let dailyTimer = null; // 自動保存の待ち時間管理

// ---------- 小道具 ----------
const $ = (id) => document.getElementById(id);

function toast(msg, ms = 2500) {
  const el = $("toast");
  el.textContent = msg;
  el.classList.remove("hidden");
  clearTimeout(el._timer);
  el._timer = setTimeout(() => el.classList.add("hidden"), ms);
}

// datetime-local用に「今」をローカル時刻の文字列にする（例: 2026-07-30T21:05）
function nowLocalString() {
  const d = new Date();
  d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0, 16);
}
function nowTimeString() {
  const d = new Date();
  return ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2);
}
function todayString() {
  const d = new Date();
  d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0, 10);
}

function gasConfigured() {
  return typeof CONFIG !== "undefined" && CONFIG.GAS_URL && !CONFIG.GAS_URL.includes("XXXXXXXX");
}

// ---------- 合言葉 ----------
function getAuth() {
  return localStorage.getItem(LS_AUTH) || "";
}
function showAuthPanel(msg) {
  if (msg) $("auth-msg").textContent = msg;
  $("auth-panel").classList.remove("hidden");
}
function hideAuthPanel() {
  $("auth-panel").classList.add("hidden");
}

/* =========================================================
   今日タブの下書き＆自動保存
   - 入力のたびに端末へ下書き保存（アプリを閉じても丸一日残る）
   - 数秒後にスプレッドシートの「今日の行」へ自動で上書き保存
   - 日付が変わると入力欄は自動で新しい日にリセットされる
   ========================================================= */
function saveDraft() {
  const draft = {
    date: todayString(),
    sleep: sleepHours,
    water: waterMl,
    exercise: selectedExercise,
    mood: selectedMood,
    stress: selectedStress,
    fullness: selectedFullness,
    periodStart: periodStart,
    obsTags: Array.from(selectedObsTags),
    meals: {
      morning: { tags: Array.from(mealTags.morning), memo: $("meal-memo-morning").value.trim() },
      noon: { tags: Array.from(mealTags.noon), memo: $("meal-memo-noon").value.trim() },
      evening: { tags: Array.from(mealTags.evening), memo: $("meal-memo-evening").value.trim() },
    },
    mentalMemo: $("mental-memo").value.trim(),
    memo: $("daily-memo").value.trim(),
  };
  localStorage.setItem(LS_DAILY, JSON.stringify(draft));
  scheduleDailySend();
}

function loadDraft() {
  try {
    const d = JSON.parse(localStorage.getItem(LS_DAILY));
    if (d && d.date === todayString()) return d; // 今日の下書きだけ有効
  } catch (e) { /* 無視 */ }
  return null;
}

// 連続入力中は待って、手が止まったらまとめて送信する
function scheduleDailySend() {
  updateSaveStatus("✏️ 入力中…");
  clearTimeout(dailyTimer);
  dailyTimer = setTimeout(sendDaily, 2500);
}

function sendDaily() {
  const draft = loadDraft();
  if (!draft) return; // 日付が変わっていたら送らない
  submitPayload({
    type: "daily",
    data: {
      date: draft.date,
      sleep: draft.sleep,
      meals: draft.meals,
      water: draft.water,
      exercise: draft.exercise || "",
      mood: draft.mood || "",
      stress: draft.stress || "",
      fullness: draft.fullness || "",
      periodStart: !!draft.periodStart,
      obsTags: draft.obsTags || [],
      mentalMemo: draft.mentalMemo || "",
      memo: draft.memo || "",
    },
  });
}

function updateSaveStatus(text) {
  const el = $("daily-save-status");
  if (el) el.textContent = text;
}

// ---------- 送信キュー ----------
function loadQueue() {
  try { return JSON.parse(localStorage.getItem(LS_QUEUE)) || []; } catch (e) { return []; }
}
function saveQueue(q) {
  localStorage.setItem(LS_QUEUE, JSON.stringify(q));
  updateQueueBadge();
}
function updateQueueBadge() {
  const n = loadQueue().length;
  const badge = $("queue-badge");
  badge.textContent = `未送信 ${n}件`;
  badge.classList.toggle("hidden", n === 0);
}

// 記録をキューに積んでから送信を試みる
async function submitPayload(payload) {
  payload.clientId = Date.now() + "-" + Math.random().toString(36).slice(2); // 重複判定用ID
  let q = loadQueue();
  // デイリーの自動保存は「同じ日の送信待ち」を置き換える（重複させない）
  if (payload.type === "daily") {
    q = q.filter((i) => !(i.type === "daily" && i.data && i.data.date === payload.data.date));
  }
  q.push(payload);
  saveQueue(q);
  await flushQueue();

  if (payload.type === "daily") {
    // 今日タブの自動保存は控えめに状態表示だけ更新（トーストは出さない）
    const pending = loadQueue().some((i) => i.type === "daily");
    updateSaveStatus(pending ? "📡 未送信（電波が戻れば自動保存されます）" : "☁️ 保存済み " + nowTimeString());
    return;
  }
  if (loadQueue().length === 0) {
    toast("✅ 記録しました");
    fetchStatus(); // 集計カード・履歴を更新
  } else {
    toast("📡 電波がないため保存しました（復帰後に自動送信）", 3500);
  }
}

// キューの中身を先頭から順に送る
async function flushQueue() {
  if (flushing || !gasConfigured()) return;
  flushing = true;
  try {
    let q = loadQueue();
    while (q.length > 0) {
      // 送信の瞬間に合言葉を添える（合言葉入力前に記録した分も後から送れるように）
      const item = Object.assign({}, q[0], { auth: getAuth() });
      let json;
      try {
        // Content-Typeを指定しない文字列POST = プリフライトなしで送れる
        const res = await fetch(CONFIG.GAS_URL, { method: "POST", body: JSON.stringify(item) });
        json = await res.json();
      } catch (e) {
        break; // 通信エラー → キューに残して後で再送
      }
      if (json && json.ok) {
        q.shift(); // 成功 → キューから削除
        saveQueue(q);
      } else if (json && json.authError) {
        // 合言葉が未入力/不一致 → 記録はキューに残したまま入力を促す
        showAuthPanel("⚠️ 合言葉が一致しません。入力し直してください（記録は消えていません）");
        break;
      } else {
        // サーバーがエラーを返した（データ不備など）→ 再送しても直らないので捨てる
        console.error("サーバーエラー:", json && json.error);
        toast("⚠️ サーバーでエラー: " + ((json && json.error) || "不明"), 4000);
        q.shift();
        saveQueue(q);
      }
    }
  } finally {
    flushing = false;
  }
}

// ---------- GASから設定と集計を取得 ----------
async function fetchStatus() {
  if (!gasConfigured()) {
    $("setup-banner").classList.remove("hidden");
    $("streak-note").textContent = "config.js の設定待ちです";
    return;
  }
  try {
    // 集計・設定の取得も合言葉つきのPOSTで行う（合言葉なしでは何も見られない）
    const res = await fetch(CONFIG.GAS_URL, {
      method: "POST",
      body: JSON.stringify({ type: "status", auth: getAuth() }),
    });
    const json = await res.json();
    if (json.authError) {
      showAuthPanel(getAuth()
        ? "⚠️ 合言葉が一致しません。入力し直してください"
        : "スプレッドシートの「設定」シートに書いた合言葉を入力してください。この端末に保存され、次回からは入力不要です。");
      $("streak-note").textContent = "合言葉の入力待ちです";
      return;
    }
    if (!json.ok) throw new Error(json.error);
    hideAuthPanel();
    settings = json.settings;
    prescriptions = json.prescriptions || [];
    todayMeds = json.todayMeds || [];
    todayEvents = json.todayEvents || [];
    // 処方もキャッシュしておく（オフライン起動時もチェックリストを出すため）
    localStorage.setItem(LS_SETTINGS, JSON.stringify({ settings: settings, prescriptions: prescriptions }));
    renderMasters();
    renderChecklist();
    renderHistory();
    renderSummary(json.summary);
  } catch (e) {
    console.warn("集計の取得に失敗（オフライン？）:", e);
    $("streak-note").textContent = "オフライン（前回の情報を表示）";
  }
}

function renderSummary(s) {
  if (!s) return;
  $("streak-days").textContent = s.constipationDays;
  $("stat-stool").textContent = s.last7.stoolCount;
  $("stat-bristol").textContent = s.last7.avgBristol != null ? s.last7.avgBristol : "–";
  $("stat-pain").textContent = s.last7.painDays;
  $("stat-med").textContent = s.last7.medCount;

  const card = $("card-streak");
  const alert = s.constipationDays >= settings.alertDays;
  card.classList.toggle("alert", alert);
  $("streak-note").textContent = alert
    ? `⚠️ 警戒ライン（${settings.alertDays}日）を超えています`
    : `警戒ラインは ${settings.alertDays}日`;
}

// ---------- 画面の組み立て ----------

// 薬ボタン・食事タグなど、設定シート由来の部品を作り直す
function renderMasters() {
  // 薬ボタン
  const medWrap = $("med-buttons");
  medWrap.innerHTML = "";
  if (settings.medicines.length === 0) {
    medWrap.innerHTML = '<div class="loading-note">薬リストが未取得です（電波のある場所で開き直してください）</div>';
  }
  settings.medicines.forEach((name) => {
    const b = document.createElement("button");
    b.className = "med-btn";
    b.textContent = name;
    b.onclick = () => {
      selectedMed = name;
      medWrap.querySelectorAll(".med-btn").forEach((x) => x.classList.toggle("selected", x === b));
      $("med-detail").classList.remove("hidden");
    };
    medWrap.appendChild(b);
  });

  // 食事タグ（朝・昼・夕それぞれ）
  buildChips_("tag-chips-morning", settings.tags, mealTags.morning, "タグが未取得です", saveDraft);
  buildChips_("tag-chips-noon", settings.tags, mealTags.noon, "タグが未取得です", saveDraft);
  buildChips_("tag-chips-evening", settings.tags, mealTags.evening, "タグが未取得です", saveDraft);
  // 観察タグ（副作用チェック用）
  buildChips_("obs-chips", settings.obsTags || [], selectedObsTags, "アップデート適用後に表示されます", saveDraft);
}

// タグ選択チップを作る共通処理
function buildChips_(containerId, names, selectedSet, emptyMsg, onToggle) {
  const wrap = $(containerId);
  wrap.innerHTML = "";
  if (!names.length) {
    wrap.innerHTML = '<div class="loading-note">' + emptyMsg + "</div>";
    return;
  }
  names.forEach((name) => {
    const c = document.createElement("button");
    c.className = "chip";
    c.textContent = name;
    if (selectedSet.has(name)) c.classList.add("selected");
    c.onclick = () => {
      if (selectedSet.has(name)) selectedSet.delete(name);
      else selectedSet.add(name);
      c.classList.toggle("selected");
      if (onToggle) onToggle();
    };
    wrap.appendChild(c);
  });
}

/* =========================================================
   今日の記録（履歴表示）
   サーバーに保存済みの分＋送信待ちの分をまとめて時刻順に表示
   ========================================================= */
function renderHistory() {
  const wrap = $("today-history");
  if (!wrap) return;
  wrap.innerHTML = "";

  const items = todayEvents.slice();
  // 送信待ちキューの中の「今日のイベント」も表示に含める
  loadQueue().forEach((qi) => {
    if (qi.type !== "event" || !qi.data) return;
    const at = new Date(qi.data.occurredAt || Date.now());
    const d = new Date(at); d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
    if (d.toISOString().slice(0, 10) !== todayString()) return;
    items.push({
      time: ("0" + at.getHours()).slice(-2) + ":" + ("0" + at.getMinutes()).slice(-2),
      kind: qi.data.kind, bristol: qi.data.bristol, pain: qi.data.pain,
      med: qi.data.med, timing: qi.data.timing || "", memo: qi.data.memo || "",
      pending: true,
    });
  });

  if (!items.length) {
    wrap.innerHTML = '<div class="loading-note">今日はまだ記録がありません</div>';
    return;
  }
  items.sort((a, b) => (a.time < b.time ? -1 : 1));
  items.forEach((ev) => {
    const div = document.createElement("div");
    div.className = "history-item";
    let text;
    if (ev.kind === "排便") {
      text = "💩 ブリストル" + (ev.bristol || "?") + "・腹痛" + (ev.pain || "なし");
    } else if (ev.kind === "服薬") {
      text = "💊 " + ev.med + (ev.timing ? "（" + ev.timing + "）" : "");
    } else {
      text = ev.kind;
    }
    div.innerHTML =
      '<span class="history-time">' + ev.time + "</span>" +
      '<span class="history-text">' + text +
      (ev.memo ? '<small>' + ev.memo + "</small>" : "") + "</span>" +
      (ev.pending ? '<span class="history-pending">⏳送信待ち</span>' : "");
    wrap.appendChild(div);
  });
}

/* =========================================================
   今日の服薬チェック（「処方」シートの現行処方から生成）
   タップで記録、チェック済みをタップで取り消し
   ========================================================= */
function isMedChecked(med, timing) {
  return todayMeds.some((m) => m.med === med && m.timing === timing);
}

function renderChecklist() {
  const wrap = $("med-checklist");
  wrap.innerHTML = "";
  if (!prescriptions.length) {
    wrap.innerHTML = '<div class="loading-note">現行処方が未設定です（スプレッドシートの「処方」シートに登録すると表示されます）</div>';
    return;
  }
  const hour = new Date().getHours();
  TIMING_ORDER.forEach((slot) => {
    const meds = prescriptions.filter((p) => (p.timings || []).indexOf(slot) >= 0);
    if (!meds.length) return;

    const slotDiv = document.createElement("div");
    slotDiv.className = "slot";
    // 夕食前の薬（グーフィス）は夕方以降・未チェックのとき強調表示
    const urgent = slot === "夕食前" && hour >= 16 && meds.some((p) => !isMedChecked(p.name, slot));
    const label = document.createElement("div");
    label.className = "slot-label" + (urgent ? " urgent" : "");
    label.textContent = slot + (urgent ? " ⚠ そろそろ時間です（食事の前に！）" : "");
    slotDiv.appendChild(label);

    meds.forEach((p) => {
      const checked = isMedChecked(p.name, slot);
      const b = document.createElement("button");
      b.className = "check-btn" + (checked ? " checked" : "");
      // 残り日数（日数指定の薬）or 数量（外用など）を小さく表示
      let sub = p.dose || "";
      if (p.remainingDays != null) sub += "　あと" + p.remainingDays + "日（" + (p.endDate || "") + "まで）";
      else if (p.qty) sub += "　" + p.qty;
      if (p.remainingDays != null && p.remainingDays <= 5) b.classList.add("low");
      b.innerHTML =
        '<span class="check-mark">' + (checked ? "✓" : "") + "</span>" +
        '<span class="check-body"><b>' + p.name + "</b><small>" + sub + "</small>" +
        (p.note ? '<small class="med-note">' + p.note + "</small>" : "") +
        "</span>";
      b.onclick = () => toggleMedCheck(p, slot);
      slotDiv.appendChild(b);
    });
    wrap.appendChild(slotDiv);
  });
}

function toggleMedCheck(p, slot) {
  // 二度押しガード: 同じボタンの連続タップは2秒間無視する
  const guardKey = p.name + "|" + slot;
  const now = Date.now();
  if (medTapGuard[guardKey] && now - medTapGuard[guardKey] < 2000) return;
  medTapGuard[guardKey] = now;

  if (isMedChecked(p.name, slot)) {
    // チェック済み → 取り消し（今日の該当行をシートから削除。オンライン時のみ）
    if (!confirm("「" + p.name + "（" + slot + "）」の今日の記録を取り消しますか？")) return;
    fetch(CONFIG.GAS_URL, {
      method: "POST",
      body: JSON.stringify({ type: "deleteMedToday", auth: getAuth(), data: { med: p.name, timing: slot } }),
    })
      .then((r) => r.json())
      .then((j) => {
        if (j.ok) {
          todayMeds = todayMeds.filter((m) => !(m.med === p.name && m.timing === slot));
          renderChecklist();
          toast("取り消しました");
        } else {
          toast("⚠️ 取り消せませんでした: " + (j.error || "不明"));
        }
      })
      .catch(() => toast("📡 オフラインのため取り消せません"));
    return;
  }
  // 未チェック → 服薬として記録（オフラインでもキューに入る）
  todayMeds.push({ med: p.name, timing: slot });
  renderChecklist();
  submitPayload({
    type: "event",
    data: { kind: "服薬", med: p.name, timing: slot, occurredAt: new Date().toISOString() },
  });
  renderHistory(); // 「今日の記録」にすぐ反映
}

// 固定選択肢のボタン群を作る（汎用）
function buildSeg(containerId, options, onSelect, labels) {
  const wrap = $(containerId);
  wrap.innerHTML = "";
  options.forEach((opt, i) => {
    const b = document.createElement("button");
    b.textContent = labels ? labels[i] : opt;
    b.dataset.val = String(opt); // 復元用（絵文字ボタンでも値で選択できるように）
    b.onclick = () => {
      wrap.querySelectorAll("button").forEach((x) => x.classList.toggle("selected", x === b));
      onSelect(opt);
    };
    wrap.appendChild(b);
  });
}

// 保存されている値からボタンの選択状態を復元する
function selectByVal(containerId, val) {
  if (val == null || val === "") return;
  $(containerId).querySelectorAll("button").forEach((b) => {
    b.classList.toggle("selected", b.dataset.val === String(val));
  });
}

function buildBristolButtons() {
  const wrap = $("bristol-buttons");
  wrap.innerHTML = "";
  BRISTOL.forEach((item) => {
    const b = document.createElement("button");
    b.className = "bristol-btn";
    b.innerHTML =
      `<span class="num" style="background:${BRISTOL_COLORS[item.n - 1]}">${item.n}</span>` +
      `<span class="desc"><b>${item.name}</b><small>${item.desc}</small></span>`;
    b.onclick = () => {
      selectedBristol = item.n;
      wrap.querySelectorAll(".bristol-btn").forEach((x) => x.classList.toggle("selected", x === b));
      $("stool-detail").classList.remove("hidden");
    };
    wrap.appendChild(b);
  });
}

// ---------- 送信処理 ----------

// ボタンを数秒間押せなくする（連打による二重送信の防止）
function lockBtn(id, ms) {
  const el = $(id);
  if (!el) return;
  el.disabled = true;
  setTimeout(() => { el.disabled = false; }, ms || 3000);
}

// 直近10分以内に同じ内容の記録がないかを調べる（送信済み＋送信待ちの両方）
function hasRecentSame(kind, med, occurredAtIso) {
  const t = new Date(occurredAtIso).getTime();
  const win = 10 * 60 * 1000;
  const p = todayString().split("-");
  for (const ev of todayEvents) {
    if (ev.kind !== kind) continue;
    if (kind === "服薬" && ev.med !== med) continue;
    const hm = String(ev.time || "").split(":");
    const d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), Number(hm[0]) || 0, Number(hm[1]) || 0);
    if (Math.abs(d.getTime() - t) <= win) return true;
  }
  for (const qi of loadQueue()) {
    if (qi.type !== "event" || !qi.data || qi.data.kind !== kind) continue;
    if (kind === "服薬" && qi.data.med !== med) continue;
    if (Math.abs(new Date(qi.data.occurredAt).getTime() - t) <= win) return true;
  }
  return false;
}

function submitStool() {
  if (!selectedBristol) { toast("ブリストルスケールを選んでください"); return; }
  const iso = new Date($("stool-time").value || Date.now()).toISOString();
  if (hasRecentSame("排便", null, iso) &&
      !confirm("⚠️ 10分以内に排便の記録がすでにあります。\n重複ではありませんか？\n\nOK＝このまま記録する ／ キャンセル＝やめる")) {
    return;
  }
  lockBtn("submit-stool");
  submitPayload({
    type: "event",
    data: {
      kind: "排便",
      occurredAt: iso,
      bristol: selectedBristol,
      pain: selectedPain,
      zanben: selectedZanben,
      memo: $("stool-memo").value.trim(),
    },
  });
  renderHistory(); // 「今日の記録」にすぐ反映
  // 入力をリセット
  selectedBristol = null;
  selectedPain = "なし";
  selectedZanben = "なし";
  $("stool-memo").value = "";
  $("stool-detail").classList.add("hidden");
  document.querySelectorAll("#bristol-buttons .bristol-btn").forEach((x) => x.classList.remove("selected"));
  buildSeg("pain-seg", PAIN_OPTIONS, (v) => (selectedPain = v));
  selectDefault("pain-seg", "なし");
  buildSeg("zanben-seg", ZANBEN_OPTIONS, (v) => (selectedZanben = v));
  selectDefault("zanben-seg", "なし");
  $("stool-time").value = nowLocalString();
}

function submitMed() {
  if (!selectedMed) { toast("薬を選んでください"); return; }
  const iso = new Date($("med-time").value || Date.now()).toISOString();
  if (hasRecentSame("服薬", selectedMed, iso) &&
      !confirm("⚠️ 10分以内に「" + selectedMed + "」の記録がすでにあります。\n重複ではありませんか？\n\nOK＝このまま記録する ／ キャンセル＝やめる")) {
    return;
  }
  lockBtn("submit-med");
  submitPayload({
    type: "event",
    data: {
      kind: "服薬",
      occurredAt: iso,
      med: selectedMed,
      memo: $("med-memo").value.trim(),
    },
  });
  renderHistory(); // 「今日の記録」にすぐ反映
  selectedMed = null;
  $("med-memo").value = "";
  $("med-detail").classList.add("hidden");
  document.querySelectorAll("#med-buttons .med-btn").forEach((x) => x.classList.remove("selected"));
  $("med-time").value = nowLocalString();
}

// ※「今日のまとめ」は保存ボタン方式をやめ、入力のたびに自動保存する
//   （saveDraft → scheduleDailySend → sendDaily の流れ）

// 指定した値のボタンを選択状態にする（デフォルト値の見た目用）
function selectDefault(containerId, value) {
  $(containerId).querySelectorAll("button").forEach((b) => {
    b.classList.toggle("selected", b.textContent === String(value));
  });
}

// ---------- 初期化 ----------
function init() {
  // サービスワーカー登録（オフラインでも画面が開くようにする）
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("service-worker.js").catch(() => {});
  }

  // タブ切り替え
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.onclick = () => {
      document.querySelectorAll(".tab-btn").forEach((x) => x.classList.toggle("active", x === btn));
      document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("active", t.id === btn.dataset.tab));
    };
  });

  // 記録タブの部品
  buildBristolButtons();
  buildSeg("pain-seg", PAIN_OPTIONS, (v) => (selectedPain = v));
  selectDefault("pain-seg", "なし");
  buildSeg("zanben-seg", ZANBEN_OPTIONS, (v) => (selectedZanben = v));
  selectDefault("zanben-seg", "なし");
  $("stool-time").value = nowLocalString();
  $("med-time").value = nowLocalString();
  $("stool-time-now").onclick = () => ($("stool-time").value = nowLocalString());
  $("med-time-now").onclick = () => ($("med-time").value = nowLocalString());
  $("submit-stool").onclick = submitStool;
  $("submit-med").onclick = submitMed;

  // 今日タブの部品
  const d = new Date();
  $("daily-date-label").textContent = (d.getMonth() + 1) + "/" + d.getDate() + "(" + ["日", "月", "火", "水", "木", "金", "土"][d.getDay()] + ")";

  // 今日の下書きがあれば復元（アプリを閉じても丸一日残る）
  const draft = loadDraft();
  if (draft) {
    if (draft.sleep != null) sleepHours = draft.sleep;
    waterMl = draft.water || 0;
    selectedExercise = draft.exercise || null;
    selectedMood = draft.mood || null;
    selectedStress = draft.stress || null;
    selectedFullness = draft.fullness || null;
    periodStart = !!draft.periodStart;
    selectedObsTags = new Set(draft.obsTags || []);
    const meals = draft.meals || {};
    ["morning", "noon", "evening"].forEach((k) => {
      mealTags[k] = new Set((meals[k] && meals[k].tags) || []);
      $("meal-memo-" + k).value = (meals[k] && meals[k].memo) || "";
    });
    $("mental-memo").value = draft.mentalMemo || "";
    $("daily-memo").value = draft.memo || "";
    updateSaveStatus("☁️ 下書きを復元しました（自動保存は有効です）");
  }

  const updateSleepView = () => { $("sleep-value").textContent = sleepHours.toFixed(1); };
  const updateWaterView = () => {
    $("water-value").textContent = waterMl;
    $("water-cups").textContent = waterMl > 0 ? "（コップ約" + Math.round(waterMl / 250) + "杯）" : "";
  };
  updateSleepView();
  updateWaterView();

  $("sleep-minus").onclick = () => { sleepHours = Math.max(0, sleepHours - 0.5); updateSleepView(); saveDraft(); };
  $("sleep-plus").onclick = () => { sleepHours = Math.min(16, sleepHours + 0.5); updateSleepView(); saveDraft(); };
  $("water-minus").onclick = () => { waterMl = Math.max(0, waterMl - 250); updateWaterView(); saveDraft(); };
  $("water-plus").onclick = () => { waterMl += 250; updateWaterView(); saveDraft(); };
  buildSeg("exercise-seg", EXERCISE_OPTIONS, (v) => { selectedExercise = v; saveDraft(); });
  buildSeg("mood-seg", [1, 2, 3, 4, 5], (v) => { selectedMood = v; saveDraft(); }, MOOD_EMOJI);
  buildSeg("stress-seg", [1, 2, 3, 4, 5], (v) => { selectedStress = v; saveDraft(); });
  buildSeg("fullness-seg", FULLNESS_OPTIONS, (v) => { selectedFullness = v; saveDraft(); });
  selectByVal("exercise-seg", selectedExercise);
  selectByVal("mood-seg", selectedMood);
  selectByVal("stress-seg", selectedStress);
  selectByVal("fullness-seg", selectedFullness);

  // 生理開始日のトグル
  const periodChip = $("period-chip");
  periodChip.classList.toggle("selected", periodStart);
  periodChip.onclick = () => {
    periodStart = !periodStart;
    periodChip.classList.toggle("selected", periodStart);
    toast(periodStart ? "🩸 今日を生理開始日として記録します" : "生理開始の記録を取り消しました");
    saveDraft();
  };

  // メモ類は入力が止まったら自動保存
  ["meal-memo-morning", "meal-memo-noon", "meal-memo-evening", "mental-memo", "daily-memo"].forEach((id) => {
    $(id).addEventListener("input", saveDraft);
  });

  // 合言葉の保存ボタン
  $("auth-save").onclick = () => {
    const v = $("auth-input").value.trim();
    if (!v) { toast("合言葉を入力してください"); return; }
    localStorage.setItem(LS_AUTH, v);
    $("auth-input").value = "";
    hideAuthPanel();
    toast("接続を確認しています…");
    fetchStatus();
    flushQueue();
  };

  // 前回取得した設定マスタ・処方があれば先に表示（オフライン起動対策）
  try {
    const cached = JSON.parse(localStorage.getItem(LS_SETTINGS));
    if (cached && cached.settings) {
      settings = cached.settings;
      prescriptions = cached.prescriptions || [];
    } else if (cached) {
      settings = cached; // 旧形式のキャッシュ
    }
  } catch (e) { /* 無視 */ }
  renderMasters();
  renderChecklist();
  renderHistory();

  // オンライン/オフライン表示と自動再送
  const updateNet = () => {
    $("net-status").classList.toggle("hidden", navigator.onLine);
    if (navigator.onLine) flushQueue().then(updateQueueBadge);
  };
  window.addEventListener("online", updateNet);
  window.addEventListener("offline", updateNet);
  updateNet();
  updateQueueBadge();

  // 最新の設定・集計を取得
  fetchStatus();
}

document.addEventListener("DOMContentLoaded", init);
