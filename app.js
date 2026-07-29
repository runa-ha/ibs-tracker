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

// ---------- 状態 ----------
let settings = { medicines: [], tags: [], alertDays: 3 }; // GASから取得（取得前はデフォルト）
let selectedBristol = null;
let selectedPain = "なし";
let selectedZanben = "なし";
let selectedMed = null;
let sleepHours = 7.0;
let waterMl = 0;
let selectedExercise = null;
let selectedMood = null;
let selectedStress = null;
let selectedTags = new Set();
let flushing = false;

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
function todayString() {
  const d = new Date();
  d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0, 10);
}

function gasConfigured() {
  return typeof CONFIG !== "undefined" && CONFIG.GAS_URL && !CONFIG.GAS_URL.includes("XXXXXXXX");
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
  const q = loadQueue();
  q.push(payload);
  saveQueue(q);
  await flushQueue();
  if (loadQueue().length === 0) {
    toast("✅ 記録しました");
    fetchStatus(); // 集計カードを更新
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
      const item = q[0];
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
    const res = await fetch(CONFIG.GAS_URL);
    const json = await res.json();
    if (!json.ok) throw new Error(json.error);
    settings = json.settings;
    localStorage.setItem(LS_SETTINGS, JSON.stringify(settings));
    renderMasters();
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

  // 食事タグ
  const tagWrap = $("tag-chips");
  tagWrap.innerHTML = "";
  if (settings.tags.length === 0) {
    tagWrap.innerHTML = '<div class="loading-note">タグが未取得です</div>';
  }
  settings.tags.forEach((name) => {
    const c = document.createElement("button");
    c.className = "chip";
    c.textContent = name;
    if (selectedTags.has(name)) c.classList.add("selected");
    c.onclick = () => {
      if (selectedTags.has(name)) selectedTags.delete(name);
      else selectedTags.add(name);
      c.classList.toggle("selected");
    };
    tagWrap.appendChild(c);
  });
}

// 固定選択肢のボタン群を作る（汎用）
function buildSeg(containerId, options, onSelect, labels) {
  const wrap = $(containerId);
  wrap.innerHTML = "";
  options.forEach((opt, i) => {
    const b = document.createElement("button");
    b.textContent = labels ? labels[i] : opt;
    b.onclick = () => {
      wrap.querySelectorAll("button").forEach((x) => x.classList.toggle("selected", x === b));
      onSelect(opt);
    };
    wrap.appendChild(b);
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

function submitStool() {
  if (!selectedBristol) { toast("ブリストルスケールを選んでください"); return; }
  submitPayload({
    type: "event",
    data: {
      kind: "排便",
      occurredAt: new Date($("stool-time").value || Date.now()).toISOString(),
      bristol: selectedBristol,
      pain: selectedPain,
      zanben: selectedZanben,
      memo: $("stool-memo").value.trim(),
    },
  });
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
  submitPayload({
    type: "event",
    data: {
      kind: "服薬",
      occurredAt: new Date($("med-time").value || Date.now()).toISOString(),
      med: selectedMed,
      memo: $("med-memo").value.trim(),
    },
  });
  selectedMed = null;
  $("med-memo").value = "";
  $("med-detail").classList.add("hidden");
  document.querySelectorAll("#med-buttons .med-btn").forEach((x) => x.classList.remove("selected"));
  $("med-time").value = nowLocalString();
}

function submitDaily() {
  if (selectedMood == null && selectedStress == null && selectedTags.size === 0 && selectedExercise == null) {
    toast("何か1つ以上入力してください");
    return;
  }
  submitPayload({
    type: "daily",
    data: {
      date: $("daily-date").value || todayString(),
      sleep: sleepHours,
      tags: Array.from(selectedTags),
      mealMemo: $("meal-memo").value.trim(),
      water: waterMl,
      exercise: selectedExercise || "",
      mood: selectedMood || "",
      stress: selectedStress || "",
      mentalMemo: $("mental-memo").value.trim(),
      memo: $("daily-memo").value.trim(),
    },
  });
}

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
  $("daily-date").value = todayString();
  $("sleep-minus").onclick = () => { sleepHours = Math.max(0, sleepHours - 0.5); $("sleep-value").textContent = sleepHours.toFixed(1); };
  $("sleep-plus").onclick = () => { sleepHours = Math.min(16, sleepHours + 0.5); $("sleep-value").textContent = sleepHours.toFixed(1); };
  $("water-minus").onclick = () => { waterMl = Math.max(0, waterMl - 250); $("water-value").textContent = waterMl; };
  $("water-plus").onclick = () => { waterMl += 250; $("water-value").textContent = waterMl; };
  buildSeg("exercise-seg", EXERCISE_OPTIONS, (v) => (selectedExercise = v));
  buildSeg("mood-seg", [1, 2, 3, 4, 5], (v) => (selectedMood = v), MOOD_EMOJI);
  buildSeg("stress-seg", [1, 2, 3, 4, 5], (v) => (selectedStress = v));
  $("submit-daily").onclick = submitDaily;

  // 前回取得した設定マスタがあれば先に表示（オフライン起動対策）
  try {
    const cached = JSON.parse(localStorage.getItem(LS_SETTINGS));
    if (cached) settings = cached;
  } catch (e) { /* 無視 */ }
  renderMasters();

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
