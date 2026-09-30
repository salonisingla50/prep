'use strict';

/* ---------- Settings ---------- */
const CFG = {
  CODE: '2911',
  START: '2026-09-29',      // Exam 1 day
  EXAM_DAY: '2026-11-29',
  TOTAL: 30,
  DEFAULT_MINUTES: 10,      // per section, if not set in the Exams sheet
  PENDING: 'tspPrep.pending.v1',
  UNLOCK: 'tspPrep.unlocked.v1',
  SYNC_URL: 'https://script.google.com/macros/s/AKfycbwkEMZvGByoay3AelNGvdeDH8N5oU1GfN8lLAzWXMYFCIQdHPDT5C3J0_C7ZHJkWaw_/exec'
};
const L = 'ABCD';

/* ---------- Helpers ---------- */
const app = document.getElementById('app');
const $ = s => document.querySelector(s);
const pad = n => String(n).padStart(2, '0');
const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pct = (c, t) => (t ? Math.round((c / t) * 100) : 0);
const ymd = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const parseYMD = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
const addDays = (s, n) => { const d = parseYMD(s); d.setDate(d.getDate() + n); return ymd(d); };
const daysBetween = (a, b) => Math.round((parseYMD(b) - parseYMD(a)) / 864e5);
const today = () => ymd(new Date());
const fmt = s => parseYMD(s).toLocaleDateString('en-IN', { day: 'numeric', month: 'short' });
const fmtDay = s => parseYMD(s).toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric', month: 'short' });
const dt = ts => new Date(ts).toLocaleString('en-IN', { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' });

const store = {
  attempts: [],
  get() { return this.attempts; },
  set(list) { this.attempts = list; },
  pending() {
    try { return JSON.parse(localStorage.getItem(CFG.PENDING)) || []; }
    catch { return []; }
  },
  setPending(list) { localStorage.setItem(CFG.PENDING, JSON.stringify(list)); },
  async request(url, options) {
    const res = await fetch(url, { cache: 'no-store', ...options });
    if (!res.ok) throw new Error('Shared history is unavailable');
    const data = await res.json();
    if (data.error || !Array.isArray(data.attempts)) throw new Error(data.error || 'Bad shared-history response');
    return data.attempts;
  },
  async sync() {
    if (!CFG.SYNC_URL) { this.set([]); syncMessage = 'Shared history is not configured.'; return false; }
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const attempts = await this.request(`${CFG.SYNC_URL}?code=${encodeURIComponent(CFG.CODE)}&v=${Date.now()}`);
        this.set(attempts);
        this.setPending(this.pending().filter(a => !attempts.some(b => String(b.id) === String(a.id))));
        syncMessage = '';
        return true;
      } catch (err) {
        if (attempt === 0 && err.message !== 'Unauthorised') {
          await new Promise(resolve => setTimeout(resolve, 700));
          continue;
        }
        console.warn('Shared history sync failed', err);
      }
    }
    this.set([]);
    syncMessage = 'Could not load shared history.';
    return false;
  },
  async add(a) {
    try {
      const attempts = await this.request(CFG.SYNC_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'save', code: CFG.CODE, attempt: a })
      });
      this.set(attempts);
      this.setPending(this.pending().filter(x => String(x.id) !== String(a.id)));
      syncMessage = this.pending().length ? 'An exam is pending upload. It is not in shared history yet.' : '';
      return true;
    } catch (err) {
      console.warn('Shared history save failed', err);
      const pending = this.pending();
      if (!pending.some(x => String(x.id) === String(a.id))) pending.push(a);
      this.setPending(pending);
      syncMessage = 'An exam is pending upload. It is not in shared history yet.';
      return false;
    }
  },
  async retryPending() {
    for (const attempt of this.pending()) {
      if (!await this.add(attempt)) return false;
    }
    return this.sync();
  }
};

let DATA = null, cur = null, timer = null, syncMessage = '';

/* ---------- Load questions and schedule ---------- */
const normRow = r => Object.fromEntries(Object.entries(r).map(([k, v]) => [k.toLowerCase().replace(/[^a-z]/g, ''), v]));

function cellToYMD(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') {                       // Excel date serial
    const d = new Date(Math.round((v - 25569) * 864e5));
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return `${m[1]}-${pad(+m[2])}-${pad(+m[3])}`;
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return `${m[3]}-${pad(+m[2])}-${pad(+m[1])}`;
  return null;
}

async function loadData() {
  let rows;
  if (CFG.SYNC_URL) {
    let data;
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        const url = `${CFG.SYNC_URL}?action=data&code=${encodeURIComponent(CFG.CODE)}&v=${Date.now()}`;
        const res = await fetch(url, { cache: 'no-store' });
        if (!res.ok) throw new Error('Google Sheet is unavailable');
        data = await res.json();
        break;
      } catch (err) {
        if (attempt === 2) throw err;
        await new Promise(resolve => setTimeout(resolve, 700));
      }
    }
    if (data.error || !Array.isArray(data.questions) || !Array.isArray(data.exams)) {
      throw new Error(data.error || 'Google Sheet question endpoint is not updated');
    }
    const records = table => {
      const [headers = [], ...items] = table;
      return items.map(cells => normRow(Object.fromEntries(headers.map((header, i) => [header, cells[i] ?? '']))));
    };
    rows = name => records(name === 'Questions' ? data.questions : data.exams);
  } else {
    const res = await fetch('questions.xlsx?v=' + Date.now());
    if (!res.ok) throw new Error('questions.xlsx not found');
    const wb = XLSX.read(await res.arrayBuffer(), { type: 'array' });
    rows = name => wb.Sheets[name] ? XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' }).map(normRow) : [];
  }

  const exams = {};
  const make = n => ({ n, unlock: addDays(CFG.START, n - 1), minutes: CFG.DEFAULT_MINUTES, sections: [] });
  for (let n = 1; n <= CFG.TOTAL; n++) exams[n] = make(n);

  rows('Exams').forEach(r => {
    const n = Number(r.exam); if (!n) return;
    exams[n] = exams[n] || make(n);
    const u = cellToYMD(r.unlockdate); if (u) exams[n].unlock = u;
    const m = Number(r.minutespersection); if (m > 0) exams[n].minutes = m;
  });

  rows('Questions').forEach((r, i) => {
    const n = Number(r.exam), section = String(r.section).trim(), q = String(r.question).trim();
    if (!Number.isInteger(n) || n < 1 || n > CFG.TOTAL || !section || !q) {
      console.warn('Skipped bad row', i + 2, r);
      return;
    }
    const opts = ['a', 'b', 'c', 'd'].map(k => String(r[k]).trim());
    while (opts.length && !opts[opts.length - 1]) opts.pop();
    const ans = L.indexOf(String(r.answer).trim().toUpperCase());
    if (opts.length < 2 || ans < 0 || ans >= opts.length) { console.warn('Skipped bad row', i + 2, r); return; }
    exams[n] = exams[n] || make(n);
    let sec = exams[n].sections.find(s => s.name === section);
    if (!sec) exams[n].sections.push(sec = { name: section, qs: [] });
    sec.qs.push({ q, opts, ans });
  });

  return { exams, total: CFG.TOTAL };
}

/* ---------- Boot + lock ---------- */
function renderLock() {
  app.innerHTML = `<div class="lock"><h1>SI Prep</h1><p class="muted">Enter access code</p>
    <input id="code" type="password" inputmode="numeric" maxlength="4" autocomplete="off" autofocus>
    <p id="err" class="err"></p></div>`;
  const i = $('#code');
  i.oninput = () => {
    if (i.value.length < 4) return;
    if (i.value === CFG.CODE) { localStorage.setItem(CFG.UNLOCK, '1'); boot(); }
    else { $('#err').textContent = 'Wrong code'; i.value = ''; }
  };
}

async function boot() {
  app.innerHTML = '<p class="center muted">Loading…</p>';
  try { DATA = await loadData(); }
  catch (e) {
    app.innerHTML = `<p class="center">Couldn't load exam questions<br><span class="muted small">${esc(e.message)}. Check the Google Sheet connection and reload.</span></p>`;
    return;
  }
  await store.sync();
  if (store.pending().length) syncMessage = 'An exam is pending upload. It is not in shared history yet.';
  go('exams');
}

/* ---------- Shell + tabs ---------- */
function stopTimer() { clearInterval(timer); timer = null; window.onbeforeunload = null; }

function daysLeft() {
  const n = daysBetween(today(), CFG.EXAM_DAY);
  return n > 1 ? `${n} days to exam` : n === 1 ? '1 day to exam' : n === 0 ? 'Exam day' : 'Exam over';
}

function shell(inner, active) {
  app.innerHTML = `<header><div class="brand">SI Prep</div><div class="days">${daysLeft()}</div></header>
    <nav>${['exams', 'history', 'progress'].map(t => `<button data-t="${t}" class="${t === active ? 'on' : ''}">${t[0].toUpperCase() + t.slice(1)}</button>`).join('')}</nav>
    <main>${syncMessage ? `<p class="err small">${esc(syncMessage)} <button class="link" id="retry-sync">${store.pending().length ? 'Retry save' : 'Retry connection'}</button></p>` : ''}${inner}</main>`;
  app.querySelectorAll('nav button').forEach(b => b.onclick = async () => {
    b.disabled = true;
    await store.sync();
    if (store.pending().length) syncMessage = 'An exam is pending upload. It is not in shared history yet.';
    go(b.dataset.t);
  });
  const retry = $('#retry-sync');
  if (retry) retry.onclick = async () => {
    retry.disabled = true;
    if (store.pending().length) await store.retryPending();
    else await store.sync();
    go(active);
  };
  window.scrollTo(0, 0);
}

function go(tab) {
  stopTimer(); cur = null;
  ({ exams: renderExams, history: renderHistory, progress: renderProgress })[tab]();
}

/* ---------- Exam list + intro ---------- */
function renderExams() {
  const T = today(), best = {};
  store.get().forEach(a => { const p = pct(a.correct, a.total); if (!(a.exam in best) || p > best[a.exam]) best[a.exam] = p; });
  let cards = '';
  for (let n = 1; n <= DATA.total; n++) {
    const ex = DATA.exams[n] || { unlock: addDays(CFG.START, n - 1), sections: [] };
    const open = T >= ex.unlock, has = ex.sections.length > 0;
    let s, locked = false;
    if (!open) { s = `Opens ${fmt(ex.unlock)}`; locked = true; }
    else if (!has) { s = 'Coming soon'; locked = true; }
    else s = n in best ? `Best ${best[n]}%` : (ex.unlock === T ? 'Today' : 'Start');
    cards += `<button class="card${locked ? ' lock' : ''}${ex.unlock === T ? ' today' : ''}${n in best ? ' done' : ''}" data-n="${n}" ${locked ? 'disabled' : ''}><span class="n">${n}</span><span class="s">${s}</span></button>`;
  }
  shell(`<div class="grid">${cards}</div>`, 'exams');
  app.querySelectorAll('.card:not(.lock)').forEach(c => c.onclick = () => showIntro(+c.dataset.n));
}

function showIntro(n) {
  const ex = DATA.exams[n];
  const total = ex.sections.reduce((s, x) => s + x.qs.length, 0);
  shell(`<h2>Exam ${n}</h2>
    <p class="muted">${ex.sections.length} sections · ${total} questions · ${ex.minutes} min per section</p>
    <ul class="secs">${ex.sections.map(s => `<li><span>${esc(s.name)}</span><span class="muted">${s.qs.length} Q</span></li>`).join('')}</ul>
    <p class="muted small">The timer starts with each section. When it ends, the section is submitted automatically.</p>
    <div class="row"><button class="btn ghost" id="back">Back</button><button class="btn" id="go">Start</button></div>`, 'exams');
  $('#back').onclick = () => go('exams');
  $('#go').onclick = () => { cur = { n, ex, si: 0, results: [], endedEarly: false }; runSection(); };
}

/* ---------- Taking a section ---------- */
function exitExam() {
  if (confirm('Exit? This exam attempt will not be saved.')) go('exams');
}

function runSection() {
  const { ex, si } = cur, sec = ex.sections[si];
  cur.answers = sec.qs.map(() => -1);
  cur.secStart = Date.now();
  cur.endAt = cur.secStart + ex.minutes * 60000;
  app.innerHTML = `<div class="exam">
    <div class="topbar">
      <div><div class="muted small">Exam ${cur.n} · Section ${si + 1} of ${ex.sections.length}</div><div class="sec">${esc(sec.name)}</div></div>
      <div><div id="timer" class="timer">--:--</div><div style="text-align:right"><button class="link" id="exit">Exit</button></div></div>
    </div>
    <div class="prog"><i style="width:${(si / ex.sections.length) * 100}%"></i></div>
    ${sec.qs.map((q, i) => `<section class="q"><p class="qt"><b>${i + 1}.</b> ${esc(q.q)}</p>
      ${q.opts.map((o, j) => `<label class="opt"><input type="radio" name="q${i}" value="${j}"><span class="l">${L[j]}</span><span>${esc(o)}</span></label>`).join('')}</section>`).join('')}
    <div class="foot"><span id="cnt" class="muted">0 / ${sec.qs.length} answered</span><button class="btn" id="sub">Submit section</button></div>
  </div>`;
  app.querySelectorAll('input[type=radio]').forEach(r => r.onchange = () => {
    cur.answers[+r.name.slice(1)] = +r.value;
    $('#cnt').textContent = `${cur.answers.filter(a => a >= 0).length} / ${sec.qs.length} answered`;
  });
  $('#sub').onclick = () => submitSection(false);
  $('#exit').onclick = exitExam;
  window.onbeforeunload = e => { e.preventDefault(); e.returnValue = ''; };
  tick(); timer = setInterval(tick, 250);
  window.scrollTo(0, 0);
}

function tick() {
  const t = $('#timer'); if (!t || !cur) return;
  const left = Math.max(0, Math.ceil((cur.endAt - Date.now()) / 1000));
  t.textContent = `${pad(Math.floor(left / 60))}:${pad(left % 60)}`;
  t.classList.toggle('low', left <= 60);
  if (left <= 0) submitSection(true);
}

function submitSection(auto) {
  const sec = cur.ex.sections[cur.si];
  const unanswered = cur.answers.filter(a => a < 0).length;
  if (!auto && unanswered && !confirm(`${unanswered} unanswered. Submit anyway?`)) return;
  stopTimer();
  const items = []; let correct = 0;
  sec.qs.forEach((q, i) => {
    if (cur.answers[i] === q.ans) correct++;
    else items.push({ q: q.q, opts: q.opts, chosen: cur.answers[i], ans: q.ans });
  });
  const r = {
    name: sec.name, total: sec.qs.length, correct, items,
    skipped: items.filter(x => x.chosen < 0).length,
    used: Math.round((Date.now() - cur.secStart) / 1000), timedOut: !!auto
  };
  cur.results.push(r);
  showSectionResult(r);
}

/* ---------- Results ---------- */
function wrongList(items) {
  if (!items.length) return '<p class="muted">Everything correct.</p>';
  return items.map(x => `<div class="wq"><p class="qt">${esc(x.q)}</p>
    <p class="ans bad">${x.chosen < 0 ? 'Not answered' : `Your answer: ${L[x.chosen]}) ${esc(x.opts[x.chosen])}`}</p>
    <p class="ans ok">Correct: ${L[x.ans]}) ${esc(x.opts[x.ans])}</p></div>`).join('');
}

function showSectionResult(r) {
  const last = cur.si === cur.ex.sections.length - 1;
  const wrong = r.total - r.correct - r.skipped;
  app.innerHTML = `<div class="exam">
    <div class="topbar"><div class="muted small">Exam ${cur.n} · Section ${cur.si + 1} of ${cur.ex.sections.length}</div><button class="link" id="exit">Exit</button></div>
    <h2>${esc(r.name)}</h2>
    <div class="score">${r.correct}<span> / ${r.total}</span></div>
    <p class="muted">${wrong} wrong · ${r.skipped} skipped${r.timedOut ? ' · time up' : ''}</p>
    ${r.items.length ? '<h3>Wrong or skipped</h3>' : ''}
    ${wrongList(r.items)}
    <div class="foot">${last ? '<span></span>' : '<button class="link" id="finish">Finish exam now</button>'}<button class="btn" id="next">${last ? 'Finish exam' : 'Next section'}</button></div></div>`;
  $('#exit').onclick = exitExam;
  $('#next').onclick = () => { if (last) finishExam(); else { cur.si++; runSection(); } };
  if (!last) $('#finish').onclick = () => {
    if (confirm('Finish now? Completed sections will be saved.')) { cur.endedEarly = true; finishExam(); }
  };
  window.scrollTo(0, 0);
}

function attemptBody(a, showScore) {
  return `${showScore ? `<div class="score">${a.correct}<span> / ${a.total}</span></div><p class="muted">${pct(a.correct, a.total)}% · ${a.total - a.correct} wrong or skipped</p>` : ''}
    ${a.sections.map(s => `<details><summary><span>${esc(s.name)}</span><span><b>${s.correct}/${s.total}</b> <span class="muted">${s.total - s.correct} wrong</span></span></summary><div>${wrongList(s.items)}</div></details>`).join('')}`;
}

async function finishExam() {
  if (cur.saving) return;
  cur.saving = true;
  const a = {
    id: Date.now(), exam: cur.n, ts: new Date().toISOString(), date: today(),
    sections: cur.results,
    total: cur.results.reduce((s, r) => s + r.total, 0),
    correct: cur.results.reduce((s, r) => s + r.correct, 0),
    endedEarly: !!cur.endedEarly
  };
  const saved = await store.add(a);
  stopTimer();
  shell(`<h2>Exam ${a.exam} ${a.endedEarly ? 'ended early' : 'complete'}</h2>
    <p class="${saved ? 'ok' : 'err'} small">${saved ? 'Saved to shared history.' : 'Not saved to the shared Sheet yet. Use Retry save above.'}</p>
    ${attemptBody(a, true)}
    <div class="row"><button class="btn" id="home">Back to exams</button></div>`, 'exams');
  cur = null;
  $('#home').onclick = () => go('exams');
}

/* ---------- History ---------- */
function renderHistory() {
  const list = store.get().slice().reverse();
  shell(list.length
    ? list.map(a => `<details><summary><span><b>Exam ${a.exam}</b> <span class="muted">· ${dt(a.ts)}${a.endedEarly ? ' · ended early' : ''}</span></span><span><b>${a.correct}/${a.total}</b> <span class="muted">${pct(a.correct, a.total)}%</span></span></summary><div>${attemptBody(a, false)}</div></details>`).join('')
    : '<p class="muted">No attempts yet.</p>', 'history');
}

/* ---------- Progress ---------- */
function renderProgress() {
  const at = store.get(), T = today();
  const byDate = {};
  at.forEach(a => (byDate[a.date] = byDate[a.date] || []).push(a));
  const agg = list => list.reduce((s, a) => { s.c += a.correct; s.t += a.total; return s; }, { c: 0, t: 0 });
  const N = daysBetween(CFG.START, CFG.EXAM_DAY) + 1;

  // stats
  const all = agg(at);
  let streak = 0, d = byDate[T] ? T : addDays(T, -1);
  while (byDate[d] && d >= CFG.START) { streak++; d = addDays(d, -1); }
  const stats = `<div class="stats">
    <div class="stat"><b>${new Set(at.map(a => a.exam)).size}/${DATA.total}</b><span>exams done</span></div>
    <div class="stat"><b>${all.t ? pct(all.c, all.t) + '%' : '–'}</b><span>accuracy</span></div>
    <div class="stat"><b>${streak}</b><span>day streak</span></div>
    <div class="stat"><b>${Math.max(0, daysBetween(T, CFG.EXAM_DAY))}</b><span>days left</span></div></div>`;

  // day strip
  let strip = '';
  for (let i = 0; i < N; i++) {
    const day = addDays(CFG.START, i), list = byDate[day];
    let cls = day > T ? 'fut' : 'past', tip = fmtDay(day);
    if (list) { const g = agg(list), p = pct(g.c, g.t); cls = p >= 75 ? 'l3' : p >= 50 ? 'l2' : 'l1'; tip += ` · ${p}%`; }
    if (day === T) cls += ' now';
    strip += `<i class="${cls}" title="${tip}"></i>`;
  }

  // subjects (weakest first)
  const subj = {};
  at.forEach(a => a.sections.forEach(s => { const o = subj[s.name] = subj[s.name] || { c: 0, t: 0 }; o.c += s.correct; o.t += s.total; }));
  const subjHtml = Object.entries(subj).sort((a, b) => pct(a[1].c, a[1].t) - pct(b[1].c, b[1].t))
    .map(([n, v]) => `<div class="sub"><span>${esc(n)}</span><em>${pct(v.c, v.t)}%</em><div class="bar"><i style="width:${pct(v.c, v.t)}%"></i></div></div>`).join('')
    || '<p class="muted">Take an exam to see subject accuracy.</p>';

  // weekly
  let weeks = '', prev = null;
  for (let w = 0; w * 7 < N; w++) {
    const ws = addDays(CFG.START, w * 7); if (ws > T) break;
    let we = addDays(ws, 6); if (we > CFG.EXAM_DAY) we = CFG.EXAM_DAY;
    const list = at.filter(a => a.date >= ws && a.date <= we), g = agg(list);
    const p = g.t ? pct(g.c, g.t) : null;
    const delta = p !== null && prev !== null ? p - prev : null;
    weeks = `<tr><td>Week ${w + 1}<br><span class="muted small">${fmt(ws)} – ${fmt(we)}</span></td>
      <td>${new Set(list.map(a => a.date)).size} days</td><td>${list.length} exams</td>
      <td>${p === null ? '–' : p + '%'}${delta ? ` <span class="${delta > 0 ? 'up' : 'down'} small">${delta > 0 ? '+' : ''}${delta}</span>` : ''}</td></tr>` + weeks;
    prev = p;
  }

  // daily
  let days = '';
  for (let i = 0, day = CFG.START; day <= T && day <= CFG.EXAM_DAY; i++, day = addDays(CFG.START, i)) {
    const list = byDate[day], g = list ? agg(list) : null;
    days = `<tr><td>${fmtDay(day)}</td><td>${list ? list.map(a => 'Exam ' + a.exam).join(', ') : '<span class="muted">–</span>'}</td>
      <td>${g ? `${g.c}/${g.t}` : ''}</td><td>${g ? pct(g.c, g.t) + '%' : ''}</td></tr>` + days;
  }

  shell(`${stats}
    <h3>Days until ${fmt(CFG.EXAM_DAY)}</h3><div class="strip">${strip}</div>
    <h3>Subjects · weakest first</h3>${subjHtml}
    <h3>Weekly report</h3><table><tbody>${weeks || '<tr><td class="muted">No weeks yet</td></tr>'}</tbody></table>
    <h3>Daily report</h3><table><thead><tr><th>Day</th><th>Exam</th><th>Score</th><th>%</th></tr></thead><tbody>${days}</tbody></table>`, 'progress');
}

/* ---------- Start ---------- */
if (localStorage.getItem(CFG.UNLOCK) === '1') boot(); else renderLock();
