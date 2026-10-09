/* TeamsBrief Web — 100% navigateur, hébergé */

const STORAGE_KEY = 'teamsbrief_meetings';
const SETTINGS_KEY = 'teamsbrief_settings';
const THEME_KEY = 'teamsbrief_theme';

const STATUS_LABELS = { ready: 'Prêt', processing: 'En cours', error: 'Erreur' };
const GROQ_FALLBACK_MODEL = 'openai/gpt-oss-20b';
let chartInstances = [];
let globalChartInstances = [];

const SUGGESTIONS = [
  "Qu'est-ce qui a été décidé ?",
  "Quelles sont les actions à faire ?",
  "Résume-moi en 3 lignes",
  "Explique-moi comme si j'étais absent",
  "Quelles questions restent en suspens ?",
];

const SUMMARY_PROMPT = (title, transcript) => `Analyse cette transcription de réunion Teams et produis un résumé structuré en JSON.

TITRE : ${title}

TRANSCRIPTION :
${transcript.slice(0, 25000)}

Réponds UNIQUEMENT en JSON valide avec cette structure :
{
  "tldr": ["point 1", "point 2", "point 3"],
  "decisions": ["décision précise avec contexte"],
  "actions": [{"who": "nom exact", "what": "action détaillée", "when": "échéance"}],
  "topics": [{"title": "thème", "summary": "synthèse", "key_points": ["détail 1", "détail 2"]}],
  "open_questions": ["question"],
  "explanation_for_absent": "paragraphe pour absent",
  "presentation": {
    "tagline": "phrase d'accroche",
    "executive_summary": "3-4 phrases pour décideurs",
    "highlights": [{"title": "point clé", "detail": "explication précise"}]
  }
}

Règles : decisions et actions précises, topics avec key_points, 3-5 highlights, ne pas inventer.`;

// ─── State ───
let meetings = loadMeetings();
let currentMeetingId = null;
let settings = loadSettings();

// ─── DOM ───
const views = {
  meetings: document.getElementById('view-meetings'),
  detail: document.getElementById('view-detail'),
  settings: document.getElementById('view-settings'),
};

// ─── Init ───
document.querySelectorAll('.nav-btn').forEach(btn => {
  btn.addEventListener('click', () => showView(btn.dataset.view));
});

document.getElementById('file-input').addEventListener('change', handleFileImport);
document.getElementById('back-btn').addEventListener('click', () => showView('meetings'));
document.getElementById('save-settings').addEventListener('click', saveSettings);
document.getElementById('chat-form').addEventListener('submit', handleChat);
document.getElementById('export-pptx-btn').addEventListener('click', exportPresentation);
document.getElementById('charts-toggle').addEventListener('click', toggleCharts);
document.getElementById('delete-meeting-btn').addEventListener('click', deleteCurrentMeeting);
document.getElementById('reprocess-btn').addEventListener('click', reprocessMeeting);
document.getElementById('copy-summary-btn').addEventListener('click', copySummary);
document.getElementById('export-md-btn').addEventListener('click', exportMarkdown);
document.getElementById('read-report-btn').addEventListener('click', openReadableReport);
document.getElementById('download-report-inline')?.addEventListener('click', downloadReportForCurrent);
document.getElementById('open-report-overlay')?.addEventListener('click', openReportOverlayForCurrent);
document.getElementById('goto-report-btn')?.addEventListener('click', openReadableReport);
document.getElementById('download-report-quick')?.addEventListener('click', downloadReportForCurrent);
document.getElementById('report-close-btn').addEventListener('click', closeReportOverlay);
document.getElementById('report-download-btn').addEventListener('click', downloadCurrentReport);
document.getElementById('report-print-btn').addEventListener('click', () => window.print());
document.getElementById('report-overlay').addEventListener('click', e => {
  if (e.target.id === 'report-overlay') closeReportOverlay();
});
document.getElementById('clear-chat-btn').addEventListener('click', clearChat);
document.getElementById('search-input').addEventListener('input', renderMeetingsList);
document.getElementById('filter-status').addEventListener('change', renderMeetingsList);
document.getElementById('transcript-search').addEventListener('input', filterTranscript);
document.getElementById('theme-toggle').addEventListener('click', toggleTheme);
document.getElementById('toggle-key').addEventListener('click', toggleKeyVisibility);
document.getElementById('backup-export').addEventListener('click', exportBackup);
document.getElementById('backup-import').addEventListener('change', importBackup);
document.getElementById('settings-backup').addEventListener('click', exportBackup);
document.getElementById('settings-restore').addEventListener('change', importBackup);
document.getElementById('onboarding-go').addEventListener('click', () => showView('settings'));

document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => switchTab(tab.dataset.tab));
});

const dropzone = document.getElementById('dropzone');
['dragenter', 'dragover'].forEach(ev => dropzone.addEventListener(ev, e => { e.preventDefault(); dropzone.classList.add('dragover'); }));
['dragleave', 'drop'].forEach(ev => dropzone.addEventListener(ev, e => { e.preventDefault(); dropzone.classList.remove('dragover'); }));
dropzone.addEventListener('drop', e => { if (e.dataTransfer.files.length) processFiles([...e.dataTransfer.files]); });

document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
    e.preventDefault();
    if (!views.meetings.classList.contains('hidden')) document.getElementById('search-input').focus();
  }
});

document.getElementById('groq-key').value = settings.apiKey || '';
const modelEl = document.getElementById('groq-model');
if (settings.model) modelEl.value = settings.model;
applyTheme(localStorage.getItem(THEME_KEY) || 'light');
updateOnboarding();
renderMeetingsList();
renderDashboard();

// ─── Navigation ───
function showView(name) {
  Object.values(views).forEach(v => v.classList.add('hidden'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));

  if (name === 'detail') {
    views.detail.classList.remove('hidden');
  } else if (name === 'settings') {
    views.settings.classList.remove('hidden');
    document.querySelector('[data-view="settings"]').classList.add('active');
  } else {
    views.meetings.classList.remove('hidden');
    document.querySelector('[data-view="meetings"]').classList.add('active');
    currentMeetingId = null;
    renderMeetingsList();
    renderDashboard();
  }
}

// ─── Storage ───
function normalizeGroqSummary(raw) {
  if (!raw || typeof raw !== 'object') return raw;
  let o = Array.isArray(raw) ? { tldr: raw.map(String) } : { ...raw };
  for (const key of ['data', 'result', 'report', 'rapport', 'summary', 'résumé', 'resume', 'réunion', 'reunion']) {
    const nest = o[key];
    if (nest && typeof nest === 'object' && !Array.isArray(nest)) o = { ...o, ...nest };
  }
  if (o.décisions && !o.decisions) o.decisions = o.décisions;
  if (o.actions_a_faire && !o.actions) o.actions = o.actions_a_faire;
  if (o.points_cles && !o.tldr) o.tldr = o.points_cles;
  if (o.questions_ouvertes && !o.open_questions) o.open_questions = o.questions_ouvertes;
  if (typeof o.presentation === 'string') o.presentation = { executive_summary: o.presentation };
  if (typeof o.synthese === 'string' && !o.presentation?.executive_summary) {
    o.presentation = { ...(o.presentation || {}), executive_summary: o.synthese };
  }
  if (Array.isArray(o.actions)) {
    o.actions = o.actions.map(a => (typeof a === 'string' ? { who: '—', what: a, when: '' } : a));
  }
  return o;
}

function coerceSummary(summary) {
  if (!summary) return null;
  if (typeof summary === 'string') {
    try {
      const t = summary.trim();
      if (t.startsWith('{')) return normalizeGroqSummary(JSON.parse(t));
      return normalizeGroqSummary({ tldr: [summary] });
    } catch {
      try { return normalizeGroqSummary(parseGroqJson(summary)); } catch { return normalizeGroqSummary({ tldr: [summary] }); }
    }
  }
  return typeof summary === 'object' ? normalizeGroqSummary(summary) : null;
}

function findMeeting(id) {
  if (id == null) return null;
  const sid = String(id);
  return meetings.find(m => String(m.id) === sid) || null;
}

function loadMeetings() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]').map(m => {
      m.summary = coerceSummary(m.summary);
      return m;
    });
  }
  catch { return []; }
}

function saveMeetings() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(meetings));
}

function loadSettings() {
  try { return JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}'); }
  catch { return {}; }
}

function saveSettings() {
  settings = {
    apiKey: document.getElementById('groq-key').value.trim(),
    model: document.getElementById('groq-model').value || 'openai/gpt-oss-20b',
  };
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  document.getElementById('settings-status').textContent = 'Paramètres enregistrés.';
  updateOnboarding();
  toast('Paramètres enregistrés', 'success');
}

function toast(msg, type = 'success') {
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = msg;
  document.getElementById('toast-container').appendChild(el);
  setTimeout(() => el.remove(), 3000);
}

function statusLabel(s) { return STATUS_LABELS[s] || s; }

function newMeetingId() {
  return crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function toggleTheme() {
  const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
  applyTheme(next);
  localStorage.setItem(THEME_KEY, next);
}

function applyTheme(theme) {
  document.documentElement.dataset.theme = theme;
  document.getElementById('theme-toggle').textContent = theme === 'dark' ? '☀️' : '🌙';
  chartInstances.forEach(c => c.destroy());
  chartInstances = [];
  destroyGlobalCharts();
  const m = findMeeting(currentMeetingId);
  if (m?.summary) renderCharts(m.summary);
  if (!views.detail.classList.contains('hidden')) return;
  renderDashboard();
}

function toggleKeyVisibility() {
  const inp = document.getElementById('groq-key');
  inp.type = inp.type === 'password' ? 'text' : 'password';
}

function updateOnboarding() {
  document.getElementById('onboarding-banner').classList.toggle('hidden', !!settings.apiKey);
}

function switchTab(name) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === name));
  document.getElementById('tab-summary').classList.toggle('hidden', name !== 'summary');
  document.getElementById('tab-report').classList.toggle('hidden', name !== 'report');
  document.getElementById('tab-transcript').classList.toggle('hidden', name !== 'transcript');
  const m = findMeeting(currentMeetingId);
  if (name === 'report' && m) renderReportTab(m);
}

function hasAiSummary(m) {
  const s = coerceSummary(m?.summary);
  if (!s) return false;
  const structured = !!(
    s.presentation?.executive_summary
    || s.presentation?.highlights?.length
    || s.tldr?.length
    || s.decisions?.length
    || s.actions?.length
    || s.topics?.length
    || s.explanation_for_absent
    || s.open_questions?.length
  );
  if (structured) return true;
  return !!(renderSummary(s, m) || '').trim();
}

/** Rapport affichable (synthèse IA ou, à défaut, transcription de l’appel). */
function hasReadableReport(m) {
  if (!m || m.status !== 'ready') return false;
  if (hasAiSummary(m)) return true;
  return (m.transcript || '').trim().length > 80;
}

function updateReportQuickBar(m) {
  const bar = document.getElementById('report-quick-bar');
  if (!bar) return;
  const show = m && m.status === 'ready' && hasReadableReport(m);
  bar.classList.toggle('hidden', !show);
}

function exportBackup() {
  const blob = new Blob([JSON.stringify({ meetings, settings, exportedAt: new Date().toISOString() }, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `teamsbrief-backup-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  toast('Sauvegarde exportée');
}

async function importBackup(e) {
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = '';
  try {
    const data = JSON.parse(await file.text());
    if (data.meetings) meetings = data.meetings;
    if (data.settings) { settings = data.settings; localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)); }
    saveMeetings();
    document.getElementById('groq-key').value = settings.apiKey || '';
    if (settings.model) document.getElementById('groq-model').value = settings.model;
    updateOnboarding();
    renderMeetingsList();
    toast(`${meetings.length} réunion(s) restaurée(s)`);
  } catch { toast('Fichier invalide', 'error'); }
}

// ─── VTT Parser ───
function parseVtt(text) {
  const lines = [];
  let time = '';
  for (const line of text.split('\n')) {
    const t = line.trim();
    if (!t || t.startsWith('WEBVTT') || t.startsWith('NOTE')) continue;
    if (t.includes('-->')) {
      const start = t.split('-->')[0].trim();
      const parts = start.split(':');
      time = parts.length >= 2 ? parts.slice(-2).join(':').split('.')[0] : start;
      continue;
    }
    if (/^\d+$/.test(t)) continue;
    lines.push(time ? `[${time}] ${t}` : t);
  }
  return lines.join('\n');
}

// ─── Groq API (via proxy hébergé ou direct) ───
function parseGroqJson(raw) {
  const text = String(raw || '').trim();
  try { return JSON.parse(text); } catch { /* continue */ }
  const fenced = text.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (fenced) {
    try { return JSON.parse(fenced[1].trim()); } catch { /* continue */ }
  }
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start >= 0 && end > start) {
    try { return JSON.parse(text.slice(start, end + 1)); } catch { /* continue */ }
  }
  throw new Error('Réponse Groq illisible. Cliquez sur Réessayer.');
}

function humanizeGroqError(msg) {
  const m = String(msg || '');
  if (m.includes('model_not_found') || m.includes('does not exist')) {
    return 'Modèle Groq introuvable. Allez dans Paramètres et choisissez « openai/gpt-oss-20b ».';
  }
  if (m.includes('invalid_api_key') || m.includes('Invalid API Key')) {
    return 'Clé Groq invalide. Vérifiez votre clé dans Paramètres (gsk_...).';
  }
  if (m.includes('rate_limit') || m.includes('429')) {
    return 'Limite Groq atteinte. Attendez 10 secondes puis réessayez.';
  }
  return m;
}

async function callGroqOnce(messages, jsonMode, model, url) {
  const payload = {
    model,
    messages,
    temperature: 0.3,
    max_tokens: 2048,
  };
  if (jsonMode) payload.response_format = { type: 'json_object' };

  const isLocalProxy = url.startsWith('/');
  const isRemoteProxy = url.includes('/api/groq');
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      ...(isLocalProxy || isRemoteProxy ? {} : { Authorization: `Bearer ${settings.apiKey}` }),
    },
    body: JSON.stringify(isLocalProxy || isRemoteProxy ? { ...payload, apiKey: settings.apiKey } : payload),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    const msg = err.error?.message || err.detail || `Erreur ${res.status}`;
    if (res.status === 429) throw new Error('Limite Groq atteinte. Attendez 10 secondes.');
    throw new Error(humanizeGroqError(msg));
  }
  const data = await res.json();
  return data.choices[0].message.content;
}

async function callGroq(messages, jsonMode = false) {
  if (!settings.apiKey) throw new Error('Configurez votre clé Groq dans Paramètres.');

  const onGithubPages = location.hostname.includes('github.io');
  const endpoints = onGithubPages
    ? [
        'https://teamsbrief-web.vercel.app/api/groq',
        'https://api.groq.com/openai/v1/chat/completions',
      ]
    : [
        '/api/groq',
        'https://teamsbrief-web.vercel.app/api/groq',
        'https://api.groq.com/openai/v1/chat/completions',
      ];
  const models = [settings.model || GROQ_FALLBACK_MODEL];
  if (!models.includes(GROQ_FALLBACK_MODEL)) models.push(GROQ_FALLBACK_MODEL);

  let lastError = null;
  for (const model of models) {
    for (const url of endpoints) {
      try {
        return await callGroqOnce(messages, jsonMode, model, url);
      } catch (e) {
        lastError = e;
        const retryable = String(e.message).includes('introuvable') || String(e.message).includes('Erreur 5');
        if (!retryable && url === endpoints[endpoints.length - 1] && model === models[models.length - 1]) break;
      }
    }
  }
  throw lastError || new Error('Impossible de contacter Groq.');
}

// ─── Import fichier ───
async function handleFileImport(e) {
  const files = [...e.target.files];
  e.target.value = '';
  if (files.length) await processFiles(files);
}

async function processFiles(files) {
  if (!settings.apiKey) {
    toast('Configurez votre clé Groq d\'abord', 'error');
    showView('settings');
    return;
  }
  const vttFiles = files.filter(f => /\.(vtt|txt)$/i.test(f.name));
  if (!vttFiles.length) { toast('Aucun fichier .vtt trouvé', 'error'); return; }

  toast(`Import de ${vttFiles.length} fichier(s)…`);
  let lastId = null;
  for (const file of vttFiles) {
    lastId = await importOneFile(file);
  }
  renderMeetingsList();
  if (lastId) openMeeting(lastId);
}

async function importOneFile(file) {
  const text = await file.text();
  const transcript = file.name.endsWith('.vtt') || text.startsWith('WEBVTT') ? parseVtt(text) : text;
  const title = file.name.replace(/\.(vtt|txt)$/i, '').replace(/^.*[\\/]/, '') || 'Réunion Teams';

  const meeting = {
    id: newMeetingId(),
    title,
    date: new Date().toISOString(),
    transcript,
    summary: null,
    status: 'processing',
    chat: [],
    doneActions: [],
  };

  meetings.unshift(meeting);
  saveMeetings();
  renderMeetingsList();

  try {
    const raw = await callGroq([
      { role: 'system', content: 'Tu réponds uniquement en JSON valide, en français.' },
      { role: 'user', content: SUMMARY_PROMPT(title, transcript) },
    ], true);
    meeting.summary = coerceSummary(parseGroqJson(raw));
    meeting.status = 'ready';
    toast(`"${title}" prête`);
  } catch (err) {
    meeting.status = 'error';
    meeting.error = humanizeGroqError(err.message);
    toast(`Erreur : ${title}`, 'error');
  }

  saveMeetings();
  return meeting.id;
}

async function reprocessMeeting() {
  const m = meetings.find(x => x.id === currentMeetingId);
  if (!m) return;
  m.status = 'processing';
  m.error = null;
  saveMeetings();
  openMeeting(m.id);
  try {
    const raw = await callGroq([
      { role: 'system', content: 'Tu réponds uniquement en JSON valide, en français.' },
      { role: 'user', content: SUMMARY_PROMPT(m.title, m.transcript) },
    ], true);
    m.summary = coerceSummary(parseGroqJson(raw));
    m.status = 'ready';
    toast('Résumé régénéré');
  } catch (err) {
    m.status = 'error';
    m.error = humanizeGroqError(err.message);
    toast(err.message, 'error');
  }
  saveMeetings();
  openMeeting(m.id);
}

function deleteMeeting(id) {
  if (!confirm('Supprimer cette réunion ?')) return;
  meetings = meetings.filter(m => String(m.id) !== String(id));
  saveMeetings();
  if (currentMeetingId === id) showView('meetings');
  else renderMeetingsList();
  toast('Réunion supprimée');
}

function deleteCurrentMeeting() {
  if (currentMeetingId) deleteMeeting(currentMeetingId);
}

function meetingMiniStats(m) {
  const s = coerceSummary(m.summary);
  if (!s) return '';
  const chips = [];
  if (s.decisions?.length) chips.push(`<span class="mini-chip chip-green">${s.decisions.length} décision(s)</span>`);
  if (s.actions?.length) chips.push(`<span class="mini-chip chip-blue">${s.actions.length} action(s)</span>`);
  if (s.open_questions?.length) chips.push(`<span class="mini-chip chip-orange">${s.open_questions.length} question(s)</span>`);
  return chips.length ? `<div class="meeting-chips">${chips.join('')}</div>` : '';
}

function computeGlobalStats() {
  const stats = {
    total: meetings.length,
    ready: 0,
    processing: 0,
    error: 0,
    decisions: 0,
    actions: 0,
    questions: 0,
    topicsCount: 0,
    actionOwners: {},
    byMonth: {},
  };
  meetings.forEach(m => {
    if (m.status === 'ready') stats.ready += 1;
    else if (m.status === 'processing') stats.processing += 1;
    else if (m.status === 'error') stats.error += 1;
    const s = coerceSummary(m.summary);
    if (!s) return;
    stats.decisions += s.decisions?.length || 0;
    stats.actions += s.actions?.length || 0;
    stats.questions += s.open_questions?.length || 0;
    (s.actions || []).forEach(a => {
      const who = (a.who || 'Non assigné').slice(0, 24);
      stats.actionOwners[who] = (stats.actionOwners[who] || 0) + 1;
    });
    const d = new Date(m.date);
    if (!Number.isNaN(d.getTime())) {
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      stats.byMonth[key] = (stats.byMonth[key] || 0) + 1;
    }
    stats.topicsCount += s.topics?.length || 0;
  });
  return stats;
}

function destroyGlobalCharts() {
  globalChartInstances.forEach(c => c.destroy());
  globalChartInstances = [];
}

function renderReportsHub() {
  const el = document.getElementById('reports-hub');
  if (!el) return;
  const reports = meetings.filter(m => hasReadableReport(m));
  const failed = meetings.filter(m => m.status === 'error');
  const pending = meetings.filter(m => m.status === 'processing');

  if (!meetings.length) {
    el.classList.add('hidden');
    return;
  }
  el.classList.remove('hidden');

  let inner = `
    <div class="dashboard-section-head">
      <h2>Rapports d'appel</h2>
      <p>Comptes-rendus de vos réunions Teams importées (.vtt)</p>
    </div>`;

  if (reports.length) {
    inner += `<ul class="reports-hub-list">${reports.slice(0, 12).map(m => `
      <li class="reports-hub-item">
        <div>
          <strong>${esc(m.title)}</strong>
          <span class="reports-hub-date">${new Date(m.date).toLocaleString('fr-FR', { dateStyle: 'medium', timeStyle: 'short' })}</span>
          ${hasAiSummary(m) ? '<span class="mini-chip chip-green">Synthèse IA</span>' : '<span class="mini-chip chip-orange">Transcription</span>'}
        </div>
        <button type="button" class="btn btn-primary btn-sm reports-hub-open" data-id="${esc(m.id)}">Ouvrir</button>
      </li>
    `).join('')}</ul>`;
    if (reports.length > 12) inner += `<p class="reports-hub-more">+ ${reports.length - 12} autre(s) rapport(s) dans la liste ci-dessous.</p>`;
  } else {
    inner += `<div class="reports-hub-empty">
      <p><strong>Aucun rapport lisible pour l’instant.</strong></p>
      ${failed.length ? `<p>${failed.length} réunion(s) en erreur — ouvrez-les et cliquez <strong>Réessayer</strong> (clé Groq + modèle <code>openai/gpt-oss-20b</code>).</p>` : ''}
      ${pending.length ? `<p>${pending.length} génération(s) en cours…</p>` : ''}
      ${!failed.length && !pending.length ? '<p>Importez un fichier .vtt ou restaurez une sauvegarde JSON (💾 / 📂).</p>' : ''}
      <p class="reports-hub-origin">Les données sont enregistrées <strong>dans ce navigateur</strong> pour cette adresse web (GitHub ≠ Vercel = listes différentes).</p>
    </div>`;
  }

  el.innerHTML = inner;
  el.querySelectorAll('.reports-hub-open').forEach(btn => {
    btn.addEventListener('click', () => {
      const id = btn.dataset.id;
      openMeeting(findMeeting(id)?.id ?? id);
      switchTab('report');
    });
  });
}

function renderDashboard() {
  const kpiEl = document.getElementById('dashboard-kpis');
  const chartsWrap = document.getElementById('dashboard-charts-wrap');
  if (!kpiEl) return;

  renderReportsHub();

  const g = computeGlobalStats();
  const subtitle = document.getElementById('meetings-stats');
  if (subtitle) {
    subtitle.textContent = g.total
      ? `${g.total} réunion(s) · ${g.ready} prête(s) · ${g.actions} action(s) à suivre`
      : 'Importez un fichier .vtt Teams pour alimenter votre tableau de bord.';
  }

  kpiEl.innerHTML = [
    { label: 'Réunions', value: g.total, tone: 'brand', icon: '📁' },
    { label: 'Prêtes', value: g.ready, tone: 'green', icon: '✓' },
    { label: 'Décisions', value: g.decisions, tone: 'violet', icon: '⚖' },
    { label: 'Actions', value: g.actions, tone: 'blue', icon: '☑' },
    { label: 'Questions', value: g.questions, tone: 'orange', icon: '?' },
    { label: 'Erreurs', value: g.error, tone: 'red', icon: '!' },
  ].map(k => `
    <div class="kpi-card kpi-${k.tone}">
      <div class="kpi-icon" aria-hidden="true">${k.icon}</div>
      <div class="kpi-body">
        <div class="kpi-value">${k.value}</div>
        <div class="kpi-label">${k.label}</div>
      </div>
    </div>
  `).join('');

  destroyGlobalCharts();
  if (!g.total || typeof Chart === 'undefined') {
    chartsWrap?.classList.add('hidden');
    document.getElementById('chart-hero-wrap')?.classList.add('hidden');
    return;
  }
  ensureChartStyle();
  const ui = getChartUi();
  renderHeroChart(g);
  chartsWrap?.classList.remove('hidden');

  const statusLabels = ['Prêtes', 'En cours', 'Erreurs'];
  const statusValues = [g.ready, g.processing, g.error];
  if (statusValues.some(v => v > 0)) {
    globalChartInstances.push(new Chart(document.getElementById('chart-global-status'), {
      type: 'doughnut',
      data: {
        labels: statusLabels,
        datasets: [doughnutDataset(statusValues, ui)],
      },
      options: {
        ...chartBaseOptions('État des réunions', ui),
        cutout: '68%',
        plugins: {
          ...chartBaseOptions('État des réunions', ui).plugins,
          doughnutCenter: {
            display: true,
            value: g.total,
            label: 'réunions',
            color: ui.text,
            subColor: ui.muted,
          },
        },
      },
    }));
  }

  const monthKeys = Object.keys(g.byMonth).sort().slice(-6);
  if (monthKeys.length) {
    const lineCanvas = document.getElementById('chart-global-volume');
    globalChartInstances.push(new Chart(lineCanvas, {
      type: 'line',
      data: {
        labels: monthKeys.map(k => {
          const [y, mo] = k.split('-');
          return `${mo}/${y.slice(2)}`;
        }),
        datasets: [{
          label: 'Réunions',
          data: monthKeys.map(k => g.byMonth[k]),
          borderColor: '#5B5FC7',
          backgroundColor: 'rgba(91, 95, 199, 0.12)',
          fill: true,
          tension: 0.42,
          borderWidth: 3,
          pointRadius: 5,
          pointHoverRadius: 7,
          pointBackgroundColor: '#fff',
          pointBorderColor: '#5B5FC7',
          pointBorderWidth: 2,
        }],
      },
      options: {
        ...chartBaseOptions('Activité (6 derniers mois)', ui),
        plugins: { ...chartBaseOptions('Activité (6 derniers mois)', ui).plugins, legend: { display: false } },
        scales: {
          x: { ticks: { color: ui.muted, font: { size: 11 } }, grid: { display: false } },
          y: { beginAtZero: true, ticks: { stepSize: 1, color: ui.muted }, grid: { color: ui.grid } },
        },
      },
    }));
  }

  const owners = Object.entries(g.actionOwners).sort((a, b) => b[1] - a[1]).slice(0, 8);
  if (owners.length) {
    const barCanvas = document.getElementById('chart-global-actions');
    globalChartInstances.push(new Chart(barCanvas, {
      type: 'bar',
      data: {
        labels: owners.map(([n]) => n),
        datasets: [{
          label: 'Actions',
          data: owners.map(([, v]) => v),
          borderRadius: 10,
          borderSkipped: false,
          backgroundColor: (ctx) => {
            const { chart } = ctx;
            const { ctx: c, chartArea } = chart;
            if (!chartArea) return '#5B5FC7';
            const hues = [['#464775', '#8B5CF6'], ['#5B5FC7', '#A5B4FC'], ['#059669', '#34D399']];
            const pair = hues[ctx.dataIndex % hues.length];
            return barGradient(c, chartArea, pair[0], pair[1]);
          },
        }],
      },
      options: {
        ...chartBaseOptions('Actions par responsable', ui),
        plugins: { ...chartBaseOptions('Actions par responsable', ui).plugins, legend: { display: false } },
        scales: {
          x: { ticks: { color: ui.muted }, grid: { display: false } },
          y: { beginAtZero: true, ticks: { stepSize: 1, color: ui.muted }, grid: { color: ui.grid } },
        },
      },
    }));
  }
}

function renderDetailKpis(m) {
  const el = document.getElementById('detail-kpi');
  if (!el) return;
  const s = coerceSummary(m.summary);
  if (!s || m.status !== 'ready') {
    el.classList.add('hidden');
    el.innerHTML = '';
    return;
  }
  const highlights = s.presentation?.highlights?.length || s.tldr?.length || 0;
  const done = (m.doneActions || []).length;
  const totalActions = s.actions?.length || 0;
  const progress = totalActions ? Math.round((done / totalActions) * 100) : 0;

  el.classList.remove('hidden');
  el.innerHTML = `
    <div class="kpi-card kpi-violet"><div class="kpi-body"><div class="kpi-value">${s.decisions?.length || 0}</div><div class="kpi-label">Décisions</div></div></div>
    <div class="kpi-card kpi-blue"><div class="kpi-body"><div class="kpi-value">${totalActions}</div><div class="kpi-label">Actions</div></div></div>
    <div class="kpi-card kpi-green"><div class="kpi-body"><div class="kpi-value">${highlights}</div><div class="kpi-label">Points clés</div></div></div>
    <div class="kpi-card kpi-orange"><div class="kpi-body"><div class="kpi-value">${s.open_questions?.length || 0}</div><div class="kpi-label">Questions</div></div></div>
    <div class="kpi-card kpi-brand kpi-progress">
      <div class="kpi-body">
        <div class="kpi-value">${progress}%</div>
        <div class="kpi-label">Actions cochées (${done}/${totalActions})</div>
        <div class="progress-bar"><div class="progress-fill" style="width:${progress}%"></div></div>
      </div>
    </div>
  `;
}

// ─── Liste ───
function renderMeetingsList() {
  const el = document.getElementById('meetings-list');
  const q = (document.getElementById('search-input')?.value || '').toLowerCase();
  const filter = document.getElementById('filter-status')?.value || 'all';

  const filtered = meetings.filter(m => {
    if (filter !== 'all' && m.status !== filter) return false;
    if (!q) return true;
    const hay = `${m.title} ${m.transcript || ''} ${JSON.stringify(m.summary || {})}`.toLowerCase();
    return hay.includes(q);
  });

  renderDashboard();

  if (!filtered.length) {
    el.innerHTML = meetings.length
      ? '<div class="empty">Aucun résultat pour cette recherche.</div>'
      : '<div class="empty"><div class="empty-icon">📋</div>Importez un fichier .vtt Teams pour commencer.</div>';
    return;
  }

  el.innerHTML = filtered.map(m => `
    <div class="meeting-item" data-id="${m.id}">
      <div class="meeting-item-accent status-${m.status}"></div>
      <div class="meeting-item-main">
        <h3>${esc(m.title)}</h3>
        <p class="meeting-date">${new Date(m.date).toLocaleString('fr-FR', { dateStyle: 'medium', timeStyle: 'short' })}</p>
        ${meetingMiniStats(m)}
      </div>
      <div class="meeting-meta">
        ${hasReadableReport(m) ? '<button type="button" class="meeting-report-btn" data-report="' + m.id + '">Rapport</button>' : ''}
        ${m.status === 'error' ? '<button type="button" class="meeting-report-btn meeting-report-retry" data-retry="' + m.id + '">Réessayer</button>' : ''}
        <span class="badge ${m.status}">${statusLabel(m.status)}</span>
        <button class="meeting-delete" data-del="${m.id}" title="Supprimer">✕</button>
      </div>
    </div>
  `).join('');

  el.querySelectorAll('.meeting-item').forEach(item => {
    item.addEventListener('click', e => {
      if (e.target.closest('.meeting-delete') || e.target.closest('.meeting-report-btn') || e.target.closest('.meeting-report-retry')) return;
      const id = item.dataset.id;
      openMeeting(meetings.find(m => String(m.id) === id)?.id ?? id);
    });
  });
  el.querySelectorAll('.meeting-report-btn[data-report]').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const id = btn.dataset.report;
      openMeeting(meetings.find(m => String(m.id) === id)?.id ?? id);
      switchTab('report');
      renderReportTab(findMeeting(id));
    });
  });
  el.querySelectorAll('.meeting-report-retry').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const id = btn.dataset.retry;
      currentMeetingId = findMeeting(id)?.id ?? id;
      openMeeting(currentMeetingId);
      if (!settings.apiKey) { showView('settings'); toast('Ajoutez votre clé Groq', 'error'); }
      else reprocessMeeting();
    });
  });
  el.querySelectorAll('.meeting-delete').forEach(btn => {
    btn.addEventListener('click', e => { e.stopPropagation(); deleteMeeting(btn.dataset.del); });
  });
}

// ─── Détail ───
function openMeeting(id) {
  currentMeetingId = id;
  const m = findMeeting(id);
  if (!m) return;
  m.summary = coerceSummary(m.summary);

  document.getElementById('detail-title').textContent = m.title;
  document.getElementById('detail-status').textContent = statusLabel(m.status);
  document.getElementById('detail-status').className = `badge ${m.status}`;

  const readyReport = m.status === 'ready' && hasReadableReport(m);
  const readyAi = m.status === 'ready' && hasAiSummary(m);
  ['read-report-btn'].forEach(btnId => { document.getElementById(btnId).hidden = !readyReport; });
  ['export-pptx-btn', 'copy-summary-btn', 'export-md-btn'].forEach(btnId => {
    document.getElementById(btnId).hidden = !readyAi;
  });
  updateReportQuickBar(m);
  document.getElementById('reprocess-btn').hidden = !(m.status === 'error' || (m.transcript && !hasAiSummary(m)));
  if (!document.getElementById('reprocess-btn').hidden) {
    document.getElementById('reprocess-btn').textContent = m.status === 'error' ? 'Réessayer' : 'Générer le rapport IA';
  }
  document.getElementById('delete-meeting-btn').hidden = false;
  document.getElementById('export-pptx-btn').disabled = false;
  document.getElementById('export-pptx-btn').textContent = 'PowerPoint';

  const chartsToggle = document.getElementById('charts-toggle');
  const chartsPanel = document.getElementById('charts-panel');
  chartsPanel.classList.add('hidden');
  chartsToggle.classList.add('hidden');
  chartsToggle.classList.remove('active');
  chartsToggle.textContent = 'Graphiques';
  renderDetailKpis(m);

  const defaultTab = (m.status === 'ready' && hasReadableReport(m)) ? 'report' : 'summary';
  switchTab(defaultTab);
  document.getElementById('transcript-content').textContent = m.transcript || '(vide)';
  document.getElementById('transcript-search').value = '';

  const panel = document.getElementById('summary-content');
  if (m.status === 'processing') {
    destroyCharts();
    renderDetailKpis(m);
    panel.innerHTML = '<div class="loading-box"><div class="spinner"></div>Génération du résumé en cours…</div>';
  } else if (m.status === 'error') {
    destroyCharts();
    renderDetailKpis(m);
    panel.innerHTML = `<div class="error-box">${esc(m.error || 'Erreur')}</div>`;
  } else if (m.summary) {
    const html = renderSummary(m.summary, m);
    panel.innerHTML = html || '<div class="empty-report">Résumé vide. Cliquez sur <strong>Réessayer</strong> pour régénérer.</div>';
    renderReportTab(m);
    try {
      renderCharts(m.summary);
      const hasCharts = !document.getElementById('charts-toggle').classList.contains('hidden');
      if (hasCharts) {
        chartsPanel.classList.remove('hidden');
        chartsToggle.classList.remove('hidden');
        chartsToggle.classList.add('active');
        chartsToggle.textContent = 'Masquer graphiques';
      }
    } catch { /* graphiques optionnels */ }
    bindActionCheckboxes(m);
  } else {
    destroyCharts();
    panel.innerHTML = '<div class="empty-report">Aucun résumé disponible. Cliquez sur <strong>Réessayer</strong>.</div>';
  }

  renderReportTab(m);
  renderChat(m);
  renderSuggestions();
  showView('detail');
  views.meetings.classList.add('hidden');
  views.detail.classList.remove('hidden');
}

function renderSummary(s, meeting) {
  const summary = coerceSummary(s) || {};
  let html = '';
  const pres = summary.presentation || {};

  if (pres.executive_summary) {
    html += `<div class="exec-summary">${esc(pres.executive_summary)}</div>`;
  }
  if (pres.highlights?.length) {
    html += `<div class="summary-section"><h4>Points clés</h4>${pres.highlights.map(h => `
      <div class="highlight-card"><strong>${esc(h.title || h)}</strong>${h.detail ? `<span>${esc(h.detail)}</span>` : ''}</div>
    `).join('')}</div>`;
  } else if (summary.tldr?.length) {
    html += `<div class="summary-section"><h4>En bref</h4><ul>${summary.tldr.map(p => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
  }
  if (summary.decisions?.length) {
    html += `<div class="summary-section"><h4>Décisions</h4><ul>${summary.decisions.map(p => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
  }
  if (summary.actions?.length) {
    html += `<div class="summary-section"><h4>Actions</h4>${summary.actions.map((a, i) => {
      const done = meeting?.doneActions?.includes(i);
      return `<div class="action-item${done ? ' done' : ''}"><input type="checkbox" data-action="${i}" ${done ? 'checked' : ''}><span><strong>${esc(a.who)}</strong> — ${esc(a.what)}${a.when && a.when !== 'non précisé' ? `<span class="action-deadline">${esc(a.when)}</span>` : ''}</span></div>`;
    }).join('')}</div>`;
  }
  if (summary.topics?.length) {
    html += `<div class="summary-section"><h4>Thèmes</h4>${summary.topics.map(t => `
      <div class="topic-card"><h5>${esc(t.title)}</h5>${t.summary ? `<p>${esc(t.summary)}</p>` : ''}${t.key_points?.length ? `<ul>${t.key_points.map(p => `<li>${esc(p)}</li>`).join('')}</ul>` : ''}</div>
    `).join('')}</div>`;
  }
  if (summary.explanation_for_absent) {
    html += `<div class="summary-section"><h4>Pour les absents</h4><p class="explanation">${esc(summary.explanation_for_absent)}</p></div>`;
  }
  if (summary.open_questions?.length) {
    html += `<div class="summary-section"><h4>Questions en suspens</h4><ul>${summary.open_questions.map(p => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
  }
  return html;
}

function renderReportTab(m) {
  const el = document.getElementById('report-tab-content');
  if (!el) return;

  if (m?.status === 'processing') {
    el.innerHTML = '<div class="loading-box"><div class="spinner"></div>Génération du rapport en cours…</div>';
    return;
  }
  if (m?.status === 'error') {
    el.innerHTML = `<div class="error-box">${esc(m.error || 'Erreur lors de la génération.')}</div>
      <p class="empty-report">Cliquez sur <strong>Réessayer</strong> en haut à droite.</p>`;
    return;
  }

  m.summary = coerceSummary(m.summary);
  if (!hasReadableReport(m)) {
    el.innerHTML = `<div class="empty-report">
      <p>Aucun rapport pour l’instant.</p>
      <p>Importez un fichier <strong>.vtt</strong> ou cliquez sur <strong>Réessayer</strong> après avoir configuré votre clé Groq.</p>
    </div>`;
    return;
  }

  const { title, dateStr, body } = buildReportContent(m);
  el.innerHTML = `
    <article class="report-article">
      <header class="report-article-head">
        <h2>${esc(title)}</h2>
        <p class="report-meta">Compte-rendu · ${esc(dateStr)} · TeamsBrief</p>
      </header>
      ${body}
    </article>
  `;
}

function bindActionCheckboxes(meeting) {
  document.querySelectorAll('[data-action]').forEach(cb => {
    cb.addEventListener('change', () => {
      if (!meeting.doneActions) meeting.doneActions = [];
      const i = +cb.dataset.action;
      if (cb.checked) { if (!meeting.doneActions.includes(i)) meeting.doneActions.push(i); }
      else meeting.doneActions = meeting.doneActions.filter(x => x !== i);
      cb.closest('.action-item').classList.toggle('done', cb.checked);
      saveMeetings();
      renderDetailKpis(meeting);
    });
  });
}

function filterTranscript() {
  const q = document.getElementById('transcript-search').value.trim();
  const m = meetings.find(x => x.id === currentMeetingId);
  if (!m) return;
  const text = m.transcript || '';
  const el = document.getElementById('transcript-content');
  if (!q) { el.textContent = text; return; }
  const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
  el.innerHTML = esc(text).replace(re, match => `<mark>${match}</mark>`);
}

function summaryToMarkdown(m) {
  const s = m.summary;
  if (!s) return '';
  let md = `# ${m.title}\n\n*${new Date(m.date).toLocaleString('fr-FR')}*\n\n`;
  if (s.presentation?.executive_summary) md += `## Synthèse\n${s.presentation.executive_summary}\n\n`;
  if (s.tldr?.length) md += `## En bref\n${s.tldr.map(p => `- ${p}`).join('\n')}\n\n`;
  if (s.decisions?.length) md += `## Décisions\n${s.decisions.map(p => `- ${p}`).join('\n')}\n\n`;
  if (s.actions?.length) md += `## Actions\n${s.actions.map(a => `- [ ] **${a.who}** — ${a.what}${a.when ? ` (${a.when})` : ''}`).join('\n')}\n\n`;
  if (s.explanation_for_absent) md += `## Pour les absents\n${s.explanation_for_absent}\n\n`;
  if (s.open_questions?.length) md += `## Questions\n${s.open_questions.map(p => `- ${p}`).join('\n')}\n`;
  return md;
}

async function copySummary() {
  const m = meetings.find(x => x.id === currentMeetingId);
  if (!m?.summary) return;
  await navigator.clipboard.writeText(summaryToMarkdown(m));
  toast('Résumé copié');
}

function exportMarkdown() {
  const m = meetings.find(x => x.id === currentMeetingId);
  if (!m?.summary) return;
  const blob = new Blob([summaryToMarkdown(m)], { type: 'text/markdown' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${m.title.slice(0, 40)}.md`;
  a.click();
  toast('Markdown exporté');
}

let reportMeetingCache = null;

function buildReportContent(m) {
  const s = coerceSummary(m.summary) || {};
  const pres = s.presentation || {};
  const dateStr = new Date(m.date).toLocaleString('fr-FR', { dateStyle: 'long', timeStyle: 'short' });
  const section = (title, body) => body
    ? `<div class="report-section"><h4>${esc(title)}</h4>${body}</div>`
    : '';

  let body = '';
  if (pres.executive_summary) {
    body += section('Synthèse exécutive', `<p class="report-lead">${esc(pres.executive_summary)}</p>`);
  }
  if (pres.highlights?.length) {
    body += section('Points clés', pres.highlights.map(h => `
      <div class="report-card"><strong>${esc(h.title || h)}</strong>${h.detail ? `<p>${esc(h.detail)}</p>` : ''}</div>
    `).join(''));
  } else if (s.tldr?.length) {
    body += section('En bref', `<ul>${s.tldr.map(p => `<li>${esc(p)}</li>`).join('')}</ul>`);
  }
  if (s.decisions?.length) {
    body += section('Décisions', `<ol>${s.decisions.map(d => `<li>${esc(d)}</li>`).join('')}</ol>`);
  }
  if (s.actions?.length) {
    body += section('Plan d\'actions', `<table><thead><tr><th>Responsable</th><th>Action</th><th>Échéance</th></tr></thead><tbody>
      ${s.actions.map(a => `<tr><td>${esc(a.who || '—')}</td><td>${esc(a.what || '')}</td><td>${esc(a.when && a.when !== 'non précisé' ? a.when : '—')}</td></tr>`).join('')}
    </tbody></table>`);
  }
  if (s.topics?.length) {
    body += section('Thèmes', s.topics.map(t => `
      <div class="report-card"><strong>${esc(t.title || 'Thème')}</strong>${t.summary ? `<p>${esc(t.summary)}</p>` : ''}
      ${t.key_points?.length ? `<ul>${t.key_points.map(p => `<li>${esc(p)}</li>`).join('')}</ul>` : ''}</div>
    `).join(''));
  }
  if (s.explanation_for_absent) {
    body += section('Pour les absents', `<p>${esc(s.explanation_for_absent)}</p>`);
  }
  if (s.open_questions?.length) {
    body += section('Questions en suspens', `<ul>${s.open_questions.map(q => `<li>${esc(q)}</li>`).join('')}</ul>`);
  }
  if (!body.trim() && (m.transcript || '').trim().length > 80) {
    const full = m.transcript.trim();
    const excerpt = full.slice(0, 14000);
    body += section(
      'Transcription de l’appel',
      `<pre class="report-transcript">${esc(excerpt)}${full.length > excerpt.length ? '\n\n[… transcription tronquée dans l’aperçu — voir l’onglet Transcription]' : ''}</pre>`
    );
    body += `<p class="report-hint">Pour un <strong>compte-rendu structuré</strong> (décisions, actions), vérifiez votre clé Groq dans Paramètres puis cliquez <strong>Réessayer</strong>.</p>`;
  }

  return {
    title: m.title,
    dateStr,
    body: body || '<p>Aucun contenu dans ce rapport.</p>',
  };
}

function buildReportHtml(m) {
  const { title, dateStr, body } = buildReportContent(m);
  return `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>${esc(title)} — Rapport</title>
<style>
  body{font-family:Segoe UI,system-ui,sans-serif;max-width:820px;margin:2rem auto;padding:0 1.5rem;color:#1e293b;line-height:1.6}
  h1{color:#464775;margin-bottom:.25rem} .meta{color:#64748b;font-size:.9rem;margin-bottom:2rem}
  h4{color:#464775;font-size:.85rem;text-transform:uppercase;letter-spacing:.05em;margin:1.5rem 0 .5rem}
  .report-lead{background:#f8fafc;border-left:4px solid #5b5fc7;padding:1rem 1.25rem;border-radius:.5rem}
  .report-card{background:#f8fafc;border:1px solid #e2e8f0;border-radius:.5rem;padding:.85rem 1rem;margin-bottom:.5rem}
  table{width:100%;border-collapse:collapse;font-size:.9rem}
  th,td{border:1px solid #e2e8f0;padding:.5rem .65rem;text-align:left;vertical-align:top}
  th{background:#464775;color:#fff} ul,ol{padding-left:1.25rem}
</style></head><body>
  <h1>${esc(title)}</h1>
  <p class="meta">Compte-rendu · ${esc(dateStr)} · Généré par TeamsBrief</p>
  ${body}
</body></html>`;
}

function showReportOverlay(m) {
  const { title, dateStr, body } = buildReportContent(m);
  reportMeetingCache = m;
  document.getElementById('report-modal-title').textContent = title;
  document.getElementById('report-body').innerHTML = `
    <h2 style="font-size:1.25rem;color:var(--brand);margin-bottom:.25rem">${esc(title)}</h2>
    <p class="report-meta">Compte-rendu · ${esc(dateStr)}</p>
    ${body}
  `;
  document.getElementById('report-overlay').classList.remove('hidden');
  document.body.style.overflow = 'hidden';
}

function closeReportOverlay() {
  document.getElementById('report-overlay').classList.add('hidden');
  document.body.style.overflow = '';
}

function downloadReportForCurrent() {
  const m = findMeeting(currentMeetingId);
  if (!m || !hasReadableReport(m)) {
    toast('Aucun rapport à télécharger', 'error');
    return;
  }
  reportMeetingCache = m;
  downloadCurrentReport();
}

function openReportOverlayForCurrent() {
  const m = findMeeting(currentMeetingId);
  if (!m || !hasReadableReport(m)) {
    toast('Aucun rapport à afficher', 'error');
    return;
  }
  showReportOverlay(m);
}

function downloadCurrentReport() {
  const m = reportMeetingCache || findMeeting(currentMeetingId);
  if (!m || !hasReadableReport(m)) return;
  const blob = new Blob([buildReportHtml(m)], { type: 'text/html;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${(m.title || 'rapport').slice(0, 40)}.html`;
  a.click();
  URL.revokeObjectURL(a.href);
  toast('Rapport téléchargé');
}

function openReadableReport() {
  const m = findMeeting(currentMeetingId);
  if (!m) return;
  switchTab('report');
  renderReportTab(m);
  document.getElementById('report-tab-content')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  if (hasReadableReport(m)) toast('Rapport affiché');
  else toast('Configurez Groq puis Réessayer pour générer le rapport', 'error');
}

function clearChat() {
  const m = meetings.find(x => x.id === currentMeetingId);
  if (!m || !confirm('Effacer la conversation ?')) return;
  m.chat = [];
  saveMeetings();
  renderChat(m);
  toast('Conversation effacée');
}

// ─── Chat ───
function renderChat(m) {
  const el = document.getElementById('chat-messages');
  if (!m.chat?.length) {
    el.innerHTML = '<div class="chat-empty">Posez une question sur cette réunion</div>';
    return;
  }
  el.innerHTML = m.chat.map(msg => `
    <div class="chat-bubble ${msg.role}">
      <div class="label">${msg.role === 'user' ? 'Vous' : 'Agent'}</div>
      <div class="content">${esc(msg.content)}</div>
    </div>
  `).join('');
  el.scrollTop = el.scrollHeight;
}

function renderSuggestions() {
  const el = document.getElementById('chat-suggestions');
  el.innerHTML = SUGGESTIONS.map(q =>
    `<button type="button" data-q="${esc(q)}">${esc(q)}</button>`
  ).join('');
  el.querySelectorAll('button').forEach(btn => {
    btn.addEventListener('click', () => {
      document.getElementById('chat-input').value = btn.dataset.q;
      document.getElementById('chat-form').dispatchEvent(new Event('submit'));
    });
  });
}

async function handleChat(e) {
  e.preventDefault();
  const input = document.getElementById('chat-input');
  const question = input.value.trim();
  if (!question || !currentMeetingId) return;

  const m = meetings.find(x => x.id === currentMeetingId);
  if (!m || m.status !== 'ready') return;

  input.value = '';
  m.chat.push({ role: 'user', content: question });
  renderChat(m);

  const bubble = document.createElement('div');
  bubble.className = 'chat-bubble assistant typing';
  bubble.innerHTML = '<div class="label">Agent</div><div class="content">Réflexion…</div>';
  document.getElementById('chat-messages').appendChild(bubble);
  const contentEl = bubble.querySelector('.content');

  try {
    const answer = await callGroq([
      {
        role: 'system',
        content: `Tu réponds aux questions sur une réunion Teams, en français, uniquement à partir du contexte fourni.

TRANSCRIPTION :
${m.transcript.slice(0, 20000)}

RÉSUMÉ :
${JSON.stringify(m.summary, null, 2)}`,
      },
      ...m.chat.slice(0, -1).map(c => ({ role: c.role, content: c.content })),
      { role: 'user', content: question },
    ]);
    bubble.classList.remove('typing');
    contentEl.textContent = answer;
    m.chat.push({ role: 'assistant', content: answer });
  } catch (err) {
    bubble.classList.remove('typing');
    contentEl.textContent = err.message || 'Erreur agent.';
  }

  saveMeetings();
}

function esc(s) {
  const d = document.createElement('div');
  d.textContent = s || '';
  return d.innerHTML;
}

// ─── Graphiques (Chart.js + données PPTX) ───
const CHART_COLORS = ['#5B5FC7', '#6264A7', '#14B8A6', '#059669', '#F59E0B', '#EC4899', '#8B5CF6', '#464775'];
let chartsStyled = false;

const doughnutCenterPlugin = {
  id: 'doughnutCenter',
  afterDraw(chart) {
    const opts = chart.options.plugins?.doughnutCenter;
    if (!opts?.display) return;
    const { ctx, chartArea } = chart;
    if (!chartArea) return;
    const x = (chartArea.left + chartArea.right) / 2;
    const y = (chartArea.top + chartArea.bottom) / 2;
    ctx.save();
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.font = `700 ${opts.valueSize || 30}px "Segoe UI", system-ui, sans-serif`;
    ctx.fillStyle = opts.color || '#464775';
    ctx.fillText(String(opts.value ?? ''), x, y - 10);
    ctx.font = `600 ${opts.labelSize || 11}px "Segoe UI", system-ui, sans-serif`;
    ctx.fillStyle = opts.subColor || '#64748b';
    ctx.fillText(opts.label || '', x, y + 18);
    ctx.restore();
  },
};

function ensureChartStyle() {
  if (chartsStyled || typeof Chart === 'undefined') return;
  chartsStyled = true;
  Chart.register(doughnutCenterPlugin);
  Chart.defaults.font.family = '"Segoe UI", system-ui, sans-serif';
  Chart.defaults.color = '#64748b';
  Chart.defaults.animation.duration = 1100;
  Chart.defaults.animation.easing = 'easeOutQuart';
}

function getChartUi() {
  const isDark = document.documentElement.dataset.theme === 'dark';
  return {
    text: isDark ? '#e2e8f0' : '#334155',
    muted: isDark ? '#94a3b8' : '#64748b',
    grid: isDark ? 'rgba(148, 163, 184, 0.15)' : 'rgba(148, 163, 184, 0.35)',
    card: isDark ? '#1a1d27' : '#ffffff',
    title: { display: true, font: { size: 13, weight: '600' }, color: isDark ? '#e2e8f0' : '#464775', padding: { bottom: 8 } },
    legend: {
      position: 'bottom',
      labels: { boxWidth: 10, boxHeight: 10, useBorderRadius: true, borderRadius: 4, padding: 14, font: { size: 11, weight: '500' } },
    },
  };
}

function chartBaseOptions(title, ui) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    layout: { padding: { top: 4, bottom: 4, left: 4, right: 4 } },
    plugins: { legend: ui.legend, title: { ...ui.title, text: title } },
  };
}

function doughnutDataset(values, ui) {
  return {
    data: values,
    backgroundColor: CHART_COLORS.slice(0, values.length),
    borderWidth: 3,
    borderColor: ui.card,
    hoverOffset: 14,
    spacing: 2,
  };
}

function barGradient(ctx, area, from, to) {
  const g = ctx.createLinearGradient(0, area.bottom, 0, area.top);
  g.addColorStop(0, from);
  g.addColorStop(1, to);
  return g;
}

function renderHeroChart(g) {
  ensureChartStyle();
  const wrap = document.getElementById('chart-hero-wrap');
  const canvas = document.getElementById('chart-hero');
  const caption = document.getElementById('chart-hero-caption');
  if (!wrap || !canvas) return;

  const items = [
    { label: 'Décisions', value: g.decisions },
    { label: 'Actions', value: g.actions },
    { label: 'Questions', value: g.questions },
    { label: 'Thèmes', value: g.topicsCount || 0 },
  ].filter(x => x.value > 0);

  if (!items.length) {
    wrap.classList.add('hidden');
    return;
  }

  wrap.classList.remove('hidden');
  const total = items.reduce((a, b) => a + b.value, 0);
  if (caption) {
    caption.textContent = `${total} éléments de compte-rendu répartis sur ${g.ready} réunion(s) prête(s). Passez la souris sur le graphique pour le détail.`;
  }

  const ui = getChartUi();
  const chart = new Chart(canvas, {
    type: 'polarArea',
    data: {
      labels: items.map(x => x.label),
      datasets: [{
        data: items.map(x => x.value),
        backgroundColor: CHART_COLORS.slice(0, items.length).map(c => c + 'D9'),
        borderColor: CHART_COLORS.slice(0, items.length),
        borderWidth: 2,
      }],
    },
    options: {
      ...chartBaseOptions('', ui),
      plugins: {
        ...chartBaseOptions('', ui).plugins,
        legend: { ...ui.legend, position: 'right' },
        title: { display: false },
      },
      scales: {
        r: {
          beginAtZero: true,
          grid: { color: ui.grid },
          ticks: { display: false },
          pointLabels: { color: ui.text, font: { size: 12, weight: '600' } },
        },
      },
    },
  });
  globalChartInstances.push(chart);
}

function destroyCharts() {
  chartInstances.forEach(c => c.destroy());
  chartInstances = [];
}

function buildChartMetrics(summary) {
  const distLabels = [];
  const distValues = [];
  const distMap = [
    ['Décisions', summary.decisions?.length || 0],
    ['Actions', summary.actions?.length || 0],
    ['Thèmes', summary.topics?.length || 0],
    ['Questions', summary.open_questions?.length || 0],
    ['Points clés', (summary.presentation?.highlights || summary.tldr || []).length],
  ];
  distMap.forEach(([l, v]) => { if (v > 0) { distLabels.push(l); distValues.push(v); } });

  const actionCounts = {};
  (summary.actions || []).forEach(a => {
    const who = (a.who || 'Non assigné').slice(0, 25);
    actionCounts[who] = (actionCounts[who] || 0) + 1;
  });

  const topics = summary.topics || [];
  const topicLabels = topics.map(t => (t.title || 'Thème').slice(0, 22));
  const topicValues = topics.map(t => (t.key_points?.length || 1));

  const highlights = summary.presentation?.highlights || summary.tldr || [];
  const hlLabels = highlights.map((h, i) => {
    const t = typeof h === 'object' ? h.title : h;
    return (t || `Point ${i + 1}`).slice(0, 28);
  });
  const hlValues = highlights.map((h, i) => {
    const d = typeof h === 'object' ? (h.detail || '') : '';
    return Math.max(1, Math.round((d.length || String(h).length) / 40));
  });

  const radarLabels = ['Décisions', 'Actions', 'Thèmes', 'Questions', 'Points clés'];
  const radarValues = [
    summary.decisions?.length || 0,
    summary.actions?.length || 0,
    summary.topics?.length || 0,
    summary.open_questions?.length || 0,
    highlights.length,
  ];

  const statusLabels = [];
  const statusValues = [];
  const withDeadline = (summary.actions || []).filter(a => a.when && a.when !== 'non précisé').length;
  const withoutDeadline = (summary.actions || []).length - withDeadline;
  if (withDeadline) { statusLabels.push('Avec échéance'); statusValues.push(withDeadline); }
  if (withoutDeadline) { statusLabels.push('Sans échéance'); statusValues.push(withoutDeadline); }
  if (summary.decisions?.length) { statusLabels.push('Décisions'); statusValues.push(summary.decisions.length); }

  return {
    distribution: { labels: distLabels, values: distValues },
    actions: { labels: Object.keys(actionCounts), values: Object.values(actionCounts) },
    topics: { labels: topicLabels, values: topicValues },
    highlights: { labels: hlLabels, values: hlValues },
    radar: { labels: radarLabels, values: radarValues },
    status: { labels: statusLabels, values: statusValues },
  };
}

function toggleCharts() {
  const panel = document.getElementById('charts-panel');
  const btn = document.getElementById('charts-toggle');
  const show = panel.classList.contains('hidden');
  panel.classList.toggle('hidden', !show);
  btn.classList.toggle('active', show);
  btn.textContent = show ? 'Masquer graphiques' : 'Graphiques';
}

function renderCharts(summary) {
  destroyCharts();
  const panel = document.getElementById('charts-panel');
  const toggle = document.getElementById('charts-toggle');
  panel.classList.add('hidden');
  toggle.classList.add('hidden');
  toggle.classList.remove('active');
  toggle.textContent = 'Graphiques';

  if (!summary || typeof Chart === 'undefined') return;

  const m = buildChartMetrics(summary);
  const hasData = m.distribution.values.length || m.actions.values.length || m.topics.values.length;
  if (!hasData) return;

  toggle.classList.remove('hidden');
  ensureChartStyle();
  const ui = getChartUi();
  const totalDist = m.distribution.values.reduce((a, b) => a + b, 0);

  if (m.distribution.values.length) {
    chartInstances.push(new Chart(document.getElementById('chart-distribution'), {
      type: 'doughnut',
      data: {
        labels: m.distribution.labels,
        datasets: [doughnutDataset(m.distribution.values, ui)],
      },
      options: {
        ...chartBaseOptions('Répartition du compte-rendu', ui),
        cutout: '62%',
        plugins: {
          ...chartBaseOptions('Répartition du compte-rendu', ui).plugins,
          doughnutCenter: { display: true, value: totalDist, label: 'éléments', color: ui.text, subColor: ui.muted },
        },
      },
    }));
  }

  if (m.actions.values.length) {
    chartInstances.push(new Chart(document.getElementById('chart-actions'), {
      type: 'bar',
      data: {
        labels: m.actions.labels,
        datasets: [{
          label: 'Actions',
          data: m.actions.values,
          borderRadius: 10,
          borderSkipped: false,
          backgroundColor: (ctx) => {
            const { chart } = ctx;
            const { ctx: c, chartArea } = chart;
            if (!chartArea) return '#5B5FC7';
            return barGradient(c, chartArea, '#464775', '#5B5FC7');
          },
        }],
      },
      options: {
        ...chartBaseOptions('Actions par responsable', ui),
        plugins: { ...chartBaseOptions('Actions par responsable', ui).plugins, legend: { display: false } },
        scales: {
          x: { ticks: { color: ui.muted }, grid: { display: false } },
          y: { beginAtZero: true, ticks: { stepSize: 1, color: ui.muted }, grid: { color: ui.grid } },
        },
      },
    }));
  }

  if (m.topics.values.length) {
    chartInstances.push(new Chart(document.getElementById('chart-topics'), {
      type: 'bar',
      data: {
        labels: m.topics.labels,
        datasets: [{
          label: 'Points abordés',
          data: m.topics.values,
          borderRadius: 10,
          borderSkipped: false,
          backgroundColor: (ctx) => {
            const { chart } = ctx;
            const { ctx: c, chartArea } = chart;
            if (!chartArea) return '#059669';
            return barGradient(c, chartArea, '#047857', '#34D399');
          },
        }],
      },
      options: {
        ...chartBaseOptions('Profondeur par thème', ui),
        indexAxis: 'y',
        plugins: { ...chartBaseOptions('Profondeur par thème', ui).plugins, legend: { display: false } },
        scales: {
          x: { beginAtZero: true, ticks: { stepSize: 1, color: ui.muted }, grid: { color: ui.grid } },
          y: { ticks: { color: ui.text, font: { weight: '600' } }, grid: { display: false } },
        },
      },
    }));
  }

  if (m.radar.values.some(v => v > 0)) {
    chartInstances.push(new Chart(document.getElementById('chart-radar'), {
      type: 'radar',
      data: {
        labels: m.radar.labels,
        datasets: [{
          label: 'Intensité',
          data: m.radar.values,
          backgroundColor: 'rgba(91, 95, 199, 0.22)',
          borderColor: '#5B5FC7',
          borderWidth: 2,
          pointBackgroundColor: '#fff',
          pointBorderColor: '#464775',
          pointBorderWidth: 2,
          pointRadius: 4,
          pointHoverRadius: 6,
        }],
      },
      options: {
        ...chartBaseOptions('Vue radar synthétique', ui),
        scales: {
          r: {
            beginAtZero: true,
            grid: { color: ui.grid },
            angleLines: { color: ui.grid },
            pointLabels: { color: ui.text, font: { size: 11, weight: '600' } },
            ticks: { display: false },
          },
        },
      },
    }));
  }

  if (m.highlights.values.length) {
    chartInstances.push(new Chart(document.getElementById('chart-highlights'), {
      type: 'polarArea',
      data: {
        labels: m.highlights.labels,
        datasets: [{
          data: m.highlights.values,
          backgroundColor: CHART_COLORS.map(c => c + 'CC'),
          borderColor: CHART_COLORS,
          borderWidth: 2,
        }],
      },
      options: {
        ...chartBaseOptions('Poids des points clés', ui),
        scales: {
          r: {
            grid: { color: ui.grid },
            ticks: { display: false },
            pointLabels: { color: ui.muted, font: { size: 10 } },
          },
        },
      },
    }));
  }

  if (m.status.values.length) {
    chartInstances.push(new Chart(document.getElementById('chart-status'), {
      type: 'doughnut',
      data: {
        labels: m.status.labels,
        datasets: [doughnutDataset(m.status.values, ui)],
      },
      options: {
        ...chartBaseOptions('Suivi & décisions', ui),
        cutout: '55%',
      },
    }));
  }
}

// ─── Export PowerPoint (v2 — pro, lisible, fiable) ───
const C = {
  blue: '464775', lightBlue: '6264A7', accent: '5B5FC7', white: 'FFFFFF',
  dark: '212C40', gray: '6B7280', lightGray: 'F8FAFC', muted: '94A3B8',
  footer: '94A3B8', green: '059669', orange: 'D97706', border: 'E2E8F0',
};

function safeFilename(title) {
  const name = (title || 'reunion').replace(/[<>:"/\\|?*]/g, '').trim() || 'reunion';
  return `${new Date().toISOString().slice(0, 10)} - ${name.slice(0, 60)}.pptx`;
}

function pptxChunk(arr, n) {
  const out = [];
  for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n));
  return out;
}

function pptxTrunc(s, max = 140) {
  const t = String(s || '').replace(/\s+/g, ' ').trim();
  return t.length > max ? t.slice(0, max - 1) + '…' : t;
}

function pptxSplitParagraph(text, maxLen = 380) {
  const t = String(text || '').trim();
  if (t.length <= maxLen) return [t];
  const parts = [];
  let rest = t;
  while (rest.length > maxLen) {
    let cut = rest.lastIndexOf('. ', maxLen);
    if (cut < maxLen * 0.4) cut = rest.lastIndexOf(' ', maxLen);
    if (cut < 1) cut = maxLen;
    parts.push(rest.slice(0, cut + (rest[cut] === '.' ? 1 : 0)).trim());
    rest = rest.slice(cut).trim();
  }
  if (rest) parts.push(rest);
  return parts;
}

function createPptxDeck() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'TeamsBrief';
  pptx.company = 'TeamsBrief';
  return pptx;
}

function pptxSlide(pptx, state) {
  const slide = pptx.addSlide();
  state.n += 1;
  slide.addText(`TeamsBrief  ·  ${state.n}`, {
    x: 0.4, y: 7.08, w: 9.2, h: 0.3, fontSize: 8, color: C.footer, align: 'right',
  });
  return slide;
}

function pptxHeader(slide, title, subtitle = '') {
  slide.addShape('rect', { x: 0, y: 0, w: '100%', h: 1.05, fill: { color: C.blue }, line: { color: C.blue } });
  slide.addText(title, {
    x: 0.55, y: 0.2, w: 8.9, h: 0.55, fontSize: 20, bold: true, color: C.white, fontFace: 'Segoe UI',
  });
  if (subtitle) {
    slide.addText(subtitle, { x: 0.55, y: 0.72, w: 8.9, h: 0.35, fontSize: 11, color: C.muted, fontFace: 'Segoe UI' });
  }
  slide.addShape('rect', { x: 0, y: 1.05, w: '100%', h: 0.05, fill: { color: C.accent }, line: { color: C.accent } });
}

function pptxBullets(slide, items, opts = {}) {
  const { numbered = false, max = 5, y = 1.35, fontSize = 15 } = opts;
  const lines = items.slice(0, max).map((item, i) => {
    const text = numbered ? `${i + 1}. ${pptxTrunc(item, opts.itemMax || 160)}` : pptxTrunc(item, opts.itemMax || 160);
    return {
      text,
      options: {
        bullet: numbered ? false : true,
        breakLine: true,
        fontSize,
        color: C.dark,
        fontFace: 'Segoe UI',
        paraSpaceAfter: 12,
      },
    };
  });
  if (!lines.length) return;
  slide.addText(lines, { x: 0.65, y, w: 8.7, h: 5.5, valign: 'top', fit: 'shrink' });
}

function pptxBody(slide, text, y = 1.4) {
  slide.addText(text, {
    x: 0.65, y, w: 8.7, h: 5.4, fontSize: 16, color: C.dark,
    fontFace: 'Segoe UI', valign: 'top', lineSpacingMultiple: 1.35, fit: 'shrink',
  });
}

function pptxStatsSlide(pptx, state, summary) {
  const slide = pptxSlide(pptx, state);
  pptxHeader(slide, 'Vue d\'ensemble');
  const stats = [
    { label: 'Décisions', val: summary.decisions?.length || 0, color: C.green },
    { label: 'Actions', val: summary.actions?.length || 0, color: C.accent },
    { label: 'Thèmes', val: summary.topics?.length || 0, color: C.lightBlue },
    { label: 'Questions', val: summary.open_questions?.length || 0, color: C.orange },
  ];
  stats.forEach((s, i) => {
    const x = 0.7 + i * 2.25;
    slide.addShape('roundRect', { x, y: 2.0, w: 2.0, h: 3.2, fill: { color: C.white }, line: { color: C.border, width: 1 }, rectRadius: 0.1 });
    slide.addShape('rect', { x, y: 2.0, w: 2.0, h: 0.12, fill: { color: s.color }, line: { color: s.color } });
    slide.addText(String(s.val), { x, y: 2.6, w: 2.0, h: 1.0, fontSize: 40, bold: true, color: s.color, align: 'center', fontFace: 'Segoe UI' });
    slide.addText(s.label, { x, y: 3.8, w: 2.0, h: 0.6, fontSize: 12, color: C.gray, align: 'center', fontFace: 'Segoe UI' });
  });
}

function pptxActionsTable(pptx, state, actions) {
  const chunks = pptxChunk(actions, 6);
  chunks.forEach((group, ci) => {
    const slide = pptxSlide(pptx, state);
    pptxHeader(slide, 'Plan d\'actions', chunks.length > 1 ? `Suite ${ci + 1}/${chunks.length}` : '');
    const rows = [
      [
        { text: 'Responsable', options: { bold: true, color: C.blue } },
        { text: 'Action', options: { bold: true, color: C.blue } },
        { text: 'Échéance', options: { bold: true, color: C.blue } },
      ],
    ];
    group.forEach(a => {
      rows.push([
        { text: pptxTrunc(a.who || '—', 28), options: { bold: true, fontSize: 11 } },
        { text: pptxTrunc(a.what || '', 100), options: { fontSize: 11 } },
        { text: pptxTrunc(a.when && a.when !== 'non précisé' ? a.when : '—', 20), options: { fontSize: 10, color: C.orange } },
      ]);
    });
    slide.addTable(rows, {
      x: 0.5, y: 1.35, w: 9.0,
      colW: [1.8, 5.5, 1.7],
      border: { type: 'solid', color: C.border, pt: 0.5 },
      fontFace: 'Segoe UI',
      fontSize: 11,
    });
  });
}

function pptxKeyPointSlide(pptx, state, index, total, title, detail) {
  const slide = pptxSlide(pptx, state);
  pptxHeader(slide, `Point clé ${index} / ${total}`);
  slide.addShape('ellipse', { x: 0.65, y: 1.5, w: 0.55, h: 0.55, fill: { color: C.accent }, line: { color: C.accent } });
  slide.addText(String(index), { x: 0.65, y: 1.58, w: 0.55, h: 0.4, fontSize: 18, bold: true, color: C.white, align: 'center', fontFace: 'Segoe UI' });
  slide.addText(pptxTrunc(title, 120), {
    x: 0.65, y: 2.3, w: 8.7, h: 1.4, fontSize: 22, bold: true, color: C.dark, fontFace: 'Segoe UI', valign: 'top', fit: 'shrink',
  });
  if (detail) {
    slide.addText(pptxTrunc(detail, 300), {
      x: 0.65, y: 3.9, w: 8.7, h: 2.5, fontSize: 14, color: C.gray, fontFace: 'Segoe UI', valign: 'top', lineSpacingMultiple: 1.3, fit: 'shrink',
    });
  }
}

function pptxTopicSlide(pptx, state, topic) {
  const slide = pptxSlide(pptx, state);
  pptxHeader(slide, pptxTrunc(topic.title || 'Thème', 60));
  if (topic.summary) {
    slide.addText(pptxTrunc(topic.summary, 220), {
      x: 0.65, y: 1.3, w: 8.7, h: 1.0, fontSize: 13, italic: true, color: C.gray, fontFace: 'Segoe UI',
    });
  }
  if (topic.key_points?.length) {
    pptxBullets(slide, topic.key_points, { y: topic.summary ? 2.45 : 1.35, max: 6, itemMax: 130 });
  } else if (topic.summary) {
    pptxBody(slide, topic.summary, 2.2);
  }
}

function generatePresentation(title, summary, dateIso) {
  if (typeof PptxGenJS === 'undefined') throw new Error('Bibliothèque PowerPoint non chargée.');

  const pptx = createPptxDeck();
  pptx.title = title;
  const pres = summary.presentation || {};
  const dateStr = dateIso
    ? new Date(dateIso).toLocaleString('fr-FR', { dateStyle: 'long', timeStyle: 'short' })
    : new Date().toLocaleDateString('fr-FR', { dateStyle: 'long' });

  const state = { n: 0 };

  // ── Titre ──
  const titleSlide = pptxSlide(pptx, state);
  titleSlide.background = { color: C.blue };
  titleSlide.addShape('rect', { x: 0.8, y: 1.7, w: 0.1, h: 3.2, fill: { color: C.accent }, line: { color: C.accent } });
  titleSlide.addText(pptxTrunc(title, 80), {
    x: 1.05, y: 1.65, w: 8.3, h: 1.6, fontSize: 30, bold: true, color: C.white, fontFace: 'Segoe UI', valign: 'top', fit: 'shrink',
  });
  titleSlide.addText('Compte-rendu de réunion', {
    x: 1.05, y: 3.2, w: 8.3, h: 0.45, fontSize: 14, color: C.muted, fontFace: 'Segoe UI',
  });
  titleSlide.addText(dateStr, {
    x: 1.05, y: 3.65, w: 8.3, h: 0.4, fontSize: 13, color: 'B8C4E0', fontFace: 'Segoe UI',
  });
  if (pres.tagline) {
    titleSlide.addText(pptxTrunc(pres.tagline, 120), {
      x: 1.05, y: 4.4, w: 8.3, h: 0.8, fontSize: 12, italic: true, color: 'AABBDD', fontFace: 'Segoe UI',
    });
  }
  titleSlide.addText('Généré par TeamsBrief', { x: 1.05, y: 6.7, w: 4, h: 0.35, fontSize: 9, color: '8899BB', fontFace: 'Segoe UI' });

  // ── Sommaire ──
  const agenda = [];
  if (pres.executive_summary) agenda.push('Synthèse exécutive');
  const highlights = pres.highlights?.length ? pres.highlights : (summary.tldr || []);
  if (highlights.length) agenda.push('Points clés');
  if (summary.decisions?.length) agenda.push('Décisions');
  if (summary.actions?.length) agenda.push('Plan d\'actions');
  if (summary.topics?.length) agenda.push('Thèmes abordés');
  if (summary.explanation_for_absent) agenda.push('Pour les absents');
  if (summary.open_questions?.length) agenda.push('Questions en suspens');

  if (agenda.length) {
    const ag = pptxSlide(pptx, state);
    pptxHeader(ag, 'Sommaire');
    agenda.forEach((item, i) => {
      const y = 1.4 + i * 0.72;
      ag.addText(`${i + 1}.  ${item}`, {
        x: 0.8, y, w: 8.4, h: 0.55, fontSize: 16, color: C.dark, fontFace: 'Segoe UI',
      });
    });
  }

  // ── Stats visuelles (pas de charts PPTX cassés) ──
  pptxStatsSlide(pptx, state, summary);

  // ── Synthèse exécutive ──
  if (pres.executive_summary) {
    pptxSplitParagraph(pres.executive_summary, 350).forEach((part, i, arr) => {
      const slide = pptxSlide(pptx, state);
      pptxHeader(slide, 'Synthèse exécutive', arr.length > 1 ? `Partie ${i + 1}/${arr.length}` : '');
      pptxBody(slide, part);
    });
  }

  // ── Points clés (1 slide chacun) ──
  if (pres.highlights?.length) {
    pres.highlights.slice(0, 5).forEach((h, i) => {
      pptxKeyPointSlide(pptx, state, i + 1, Math.min(pres.highlights.length, 5),
        typeof h === 'object' ? h.title : h,
        typeof h === 'object' ? h.detail : '');
    });
  } else if (summary.tldr?.length) {
    summary.tldr.slice(0, 4).forEach((p, i) => {
      pptxKeyPointSlide(pptx, state, i + 1, Math.min(summary.tldr.length, 4), p, '');
    });
  }

  // ── Décisions (max 3 par slide, texte tronqué) ──
  if (summary.decisions?.length) {
    pptxChunk(summary.decisions, 3).forEach((group, i, arr) => {
      const slide = pptxSlide(pptx, state);
      pptxHeader(slide, 'Décisions prises', arr.length > 1 ? `${i + 1} / ${arr.length}` : '');
      pptxBullets(slide, group, { numbered: true, max: 3, itemMax: 180, fontSize: 15 });
    });
  }

  // ── Actions (tableau pro) ──
  if (summary.actions?.length) {
    pptxActionsTable(pptx, state, summary.actions);
  }

  // ── Thèmes ──
  (summary.topics || []).slice(0, 5).forEach(topic => {
    if (topic.summary || topic.key_points?.length) pptxTopicSlide(pptx, state, topic);
  });

  // ── Pour les absents ──
  if (summary.explanation_for_absent) {
    pptxSplitParagraph(summary.explanation_for_absent, 380).forEach((part, i, arr) => {
      const slide = pptxSlide(pptx, state);
      pptxHeader(slide, 'Pour les absents', arr.length > 1 ? `Partie ${i + 1}/${arr.length}` : '');
      pptxBody(slide, part);
    });
  }

  // ── Questions ──
  if (summary.open_questions?.length) {
    pptxChunk(summary.open_questions, 4).forEach((group, i, arr) => {
      const slide = pptxSlide(pptx, state);
      pptxHeader(slide, 'Questions en suspens', arr.length > 1 ? `${i + 1} / ${arr.length}` : '');
      pptxBullets(slide, group, { max: 4, itemMax: 150 });
    });
  }

  // ── Clôture ──
  const closing = pptxSlide(pptx, state);
  closing.background = { color: C.blue };
  closing.addText('Merci', {
    x: 0.5, y: 2.9, w: 9, h: 1.0, fontSize: 42, bold: true, color: C.white, align: 'center', fontFace: 'Segoe UI',
  });
  closing.addText('TeamsBrief — Synthèse de réunion Teams', {
    x: 0.5, y: 4.1, w: 9, h: 0.5, fontSize: 13, color: 'AABBDD', align: 'center', fontFace: 'Segoe UI',
  });

  return pptx;
}

async function exportPresentation() {
  const m = meetings.find(x => x.id === currentMeetingId);
  if (!m?.summary) return;

  const btn = document.getElementById('export-pptx-btn');
  btn.disabled = true;
  btn.textContent = 'Génération...';

  try {
    const pptx = generatePresentation(m.title, m.summary, m.date);
    await pptx.writeFile({ fileName: safeFilename(m.title) });
    toast('PowerPoint téléchargé');
  } catch (err) {
    toast(err.message || 'Erreur PowerPoint', 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = 'PowerPoint';
  }
}
