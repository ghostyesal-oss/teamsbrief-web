/* TeamsBrief Web — 100% navigateur, hébergé */

const STORAGE_KEY = 'teamsbrief_meetings';
const SETTINGS_KEY = 'teamsbrief_settings';
const THEME_KEY = 'teamsbrief_theme';

const STATUS_LABELS = { ready: 'Prêt', processing: 'En cours', error: 'Erreur' };
const GROQ_FALLBACK_MODEL = 'openai/gpt-oss-20b';
const GROQ_FETCH_TIMEOUT_MS = 90_000;
const PROCESSING_STALE_MS = 2 * 60 * 1000;
let generationResumeRunning = false;
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
  backup: document.getElementById('view-backup'),
  settings: document.getElementById('view-settings'),
};

// ─── Init ───
document.querySelectorAll('.nav-btn[data-view]').forEach(btn => {
  btn.addEventListener('click', () => showView(btn.dataset.view));
});
document.querySelectorAll('.inline-link[data-view]').forEach(btn => {
  btn.addEventListener('click', () => showView(btn.dataset.view));
});

document.getElementById('file-input').addEventListener('change', handleFileImport);
document.querySelectorAll('.hero-file-input').forEach(inp => {
  inp.addEventListener('change', handleFileImport);
});
document.getElementById('back-btn').addEventListener('click', () => showView('meetings'));
document.getElementById('save-settings').addEventListener('click', saveSettings);
document.getElementById('chat-form').addEventListener('submit', handleChat);
document.getElementById('export-pptx-btn').addEventListener('click', exportPresentation);
document.getElementById('delete-meeting-btn').addEventListener('click', deleteCurrentMeeting);
document.getElementById('download-report-inline')?.addEventListener('click', downloadReportForCurrent);
document.getElementById('open-report-overlay')?.addEventListener('click', openReportOverlayForCurrent);
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
reconcileMeetingStatuses();
renderMeetingsList();
renderDashboard();
resumeInterruptedGenerations();

// ─── Navigation ───
function showView(name) {
  Object.values(views).forEach(v => v.classList.add('hidden'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));

  if (name === 'detail') {
    views.detail.classList.remove('hidden');
    views.detail.classList.add('view-enter');
    requestAnimationFrame(() => views.detail.classList.remove('view-enter'));
  } else if (name === 'backup') {
    views.backup.classList.remove('hidden');
    views.backup.classList.add('view-enter');
    requestAnimationFrame(() => views.backup.classList.remove('view-enter'));
    document.querySelector('[data-view="backup"]').classList.add('active');
    renderBackupSummary();
  } else if (name === 'settings') {
    views.settings.classList.remove('hidden');
    views.settings.classList.add('view-enter');
    requestAnimationFrame(() => views.settings.classList.remove('view-enter'));
    document.querySelector('[data-view="settings"]').classList.add('active');
  } else if (name === 'meetings') {
    views.meetings.classList.remove('hidden');
    views.meetings.classList.add('view-enter');
    requestAnimationFrame(() => views.meetings.classList.remove('view-enter'));
    document.querySelector('[data-view="meetings"]').classList.add('active');
    currentMeetingId = null;
    renderMeetingsList();
    renderDashboard();
  } else {
    showView('meetings');
  }
}

function renderBackupSummary() {
  const el = document.getElementById('backup-summary');
  if (!el) return;
  const n = meetings.length;
  const ready = meetings.filter(m => hasReadableReport(m)).length;
  el.innerHTML = n
    ? `<p><strong>${n}</strong> réunion(s) enregistrée(s) localement · <strong>${ready}</strong> rapport(s) lisible(s).</p>`
    : '<p>Aucune réunion pour l’instant. Importez des fichiers .vtt depuis l’onglet Réunions.</p>';
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

function processingAgeMs(m) {
  const t = m.processingStartedAt || m.date;
  const ms = new Date(t).getTime();
  return Number.isNaN(ms) ? PROCESSING_STALE_MS + 1 : Date.now() - ms;
}

/** Corrige les statuts « En cours » orphelins (onglet fermé, rechargement, etc.). */
function reconcileMeetingStatuses() {
  let changed = false;
  for (const m of meetings) {
    if (m.status !== 'processing') continue;
    if (hasAiSummary(m)) {
      m.status = 'ready';
      delete m.processingStartedAt;
      changed = true;
      continue;
    }
    if (processingAgeMs(m) > PROCESSING_STALE_MS) {
      m.status = 'error';
      m.error = m.error || 'Génération interrompue (page fermée ou délai dépassé). Cliquez sur Réessayer.';
      delete m.processingStartedAt;
      changed = true;
    }
  }
  if (changed) saveMeetings();
  return changed;
}

async function fetchWithTimeout(url, options = {}, ms = GROQ_FETCH_TIMEOUT_MS) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), ms);
  try {
    return await fetch(url, { ...options, signal: ctrl.signal });
  } catch (e) {
    if (e.name === 'AbortError') {
      throw new Error('Délai Groq dépassé (90 s). Réessayez ou raccourcissez la transcription.');
    }
    throw e;
  } finally {
    clearTimeout(timer);
  }
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

function meetingInitials(title) {
  const words = String(title || 'R').trim().split(/\s+/).filter(Boolean);
  if (words.length >= 2) return (words[0][0] + words[words.length - 1][0]).toUpperCase();
  return (words[0] || 'R').slice(0, 2).toUpperCase();
}

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
  document.getElementById('theme-toggle').textContent = theme === 'dark' ? 'Clair' : 'Sombre';
  destroyGlobalCharts();
  if (!views.detail.classList.contains('hidden')) return;
  renderDashboard();
}

function toggleKeyVisibility() {
  const inp = document.getElementById('groq-key');
  const btn = document.getElementById('toggle-key');
  const show = inp.type === 'password';
  inp.type = show ? 'text' : 'password';
  if (btn) btn.textContent = show ? 'Masquer' : 'Afficher';
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

function updateReportQuickBar() { /* barre retirée — rapport dans l’onglet principal */ }

function exportBackup() {
  const blob = new Blob([JSON.stringify({ meetings, settings, exportedAt: new Date().toISOString() }, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `teamsbrief-backup-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  toast('Sauvegarde exportée');
  renderBackupSummary();
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
    renderBackupSummary();
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
  if (m.includes('TimeoutError') || m.includes('timed out') || m.includes('aborted')) {
    return 'Délai Groq dépassé (2 min). Vérifiez votre connexion puis réessayez.';
  }
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
  const res = await fetchWithTimeout(url, {
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

async function generateMeetingSummary(meeting, { fresh = false } = {}) {
  if (!settings.apiKey) {
    meeting.status = 'error';
    meeting.error = 'Clé Groq manquante. Paramètres → clé gsk_… → Enregistrer, puis Réessayer.';
    delete meeting.processingStartedAt;
    saveMeetings();
    return false;
  }
  meeting.status = 'processing';
  meeting.error = null;
  if (fresh || !meeting.processingStartedAt) {
    meeting.processingStartedAt = new Date().toISOString();
  }
  saveMeetings();
  renderMeetingsList();

  try {
    const raw = await callGroq([
      { role: 'system', content: 'Tu réponds uniquement en JSON valide, en français.' },
      { role: 'user', content: SUMMARY_PROMPT(meeting.title, meeting.transcript) },
    ], true);
    meeting.summary = coerceSummary(parseGroqJson(raw));
    meeting.status = 'ready';
    delete meeting.processingStartedAt;
    return true;
  } catch (err) {
    meeting.status = 'error';
    meeting.error = humanizeGroqError(err.message);
    delete meeting.processingStartedAt;
    return false;
  } finally {
    saveMeetings();
    renderMeetingsList();
  }
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
    processingStartedAt: new Date().toISOString(),
    chat: [],
    doneActions: [],
  };

  meetings.unshift(meeting);
  saveMeetings();
  renderMeetingsList();

  const ok = await generateMeetingSummary(meeting, { fresh: true });
  toast(ok ? `"${title}" prête` : `Erreur : ${title}`, ok ? 'success' : 'error');
  return meeting.id;
}

async function reprocessMeeting() {
  const m = meetings.find(x => x.id === currentMeetingId);
  if (!m) return;
  openMeeting(m.id);
  const ok = await generateMeetingSummary(m, { fresh: true });
  toast(ok ? 'Résumé régénéré' : (m.error || 'Échec'), ok ? 'success' : 'error');
  openMeeting(m.id);
}

async function resumeAllProcessing() {
  reconcileMeetingStatuses();
  const pending = meetings.filter(m =>
    m.status === 'processing'
    && !hasAiSummary(m)
    && (m.transcript || '').trim().length > 0
  );
  if (!pending.length) {
    toast('Aucune génération en attente', 'error');
    return;
  }
  if (!settings.apiKey) {
    toast('Ajoutez votre clé Groq dans Paramètres', 'error');
    showView('settings');
    return;
  }
  await runResumeInterruptedGenerations(pending);
}

async function runResumeInterruptedGenerations(pending) {
  if (generationResumeRunning) return;
  generationResumeRunning = true;
  toast(`Génération de ${pending.length} rapport(s)…`);
  try {
    for (const m of pending) {
      const ok = await generateMeetingSummary(m, { fresh: false });
      if (!ok) toast(`${m.title} : ${m.error || 'erreur'}`, 'error');
    }
    renderDashboard();
  } finally {
    generationResumeRunning = false;
  }
}

async function resumeInterruptedGenerations() {
  reconcileMeetingStatuses();
  if (!settings.apiKey) {
    let changed = false;
    for (const m of meetings) {
      if (m.status !== 'processing') continue;
      m.status = 'error';
      m.error = 'Clé Groq manquante. Paramètres → enregistrez votre clé puis Réessayer.';
      delete m.processingStartedAt;
      changed = true;
    }
    if (changed) {
      saveMeetings();
      renderMeetingsList();
      renderDashboard();
    }
    return;
  }
  const pending = meetings.filter(m =>
    m.status === 'processing'
    && !hasAiSummary(m)
    && (m.transcript || '').trim().length > 0
    && processingAgeMs(m) < PROCESSING_STALE_MS
  );
  if (!pending.length) return;
  await runResumeInterruptedGenerations(pending);
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
    inner += `<div class="reports-hub-grid">${reports.slice(0, 8).map(m => `
      <article class="report-card">
        <div class="report-card-avatar" aria-hidden="true">${esc(meetingInitials(m.title))}</div>
        <div class="report-card-body">
          <h3 class="report-card-title">${esc(m.title)}</h3>
          <p class="report-card-date">${new Date(m.date).toLocaleString('fr-FR', { dateStyle: 'medium', timeStyle: 'short' })}</p>
          ${hasAiSummary(m) ? '<span class="mini-chip chip-green">Synthèse IA</span>' : '<span class="mini-chip chip-orange">Transcription</span>'}
        </div>
        <button type="button" class="btn btn-primary btn-sm reports-hub-open" data-id="${esc(m.id)}">Ouvrir</button>
      </article>
    `).join('')}</div>`;
    if (reports.length > 8) inner += `<p class="reports-hub-more">+ ${reports.length - 8} autre(s) rapport(s) dans la liste ci-dessous.</p>`;
  } else {
    inner += `<div class="reports-hub-empty">
      <p><strong>Aucun rapport lisible pour l’instant.</strong></p>
      ${failed.length ? `<p>${failed.length} réunion(s) en erreur.</p>
        <button type="button" class="btn btn-ghost btn-sm" id="retry-all-failed">Réessayer tout</button>` : ''}
      ${pending.length ? `<p>${pending.length} génération(s) bloquée(s) ou en cours.</p>
        <button type="button" class="btn btn-primary btn-sm" id="resume-all-processing">Relancer la génération</button>` : ''}
      ${!failed.length && !pending.length ? '<p>Importez un fichier .vtt ou restaurez une sauvegarde via l’onglet <strong>Backup</strong>.</p>' : ''}
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
  document.getElementById('resume-all-processing')?.addEventListener('click', () => {
    resumeAllProcessing();
  });
  document.getElementById('retry-all-failed')?.addEventListener('click', () => {
    retryAllFailedMeetings();
  });
}

async function retryAllFailedMeetings() {
  if (!settings.apiKey) {
    toast('Ajoutez votre clé Groq dans Paramètres', 'error');
    showView('settings');
    return;
  }
  const failed = meetings.filter(m =>
    m.status === 'error' && (m.transcript || '').trim().length > 0
  );
  if (!failed.length) {
    toast('Aucune réunion en erreur', 'error');
    return;
  }
  toast(`Nouvelle tentative pour ${failed.length} réunion(s)…`);
  for (const m of failed) {
    await generateMeetingSummary(m, { fresh: true });
  }
  renderDashboard();
}

function renderDashboard() {
  const kpiEl = document.getElementById('dashboard-kpis');
  if (!kpiEl) return;

  renderReportsHub();

  const g = computeGlobalStats();
  const reportsReady = meetings.filter(m => hasReadableReport(m)).length;
  const subtitle = document.getElementById('meetings-stats');
  if (subtitle) {
    subtitle.textContent = g.total
      ? `${g.total} réunion(s) · ${reportsReady} rapport(s) · ${g.actions} action(s)`
      : 'Importez une transcription Teams (.vtt) pour générer vos rapports d’appel.';
  }

  const kpiDefs = [
    { label: 'Réunions', value: g.total, tone: 'brand', hint: 'importées' },
    { label: 'Rapports', value: reportsReady, tone: 'green', hint: 'lisibles' },
    { label: 'Actions', value: g.actions, tone: 'blue', hint: 'à suivre' },
    { label: 'Erreurs', value: g.error, tone: 'red', hint: 'à corriger' },
  ];
  kpiEl.innerHTML = kpiDefs.map((k, i) => `
    <div class="kpi-card kpi-${k.tone}" style="--stagger:${i}">
      <div class="kpi-icon-wrap" aria-hidden="true"></div>
      <div class="kpi-body">
        <div class="kpi-value">${k.value}</div>
        <div class="kpi-label">${k.label}</div>
        <div class="kpi-hint">${k.hint}</div>
      </div>
    </div>
  `).join('');
}

function renderDetailKpis() {}

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
      : '<div class="empty"><div class="empty-icon" aria-hidden="true"></div><p>Importez une transcription Teams (.vtt) pour générer votre premier rapport.</p></div>';
    return;
  }

  el.innerHTML = filtered.map((m, i) => `
    <div class="meeting-item" data-id="${m.id}" style="--stagger:${i % 12}">
      <div class="meeting-item-accent status-${m.status}"></div>
      <div class="meeting-avatar" aria-hidden="true">${esc(meetingInitials(m.title))}</div>
      <div class="meeting-item-main">
        <h3>${esc(m.title)}</h3>
        <p class="meeting-date">${new Date(m.date).toLocaleString('fr-FR', { dateStyle: 'medium', timeStyle: 'short' })}</p>
        ${meetingMiniStats(m)}
      </div>
      <div class="meeting-meta">
        ${hasReadableReport(m) ? '<button type="button" class="meeting-report-btn" data-report="' + m.id + '">Rapport</button>' : ''}
        ${m.status === 'error' || m.status === 'processing' ? '<button type="button" class="meeting-report-btn meeting-report-retry" data-retry="' + m.id + '">' + (m.status === 'processing' ? 'Relancer' : 'Réessayer') + '</button>' : ''}
        <span class="badge ${m.status}">${statusLabel(m.status)}</span>
        <button class="meeting-delete" data-del="${m.id}" title="Supprimer" aria-label="Supprimer">×</button>
        <span class="meeting-chevron" aria-hidden="true"></span>
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
  if (m.status === 'processing' && hasAiSummary(m)) {
    m.status = 'ready';
    delete m.processingStartedAt;
    saveMeetings();
  }

  document.getElementById('detail-title').textContent = m.title;
  const metaEl = document.getElementById('detail-meta');
  if (metaEl) {
    metaEl.textContent = new Date(m.date).toLocaleString('fr-FR', { dateStyle: 'long', timeStyle: 'short' });
  }
  document.getElementById('detail-status').textContent = statusLabel(m.status);
  document.getElementById('detail-status').className = `badge ${m.status}`;

  const readyAi = m.status === 'ready' && hasAiSummary(m);
  const pptxBtn = document.getElementById('export-pptx-btn');
  if (pptxBtn) {
    pptxBtn.hidden = !readyAi;
    pptxBtn.disabled = !readyAi;
  }
  updateReportQuickBar(m);

  const defaultTab = (m.status === 'ready' && hasReadableReport(m)) ? 'report' : 'summary';
  switchTab(defaultTab);
  document.getElementById('transcript-content').textContent = m.transcript || '(vide)';
  document.getElementById('transcript-search').value = '';

  const panel = document.getElementById('summary-content');
  if (m.status === 'processing' && !hasAiSummary(m)) {
    panel.innerHTML = '<div class="loading-box"><div class="spinner"></div>Génération du résumé en cours…</div>';
  } else if (m.status === 'error') {
    panel.innerHTML = `<div class="error-box">${esc(m.error || 'Erreur')}</div><p class="report-hint">Relancez depuis la liste avec le bouton <strong>Réessayer</strong>.</p>`;
  } else if (m.summary) {
    const html = renderSummary(m.summary, m);
    panel.innerHTML = html || '<div class="empty-report">Résumé vide. Relancez la génération depuis la liste des réunions.</div>';
    renderReportTab(m);
    bindActionCheckboxes(m);
  } else {
    panel.innerHTML = '<div class="empty-report">Aucun résumé disponible. Configurez Groq dans Paramètres puis relancez depuis la liste.</div>';
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
