/* TeamsBrief Web — 100% navigateur, hébergé */

const STORAGE_KEY = 'teamsbrief_meetings';
const SETTINGS_KEY = 'teamsbrief_settings';
const THEME_KEY = 'teamsbrief_theme';

const STATUS_LABELS = { ready: 'Prêt', processing: 'En cours', error: 'Erreur' };

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
  }
}

// ─── Storage ───
function loadMeetings() {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]'); }
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
  const m = meetings.find(x => x.id === currentMeetingId);
  if (m?.summary) renderCharts(m.summary);
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
  document.getElementById('tab-transcript').classList.toggle('hidden', name !== 'transcript');
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
async function callGroq(messages, jsonMode = false) {
  if (!settings.apiKey) throw new Error('Configurez votre clé Groq dans Paramètres.');

  const payload = {
    model: settings.model || 'openai/gpt-oss-20b',
    messages,
    temperature: 0.3,
    max_tokens: 2048,
  };
  if (jsonMode) payload.response_format = { type: 'json_object' };

  // Proxy : même domaine, puis Vercel (pour GitHub Pages / autres hôtes statiques)
  const endpoints = [
    '/api/groq',
    'https://teamsbrief-web.vercel.app/api/groq',
    'https://api.groq.com/openai/v1/chat/completions',
  ];

  for (const url of endpoints) {
    try {
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
        if (url !== endpoints[endpoints.length - 1]) continue;
        throw new Error(msg);
      }
      const data = await res.json();
      return data.choices[0].message.content;
    } catch (e) {
      if (url === endpoints[endpoints.length - 1]) throw e;
    }
  }
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
    meeting.summary = JSON.parse(raw);
    meeting.status = 'ready';
    toast(`"${title}" prête`);
  } catch (err) {
    meeting.status = 'error';
    meeting.error = err.message;
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
    m.summary = JSON.parse(raw);
    m.status = 'ready';
    toast('Résumé régénéré');
  } catch (err) {
    m.status = 'error';
    m.error = err.message;
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

  const ready = meetings.filter(m => m.status === 'ready').length;
  document.getElementById('meetings-stats').textContent = meetings.length
    ? `${meetings.length} réunion(s) · ${ready} prête(s)`
    : '';

  if (!filtered.length) {
    el.innerHTML = meetings.length
      ? '<div class="empty">Aucun résultat pour cette recherche.</div>'
      : '<div class="empty"><div class="empty-icon">📋</div>Importez un fichier .vtt Teams pour commencer.</div>';
    return;
  }

  el.innerHTML = filtered.map(m => `
    <div class="meeting-item" data-id="${m.id}">
      <div class="meeting-item-main">
        <h3>${esc(m.title)}</h3>
        <p>${new Date(m.date).toLocaleString('fr-FR')}${m.summary?.decisions?.length ? ` · ${m.summary.decisions.length} décision(s)` : ''}</p>
      </div>
      <div class="meeting-meta">
        <span class="badge ${m.status}">${statusLabel(m.status)}</span>
        <button class="meeting-delete" data-del="${m.id}" title="Supprimer">✕</button>
      </div>
    </div>
  `).join('');

  el.querySelectorAll('.meeting-item').forEach(item => {
    item.addEventListener('click', e => {
      if (e.target.closest('.meeting-delete')) return;
      const id = item.dataset.id;
      openMeeting(meetings.find(m => String(m.id) === id)?.id ?? id);
    });
  });
  el.querySelectorAll('.meeting-delete').forEach(btn => {
    btn.addEventListener('click', e => { e.stopPropagation(); deleteMeeting(btn.dataset.del); });
  });
}

// ─── Détail ───
function openMeeting(id) {
  currentMeetingId = id;
  const m = meetings.find(x => x.id === id);
  if (!m) return;

  document.getElementById('detail-title').textContent = m.title;
  document.getElementById('detail-status').textContent = statusLabel(m.status);
  document.getElementById('detail-status').className = `badge ${m.status}`;

  const ready = m.status === 'ready' && m.summary;
  ['export-pptx-btn', 'copy-summary-btn', 'export-md-btn'].forEach(id => {
    document.getElementById(id).hidden = !ready;
  });
  document.getElementById('reprocess-btn').hidden = m.status !== 'error';
  document.getElementById('delete-meeting-btn').hidden = false;
  document.getElementById('export-pptx-btn').disabled = false;
  document.getElementById('export-pptx-btn').textContent = 'PowerPoint';

  const chartsToggle = document.getElementById('charts-toggle');
  const chartsPanel = document.getElementById('charts-panel');
  chartsPanel.classList.add('hidden');
  chartsToggle.classList.add('hidden');
  chartsToggle.classList.remove('active');
  chartsToggle.textContent = 'Graphiques';

  switchTab('summary');
  document.getElementById('transcript-content').textContent = m.transcript || '(vide)';
  document.getElementById('transcript-search').value = '';

  const panel = document.getElementById('summary-content');
  if (m.status === 'processing') {
    destroyCharts();
    panel.innerHTML = '<div class="loading-box"><div class="spinner"></div>Génération du résumé en cours…</div>';
  } else if (m.status === 'error') {
    destroyCharts();
    panel.innerHTML = `<div class="error-box">${esc(m.error || 'Erreur')}</div>`;
  } else if (m.summary) {
    panel.innerHTML = renderSummary(m, m);
    renderCharts(m.summary);
    bindActionCheckboxes(m);
  } else {
    destroyCharts();
  }

  renderChat(m);
  renderSuggestions();
  showView('detail');
  views.meetings.classList.add('hidden');
  views.detail.classList.remove('hidden');
}

function renderSummary(s, meeting) {
  let html = '';
  const pres = s.presentation || {};

  if (pres.executive_summary) {
    html += `<div class="exec-summary">${esc(pres.executive_summary)}</div>`;
  }
  if (pres.highlights?.length) {
    html += `<div class="summary-section"><h4>Points clés</h4>${pres.highlights.map(h => `
      <div class="highlight-card"><strong>${esc(h.title || h)}</strong>${h.detail ? `<span>${esc(h.detail)}</span>` : ''}</div>
    `).join('')}</div>`;
  } else if (s.tldr?.length) {
    html += `<div class="summary-section"><h4>En bref</h4><ul>${s.tldr.map(p => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
  }
  if (s.decisions?.length) {
    html += `<div class="summary-section"><h4>Décisions</h4><ul>${s.decisions.map(p => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
  }
  if (s.actions?.length) {
    html += `<div class="summary-section"><h4>Actions</h4>${s.actions.map((a, i) => {
      const done = meeting?.doneActions?.includes(i);
      return `<div class="action-item${done ? ' done' : ''}"><input type="checkbox" data-action="${i}" ${done ? 'checked' : ''}><span><strong>${esc(a.who)}</strong> — ${esc(a.what)}${a.when && a.when !== 'non précisé' ? `<span class="action-deadline">${esc(a.when)}</span>` : ''}</span></div>`;
    }).join('')}</div>`;
  }
  if (s.topics?.length) {
    html += `<div class="summary-section"><h4>Thèmes</h4>${s.topics.map(t => `
      <div class="topic-card"><h5>${esc(t.title)}</h5>${t.summary ? `<p>${esc(t.summary)}</p>` : ''}${t.key_points?.length ? `<ul>${t.key_points.map(p => `<li>${esc(p)}</li>`).join('')}</ul>` : ''}</div>
    `).join('')}</div>`;
  }
  if (s.explanation_for_absent) {
    html += `<div class="summary-section"><h4>Pour les absents</h4><p class="explanation">${esc(s.explanation_for_absent)}</p></div>`;
  }
  if (s.open_questions?.length) {
    html += `<div class="summary-section"><h4>Questions en suspens</h4><ul>${s.open_questions.map(p => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
  }
  return html;
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
const CHART_COLORS = ['#464775', '#6264A7', '#5B5FC7', '#059669', '#D97706', '#DC2626', '#2563EB', '#7C3AED'];
let chartInstances = [];

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

  const base = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 10 } } } },
  };

  if (m.distribution.values.length) {
    chartInstances.push(new Chart(document.getElementById('chart-distribution'), {
      type: 'doughnut',
      data: {
        labels: m.distribution.labels,
        datasets: [{ data: m.distribution.values, backgroundColor: CHART_COLORS, borderWidth: 2, borderColor: '#fff' }],
      },
      options: { ...base, plugins: { ...base.plugins, title: { display: true, text: 'Répartition du compte-rendu', font: { size: 12 } } } },
    }));
  }

  if (m.actions.values.length) {
    chartInstances.push(new Chart(document.getElementById('chart-actions'), {
      type: 'bar',
      data: {
        labels: m.actions.labels,
        datasets: [{ label: 'Actions', data: m.actions.values, backgroundColor: '#5B5FC7', borderRadius: 6 }],
      },
      options: {
        ...base,
        plugins: { ...base.plugins, title: { display: true, text: 'Actions par responsable', font: { size: 12 } }, legend: { display: false } },
        scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } },
      },
    }));
  }

  if (m.topics.values.length) {
    chartInstances.push(new Chart(document.getElementById('chart-topics'), {
      type: 'bar',
      data: {
        labels: m.topics.labels,
        datasets: [{ label: 'Points abordés', data: m.topics.values, backgroundColor: '#059669', borderRadius: 6 }],
      },
      options: {
        ...base,
        indexAxis: 'y',
        plugins: { ...base.plugins, title: { display: true, text: 'Profondeur par thème', font: { size: 12 } }, legend: { display: false } },
        scales: { x: { beginAtZero: true, ticks: { stepSize: 1 } } },
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
          backgroundColor: 'rgba(91, 95, 199, 0.2)',
          borderColor: '#5B5FC7',
          pointBackgroundColor: '#464775',
        }],
      },
      options: { ...base, plugins: { ...base.plugins, title: { display: true, text: 'Vue radar synthétique', font: { size: 12 } } }, scales: { r: { beginAtZero: true } } },
    }));
  }

  if (m.highlights.values.length) {
    chartInstances.push(new Chart(document.getElementById('chart-highlights'), {
      type: 'polarArea',
      data: {
        labels: m.highlights.labels,
        datasets: [{ data: m.highlights.values, backgroundColor: CHART_COLORS.map(c => c + 'CC') }],
      },
      options: { ...base, plugins: { ...base.plugins, title: { display: true, text: 'Poids des points clés', font: { size: 12 } } } },
    }));
  }

  if (m.status.values.length) {
    chartInstances.push(new Chart(document.getElementById('chart-status'), {
      type: 'pie',
      data: {
        labels: m.status.labels,
        datasets: [{ data: m.status.values, backgroundColor: ['#059669', '#D97706', '#464775'], borderWidth: 2, borderColor: '#fff' }],
      },
      options: { ...base, plugins: { ...base.plugins, title: { display: true, text: 'Suivi & décisions', font: { size: 12 } } } },
    }));
  }
}

// ─── Export PowerPoint (design premium + animations) ───
const C = {
  blue: '464775', lightBlue: '6264A7', accent: '5B5FC7', white: 'FFFFFF',
  dark: '212C40', gray: '6B7280', lightGray: 'F3F4F6', muted: 'CCCCDD',
  footer: '9999BB', green: '059669', orange: 'D97706',
};
const TRANS = ['fade', 'push', 'wipe', 'dissolve'];

function safeFilename(title) {
  const name = (title || 'reunion').replace(/[<>:"/\\|?*]/g, '').trim() || 'reunion';
  return `${new Date().toISOString().slice(0, 10)} - ${name.slice(0, 60)}.pptx`;
}

function chunk(arr, n = 4) {
  const out = [];
  for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n));
  return out;
}

function setTrans(slide, idx) {
  slide.transition = { type: TRANS[idx % TRANS.length], durationMs: 800 };
}

function addFooter(slide, num) {
  slide.addText(`TeamsBrief  ·  Slide ${num}`, {
    x: 0.5, y: 7.05, w: 9, h: 0.35, fontSize: 9, color: C.footer, align: 'right',
  });
}

function addHeader(pptx, slide, title, icon = '', color = C.lightBlue) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 1.15, fill: { color }, line: { color } });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 1.15, w: '100%', h: 0.06, fill: { color: C.accent }, line: { color: C.accent } });
  slide.addText(`${icon ? icon + '  ' : ''}${title}`, {
    x: 0.55, y: 0.22, w: 8.9, h: 0.75, fontSize: 22, bold: true, color: C.white,
  });
}

function addDividerSlide(pptx, title, icon, slideNum) {
  const slide = pptx.addSlide({ background: { color: C.blue } });
  setTrans(slide, slideNum);
  slide.addText(`${icon ? icon + '  ' : ''}${title}`, {
    x: 0.8, y: 3.0, w: 8.4, h: 1.2, fontSize: 36, bold: true, color: C.white, align: 'center',
  });
  addFooter(slide, slideNum);
}

function addSectionSlide(pptx, heading, items, icon, numbered, slideNum) {
  const slide = pptx.addSlide({ background: { color: C.lightGray } });
  setTrans(slide, slideNum);
  addHeader(pptx, slide, heading, icon);
  items.forEach((item, i) => {
    const y = 1.55 + i * 1.1;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.75, y, w: 8.5, h: 0.95, fill: { color: C.white }, line: { color: 'E5E7EB', width: 0.75 }, rectRadius: 0.08,
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.75, y, w: 0.1, h: 0.95, fill: { color: numbered ? C.green : C.accent }, line: { color: numbered ? C.green : C.accent },
    });
    slide.addText(`${numbered ? i + 1 + '. ' : '→ '}${item}`, {
      x: 1.05, y: y + 0.12, w: 8.0, h: 0.75, fontSize: 16, color: C.dark,
    });
  });
  addFooter(slide, slideNum);
}

function addHighlightSlide(pptx, index, total, text, detail, slideNum) {
  const slide = pptx.addSlide({ background: { color: C.white } });
  setTrans(slide, slideNum);
  slide.addShape(pptx.ShapeType.ellipse, { x: 0.7, y: 1.5, w: 1.1, h: 1.1, fill: { color: C.accent }, line: { color: C.accent } });
  slide.addText(String(index), { x: 0.7, y: 1.65, w: 1.1, h: 0.8, fontSize: 32, bold: true, color: C.white, align: 'center' });
  slide.addText(`Point clé ${index}/${total}`, { x: 2.1, y: 1.55, w: 7.0, h: 0.5, fontSize: 13, color: C.gray });
  slide.addText(text, { x: 0.9, y: 2.9, w: 8.2, h: 2.0, fontSize: 26, bold: true, color: C.dark, valign: 'top' });
  if (detail) slide.addText(detail, { x: 0.9, y: 4.5, w: 8.2, h: 1.5, fontSize: 15, color: C.gray, valign: 'top' });
  slide.addShape(pptx.ShapeType.rect, { x: 0.9, y: 5.8, w: 2.5, h: 0.08, fill: { color: C.accent }, line: { color: C.accent } });
  addFooter(slide, slideNum);
}

function addActionsSlide(pptx, actions, slideNum) {
  const slide = pptx.addSlide({ background: { color: C.lightGray } });
  setTrans(slide, slideNum);
  addHeader(pptx, slide, 'Actions à faire', '✅', C.green);
  actions.forEach((a, i) => {
    const y = 1.5 + i * 1.2;
    slide.addShape(pptx.ShapeType.roundRect, { x: 0.7, y, w: 8.6, h: 1.05, fill: { color: C.white }, line: { color: 'E5E7EB' }, rectRadius: 0.06 });
    slide.addText('☐', { x: 0.95, y: y + 0.2, w: 0.5, h: 0.5, fontSize: 20, color: C.accent });
    slide.addText(a.who || '?', { x: 1.5, y: y + 0.15, w: 2.2, h: 0.45, fontSize: 13, bold: true, color: C.accent });
    slide.addText(a.what || '', { x: 3.8, y: y + 0.15, w: 4.0, h: 0.75, fontSize: 14, color: C.dark });
    if (a.when && a.when !== 'non précisé') {
      slide.addShape(pptx.ShapeType.roundRect, { x: 8.0, y: y + 0.25, w: 1.2, h: 0.5, fill: { color: 'FEF3C7' }, line: { color: 'FEF3C7' } });
      slide.addText(a.when, { x: 8.0, y: y + 0.3, w: 1.2, h: 0.4, fontSize: 9, color: C.orange, align: 'center' });
    }
  });
  addFooter(slide, slideNum);
}

function addChartSlide(pptx, heading, chartType, chartData, slideNum, opts = {}) {
  const slide = pptx.addSlide({ background: { color: C.white } });
  setTrans(slide, slideNum);
  addHeader(pptx, slide, heading, '📊', C.blue);
  slide.addChart(chartType, chartData, {
    x: 0.6, y: 1.45, w: 8.8, h: 5.1,
    showLegend: true,
    legendPos: 'b',
    legendFontSize: 9,
    chartColors: CHART_COLORS.map(c => c.replace('#', '')),
    showTitle: false,
    barDir: opts.barDir || 'col',
    ...opts,
  });
  addFooter(slide, slideNum);
}

function addTopicSlide(pptx, topic, slideNum) {
  const slide = pptx.addSlide({ background: { color: C.white } });
  setTrans(slide, slideNum);
  addHeader(pptx, slide, topic.title || 'Thème', '💡', C.blue);
  if (topic.summary) {
    slide.addText(topic.summary, { x: 0.8, y: 1.45, w: 8.4, h: 1.2, fontSize: 15, italic: true, color: C.gray });
  }
  (topic.key_points || []).slice(0, 5).forEach((pt, i) => {
    const y = (topic.summary ? 2.8 : 1.55) + i * 0.75;
    slide.addShape(pptx.ShapeType.ellipse, { x: 0.95, y: y + 0.08, w: 0.18, h: 0.18, fill: { color: C.accent }, line: { color: C.accent } });
    slide.addText(pt, { x: 1.3, y, w: 8.0, h: 0.7, fontSize: 15, color: C.dark });
  });
  addFooter(slide, slideNum);
}

function generatePresentation(title, summary, dateIso) {
  if (typeof PptxGenJS === 'undefined') throw new Error('Bibliothèque PowerPoint non chargée.');

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'TeamsBrief';
  pptx.title = title;

  const pres = summary.presentation || {};
  const dateStr = dateIso
    ? new Date(dateIso).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' })
    : new Date().toLocaleDateString('fr-FR');

  let sn = 1;

  const titleSlide = pptx.addSlide({ background: { color: C.blue } });
  setTrans(titleSlide, 0);
  titleSlide.addShape(pptx.ShapeType.ellipse, { x: 7.5, y: -0.5, w: 3.5, h: 3.5, fill: { color: C.lightBlue, transparency: 92 }, line: { color: C.lightBlue } });
  titleSlide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.9, w: 0.12, h: 2.8, fill: { color: C.accent }, line: { color: C.accent } });
  titleSlide.addText(title, { x: 1.1, y: 1.85, w: 8.2, h: 1.8, fontSize: 34, bold: true, color: C.white });
  titleSlide.addText(`Compte-rendu · ${dateStr}`, { x: 1.1, y: 3.6, w: 8.2, h: 0.7, fontSize: 15, color: C.muted });
  if (pres.tagline) titleSlide.addText(pres.tagline, { x: 1.1, y: 4.5, w: 8.2, h: 1.0, fontSize: 13, italic: true, color: 'AABBDD' });
  sn++;

  const sections = [];
  if (pres.executive_summary || summary.tldr?.length) sections.push('Synthèse exécutive');
  if (summary.decisions?.length) sections.push('Décisions prises');
  if (summary.actions?.length) sections.push("Plan d'actions");
  if (summary.topics?.length) sections.push('Thèmes abordés');
  if (summary.explanation_for_absent) sections.push('Pour les absents');
  if (summary.open_questions?.length) sections.push('Questions en suspens');

  if (sections.length) {
    const agSlide = pptx.addSlide({ background: { color: C.lightGray } });
    setTrans(agSlide, 1);
    addHeader(pptx, agSlide, 'Sommaire', '📋');
    sections.forEach((s, i) => {
      const y = 1.55 + i * 0.88;
      agSlide.addShape(pptx.ShapeType.roundRect, { x: 0.9, y, w: 8.2, h: 0.72, fill: { color: C.white }, line: { color: 'E5E7EB' }, rectRadius: 0.06 });
      agSlide.addShape(pptx.ShapeType.ellipse, { x: 1.15, y: y + 0.12, w: 0.48, h: 0.48, fill: { color: C.accent }, line: { color: C.accent } });
      agSlide.addText(String(i + 1), { x: 1.15, y: y + 0.15, w: 0.48, h: 0.42, fontSize: 14, bold: true, color: C.white, align: 'center' });
      agSlide.addText(s, { x: 1.85, y: y + 0.1, w: 7.0, h: 0.55, fontSize: 16, color: C.dark });
    });
    addFooter(agSlide, sn++);
  }

  const metrics = buildChartMetrics(summary);

  // Slide graphiques — répartition (donut)
  if (metrics.distribution.values.length) {
    addChartSlide(pptx, 'Répartition du compte-rendu', pptx.ChartType.doughnut, [{
      name: 'Contenu',
      labels: metrics.distribution.labels,
      values: metrics.distribution.values,
    }], sn++, { showPercent: true });
  }

  // Actions par responsable (barres)
  if (metrics.actions.values.length) {
    addChartSlide(pptx, 'Actions par responsable', pptx.ChartType.bar, [{
      name: 'Actions',
      labels: metrics.actions.labels,
      values: metrics.actions.values,
    }], sn++, { barDir: 'col', showValue: true });
  }

  // Profondeur par thème (barres horizontales)
  if (metrics.topics.values.length) {
    addChartSlide(pptx, 'Profondeur par thème', pptx.ChartType.bar, [{
      name: 'Points',
      labels: metrics.topics.labels,
      values: metrics.topics.values,
    }], sn++, { barDir: 'bar', showValue: true });
  }

  // Points clés (barres)
  if (metrics.highlights.values.length) {
    addChartSlide(pptx, 'Poids des points clés', pptx.ChartType.bar, [{
      name: 'Importance',
      labels: metrics.highlights.labels.map((l, i) => `P${i + 1}`),
      values: metrics.highlights.values,
    }], sn++, { barDir: 'col', showValue: true });
  }

  // Suivi décisions / échéances (secteurs)
  if (metrics.status.values.length) {
    addChartSlide(pptx, 'Suivi & décisions', pptx.ChartType.pie, [{
      name: 'Statut',
      labels: metrics.status.labels,
      values: metrics.status.values,
    }], sn++, { showPercent: true });
  }

  // Vue d'ensemble comparative (barres groupées)
  if (metrics.radar.values.some(v => v > 0)) {
    addChartSlide(pptx, "Vue d'ensemble comparative", pptx.ChartType.bar, [{
      name: 'Volume',
      labels: metrics.radar.labels,
      values: metrics.radar.values,
    }], sn++, { barDir: 'col', showValue: true });
  }

  if (pres.executive_summary) {
    const slide = pptx.addSlide({ background: { color: C.white } });
    setTrans(slide, sn);
    addHeader(pptx, slide, 'Synthèse exécutive', '🎯');
    slide.addShape(pptx.ShapeType.roundRect, { x: 0.7, y: 1.5, w: 8.6, h: 5.2, fill: { color: C.lightGray }, line: { color: 'E5E7EB' } });
    slide.addText(pres.executive_summary, { x: 1.0, y: 1.8, w: 8.0, h: 4.6, fontSize: 17, color: C.dark, valign: 'top', lineSpacingMultiple: 1.5 });
    addFooter(slide, sn++);
  }

  const highlights = pres.highlights || [];
  const tldr = summary.tldr || [];
  if (highlights.length) {
    highlights.slice(0, 5).forEach((h, i) => {
      addHighlightSlide(pptx, i + 1, highlights.length, typeof h === 'object' ? h.title : h, typeof h === 'object' ? h.detail : '', sn++);
    });
  } else if (tldr.length) {
    tldr.slice(0, 5).forEach((p, i) => addHighlightSlide(pptx, i + 1, tldr.length, p, '', sn++));
  }

  if (summary.decisions?.length) {
    addDividerSlide(pptx, 'Décisions', '✅', sn++);
    chunk(summary.decisions, 4).forEach(c => addSectionSlide(pptx, 'Décisions prises', c, '✅', true, sn++));
  }
  if (summary.actions?.length) {
    addDividerSlide(pptx, 'Actions', '📋', sn++);
    chunk(summary.actions, 4).forEach(c => addActionsSlide(pptx, c, sn++));
  }
  if (summary.topics?.length) {
    addDividerSlide(pptx, 'Thèmes abordés', '💡', sn++);
    summary.topics.slice(0, 5).forEach(t => { if (t.summary || t.key_points?.length) addTopicSlide(pptx, t, sn++); });
  }
  if (summary.explanation_for_absent) {
    const slide = pptx.addSlide({ background: { color: C.white } });
    setTrans(slide, sn);
    addHeader(pptx, slide, 'Pour les absents', '👤');
    slide.addShape(pptx.ShapeType.roundRect, { x: 0.7, y: 1.5, w: 8.6, h: 5.2, fill: { color: C.lightGray }, line: { color: 'E5E7EB' } });
    slide.addText(summary.explanation_for_absent, { x: 1.0, y: 1.8, w: 8.0, h: 4.6, fontSize: 17, color: C.dark, valign: 'top', lineSpacingMultiple: 1.5 });
    addFooter(slide, sn++);
  }
  if (summary.open_questions?.length) {
    chunk(summary.open_questions, 4).forEach(c => addSectionSlide(pptx, 'Questions en suspens', c, '❓', false, sn++));
  }

  const closing = pptx.addSlide({ background: { color: C.blue } });
  setTrans(closing, 0);
  closing.addShape(pptx.ShapeType.ellipse, { x: 3.5, y: 1.5, w: 3, h: 3, fill: { color: C.lightBlue, transparency: 30 }, line: { color: C.lightBlue } });
  closing.addText('Merci', { x: 0.8, y: 2.8, w: 8.4, h: 1.2, fontSize: 44, bold: true, color: C.white, align: 'center' });
  closing.addText('TeamsBrief — Synthèse de réunion Teams', { x: 0.8, y: 4.3, w: 8.4, h: 0.6, fontSize: 14, color: 'AAAACC', align: 'center' });

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
