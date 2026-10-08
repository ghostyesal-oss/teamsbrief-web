/* TeamsBrief Web — 100% navigateur, hébergé */

const STORAGE_KEY = 'teamsbrief_meetings';
const SETTINGS_KEY = 'teamsbrief_settings';

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
  "decisions": ["décision 1"],
  "actions": [{"who": "nom", "what": "action", "when": "échéance"}],
  "open_questions": ["question"],
  "explanation_for_absent": "paragraphe pour absent"
}`;

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

document.getElementById('groq-key').value = settings.apiKey || '';
document.getElementById('groq-model').value = settings.model || 'openai/gpt-oss-20b';

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
    model: document.getElementById('groq-model').value.trim() || 'openai/gpt-oss-20b',
  };
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  document.getElementById('settings-status').textContent = 'Paramètres enregistrés.';
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
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = '';

  if (!settings.apiKey) {
    alert('Configurez votre clé Groq dans Paramètres d\'abord.');
    showView('settings');
    return;
  }

  const text = await file.text();
  const transcript = file.name.endsWith('.vtt') || text.startsWith('WEBVTT') ? parseVtt(text) : text;
  const title = file.name.replace(/\.(vtt|txt)$/i, '').replace(/^.*[\\/]/, '') || 'Réunion Teams';

  const meeting = {
    id: Date.now(),
    title,
    date: new Date().toISOString(),
    transcript,
    summary: null,
    status: 'processing',
    chat: [],
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
  } catch (err) {
    meeting.status = 'error';
    meeting.error = err.message;
  }

  saveMeetings();
  renderMeetingsList();
  openMeeting(meeting.id);
}

// ─── Liste ───
function renderMeetingsList() {
  const el = document.getElementById('meetings-list');
  if (!meetings.length) {
    el.innerHTML = '<div class="empty">Aucune réunion. Importez un fichier .vtt Teams.</div>';
    return;
  }
  el.innerHTML = meetings.map(m => `
    <div class="meeting-item" data-id="${m.id}">
      <div style="display:flex;justify-content:space-between;align-items:center">
        <div>
          <h3>${esc(m.title)}</h3>
          <p>${new Date(m.date).toLocaleString('fr-FR')}</p>
        </div>
        <span class="badge ${m.status}">${m.status}</span>
      </div>
    </div>
  `).join('');

  el.querySelectorAll('.meeting-item').forEach(item => {
    item.addEventListener('click', () => openMeeting(+item.dataset.id));
  });
}

// ─── Détail ───
function openMeeting(id) {
  currentMeetingId = id;
  const m = meetings.find(x => x.id === id);
  if (!m) return;

  document.getElementById('detail-title').textContent = m.title;
  document.getElementById('detail-status').textContent = m.status;
  document.getElementById('detail-status').className = `badge ${m.status}`;

  const exportBtn = document.getElementById('export-pptx-btn');
  const canExport = m.status === 'ready' && m.summary;
  exportBtn.hidden = !canExport;
  exportBtn.disabled = false;
  exportBtn.textContent = '📊 PowerPoint';

  const panel = document.getElementById('summary-content');
  if (m.status === 'processing') {
    panel.innerHTML = '<p class="hint">Génération du résumé en cours...</p>';
  } else if (m.status === 'error') {
    panel.innerHTML = `<p style="color:#991b1b">${esc(m.error || 'Erreur')}</p>`;
  } else if (m.summary) {
    panel.innerHTML = renderSummary(m.summary);
  }

  renderChat(m);
  renderSuggestions();
  showView('detail');
  views.meetings.classList.add('hidden');
  views.detail.classList.remove('hidden');
}

function renderSummary(s) {
  let html = '';
  if (s.tldr?.length) {
    html += `<div class="summary-section"><h4>En bref</h4><ul>${s.tldr.map(p => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
  }
  if (s.decisions?.length) {
    html += `<div class="summary-section"><h4>Décisions</h4><ul>${s.decisions.map(p => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
  }
  if (s.actions?.length) {
    html += `<div class="summary-section"><h4>Actions</h4><ul>${s.actions.map(a => `<li>☐ ${esc(a.who)} — ${esc(a.what)}${a.when ? ` (${esc(a.when)})` : ''}</li>`).join('')}</ul></div>`;
  }
  if (s.explanation_for_absent) {
    html += `<div class="summary-section"><h4>Pour les absents</h4><p class="explanation">${esc(s.explanation_for_absent)}</p></div>`;
  }
  if (s.open_questions?.length) {
    html += `<div class="summary-section"><h4>Questions en suspens</h4><ul>${s.open_questions.map(p => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
  }
  return html;
}

// ─── Chat ───
function renderChat(m) {
  const el = document.getElementById('chat-messages');
  el.innerHTML = (m.chat || []).map(msg => `
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
  bubble.className = 'chat-bubble assistant';
  bubble.innerHTML = '<div class="label">Agent</div><div class="content">Réflexion...</div>';
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
    contentEl.textContent = answer;
    m.chat.push({ role: 'assistant', content: answer });
  } catch (err) {
    contentEl.textContent = err.message || 'Erreur agent.';
  }

  saveMeetings();
}

function esc(s) {
  const d = document.createElement('div');
  d.textContent = s || '';
  return d.innerHTML;
}

// ─── Export PowerPoint ───
const PPTX_COLORS = {
  blue: '464775',
  lightBlue: '6264A7',
  white: 'FFFFFF',
  dark: '212C40',
  muted: 'CCCCDD',
  footer: '9999BB',
};

function safeFilename(title) {
  const name = (title || 'reunion').replace(/[<>:"/\\|?*]/g, '').trim() || 'reunion';
  const date = new Date().toISOString().slice(0, 10);
  return `${date} - ${name.slice(0, 60)}.pptx`;
}

function addSectionSlide(pptx, heading, items, numbered = false) {
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: '100%', h: 1.2,
    fill: { color: PPTX_COLORS.lightBlue },
    line: { color: PPTX_COLORS.lightBlue },
  });
  slide.addText(heading, {
    x: 0.6, y: 0.25, w: 8.8, h: 0.7,
    fontSize: 24, bold: true, color: PPTX_COLORS.white,
  });
  const lines = items.map((item, i) => {
    const prefix = numbered ? `${i + 1}. ` : '→ ';
    return `${prefix}${item}`;
  });
  slide.addText(lines.join('\n'), {
    x: 0.8, y: 1.6, w: 8.4, h: 5.2,
    fontSize: 18, color: PPTX_COLORS.dark, valign: 'top',
    lineSpacingMultiple: 1.2,
  });
}

function addTextSlide(pptx, heading, bodyText) {
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: '100%', h: 1.2,
    fill: { color: PPTX_COLORS.lightBlue },
    line: { color: PPTX_COLORS.lightBlue },
  });
  slide.addText(heading, {
    x: 0.6, y: 0.25, w: 8.8, h: 0.7,
    fontSize: 24, bold: true, color: PPTX_COLORS.white,
  });
  slide.addText(bodyText, {
    x: 0.8, y: 1.6, w: 8.4, h: 5.2,
    fontSize: 16, color: PPTX_COLORS.dark, valign: 'top',
    lineSpacingMultiple: 1.4,
  });
}

function generatePresentation(title, summary, dateIso) {
  if (typeof PptxGenJS === 'undefined') {
    throw new Error('Bibliothèque PowerPoint non chargée. Rechargez la page.');
  }

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'TeamsBrief';
  pptx.title = title;

  const dateStr = dateIso
    ? new Date(dateIso).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' })
    : new Date().toLocaleDateString('fr-FR');

  const titleSlide = pptx.addSlide();
  titleSlide.background = { color: PPTX_COLORS.blue };
  titleSlide.addText(title, {
    x: 0.8, y: 2.2, w: 8.4, h: 1.5,
    fontSize: 32, bold: true, color: PPTX_COLORS.white,
  });
  titleSlide.addText(`Compte-rendu · ${dateStr}`, {
    x: 0.8, y: 3.8, w: 8.4, h: 0.8,
    fontSize: 16, color: PPTX_COLORS.muted,
  });
  titleSlide.addText('Généré par TeamsBrief', {
    x: 0.8, y: 6.8, w: 4, h: 0.4,
    fontSize: 11, color: PPTX_COLORS.footer,
  });

  if (summary.tldr?.length) addSectionSlide(pptx, 'En bref', summary.tldr);
  if (summary.decisions?.length) addSectionSlide(pptx, 'Décisions prises', summary.decisions, true);

  if (summary.actions?.length) {
    const actionLines = summary.actions.map(a => {
      let line = `☐  ${a.who || '?'} — ${a.what || ''}`;
      if (a.when && a.when !== 'non précisé') line += `  (${a.when})`;
      return line;
    });
    addSectionSlide(pptx, 'Actions à faire', actionLines);
  }

  if (summary.explanation_for_absent) {
    addTextSlide(pptx, 'Pour les absents', summary.explanation_for_absent);
  }

  if (summary.open_questions?.length) {
    addSectionSlide(pptx, 'Questions en suspens', summary.open_questions);
  }

  const closing = pptx.addSlide();
  closing.background = { color: PPTX_COLORS.blue };
  closing.addText('Merci', {
    x: 0.8, y: 2.8, w: 8.4, h: 1.2,
    fontSize: 40, bold: true, color: PPTX_COLORS.white, align: 'center',
  });
  closing.addText('TeamsBrief — Synthèse de réunion Teams', {
    x: 0.8, y: 4.2, w: 8.4, h: 0.6,
    fontSize: 14, color: 'AAAACC', align: 'center',
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
  } catch (err) {
    alert(err.message || 'Erreur lors de la génération PowerPoint.');
  } finally {
    btn.disabled = false;
    btn.textContent = '📊 PowerPoint';
  }
}
