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

  const stats = [[summary.decisions?.length || 0, 'Décisions', '✅'], [summary.actions?.length || 0, 'Actions', '📋'], [summary.topics?.length || 0, 'Thèmes', '💡']];
  const stSlide = pptx.addSlide({ background: { color: C.lightGray } });
  setTrans(stSlide, 2);
  addHeader(pptx, stSlide, "Vue d'ensemble", '📊', C.blue);
  stats.forEach(([val, lab, em], i) => {
    const x = 1.2 + i * 2.8;
    stSlide.addShape(pptx.ShapeType.roundRect, { x, y: 2.2, w: 2.2, h: 3.5, fill: { color: C.white }, line: { color: 'E5E7EB' }, rectRadius: 0.1 });
    stSlide.addText(em, { x, y: 2.45, w: 2.2, h: 0.6, fontSize: 28, align: 'center' });
    stSlide.addText(String(val), { x, y: 3.2, w: 2.2, h: 0.9, fontSize: 36, bold: true, color: C.accent, align: 'center' });
    stSlide.addText(lab, { x, y: 4.2, w: 2.2, h: 0.8, fontSize: 12, color: C.gray, align: 'center' });
  });
  addFooter(stSlide, sn++);

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
  } catch (err) {
    alert(err.message || 'Erreur lors de la génération PowerPoint.');
  } finally {
    btn.disabled = false;
    btn.textContent = '📊 PowerPoint';
  }
}
