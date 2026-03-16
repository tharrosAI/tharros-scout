const STAGES = [
  { name: '1st Contact', short: '1st Contact', code: 'S1' },
  { name: 'Demo Permission', short: 'Demo', code: 'S2' },
  { name: 'Demo Successful', short: 'Demo Done', code: 'S3' },
  { name: 'Activation', short: 'Activation', code: 'S4' },
  { name: 'Payment + Onboarding', short: 'Payment', code: 'S5' },
];

const BADGE_CLASS = ['badge-s1','badge-s2','badge-s3','badge-s4','badge-s5'];

const MESSAGES = {
  call_intro: {
    code:'001', type:'call', title:'Intro Pitch Call', stage:0,
    content: `Hey — quick question.\n\nWhen someone calls your moving company after hours or when you're out on a job, what usually happens?\n\n[let them answer]\n\nGot it.\n\nThe reason I'm calling is we built a simple AI call assistant that answers those missed moving calls and books quote requests automatically. It basically acts like a 24/7 receptionist for moving companies.\n\n[pause]\n\nWould it be alright if I text you a demo number? Takes about 20 seconds to try.\n\n[If yes:] Cool — I'll text it over right now.`
  },
  call_voicemail: {
    code:'002', type:'call', title:'Voicemail Script', stage:0,
    content: `Hey [Contact], this is [Your Name].\n\nQuick reason for the call — we built an AI call assistant that answers missed calls for moving companies and captures quote requests automatically.\n\nI'll text you a demo number so you can try it in about 20 seconds.\n\nTalk soon.\n\n[→ Send Demo Text immediately after]`
  },
  demo_text: {
    code:'0100', type:'text', title:'Demo Text', stage:1, timing:'Send immediately after call',
    content: `Hey [Contact] — this is the demo we talked about.\n\nCall this number and pretend you need a moving quote.\n\n(828) 623-1806\n\nIt will answer like a receptionist and ask about the move.\n\nCurious what you think.`
  },
  demo_followup: {
    code:'0201', type:'text', title:'Follow Up', stage:1, timing:'4 hours later',
    content: `Hey [Contact] — did you get a chance to try the demo number yet?\n\nMost moving companies test it by saying something like:\n\n"I need a quote for a move next week."\n\nCurious what you think.`
  },
  demo_bump: {
    code:'0202', type:'text', title:'Bump', stage:1, timing:'Next day',
    content: `Quick bump on this — the demo number is still live if you want to try it.\n\n(828) 623-1806\n\nTakes about 20 seconds.`
  },
  demo_final: {
    code:'0203', type:'text', title:'Final', stage:1, timing:'3 days later',
    content: `Last message on this — totally fine if now isn't the right time.\n\nBut if you're curious how it works for moving companies, the demo is still here:\n\n(828) 623-1806\n\nHappy to show how companies are using it to catch missed quote calls.`
  },
  seen_main: {
    code:'0300', type:'text', title:'Demo Seen — Main Message', stage:2, timing:'Send when they respond positively to demo',
    content: `Glad you tried it.\n\nThe main thing moving companies use it for is catching missed calls while crews are on jobs or after hours.\n\nIt collects the move details and sends you the lead instantly.\n\nIf you'd like I can show you how it would work specifically for your company.`
  },
  seen_followup: {
    code:'0301', type:'text', title:'Follow Up', stage:2, timing:'24 hours',
    content: `Quick question — roughly how many calls do you think you miss in a typical week while crews are out on jobs?\n\nJust trying to see if this would actually help you.`
  },
  seen_bump: {
    code:'0302', type:'text', title:'Bump', stage:2, timing:'2 days',
    content: `Most moving companies we talk to miss about 10–20 quote calls a week while they're on jobs.\n\nThat's usually where this helps.\n\nHappy to show how it would work for your setup if you're curious.`
  },
  seen_final: {
    code:'0303', type:'text', title:'Final', stage:2, timing:'4 days',
    content: `Last ping on this.\n\nIf catching missed moving quote calls is ever something you want help with, feel free to reach out.\n\nOtherwise I won't bug you again.`
  },
  act_main: {
    code:'0400', type:'text', title:'Activation Message', stage:3, timing:'Send when they want to move forward',
    content: `Awesome.\n\nThe way we usually do it is set up a version trained for your moving company.\n\nIt asks about things like:\n• move date\n• origin/destination\n• home size\n• contact info\n\nThen sends the lead directly to you.\n\nTakes about 15 minutes to set up.`
  },
  act_followup: {
    code:'0401', type:'text', title:'Follow Up', stage:3, timing:'Same day',
    content: `If you're good with it we can get your version live pretty quickly.\n\nI'll send over the setup link and we can launch it this week.`
  },
  act_bump: {
    code:'0402', type:'text', title:'Bump', stage:3, timing:'24 hours',
    content: `Just checking back — still happy to get your moving call assistant set up if you want to try it out.`
  },
  act_final: {
    code:'0403', type:'text', title:'Final', stage:3, timing:'3 days',
    content: `No worries if timing isn't right.\n\nIf you ever want help catching missed moving quote calls just shoot me a message.`
  },
  payment: {
    code:'100', type:'text', title:'Payment + Onboarding Link', stage:4, timing:'After activation confirmation',
    content: `Awesome — here's the setup link.\n\nOnce it's completed we'll configure the assistant for your moving company.\n\n[Stripe Link]\n\nRight after payment you'll be redirected to a short onboarding form so we can train it for your business.`
  },
};

const STAGE_SEQUENCES = [
  ['call_intro', 'call_voicemail'],
  ['demo_text', 'demo_followup', 'demo_bump', 'demo_final'],
  ['seen_main', 'seen_followup', 'seen_bump', 'seen_final'],
  ['act_main', 'act_followup', 'act_bump', 'act_final'],
  ['payment'],
];

const WEBHOOK_URL = 'https://tharros.app.n8n.cloud/webhook/e196fe46-4e43-474d-844a-b0f7cce8eab3';
const FETCH_WEBHOOK_URL = 'https://tharros.app.n8n.cloud/webhook/2e6968bb-c9a1-4150-9516-e845842b37cc';
const DELETE_WEBHOOK_URL = 'https://tharros.app.n8n.cloud/webhook/1456e7f1-fbb4-4d32-975f-0b5a5f8132cd';
const UPDATE_WEBHOOK_URL = 'https://tharros.app.n8n.cloud/webhook/064cad6f-d3d5-4e59-a852-6afc86e9a283';
let scriptOverrides = {};
let leads = [];
let currentLeadId = null;
let currentFilter = 'all';
let activeScriptKey = Object.keys(MESSAGES)[0] || '';
const METRIC_OPTIONS = [
  { id: 'leads', label: 'Leads Added' },
  { id: 'actions', label: 'Actions Logged' },
  { id: 'advances', label: 'Stage Advances' },
];
let overviewMetric = 'actions';

function save() {
  const payload = {
    leads,
    overrides: scriptOverrides,
    timestamp: new Date().toISOString(),
  };
  fetch(WEBHOOK_URL, {
    method: 'POST',
    mode: 'cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  }).catch(err => {
    console.error('Sync request failed', err);
    showToast('Could not sync to n8n', 'error');
  });
}

function syncLeadUpdate(lead) {
  if (!lead || !lead.name) return Promise.resolve();
  const sentMessagesPayload = JSON.stringify(Array.isArray(lead.sentMessages) ? lead.sentMessages : []);
  const payload = {
    'Business Name': lead.name,
    Website: lead.website || '',
    City: lead.city || '',
    State: lead.state || '',
    Email: lead.email || '',
    Notes: lead.notes || '',
    'Sent Messages': sentMessagesPayload,
  };
  return fetch(UPDATE_WEBHOOK_URL, {
    method: 'POST',
    mode: 'cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  }).then(res => {
    if (!res.ok) throw new Error('Lead sync failed');
    return res;
  });
}

function fetchLeads() {
  const empty = document.getElementById('empty-state');
  const detail = document.getElementById('lead-detail');
  leads = [];
  currentLeadId = null;
  renderLeadList();
  if (detail) detail.style.display = 'none';
  if (empty) empty.style.display = 'flex';
  fetch(FETCH_WEBHOOK_URL, { method: 'GET', mode: 'cors', headers: { 'Accept': 'application/json' } })
    .then(res => {
      if (!res.ok) throw new Error('Network response was not ok');
      return res.json();
    })
    .then(data => {
      const rows = Array.isArray(data)
        ? data
        : data?.leads || data?.rows || data?.items || data?.data || [];
      leads = rows.map(normalizeLeadRow);
      renderLeadList();
      renderOverview();
      if (empty) empty.style.display = leads.length ? 'none' : 'flex';
      showToast('Leads refreshed');
    })
    .catch(err => {
      console.error('fetchLeads error', err);
      showToast('Unable to refresh leads', 'error');
    });
}

function normalizeLeadRow(row) {
  const businessName = resolveField(row, 'Business Name');
  const contactName = resolveField(row, 'Contact Name');
  const notes = resolveField(row, 'Notes');
  const phone = resolveField(row, 'Phone');
  const sentMessagesRaw = resolveField(row, 'Sent Messages');
  const stageValue = resolveField(row, 'Stage');
  const lastMessageRaw = resolveField(row, 'lastMessageAt');
  return {
    id: resolveField(row, 'id') || genId(),
    name: businessName || contactName || 'Untitled Lead',
    phone: phone || '',
    contact: contactName || '',
    stage: parseStage(stageValue),
    sentMessages: normalizeSentMessages(sentMessagesRaw),
    history: normalizeHistory(resolveField(row, 'history')),
    notes: notes || '',
    website: resolveField(row, 'Website') || '',
    city: resolveField(row, 'City') || '',
    state: resolveField(row, 'State') || '',
    email: resolveField(row, 'Email') || '',
    createdAt: parseTimestamp(resolveField(row, 'time')) || now(),
    lastMessageAt: parseTimestamp(lastMessageRaw),
  };
}

function resolveField(row, key) {
  if (!row) return undefined;
  if (Object.prototype.hasOwnProperty.call(row, key)) return row[key];
  const normalizedKey = key.toLowerCase().replace(/\s+/g, '');
  for (const entryKey of Object.keys(row)) {
    const normalizedEntry = entryKey.toLowerCase().replace(/\s+/g, '');
    if (normalizedEntry === normalizedKey) {
      return row[entryKey];
    }
  }
  return undefined;
}

function parseTimestamp(value) {
  if (!value) return null;
  if (typeof value === 'number') return value;
  const parsed = Date.parse(value);
  return Number.isNaN(parsed) ? null : parsed;
}

function normalizeHistory(raw) {
  if (Array.isArray(raw)) return raw;
  if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return parsed;
    } catch (_) {}
  }
  return [];
}

function normalizeSentMessages(value) {
  if (Array.isArray(value)) return value;
  if (typeof value === 'string') {
    try {
      const parsed = JSON.parse(value);
      if (Array.isArray(parsed)) return parsed;
    } catch (_) {}
    return value.split(',').map(item => item.trim()).filter(Boolean);
  }
  return [];
}

function parseStage(value) {
  const numeric = parseInt(value, 10);
  if (!Number.isNaN(numeric)) {
    return Math.min(Math.max(numeric, 0), STAGES.length - 1);
  }
  return 0;
}

function genId() { return Date.now().toString(36) + Math.random().toString(36).slice(2); }
function now() { return Date.now(); }
function fmtTime(ts) {
  if (!ts) return 'never';
  const d = new Date(ts);
  const diff = (now() - d) / 1000;
  if (diff < 60) return 'just now';
  if (diff < 3600) return Math.floor(diff/60) + 'm ago';
  if (diff < 86400) return Math.floor(diff/3600) + 'h ago';
  return Math.floor(diff/86400) + 'd ago';
}
function fmtAbsTime(ts) {
  if (!ts) return 'unknown';
  return new Date(ts).toLocaleString('en-US', {month:'short',day:'numeric',hour:'numeric',minute:'2-digit'});
}

function getMessageDefinition(key) {
  const base = MESSAGES[key];
  if (!base) return null;
  const override = scriptOverrides[key] || {};
  return {
    ...base,
    title: override.title || base.title,
    timing: override.timing !== undefined ? override.timing : base.timing,
    content: override.content || base.content,
  };
}

function getNextAction(lead) {
  const stageSeq = STAGE_SEQUENCES[lead.stage] || [];
  const sent = lead.sentMessages || [];
  for (let key of stageSeq) {
    if (!sent.includes(key)) {
      return { key, msg: getMessageDefinition(key) };
    }
  }
  return null;
}

function isActionDue(lead) {
  const next = getNextAction(lead);
  if (!next) return false;
  if (lead.stage === 4 && (lead.sentMessages||[]).includes('payment')) return false;
  const msg = next.msg;
  if (!msg || !msg.timing) return true;
  const lastSent = lead.lastMessageAt || lead.createdAt || now();
  const elapsed = (now() - lastSent) / 3600000;
  const timing = msg.timing.toLowerCase();
  if (timing.includes('4 hours')) return elapsed >= 4;
  if (timing.includes('same day')) return elapsed >= 1;
  if (timing.includes('24 hours')) return elapsed >= 24;
  if (timing.includes('next day')) return elapsed >= 20;
  if (timing.includes('2 days')) return elapsed >= 48;
  if (timing.includes('3 days')) return elapsed >= 72;
  if (timing.includes('4 days')) return elapsed >= 96;
  return true;
}

function renderLeadList() {
  const list = document.getElementById('lead-list');
  const search = document.getElementById('search-input').value.toLowerCase();
  let filtered = leads.filter(l => {
    if (search && !l.name.toLowerCase().includes(search) && !(l.phone || '').includes(search)) return false;
    if (currentFilter === 'action') return isActionDue(l) && l.stage < 5;
    if (currentFilter !== 'all' && currentFilter !== 'action') {
      return l.stage === parseInt(currentFilter, 10);
    }
    return true;
  });

  filtered.sort((a,b) => {
    const aA = isActionDue(a); const bA = isActionDue(b);
    if (aA && !bA) return -1;
    if (!aA && bA) return 1;
    return (b.lastMessageAt||b.createdAt) - (a.lastMessageAt||a.createdAt);
  });

  document.getElementById('hdr-total').textContent = leads.length;
  const needAction = leads.filter(l => isActionDue(l) && l.stage < 5).length;
  document.getElementById('hdr-action').textContent = needAction;

  if (!filtered.length) {
    list.innerHTML = `<div style="padding:32px 20px;text-align:center;color:var(--text3);font-size:13px">No leads found</div>`;
    return;
  }

  list.innerHTML = filtered.map((lead,i) => {
    const active = lead.id === currentLeadId;
    const action = isActionDue(lead) && lead.stage < 5;
    const bc = BADGE_CLASS[lead.stage] || 'badge-s1';
    return `
      <div class="lead-item ${active?'active':''}" onclick="selectLead('${lead.id}')">
        <div class="lead-item-top">
          <span class="lead-name">${lead.name}</span>
          <span class="lead-time">${fmtTime(lead.lastMessageAt||lead.createdAt)}</span>
        </div>
        <div class="lead-phone">${lead.phone || '—'}</div>
        <span class="lead-stage-badge ${bc}">${STAGES[lead.stage]?.short||''}</span>
        ${action ? '<div class="action-dot"></div>' : ''}
      </div>
      ${i < filtered.length-1 ? '<div class="lead-item-divider"></div>' : ''}
    `;
  }).join('');
}

function hidePanels() {
  document.getElementById('lead-detail').style.display = 'none';
  document.getElementById('overview-area').style.display = 'none';
  document.getElementById('script-area').style.display = 'none';
}

function selectLead(id) {
  currentLeadId = id;
  renderLeadList();
  renderDetail();
  document.getElementById('empty-state').style.display = 'none';
  hidePanels();
  document.getElementById('lead-detail').style.display = 'block';
}

function renderDetail() {
  const lead = leads.find(l => l.id === currentLeadId);
  if (!lead) return;
  const el = document.getElementById('lead-detail');
  const nextAction = getNextAction(lead);
  const sentMessages = lead.sentMessages || [];
  const isComplete = lead.stage >= 5 || (lead.stage === 4 && sentMessages.includes('payment'));

  const progressPct = STAGES.length > 1 ? (lead.stage / (STAGES.length - 1)) * 100 : 0;
  const nodesHTML = STAGES.map((stage, index) => {
    const done = index < lead.stage;
    const current = index === lead.stage;
    const statusClass = done ? 'done' : current ? 'current' : '';
    const left = STAGES.length > 1 ? (index / (STAGES.length - 1)) * 100 : 0;
    return `
      <div class="pipeline-node ${statusClass}" style="left:${left}%">
        <div class="pipeline-node-circle"></div>
        <span class="pipeline-node-label">${stage.short}</span>
      </div>
    `;
  }).join('');
  const pipelineHTML = `
    <div class="pipeline">
      <div class="pipeline-visual">
        <div class="pipeline-label">Pipeline Stage</div>
        <div class="pipeline-track-wrap">
          <div class="pipeline-track-line"></div>
          <div class="pipeline-track-progress" style="width:${progressPct}%"></div>
          ${nodesHTML}
        </div>
      </div>
    </div>
  `;

  let actionHTML = '';
  if (isComplete) {
    actionHTML = `<div class="next-ribbon"><div class="next-icon"></div><div class="next-text"><div class="next-label">Complete</div><div class="next-desc">This lead finished the pipeline.</div></div></div>`;
  } else if (nextAction && nextAction.msg) {
    const msg = nextAction.msg;
    const isCall = msg.type === 'call';
    actionHTML = `
      <div class="next-ribbon">
        <div class="next-icon"></div>
        <div class="next-text">
          <div class="next-label">Next Action</div>
          <div class="next-desc">${isCall ? 'Call' : 'Text'}: <strong>${msg.title}</strong>${msg.timing ? ` · <span style="color:var(--text3)">${msg.timing}</span>` : ''}</div>
        </div>
        <span style="font-size:11px;color:var(--text3);font-family:'DM Mono',monospace;background:var(--bg4);border:1px solid var(--border);border-radius:4px;padding:3px 8px">${msg.code}</span>
      </div>
    `;
  } else {
    actionHTML = `<div class="next-ribbon"><div class="next-icon"></div><div class="next-text"><div class="next-label">Waiting</div><div class="next-desc">All messages sent — advance the lead when they reply.</div></div></div>`;
  }

  let messagesHTML = '';
  if (!isComplete && lead.stage < STAGE_SEQUENCES.length) {
    const stageSeq = STAGE_SEQUENCES[lead.stage];
    messagesHTML = stageSeq.map(key => {
      const msg = getMessageDefinition(key);
      if (!msg) return '';
      const isSent = sentMessages.includes(key);
      const isNext = nextAction && nextAction.key === key;
      const isCall = msg.type === 'call';
      const content = msg.content.replace(/\[Contact\]/g, lead.contact || 'there').replace(/\[Your Name\]/g, 'Your Name');
      return `
        <div class="action-card" style="${isSent ? 'opacity:0.45;' : ''}">
          <div class="action-card-label" style="${isCall?'color:var(--blue)':''}">
            ${isCall ? 'Call Script' : 'Text Message'} · ${msg.code} ${isSent?'<span style="color:var(--green);margin-left:8px">✓ Sent</span>':''}${isNext ? '<span style="margin-left:8px;font-size:11px;color:var(--accent)">Next</span>' : ''}
          </div>
          <div class="action-card-title">${msg.title}</div>
          ${msg.timing ? `<div class="action-card-timing"><span class="timing-badge">${msg.timing}</span></div>` : ''}
          <div class="message-box" style="position:relative">
            ${escHTML(content).replace(/\[([^\]]+)\]/g,'<span class="placeholder">[$1]</span>')}
            <button class="copy-btn" onclick="copyMsg(this,'${key}','${lead.id}')">Copy</button>
          </div>
          ${!isSent ? `
          <div class="stage-actions">
            <button class="stage-btn advance" onclick="markSent('${lead.id}','${key}')">✓ Mark as Sent</button>
          </div>` : ''}
        </div>
      `;
    }).join('');
  }

  const advanceHTML = lead.stage < STAGES.length ? `
    <div style="margin-top:16px;padding-top:16px;border-top:1px solid var(--border);display:flex;gap:8px;flex-wrap:wrap;align-items:center">
      <span style="font-size:11px;color:var(--text3);letter-spacing:1px;text-transform:uppercase;font-weight:600">Advance to:</span>
      ${STAGES.map((s,i) => i > lead.stage ? `<button class="stage-btn" onclick="moveToStage(${i})">${s.name} →</button>` : '').filter(Boolean).join('')}
    </div>
  ` : '';

  const notesHTML = `
    <div class="notes-area">
      <label class="widget-subtext">Lead Notes</label>
      <textarea id="lead-notes-input" placeholder="Capture context that matters">${lead.notes || ''}</textarea>
          <div class="notes-actions">
            <button class="btn save-notes-btn" onclick="saveLeadNotes('${lead.id}', this)">Save notes</button>
          </div>
    </div>
  `;

  const websiteLink = lead.website ? `<a href="${escHTML(lead.website)}" target="_blank" rel="noreferrer">${escHTML(lead.website)}</a>` : '';
  const location = [lead.city, lead.state].filter(Boolean).join(', ');
  const emailLine = lead.email ? `<div class="detail-meta-item">Email · <a href="mailto:${escHTML(lead.email)}">${escHTML(lead.email)}</a></div>` : '';
  const websiteLine = lead.website ? `<div class="detail-meta-item">Website · ${websiteLink}</div>` : '';
  const locationLine = location ? `<div class="detail-meta-item">Location · ${escHTML(location)}</div>` : '';
  el.innerHTML = `
    <div class="lead-detail">
      <div class="detail-header">
        <div>
          <div class="detail-name">${lead.name}</div>
          <div class="detail-phone">${lead.phone}${lead.contact ? ' · ' + lead.contact : ''}</div>
          <div class="detail-meta">
            ${websiteLine}
            ${locationLine}
            ${emailLine}
          </div>
        </div>
        <div class="detail-actions">
          <span style="font-size:11px;color:var(--text3);font-family:'DM Mono',monospace">Added ${fmtAbsTime(lead.createdAt)}</span>
          <button class="btn btn-danger" onclick="deleteLead('${lead.id}')">Delete</button>
        </div>
      </div>
      ${pipelineHTML}
      ${actionHTML}
      ${advanceHTML}
      ${notesHTML}
      ${messagesHTML}
    </div>
  `;
}

function escHTML(str) {
  return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>');
}

function markSent(leadId, msgKey) {
  const lead = leads.find(l => l.id === leadId);
  if (!lead) return;
  if (!lead.sentMessages) lead.sentMessages = [];
  if (!lead.sentMessages.includes(msgKey)) {
    lead.sentMessages.push(msgKey);
    const msg = getMessageDefinition(msgKey);
    if (!lead.history) lead.history = [];
    lead.history.push({ code: msg.code, label: `${msg.type==='call'?'Called':'Texted'}: ${msg.title}`, ts: now() });
    lead.lastMessageAt = now();
    save();
    syncLeadUpdate(lead).catch(err => {
      console.error('markSent sync error', err);
      showToast('Could not sync sent message', 'error');
    });
    renderLeadList();
    renderDetail();
    showToast('Marked as sent');
  }
}

function moveToStage(stageIdx) {
  const lead = leads.find(l => l.id === currentLeadId);
  if (!lead || stageIdx <= lead.stage) return;
  lead.stage = stageIdx;
  lead.sentMessages = [];
  if (!lead.history) lead.history = [];
  lead.history.push({ code: `S${stageIdx+1}`, label: `Advanced to: ${STAGES[stageIdx].name}`, ts: now() });
  save();
  renderLeadList();
  renderDetail();
  showToast(`Moved to ${STAGES[stageIdx].name}`);
}

function copyMsg(btn, msgKey, leadId) {
  const lead = leads.find(l => l.id === leadId);
  const msg = getMessageDefinition(msgKey);
  if (!msg || !lead) return;
  const content = msg.content
    .replace(/\[Contact\]/g, lead.contact || 'there')
    .replace(/\[Your Name\]/g, 'Your Name');
  navigator.clipboard.writeText(content).then(() => {
    btn.textContent = 'Copied';
    btn.classList.add('copied');
    setTimeout(() => { btn.textContent = 'Copy'; btn.classList.remove('copied'); }, 2000);
  });
}

function deleteLead(id) {
  const lead = leads.find(l => l.id === id);
  if (!lead) return;
  if (!confirm(`Are you sure you want to delete ${lead.name || 'this lead'}?`)) return;
  fetch(DELETE_WEBHOOK_URL, {
    method: 'POST',
    mode: 'cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ 'Business Name': lead.name }),
  })
    .then(res => {
      if (!res.ok) throw new Error('Delete request failed');
      leads = leads.filter(l => l.id !== id);
      currentLeadId = null;
      save();
      renderLeadList();
      if (document.getElementById('lead-detail')) document.getElementById('lead-detail').style.display = 'none';
      document.getElementById('empty-state').style.display = 'flex';
      showToast('Lead deleted');
    })
    .catch(err => {
      console.error('deleteLead error', err);
      showToast('Could not delete lead', 'error');
    });
}

function saveLeadNotes(id, btn) {
  const textarea = document.getElementById('lead-notes-input');
  if (!textarea) return;
  const lead = leads.find(l => l.id === id);
  if (!lead) return;
  const notes = textarea.value.trim();
  lead.notes = notes;
  syncLeadUpdate(lead)
    .then(() => {
      save();
      if (btn) {
        const original = btn.dataset.originalText || btn.textContent;
        btn.dataset.originalText = original;
        btn.textContent = 'Saved!';
        btn.classList.add('saved-note');
        setTimeout(() => {
          btn.textContent = btn.dataset.originalText || 'Save notes';
          btn.classList.remove('saved-note');
        }, 1600);
      }
      showToast('Notes saved');
    })
    .catch(err => {
      console.error('saveLeadNotes error', err);
      showToast('Unable to sync notes', 'error');
    });
}

function filterLeads() { renderLeadList(); }
function setFilter(f) {
  currentFilter = f;
  document.querySelectorAll('.filter-tab').forEach(t => t.classList.toggle('active', t.dataset.filter === f));
  renderLeadList();
}

function openModal() {
  document.getElementById('modal').classList.add('open');
  setTimeout(() => document.getElementById('inp-name').focus(), 50);
}

function closeModal() {
  document.getElementById('modal').classList.remove('open');
  ['inp-name','inp-phone','inp-contact','inp-website','inp-city','inp-state','inp-email','inp-notes'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
}

function addLead() {
  const name = document.getElementById('inp-name').value.trim();
  const phone = document.getElementById('inp-phone').value.trim();
  const contact = document.getElementById('inp-contact').value.trim();
  const website = document.getElementById('inp-website').value.trim();
  const city = document.getElementById('inp-city').value.trim();
  const state = document.getElementById('inp-state').value.trim();
  const email = document.getElementById('inp-email').value.trim();
  const notes = document.getElementById('inp-notes').value.trim();
  if (!name || !phone) { showToast('Name and phone required','error'); return; }
  const lead = {
    id: genId(),
    name, phone, contact,
    stage: 0,
    sentMessages: [],
    history: [{ code:'NEW', label:'Lead created', ts: now() }],
    createdAt: now(),
    lastMessageAt: null,
    notes,
    website,
    city,
    state,
    email,
  };
  leads.unshift(lead);
  save();
  closeModal();
  renderLeadList();
  selectLead(lead.id);
  showToast('Lead added');
}

function showOverview() {
  currentLeadId = null;
  document.getElementById('empty-state').style.display = 'none';
  hidePanels();
  document.getElementById('overview-area').style.display = 'block';
  renderLeadList();
  renderOverview();
}

function showScriptLibrary() {
  currentLeadId = null;
  document.getElementById('empty-state').style.display = 'none';
  hidePanels();
  document.getElementById('script-area').style.display = 'block';
  renderScriptEditor();
}

function setOverviewMetric(metric) {
  overviewMetric = metric;
  renderOverview();
}

function collectHistoryEvents() {
  return leads.flatMap(lead => (lead.history || []).map(evt => ({
    ...evt,
    leadId: lead.id,
    leadName: lead.name,
    ts: evt.ts || lead.createdAt || now(),
  })));
}

function getOverviewBuckets(events) {
  const buckets = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  for (let offset = 6; offset >= 0; offset--) {
    const bucketDate = new Date(today);
    bucketDate.setDate(bucketDate.getDate() - offset);
    const start = bucketDate.getTime();
    const end = start + 86400000 - 1;
    buckets.push({
      label: bucketDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }),
      start,
      end,
      leads: leads.filter(lead => lead.createdAt >= start && lead.createdAt <= end).length,
      actions: 0,
      advances: 0,
    });
  }
  events.forEach(event => {
    buckets.forEach(bucket => {
      if (event.ts >= bucket.start && event.ts <= bucket.end) {
        if (event.code !== 'NEW') bucket.actions++;
        if (event.code && event.code.startsWith('S')) bucket.advances++;
      }
    });
  });
  return buckets;
}

function renderOverview() {
  const ov = document.getElementById('overview-area');
  const colors = ['var(--accent)','var(--accent-soft)','var(--blue)','var(--green)','var(--accent-soft)'];
  const counts = STAGES.map((_,i) => leads.filter(l => l.stage === i).length);
  const actionItems = leads
    .filter(l => isActionDue(l) && l.stage < 5)
    .sort((a,b) => (a.lastMessageAt||a.createdAt) - (b.lastMessageAt||b.createdAt));

  const historyEvents = collectHistoryEvents();
  const buckets = getOverviewBuckets(historyEvents);
  const metricOption = METRIC_OPTIONS.find(opt => opt.id === overviewMetric) || METRIC_OPTIONS[0];
  const metricLabel = metricOption.label;
  const bucketMetricTotal = buckets.reduce((sum, bucket) => sum + (bucket[overviewMetric] || 0), 0);
  const maxMetricValue = Math.max(...buckets.map(bucket => bucket[overviewMetric] || 0));
  const scalingMax = maxMetricValue || 1;
  const sparkBars = buckets.map(bucket => {
    const value = bucket[overviewMetric] || 0;
    const height = value === 0 ? 6 : Math.max((value / scalingMax) * 100, 12);
    return `
      <div class="spark-bar" title="${bucket.label} · ${value}">
        <div class="spark-bar-track">
          <div class="spark-bar-fill" style="height:${height}%"></div>
        </div>
        <div class="spark-bar-label">${bucket.label}</div>
        <div class="spark-bar-value">${value}</div>
      </div>
    `;
  }).join('');
  const recentEvents = [...historyEvents].sort((a,b) => b.ts - a.ts).slice(0, 8);
  const activityHTML = recentEvents.length ? `
    <div class="activity-list">
      ${recentEvents.map(evt => `
        <div class="activity-item" ${evt.leadId ? `onclick="selectLead('${evt.leadId}')"` : ''}>
          <div class="activity-icon"></div>
          <div class="activity-text">
            <div class="activity-title">${escHTML(evt.leadName || 'Lead')}</div>
            <div class="activity-label">${escHTML(evt.label)}</div>
          </div>
          <div class="activity-time">${fmtTime(evt.ts)}</div>
        </div>
      `).join('')}
    </div>
  ` : `<div class="ov-empty">No activity recorded yet. Start logging actions to build a timeline.</div>`;
  const historicalSection = `
    <div class="historical-panel">
      <div class="historical-head">
        <div>
          <div class="section-label">Historical Metrics</div>
          <div class="section-headline">Last 7 days · ${metricLabel}</div>
        </div>
        <div class="metric-switch">
          ${METRIC_OPTIONS.map(option => `
            <button class="metric-btn ${overviewMetric === option.id ? 'active' : ''}" onclick="setOverviewMetric('${option.id}')">${option.label}</button>
          `).join('')}
        </div>
      </div>
      <div class="metric-summary">Week total: <strong>${bucketMetricTotal}</strong></div>
      <div class="sparkline">
        ${sparkBars}
      </div>
    </div>
  `;
  const activitySection = `
    <div class="recent-activity-panel">
      <div class="recent-activity-header">
        <div class="section-label">Recent Activity</div>
        <div class="section-headline">Tap an event to open the lead</div>
      </div>
      ${activityHTML}
    </div>
  `;

  ov.innerHTML = `
    <div class="overview">
      <div class="ov-title">PIPELINE OVERVIEW</div>
      <div class="ov-grid">
        ${STAGES.map((s,i) => `
          <div class="ov-card">
            <div class="ov-card-num" style="color:${colors[i]}">${counts[i]}</div>
            <div class="ov-card-label">${s.short}</div>
          </div>
        `).join('')}
      </div>
      <div class="section-label">Action Queue <span style="color:var(--accent);font-style:normal">${actionItems.length} needed</span></div>
      ${actionItems.length ? `
        <div class="action-queue">
          ${actionItems.map(l => {
            const next = getNextAction(l);
            const elapsed = (now() - (l.lastMessageAt||l.createdAt)) / 3600000;
            const urgClass = elapsed > 24 ? 'urg-now' : elapsed > 4 ? 'urg-soon' : 'urg-later';
            return `
              <div class="queue-item" onclick="selectLead('${l.id}')">
                <div class="queue-urgency ${urgClass}"></div>
                <div class="queue-info">
                  <div class="queue-name">${l.name}</div>
                  <div class="queue-action">${next ? (next.msg.type==='call'?'Call: '+next.msg.title:'Text: '+next.msg.title) : ''} · ${STAGES[l.stage]?.name}</div>
                </div>
                <div class="queue-time">${fmtTime(l.lastMessageAt||l.createdAt)}</div>
              </div>
            `;
          }).join('')}
        </div>
      ` : `<div style="color:var(--text3);font-size:13px;padding:20px 0">No actions needed right now. You're all caught up.</div>`}
      ${historicalSection}
      ${activitySection}
    </div>
  `;
}

function renderScriptEditor() {
  const select = document.getElementById('script-select');
  if (!select) return;
  select.innerHTML = Object.entries(MESSAGES).map(([key,msg]) => {
    const stage = STAGES[msg.stage]?.short || 'Other';
    const label = `${stage} · ${msg.code} · ${msg.title}`;
    return `<option value="${key}">${label}</option>`;
  }).join('');
  if (!activeScriptKey) activeScriptKey = Object.keys(MESSAGES)[0] || '';
  select.value = activeScriptKey;
  loadScriptEditor(activeScriptKey);
}

function loadScriptEditor(key) {
  if (!key) return;
  activeScriptKey = key;
  const msg = getMessageDefinition(key);
  if (!msg) return;
  const titleInput = document.getElementById('script-title-input');
  const timingInput = document.getElementById('script-timing-input');
  const contentInput = document.getElementById('script-content-input');
  if (titleInput) titleInput.value = msg.title;
  if (timingInput) timingInput.value = msg.timing || '';
  if (contentInput) contentInput.value = msg.content;
}

function saveScriptEdit() {
  const titleInput = document.getElementById('script-title-input');
  const timingInput = document.getElementById('script-timing-input');
  const contentInput = document.getElementById('script-content-input');
  if (!activeScriptKey || !titleInput || !contentInput) return;
  scriptOverrides[activeScriptKey] = {
    title: titleInput.value.trim(),
    timing: timingInput.value.trim(),
    content: contentInput.value.trim(),
  };
  save();
  showToast('Script saved');
  renderLeadList();
  renderDetail();
  loadScriptEditor(activeScriptKey);
}

function resetScriptEdit() {
  if (!activeScriptKey) return;
  delete scriptOverrides[activeScriptKey];
  save();
  loadScriptEditor(activeScriptKey);
  showToast('Script reset');
}

function escapeXml(value) {
  if (value === undefined || value === null) return '';
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function buildExcelSheet(name, header, rows) {
  const headerRow = header.map(cell => `<Cell><Data ss:Type="String">${escapeXml(cell)}</Data></Cell>`).join('');
  const bodyRows = rows.map(row => {
    const cells = row.map(cell => `<Cell><Data ss:Type="String">${escapeXml(cell)}</Data></Cell>`).join('');
    return `<Row>${cells}</Row>`;
  }).join('');
  return `
    <Worksheet ss:Name="${escapeXml(name)}">
      <Table>
        <Row>${headerRow}</Row>
        ${bodyRows}
      </Table>
    </Worksheet>
  `;
}

function buildExcelWorkbook(sheets) {
  return `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n${sheets.map(sheet => buildExcelSheet(sheet.name, sheet.header, sheet.rows)).join('')}\n</Workbook>`;
}

function exportLeads() {
  if (!leads.length) {
    showToast('No leads to export', 'error');
    return;
  }
  const leadHeaders = ['Lead Name','Phone','Contact','Stage','Stage Short','Status','Notes','Created At','Last Message At','History Count'];
  const leadRows = leads.map(lead => {
    const stageInfo = STAGES[lead.stage] || {};
    const createdAt = lead.createdAt ? new Date(lead.createdAt).toLocaleString('en-US') : '';
    const lastMessageAt = lead.lastMessageAt ? new Date(lead.lastMessageAt).toLocaleString('en-US') : '';
    const historyCount = (lead.history || []).length;
    const isComplete = lead.stage >= STAGES.length || (lead.stage === STAGES.length - 1 && (lead.sentMessages || []).includes('payment'));
    const statusLabel = isComplete ? 'Complete' : 'In Progress';
    return [
      lead.name || '',
      lead.phone || '',
      lead.contact || '',
      stageInfo.name || '',
      stageInfo.short || '',
      statusLabel,
      lead.notes || '',
      createdAt,
      lastMessageAt,
      historyCount.toString(),
    ];
  });
  const historyHeaders = ['Lead Name','Event','Code','Timestamp'];
  const historyRows = leads.flatMap(lead => (lead.history || []).map(evt => [
    lead.name || '',
    evt.label || '',
    evt.code || '',
    evt.ts ? new Date(evt.ts).toLocaleString('en-US') : '',
  ]));
  if (!historyRows.length) {
    historyRows.push(['', 'No history events recorded', '', '']);
  }
  const workbookXml = buildExcelWorkbook([
    { name: 'Leads', header: leadHeaders, rows: leadRows },
    { name: 'History', header: historyHeaders, rows: historyRows },
  ]);
  const blob = new Blob([workbookXml], { type: 'application/vnd.ms-excel' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  const filenameDate = new Date().toISOString().slice(0,10);
  link.download = `TharrosScout-Leads-${filenameDate}.xls`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
  showToast('Backup downloaded');
}

function showToast(msg, type) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  const isError = type === 'error';
  t.style.borderColor = isError ? 'var(--red)' : 'var(--green)';
  t.style.color = isError ? 'var(--red)' : 'var(--green)';
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2200);
}

document.getElementById('modal').addEventListener('click', function(e) {
  if (e.target === this) closeModal();
});
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });
document.getElementById('inp-contact').addEventListener('keydown', e => { if(e.key==='Enter') addLead(); });

document.getElementById('filter-tabs').addEventListener('click', function(e) {
  if (e.target.classList.contains('filter-tab')) {
    setFilter(e.target.dataset.filter);
  }
});

renderLeadList();
renderScriptEditor();
fetchLeads();
