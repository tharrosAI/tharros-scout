const STAGES = [
  { name: '1st Contact', short: '1st Contact', code: 'S1' },
  { name: 'Demo Permission', short: 'Demo', code: 'S2' },
  { name: 'Demo Complete', short: 'Demo Done', code: 'S3' },
  { name: 'Activation', short: 'Activation', code: 'S4' },
  { name: 'Payment + Onboarding', short: 'Payment', code: 'S5' },
  { name: 'Archive', short: 'Archive', code: 'ARCH' },
];

const ARCHIVE_STAGE = STAGES.length - 1;
const BADGE_CLASS = ['badge-s1','badge-s2','badge-s3','badge-s4','badge-s5'];

const VERTICALS = [
  { id: '1', label: 'Moving' },
  { id: '2', label: 'Plumbing' },
];

const PRODUCTS = [
  { id: '1', label: 'AI Receptionist' },
  { id: '2', label: 'Web Dev' },
];

const FUNNELS = [
  { id: '1', label: 'Web Dev Funnel', productId: '2', channel: 'Email', aliases: ['Web Dev', 'Website Funnel'] },
  { id: '2', label: 'High Intent Funnel', productId: '1', channel: 'SMS', aliases: ['High Intent', 'High Intent SMS'] },
  { id: '3', label: 'SMS Campaign', productId: '1', channel: 'SMS', aliases: ['SMS Campaign'] },
  { id: '4', label: 'Instagram DM', productId: '1', channel: 'Instagram DM', aliases: ['Instagram', 'IG DM'] },
  { id: '5', label: 'Facebook DM', productId: '1', channel: 'Facebook DM', aliases: ['Facebook', 'FB DM'] },
  { id: '6', label: 'Email Campaign', productId: '1', channel: 'Email', aliases: ['Email Blast', 'Email Funnel'] },
  { id: '7', label: 'Cold Call', productId: '1', channel: 'Call', aliases: ['Cold Call + SMS Follow-Up', 'Cold Call Funnel'] },
];

const FUNNELS_BY_ID = FUNNELS.reduce((acc, funnel) => {
  acc[funnel.id] = funnel;
  return acc;
}, {});

const FUNNELS_BY_PRODUCT = PRODUCTS.reduce((acc, product) => {
  acc[product.id] = FUNNELS.filter(funnel => funnel.productId === product.id);
  return acc;
}, {});

const normalizeText = value => String(value || '').trim().toLowerCase();

function findOption(list, rawValue) {
  const normalized = normalizeText(rawValue);
  if (!normalized) return undefined;
  for (const item of list) {
    if (normalizeText(item.id) === normalized) return item;
    if (normalizeText(item.label) === normalized) return item;
    if ((item.aliases || []).some(alias => {
      const aliasNorm = normalizeText(alias);
      return aliasNorm === normalized || aliasNorm.includes(normalized) || normalized.includes(aliasNorm);
    })) return item;
  }
  return undefined;
}

function resolveOptionId(list, rawValue, fallback) {
  const candidate = findOption(list, rawValue);
  if (candidate) return candidate.id;
  if (fallback) return fallback;
  return list[0]?.id || '';
}

function resolveVerticalId(rawValue) {
  return resolveOptionId(VERTICALS, rawValue, VERTICALS[0]?.id || '');
}

function resolveProductId(rawValue) {
  return resolveOptionId(PRODUCTS, rawValue, PRODUCTS[0]?.id || '');
}

function resolveFunnelId(rawValue, productId) {
  const funnelList = FUNNELS_BY_PRODUCT[productId] || FUNNELS;
  const fallbackId = funnelList[0]?.id || FUNNELS[0]?.id || '';
  const match = findOption(funnelList, rawValue) || findOption(FUNNELS, rawValue);
  if (match) return match.id;
  return fallbackId;
}

const ASSET_STEPS = [
  { id: '1', label: 'Primary Message' },
  { id: '2', label: 'Follow-Up' },
  { id: '3', label: 'Bump' },
  { id: '4', label: 'Close' },
];

const FILTER_STATE = {
  stage: 'all',
  vertical: '',
  product: '',
  funnel: '',
  channel: '',
  search: '',
};

const LEAD_SYNC_WEBHOOK = 'https://automation.tharrosai.com/webhook/bc225c06-6ba3-49e7-9169-d5c49729591b';
const FETCH_LEADS_WEBHOOK = 'https://automation.tharrosai.com/webhook/127e7f4a-dd40-4f7c-b995-05efb0977b9a';
const DELETE_WEBHOOK_URL = 'https://tharros.app.n8n.cloud/webhook/1456e7f1-fbb4-4d32-975f-0b5a5f8132cd';
const ASSET_FETCH_URL = 'https://automation.tharrosai.com/webhook/907069ea-0044-4bda-b0c3-2529dd377014';
const ASSET_UPDATE_URL = 'https://automation.tharrosai.com/webhook/a90b8352-a749-499a-b711-742fc7e62814';
const METRIC_FETCH_URL = 'https://automation.tharrosai.com/webhook/71c97286-b92b-4995-82e3-c665468c2aa';
const METRIC_UPDATE_URL = 'https://automation.tharrosai.com/webhook/b22e1cc7-a530-42bc-b7df-b2e8324909ce';

const METRIC_OPTIONS = [
  { id: 'leads', label: 'Leads Added', color: '#0df2ff' },
  { id: 'actions', label: 'Actions Logged', color: '#8c5cff' },
  { id: 'advances', label: 'Stages Advanced', color: '#43ffa1' },
  { id: 'customers', label: 'Customers', color: '#f7b500' },
];
let leads = [];
let currentLeadId = null;
let assetsById = {};
let metricsData = [];
let activeAssetSelection = {
  vertical: VERTICALS[0]?.id || '1',
  product: PRODUCTS[0]?.id || '1',
  funnel: FUNNELS_BY_PRODUCT[PRODUCTS[0]?.id]?.[0]?.id || FUNNELS[0]?.id || '1',
  stage: '1',
  step: '1',
};
let stageFilter = 'all';
let pendingLeadSelection = null;
let pendingLeads = [];

function updateLead(lead) {
  if (!lead || !lead.name) return Promise.resolve();
  setSyncState('start');
  const payload = {
    row_number: lead.rowNumber || lead.row_number || lead.row || '',
    'Business Name': lead.name,
    Phone: lead.phone || '',
    'Contact Name': lead.contact || '',
    Website: lead.website || '',
    City: lead.city || '',
    State: lead.state || '',
    Email: lead.email || '',
    Notes: lead.notes || '',
    Vertical: getVerticalLabel(lead.verticalId),
    Product: getProductLabel(lead.productId),
    Funnel: getFunnelLabel(lead.funnelId),
    Channel: lead.channel || FUNNELS_BY_ID[lead.funnelId]?.channel || '',
    Stage: lead.stage,
    'Sent Messages': JSON.stringify(Array.isArray(lead.sentMessages) ? lead.sentMessages : []),
  };
  return fetch(LEAD_SYNC_WEBHOOK, {
    method: 'POST',
    mode: 'cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
    .then(res => {
      if (!res.ok) throw new Error('Lead update failed');
      setSyncState('success');
      return res;
    })
    .catch(err => {
      setSyncState('failure');
      throw err;
    });
}

function setSyncState(state) {
  const btn = document.querySelector('.btn-refresh');
  if (!btn) return;
  btn.classList.remove('syncing', 'synced', 'failed');
  if (state === 'start') {
    btn.classList.add('syncing');
    return;
  }
  if (state === 'success') {
    btn.classList.add('synced');
    setTimeout(() => btn.classList.remove('synced'), 900);
    return;
  }
  if (state === 'failure') {
    btn.classList.add('failed');
    setTimeout(() => btn.classList.remove('failed'), 900);
  }
}

function fetchLeads(options = {}) {
  const { preserveSelection = false } = options;
  const empty = document.getElementById('empty-state');
  const detail = document.getElementById('lead-detail');
  if (!preserveSelection) {
    leads = [];
    currentLeadId = null;
    renderLeadList();
    if (detail) detail.classList.add('hidden');
  }
  if (empty) empty.classList.add('hidden');
  fetch(FETCH_LEADS_WEBHOOK, { method: 'GET', mode: 'cors', headers: { 'Accept': 'application/json' } })
    .then(res => {
      if (!res.ok) throw new Error('Network response was not ok');
      return res.json();
    })
    .then(data => {
      const rows = extractRows(data);
      leads = rows.map(normalizeLeadRow);
      mergePendingLeads();
      renderLeadList();
      if (leads.length) {
        const focused = resolvePendingLead();
        if (!focused && !preserveSelection) showOverview();
      } else {
        showEmptyState();
      }
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
  const lastMessageRaw = resolveField(row, 'lastMessageAt') || resolveField(row, 'Last Message At');
  const rowNumber = resolveField(row, 'row_number') || resolveField(row, 'rowNumber') || resolveField(row, 'row');
  const verticalRaw = resolveField(row, 'Vertical');
  const verticalId = resolveVerticalId(verticalRaw);
  const productRaw = resolveField(row, 'Product');
  const productId = resolveProductId(productRaw);
  const funnelValue = resolveField(row, 'Funnel') || resolveField(row, 'Funnel Name');
  const funnelId = resolveFunnelId(funnelValue, productId);
  const funnel = FUNNELS_BY_ID[funnelId];
  const channel = funnel?.channel || resolveField(row, 'Channel') || '';
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
    createdAt: parseTimestamp(resolveField(row, 'time')) || parseTimestamp(resolveField(row, 'Created At')) || now(),
    lastMessageAt: parseTimestamp(lastMessageRaw),
    rowNumber: Number(rowNumber) || null,
    verticalId,
    productId,
    funnelId,
    verticalLabel: getVerticalLabel(verticalId, verticalRaw || 'Moving'),
    productLabel: getProductLabel(productId, productRaw || 'AI Receptionist'),
    funnelLabel: getFunnelLabel(funnelId, funnelValue || 'General'),
    channel,
  };
}

function extractRows(data) {
  if (Array.isArray(data)) return data;
  return data?.leads || data?.rows || data?.items || data?.data || [];
}

function findFunnelId(value, productId) {
  return resolveFunnelId(value, productId);
}

function getVerticalLabel(id, fallback = 'Moving') {
  if (!id) return fallback;
  const match = VERTICALS.find(item => item.id === id);
  if (match) return match.label;
  return id;
}

function getProductLabel(id, fallback = 'AI Receptionist') {
  if (!id) return fallback;
  const match = PRODUCTS.find(item => item.id === id);
  if (match) return match.label;
  return id;
}

function getFunnelLabel(id, fallback = 'General') {
  if (!id) return fallback;
  const match = FUNNELS_BY_ID[id];
  if (match) return match.label;
  return id;
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

function buildAssetId({ verticalId, productId, funnelId, stageIdx, stepId }) {
  const safeVertical = verticalId || VERTICALS[0]?.id || '1';
  const safeProduct = productId || PRODUCTS[0]?.id || '1';
  const safeFunnel = funnelId || FUNNELS_BY_PRODUCT[safeProduct]?.[0]?.id || FUNNELS[0]?.id || '1';
  const boundedStage = Math.min(Math.max(stageIdx, 0), STAGES.length - 2);
  const stageDigit = (boundedStage + 1).toString();
  return `${safeVertical}${safeProduct}${safeFunnel}${stageDigit}${stepId}`;
}

function getStageAssets(lead) {
  if (!lead || lead.stage >= ARCHIVE_STAGE) return [];
  return ASSET_STEPS.map(step => buildAssetId({
    verticalId: lead.verticalId,
    productId: lead.productId,
    funnelId: lead.funnelId,
    stageIdx: Math.min(Math.max(lead.stage, 0), STAGES.length - 2),
    stepId: step.id,
  }));
}

function getAssetStepLabel(stepId) {
  return ASSET_STEPS.find(step => step.id === stepId)?.label || 'Message';
}

function getAssetDefinition(assetId, lead) {
  const content = (assetsById[assetId] || '').trim();
  const stepId = assetId?.slice(-1);
  const stageDigit = assetId?.slice(-2, -1);
  const stageIdx = Number(stageDigit) - 1;
  return {
    id: assetId,
    content,
    stepLabel: getAssetStepLabel(stepId),
    stageLabel: STAGES[stageIdx]?.short || STAGES[lead.stage]?.short || '',
    channel: lead?.channel || FUNNELS_BY_ID[lead?.funnelId]?.channel || '',
  };
}

function getNextAction(lead) {
  if (!lead || lead.stage >= ARCHIVE_STAGE) return null;
  const sent = lead.sentMessages || [];
  for (const assetId of getStageAssets(lead)) {
    if (!assetId || sent.includes(assetId)) continue;
    return { assetId, definition: getAssetDefinition(assetId, lead) };
  }
  return null;
}

function isActionDue(lead) {
  return Boolean(getNextAction(lead));
}

function renderLeadList() {
  const list = document.getElementById('lead-list');
  if (!list) return;
  const search = (FILTER_STATE.search || '').toLowerCase();
  let filtered = leads.filter(lead => {
    if (stageFilter === 'archived') {
      return lead.stage === ARCHIVE_STAGE;
    }
    if (lead.stage === ARCHIVE_STAGE) return false;
    if (stageFilter === 'action' && !isActionDue(lead)) return false;
    if (stageFilter !== 'all' && stageFilter !== 'action') {
      return lead.stage === parseInt(stageFilter, 10);
    }
    if (FILTER_STATE.vertical && lead.verticalId !== FILTER_STATE.vertical) return false;
    if (FILTER_STATE.product && lead.productId !== FILTER_STATE.product) return false;
    if (FILTER_STATE.funnel && lead.funnelId !== FILTER_STATE.funnel) return false;
    if (FILTER_STATE.channel && lead.channel && lead.channel.toLowerCase() !== FILTER_STATE.channel.toLowerCase()) return false;
    if (search) {
      const haystack = `${lead.name} ${lead.contact || ''} ${lead.phone || ''}`.toLowerCase();
      if (!haystack.includes(search)) return false;
    }
    return true;
  });

  filtered.sort((a,b) => {
    const aUrgent = isActionDue(a); const bUrgent = isActionDue(b);
    if (aUrgent && !bUrgent) return -1;
    if (!aUrgent && bUrgent) return 1;
    return (b.lastMessageAt||b.createdAt) - (a.lastMessageAt||a.createdAt);
  });

  document.getElementById('hdr-total').textContent = leads.length;
  const needAction = leads.filter(l => isActionDue(l) && l.stage < ARCHIVE_STAGE).length;
  document.getElementById('hdr-action').textContent = needAction;

  if (!filtered.length) {
    list.innerHTML = `<div style="padding:32px 20px;text-align:center;color:var(--text3);font-size:13px">No leads found</div>`;
    return;
  }

  list.innerHTML = filtered.map((lead,i) => {
    const active = lead.id === currentLeadId;
    const action = isActionDue(lead);
    const bc = BADGE_CLASS[lead.stage] || 'badge-s1';
    const metaParts = [lead.verticalLabel, lead.productLabel, lead.funnelLabel, lead.channel].filter(Boolean);
    const metaHTML = metaParts.map((part,idx) => `${idx ? '<span class="meta-dot">·</span>' : ''}<span>${escHTML(part)}</span>`).join('');
    return `
      <div class="lead-item ${active?'active':''}" onclick="selectLead('${lead.id}')">
        <div class="lead-item-top">
          <span class="lead-name">${lead.name}</span>
          <span class="lead-time">${fmtTime(lead.lastMessageAt||lead.createdAt)}</span>
        </div>
        <div class="lead-phone">${lead.phone || '—'}</div>
        <div class="lead-meta">
          ${metaHTML}
        </div>
        <span class="lead-stage-badge ${bc}">${STAGES[lead.stage]?.short||''}</span>
        ${action ? '<div class="action-dot"></div>' : ''}
      </div>
      ${i < filtered.length-1 ? '<div class="lead-item-divider"></div>' : ''}
    `;
  }).join('');
}

function resolvePendingLead() {
  if (!pendingLeadSelection) return false;
  const match = leads.find(lead => !lead.isTemp && lead.phone === pendingLeadSelection.phone && lead.name === pendingLeadSelection.name);
  if (match) {
    pendingLeadSelection = null;
    selectLead(match.id);
    return true;
  }
  return false;
}

function cleanPhone(value) {
  if (!value) return '';
  return String(value).replace(/\D+/g, '');
}

function getLeadKey(lead = {}) {
  const name = (lead.name || lead['Business Name'] || lead.contact || '').toString().toLowerCase().trim();
  const phone = cleanPhone(lead.phone || lead.Phone || lead.contactPhone || '');
  return `${name}::${phone}`;
}

function queuePendingLead(lead) {
  const key = getLeadKey(lead);
  pendingLeads = pendingLeads.filter(entry => entry.key !== key);
  lead.pendingKey = key;
  lead.isTemp = true;
  pendingLeads.push({ key, lead });
}

function mergePendingLeads() {
  if (!pendingLeads.length) return;
  let combined = [...leads];
  pendingLeads = pendingLeads.filter(entry => {
    const key = entry.key;
    const hasActual = combined.some(item => !item.isTemp && getLeadKey(item) === key);
    if (hasActual) {
      combined = combined.filter(item => !(item.isTemp && item.pendingKey === key));
      return false;
    }
    const alreadyInserted = combined.some(item => item.pendingKey === key && item.isTemp);
    if (!alreadyInserted) {
      combined = [entry.lead, ...combined];
    }
    return true;
  });
  leads = combined;
}

function insertTemporaryLead(lead) {
  leads = [lead, ...leads];
  selectLead(lead.id);
}

function hidePanels() {
  ['lead-detail', 'overview-area', 'metrics-area', 'script-area'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.add('hidden');
  });
}

function selectLead(id) {
  currentLeadId = id;
  renderLeadList();
  renderDetail();
  const empty = document.getElementById('empty-state');
  if (empty) empty.classList.add('hidden');
  hidePanels();
  const detail = document.getElementById('lead-detail');
  if (detail) detail.classList.remove('hidden');
}

function renderDetail() {
  const lead = leads.find(l => l.id === currentLeadId);
  if (!lead) return;
  const el = document.getElementById('lead-detail');
  const nextAction = getNextAction(lead);
  const stageAssets = getStageAssets(lead);
  const allAssetsSent = stageAssets.length > 0 && stageAssets.every(id => lead.sentMessages?.includes(id));
  const isComplete = lead.stage >= ARCHIVE_STAGE || allAssetsSent;

  const timelineStages = STAGES.slice(0, -1);
  const maxTimelineIndex = timelineStages.length - 1;
  const stageForProgress = Math.min(lead.stage, maxTimelineIndex);
  const progressPct = maxTimelineIndex > 0 ? (stageForProgress / maxTimelineIndex) * 100 : 0;
  const nodesHTML = timelineStages.map((stage, index) => {
    const done = index < lead.stage;
    const current = index === lead.stage;
    const statusClass = done ? 'done' : current ? 'current' : '';
    return `
      <div class="pipeline-node ${statusClass}">
        <div class="pipeline-node-circle">${done ? '✓' : ''}</div>
        <span class="pipeline-node-label">${stage.short}</span>
      </div>
    `;
  }).join('');
  const pipelineHTML = `
    <div class="pipeline">
      <div class="pipeline-visual">
        <div class="pipeline-track-wrap">
          <div class="pipeline-track-line"></div>
          <div class="pipeline-track-progress" style="width:${progressPct}%"></div>
        </div>
        <div class="pipeline-nodes">
          ${nodesHTML}
        </div>
      </div>
    </div>
  `;

  let actionHTML = '';
  if (lead.stage >= ARCHIVE_STAGE) {
    actionHTML = `<div class="next-ribbon"><div class="next-icon"></div><div class="next-text"><div class="next-label">Archived</div><div class="next-desc">This lead was archived.</div></div></div>`;
  } else if (nextAction && nextAction.definition) {
    const def = nextAction.definition;
    actionHTML = `
      <div class="next-ribbon">
        <div class="next-icon"></div>
        <div class="next-text">
          <div class="next-label">Next Action</div>
          <div class="next-desc">${def.stepLabel} · ${def.stageLabel}${def.channel ? ` · ${def.channel}` : ''}</div>
        </div>
        <span class="asset-id-badge">${def.id}</span>
      </div>
    `;
  } else {
    actionHTML = `<div class="next-ribbon"><div class="next-icon"></div><div class="next-text"><div class="next-label">Waiting</div><div class="next-desc">All messages sent — advance the lead when they reply.</div></div></div>`;
  }

  const messagesHTML = stageAssets.map(assetId => {
    const asset = getAssetDefinition(assetId, lead);
    if (!asset) return '';
    const isSent = lead.sentMessages?.includes(assetId);
    const isNext = nextAction?.assetId === assetId;
    const content = asset.content || 'Script missing for this combination.';
    return `
      <div class="action-card ${isSent ? 'sent-card' : ''}">
        <div class="action-card-label">
          ${asset.stepLabel}
          ${isSent ? '<span class="sent-pill">✓ Sent</span>' : ''}
          ${isNext ? '<span class="next-pill">Next</span>' : ''}
        </div>
        <div class="action-card-title">${asset.stageLabel}</div>
        <div class="message-box" style="position:relative">
          ${escHTML(content)}
          <button class="copy-btn" onclick="copyMsg(this,'${assetId}','${lead.id}')">Copy</button>
        </div>
        ${!isSent ? `
          <div class="stage-actions">
            <button class="stage-btn advance" onclick="markSent('${lead.id}','${assetId}')">✓ Mark as Sent</button>
          </div>` : ''}
      </div>
    `;
  }).join('');

  const futureStages = STAGES
    .map((stage, index) => ({ stage, index }))
    .filter(({ index }) => index > lead.stage && index < ARCHIVE_STAGE);
  const stageButtons = futureStages
    .map(({ stage, index }) => `<button class="stage-pill" onclick="moveToStage(${index})">${stage.name} →</button>`)
    .join('\n');
  const stageMessage = allAssetsSent ? 'All messages sent · move them once you hear back' : 'Next asset ready · reach out now';
  const currentStage = STAGES[lead.stage] || STAGES[0];
  const stageProgressionHTML = `
    <div class="stage-progress-top">
      <div>
        <div class="stage-progression-head">Current stage</div>
        <div class="stage-progress-current">${currentStage.name}</div>
      </div>
      <div class="stage-progress-status">${stageMessage}</div>
    </div>
    <div class="stage-progression-actions">
      ${stageButtons || '<span class="stage-empty">No further stages · lead is ready</span>'}
    </div>
  `;

  const notesHTML = `
    <div class="notes-area">
      <label class="widget-subtext">Lead Notes</label>
      <textarea id="lead-notes-input" placeholder="Capture context that matters">${lead.notes || ''}</textarea>
      <div class="notes-actions">
        <button class="btn save-notes-btn" onclick="saveLeadNotes('${lead.id}', this)">Save notes</button>
      </div>
    </div>
  `;

  const location = [lead.city, lead.state].filter(Boolean).join(', ');
  const emailLine = lead.email ? `<div class="detail-meta-item">Email · <a href="mailto:${escHTML(lead.email)}">${escHTML(lead.email)}</a></div>` : '';
  const websiteLine = lead.website ? `<div class="detail-meta-item">Website · <a href="${escHTML(lead.website)}" target="_blank" rel="noreferrer">${escHTML(lead.website)}</a></div>` : '';
  const locationLine = location ? `<div class="detail-meta-item">Location · ${escHTML(location)}</div>` : '';
  const tagHTML = `
    <div class="detail-tags">
      ${[lead.verticalLabel, lead.productLabel, lead.funnelLabel, lead.channel]
        .filter(Boolean)
        .map(tag => `<span>${escHTML(tag)}</span>`)
        .join('')}
    </div>
  `;
  const actionCard = `
    <div class="detail-card detail-next-card">
      <div class="detail-card-title">Next Action</div>
      ${actionHTML}
    </div>
  `;
  const notesCard = `
    <div class="detail-card detail-notes-card">
      <div class="detail-card-title">Lead Notes</div>
      ${notesHTML}
    </div>
  `;
  const stageProgressionCard = `
    <div class="detail-card detail-advance-card">
      <div class="detail-card-title">Stage Progression</div>
      ${stageProgressionHTML}
    </div>
  `;
  const pipelineRow = `<div class="pipeline-row">${pipelineHTML}</div>`;
  el.innerHTML = `
    <div class="detail-content">
      <div class="detail-header">
        <div>
          <div class="detail-name">${lead.name}</div>
          <div class="detail-phone">${lead.phone}${lead.contact ? ' · ' + lead.contact : ''}</div>
          ${tagHTML}
          <div class="detail-contact-meta">
            ${websiteLine}
            ${locationLine}
            ${emailLine}
          </div>
        </div>
        <div class="detail-cta">
          <div class="detail-added">Added ${fmtAbsTime(lead.createdAt)}</div>
          ${lead.stage !== ARCHIVE_STAGE ? `<button class="btn btn-outline detail-archive" onclick="archiveLead('${lead.id}')">Archive</button>` : ''}
        </div>
      </div>
      ${pipelineRow}
      <div class="detail-grid">
        ${actionCard}
        ${stageProgressionCard}
      </div>
      <div class="detail-grid notes-grid">
        ${notesCard}
      </div>
      <div class="detail-messages">
        ${messagesHTML}
      </div>
    </div>
  `;
}

function escHTML(str = '') {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>');
}

function markSent(leadId, assetId) {
  const lead = leads.find(l => l.id === leadId);
  if (!lead || !assetId) return;
  if (!lead.sentMessages) lead.sentMessages = [];
  if (lead.sentMessages.includes(assetId)) return;
  lead.sentMessages.push(assetId);
  if (!lead.history) lead.history = [];
  const asset = getAssetDefinition(assetId, lead);
  lead.history.push({ code: asset.stepLabel, label: `Sent ${asset.stepLabel}`, ts: now() });
  lead.lastMessageAt = now();
  updateLead(lead)
    .then(() => logMetricEvent('actions'))
    .then(() => {
      renderLeadList();
      renderDetail();
      showToast('Marked as sent');
    })
    .catch(err => {
      console.error('markSent sync error', err);
      showToast('Could not sync sent message', 'error');
    });
}

function moveToStage(stageIdx) {
  const lead = leads.find(l => l.id === currentLeadId);
  if (!lead || stageIdx <= lead.stage) return;
  lead.stage = stageIdx;
  lead.sentMessages = [];
  if (!lead.history) lead.history = [];
  lead.history.push({ code: `S${stageIdx+1}`, label: `Advanced to: ${STAGES[stageIdx].name}`, ts: now() });
  lead.lastMessageAt = now();
  updateLead(lead)
    .then(() => {
      logMetricEvent('advances');
      if (stageIdx === 4) logMetricEvent('customers');
      renderLeadList();
      renderDetail();
      showToast(`Moved to ${STAGES[stageIdx].name}`);
    })
    .catch(err => {
      console.error('moveToStage sync error', err);
      showToast('Could not sync stage change', 'error');
    });
}

function copyMsg(btn, assetId, leadId) {
  const lead = leads.find(l => l.id === leadId);
  if (!lead || !assetId) return;
  const asset = getAssetDefinition(assetId, lead);
  if (!asset || !asset.content) return;
  navigator.clipboard.writeText(asset.content).then(() => {
    btn.textContent = 'Copied';
    btn.classList.add('copied');
    setTimeout(() => { btn.textContent = 'Copy'; btn.classList.remove('copied'); }, 2000);
  });
}

function archiveLead(id) {
  const lead = leads.find(l => l.id === id);
  if (!lead || lead.stage === ARCHIVE_STAGE) return;
  lead.stage = ARCHIVE_STAGE;
  if (!lead.history) lead.history = [];
  lead.history.push({ code: 'ARCH', label: 'Lead archived', ts: now() });
  updateLead(lead)
    .then(() => {
      showToast('Lead archived');
    })
    .catch(err => {
      console.error('archiveLead error', err);
      showToast('Could not sync archive', 'error');
    });
  renderLeadList();
  renderDetail();
}

function saveLeadNotes(id, btn) {
  const textarea = document.getElementById('lead-notes-input');
  if (!textarea) return;
  const lead = leads.find(l => l.id === id);
  if (!lead) return;
  const notes = textarea.value.trim();
  lead.notes = notes;
  updateLead(lead)
    .then(() => {
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

function filterLeads() {
  FILTER_STATE.search = document.getElementById('search-input').value || '';
  renderLeadList();
}

function setFilter(f) {
  stageFilter = f;
  document.querySelectorAll('.filter-tab').forEach(t => t.classList.toggle('active', t.dataset.filter === f));
  renderLeadList();
}

function setDropdownFilter(type, value) {
  FILTER_STATE[type] = value;
  renderLeadList();
}

function resetFilters() {
  stageFilter = 'all';
  FILTER_STATE.vertical = '';
  FILTER_STATE.product = '';
  FILTER_STATE.funnel = '';
  FILTER_STATE.channel = '';
  FILTER_STATE.search = '';
  document.getElementById('search-input').value = '';
  ['filter-vertical','filter-product','filter-funnel','filter-channel'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  document.querySelectorAll('.filter-tab').forEach(t => t.classList.toggle('active', t.dataset.filter === 'all'));
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
  const verticalId = document.getElementById('inp-vertical').value || VERTICALS[0].id;
  const productId = document.getElementById('inp-product').value || PRODUCTS[0].id;
  const funnelId = document.getElementById('inp-funnel').value || FUNNELS_BY_PRODUCT[productId]?.[0]?.id || FUNNELS[0].id;
  const funnel = FUNNELS_BY_ID[funnelId];
  const channel = funnel?.channel || '';
  if (!name || !phone) { showToast('Name and phone required','error'); pendingLeadSelection = null; return; }
  const nowTs = now();
  const tempLead = {
    id: genId(),
    name,
    phone,
    contact,
    website,
    city,
    state,
    email,
    notes,
    stage: 0,
    sentMessages: [],
    history: [],
    verticalId,
    productId,
    funnelId,
    verticalLabel: getVerticalLabel(verticalId),
    productLabel: getProductLabel(productId),
    funnelLabel: getFunnelLabel(funnelId),
    channel,
    createdAt: nowTs,
    lastMessageAt: nowTs,
  };
  queuePendingLead(tempLead);
  insertTemporaryLead(tempLead);
  pendingLeadSelection = { name, phone, timestamp: Date.now() };
  const payload = {
    Vertical: tempLead.verticalLabel,
    Product: tempLead.productLabel,
    Funnel: tempLead.funnelLabel,
    Channel: channel,
    'Business Name': name,
    'Contact Name': contact,
    Phone: phone,
    Email: email,
    Website: website,
    City: city,
    State: state,
    Notes: notes,
    Stage: 0,
    'Sent Messages': JSON.stringify([]),
  };
  fetch(LEAD_SYNC_WEBHOOK, {
    method: 'POST',
    mode: 'cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
    .then(res => {
      if (!res.ok) throw new Error('Lead creation failed');
      return res.json();
    })
    .then(() => {
      closeModal();
      showToast('Lead added');
      logMetricEvent('leads');
      fetchLeads({ preserveSelection: true });
    })
    .catch(err => {
      console.error('addLead error', err);
      showToast('Could not add lead', 'error');
      pendingLeadSelection = null;
    });
}

function showOverview() {
  currentLeadId = null;
  hidePanels();
  const overview = document.getElementById('overview-area');
  if (overview) overview.classList.remove('hidden');
  const empty = document.getElementById('empty-state');
  if (empty) empty.classList.add('hidden');
  renderLeadList();
  renderOverview();
}

function showHistoricalMetrics() {
  currentLeadId = null;
  hidePanels();
  const metrics = document.getElementById('metrics-area');
  if (metrics) metrics.classList.remove('hidden');
  const empty = document.getElementById('empty-state');
  if (empty) empty.classList.add('hidden');
  renderHistoricalMetrics();
}

function showScriptLibrary() {
  currentLeadId = null;
  hidePanels();
  const scriptArea = document.getElementById('script-area');
  if (scriptArea) scriptArea.classList.remove('hidden');
  const empty = document.getElementById('empty-state');
  if (empty) empty.classList.add('hidden');
  renderScriptEditor();
}

function showEmptyState() {
  currentLeadId = null;
  hidePanels();
  const empty = document.getElementById('empty-state');
  if (empty) empty.classList.remove('hidden');
}

function collectHistoryEvents() {
  return leads.flatMap(lead => (lead.history || []).map(evt => ({
    ...evt,
    leadId: lead.id,
    leadName: lead.name,
    ts: evt.ts || lead.createdAt || now(),
  })));
}

function renderOverview() {
  const ov = document.getElementById('overview-area');
  if (!ov) return;
  const stageCounts = STAGES.map((_, i) => leads.filter(l => l.stage === i).length);
  const stageCards = STAGES.slice(0, -1)
    .map((stage, index) => {
      const count = stageCounts[index];
      const waiting = leads.filter(l => l.stage === index && isActionDue(l)).length;
      return `
        <div class="stage-card">
          <div class="stage-card-title">${stage.short}</div>
          <div class="stage-card-count">${count}<span>leads</span></div>
          <div class="stage-card-sub">${waiting} need action</div>
        </div>
      `;
    })
    .join('');
  const actionItems = leads
    .filter(l => l.stage < ARCHIVE_STAGE && isActionDue(l))
    .sort((a, b) => (a.lastMessageAt || a.createdAt) - (b.lastMessageAt || b.createdAt));
  const actionQueue = actionItems.length
    ? `
      <div class="queue-panel">
        <div class="panel-heading">
          <div class="panel-title">Action Queue</div>
          <div class="panel-sub">Prioritized by last contact</div>
        </div>
        <div class="action-queue">
          ${actionItems
            .map(l => {
              const next = getNextAction(l);
              const label = next?.definition
                ? `${next.definition.stepLabel} · ${next.definition.stageLabel}`
                : `${STAGES[l.stage]?.short || 'Pending'}`;
              return `
                <div class="queue-item" onclick="selectLead('${l.id}')">
                  <div class="queue-urgency"></div>
                  <div class="queue-info">
                    <div class="queue-name">${l.name}</div>
                    <div class="queue-action">${label}</div>
                  </div>
                  <div class="queue-time">${fmtTime(l.lastMessageAt || l.createdAt)}</div>
                </div>
              `;
            })
            .join('')}
        </div>
      </div>
    `
    : `<div class="queue-panel empty">No actions queued yet. Open a lead to continue.</div>`;
  const historyEvents = collectHistoryEvents();
  const activityHTML = historyEvents.length
    ? `
      <div class="activity-panel">
        <div class="panel-heading">
          <div class="panel-title">Recent Activity</div>
          <div class="panel-sub">Tap an event to open that lead</div>
        </div>
        <div class="activity-list">
          ${historyEvents
            .slice(-6)
            .map(evt => `
              <div class="activity-item" ${evt.leadId ? `onclick="selectLead('${evt.leadId}')"` : ''}>
                <div class="activity-icon"></div>
                <div class="activity-text">
                  <div class="activity-title">${escHTML(evt.leadName || 'Lead')}</div>
                  <div class="activity-label">${escHTML(evt.label)}</div>
                </div>
                <div class="activity-time">${fmtTime(evt.ts)}</div>
              </div>
            `)
            .join('')}
        </div>
      </div>
    `
    : `<div class="activity-panel empty">No activity recorded yet. Send a message to start tracking.</div>`;
  ov.innerHTML = `
    <div class="war-room">
      <div class="overview-head">
        <div>
          <div class="section-label">War Room</div>
          <div class="section-headline">Review your pipeline and decide the next steps</div>
        </div>
        <div class="overview-cta">
          <div class="overview-cta-label">Action needed</div>
          <div class="overview-cta-value">${actionItems.length}</div>
        </div>
      </div>
      <div class="stage-grid">
        ${stageCards}
      </div>
      <div class="overview-body">
        ${actionQueue}
        ${activityHTML}
      </div>
    </div>
  `;
}

function renderHistoricalMetrics() {
  const panel = document.getElementById('metrics-area');
  if (!panel) return;
  const rows = metricsData.length ? metricsData.slice(-7) : buildEmptyMetricSeries();
  const summaryChips = METRIC_OPTIONS.map(option => {
    const latestValue = rows[rows.length - 1]?.[option.id] || 0;
    return `
      <div class="metric-summary-chip" style="--chip-accent:${option.color};">
        <div class="chip-label">${option.label}</div>
        <div class="chip-value">${latestValue}</div>
      </div>
    `;
  }).join('');
  const detailCards = METRIC_OPTIONS.map(option => {
    const cardMax = Math.max(6, ...rows.map(row => row[option.id] || 0));
    const bars = rows.map(row => {
      const value = row[option.id] || 0;
      const height = value === 0 ? 6 : Math.max((value / cardMax) * 100, 12);
      return `
        <div class="metric-detail-bar" title="${formatMetricLabel(row.Date)} · ${value}">
          <div class="metric-detail-bar-track">
            <div class="metric-detail-bar-fill" style="height:${height}%; background:${option.color};"></div>
          </div>
          <div class="metric-detail-bar-label">${formatMetricLabel(row.Date)}</div>
        </div>
      `;
    }).join('');
    const latestValue = rows[rows.length - 1]?.[option.id] || 0;
    const total = rows.reduce((sum, row) => sum + (row[option.id] || 0), 0);
    return `
      <div class="metric-detail-card">
        <div class="metric-detail-top">
          <span class="metric-detail-label" style="color:${option.color};">${option.label}</span>
          <span class="metric-detail-value">${latestValue}</span>
        </div>
        <div class="metric-detail-chart">
          ${bars}
        </div>
        <div class="metric-detail-footer">
          <span>Week total <strong>${total}</strong></span>
          <span>Latest <strong>${latestValue}</strong></span>
        </div>
      </div>
    `;
  }).join('');
  panel.innerHTML = `
    <div class="historical-panel">
      <div class="historical-head">
        <div>
          <div class="section-label">Historical Metrics</div>
          <div class="section-headline">Track every KPI at a glance</div>
        </div>
      </div>
      <div class="metric-summary-grid">
        ${summaryChips}
      </div>
      <div class="metric-detail-grid">
        ${detailCards}
      </div>
      ${!metricsData.length ? '<div class="metric-empty">No historical data yet. Start logging actions to build this chart.</div>' : ''}
    </div>
  `;
}

function buildEmptyMetricSeries() {
  const series = [];
  const today = new Date();
  for (let offset = 6; offset >= 0; offset--) {
    const date = new Date(today);
    date.setDate(date.getDate() - offset);
    series.push({
      Date: date.toISOString().slice(0, 10),
      leads: 0,
      actions: 0,
      advances: 0,
      customers: 0,
    });
  }
  return series;
}
function renderScriptEditor() {
  populateSelect('script-vertical', VERTICALS, 'Choose vertical', activeAssetSelection.vertical);
  populateSelect('script-product', PRODUCTS, 'Choose product', activeAssetSelection.product);
  updateScriptFunnelOptions(activeAssetSelection.product);
  const stageOptions = STAGES.slice(0, -1).map((stage, index) => ({ id: String(index + 1), label: stage.short }));
  populateSelect('script-stage', stageOptions, 'Stage', activeAssetSelection.stage);
  populateSelect('script-step', ASSET_STEPS, 'Asset step', activeAssetSelection.step);
  loadActiveScriptContent();
}

function handleScriptSelectionChange(key, value) {
  if (!value) return;
  activeAssetSelection[key] = value;
  if (key === 'product') {
    updateScriptFunnelOptions(value);
  }
  loadActiveScriptContent();
}

function getActiveAssetId() {
  const stageIdx = Math.max(0, Number(activeAssetSelection.stage) - 1);
  return buildAssetId({
    verticalId: activeAssetSelection.vertical,
    productId: activeAssetSelection.product,
    funnelId: activeAssetSelection.funnel,
    stageIdx,
    stepId: activeAssetSelection.step,
  });
}

function loadActiveScriptContent() {
  const assetId = getActiveAssetId();
  const contentInput = document.getElementById('script-content-input');
  if (contentInput) contentInput.value = assetsById[assetId] || '';
  const idDisplay = document.getElementById('script-id-display');
  if (idDisplay) idDisplay.textContent = assetId || '—';
  const channelDisplay = document.getElementById('script-channel-display');
  if (channelDisplay) channelDisplay.textContent = FUNNELS_BY_ID[activeAssetSelection.funnel]?.channel || 'Channel';
}

function fetchAssets() {
  fetch(ASSET_FETCH_URL, { method: 'GET', mode: 'cors', headers: { 'Accept': 'application/json' } })
    .then(res => { if (!res.ok) throw new Error('Assets fetch failed'); return res.json(); })
    .then(data => {
      const rows = extractRows(data);
      assetsById = rows.reduce((acc, row) => {
        const assetId = String(resolveField(row, 'Asset_ID') || resolveField(row, 'Asset Id') || resolveField(row, 'AssetId') || resolveField(row, 'asset_id') || resolveField(row, 'id')).trim();
        const assetContent = resolveField(row, 'Asset') || resolveField(row, 'Script') || '';
        if (assetId) acc[assetId] = assetContent;
        return acc;
      }, {});
      renderScriptEditor();
      renderDetail();
    })
    .catch(err => {
      console.error('fetchAssets error', err);
      showToast('Unable to load script assets', 'error');
    });
}

function fetchMetrics() {
  fetch(METRIC_FETCH_URL, { method: 'GET', mode: 'cors', headers: { 'Accept': 'application/json' } })
    .then(res => { if (!res.ok) throw new Error('Metrics fetch failed'); return res.json(); })
    .then(data => {
      const rows = extractRows(data);
      metricsData = rows.map(row => ({
        Date: resolveField(row, 'Date'),
        leads: Number(resolveField(row, 'Leads Added') || 0),
        actions: Number(resolveField(row, 'Actions Logged') || 0),
        advances: Number(resolveField(row, 'Stages Advanced') || 0),
        customers: Number(resolveField(row, 'Customers') || 0),
      })).sort((a,b) => new Date(a.Date) - new Date(b.Date));
      renderOverview();
      renderHistoricalMetrics();
    })
    .catch(err => {
      console.error('fetchMetrics error', err);
      showToast('Unable to load metrics', 'error');
    });
}

function logMetricEvent(metricKey, amount = 1) {
  const columnMap = {
    leads: 'Leads Added',
    actions: 'Actions Logged',
    advances: 'Stages Advanced',
    customers: 'Customers',
  };
  const column = columnMap[metricKey];
  if (!column) return;
  const payload = { Date: new Date().toISOString().slice(0, 10), [column]: amount };
  fetch(METRIC_UPDATE_URL, {
    method: 'POST',
    mode: 'cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
    .then(res => {
      if (!res.ok) throw new Error('Metric update failed');
      return res.json();
    })
    .then(() => fetchMetrics())
    .catch(err => {
      console.error('logMetricEvent error', err);
    });
}

function populateFilterControls() {
  populateSelect('filter-vertical', VERTICALS, 'All verticals');
  populateSelect('filter-product', PRODUCTS, 'All products');
  populateSelect('filter-funnel', FUNNELS, 'All funnels');
  const uniqueChannels = [...new Set(FUNNELS.map(f => f.channel).filter(Boolean))].map(channel => ({ id: channel, label: channel }));
  populateSelect('filter-channel', uniqueChannels, 'All channels');
}

function populateLeadFormControls() {
  populateSelect('inp-vertical', VERTICALS, 'Select vertical', VERTICALS[0]?.id);
  populateSelect('inp-product', PRODUCTS, 'Select product', PRODUCTS[0]?.id);
  updateModalFunnelOptions(PRODUCTS[0]?.id);
  updateModalChannelPreview();
}

function populateSelect(id, items, blankLabel = '', selectedValue = '', valueKey = 'id', labelKey = 'label') {
  const select = document.getElementById(id);
  if (!select) return;
  const options = items.map(item => `<option value="${item[valueKey]}"${item[valueKey] === selectedValue ? ' selected' : ''}>${item[labelKey]}</option>`).join('');
  select.innerHTML = `${blankLabel ? `<option value="">${blankLabel}</option>` : ''}${options}`;
  if (selectedValue) select.value = selectedValue;
}

function updateModalFunnelOptions(productId) {
  const funnels = FUNNELS_BY_PRODUCT[productId] || [];
  populateSelect('inp-funnel', funnels, 'Select funnel', funnels[0]?.id);
  updateModalChannelPreview();
}

function updateModalChannelPreview() {
  const funnelId = document.getElementById('inp-funnel')?.value;
  const channel = FUNNELS_BY_ID[funnelId]?.channel || '';
  const channelEl = document.getElementById('inp-channel');
  if (channelEl) channelEl.value = channel;
}

function updateScriptFunnelOptions(productId) {
  const funnels = FUNNELS_BY_PRODUCT[productId] || [];
  const select = document.getElementById('script-funnel');
  if (!select) return;
  const targetId = funnels.find(f => f.id === activeAssetSelection.funnel)?.id || funnels[0]?.id || '';
  select.innerHTML = funnels.map(f => `<option value="${f.id}">${f.label}</option>`).join('');
  select.value = targetId;
  activeAssetSelection.funnel = targetId;
}

function formatMetricLabel(value) {
  const date = new Date(value);
  if (isNaN(date)) return value;
  return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

function exportLeads() {
  if (!leads.length) {
    showToast('No leads to export', 'error');
    return;
  }
  const leadHeaders = ['Lead Name','Phone','Contact','Vertical','Product','Funnel','Channel','Stage','Stage Short','Status','Notes','Created At','Last Message At','History Count'];
  const leadRows = leads.map(lead => {
    const stageInfo = STAGES[lead.stage] || {};
    const createdAt = lead.createdAt ? new Date(lead.createdAt).toLocaleString('en-US') : '';
    const lastMessageAt = lead.lastMessageAt ? new Date(lead.lastMessageAt).toLocaleString('en-US') : '';
    const historyCount = (lead.history || []).length;
    const isComplete = lead.stage >= STAGES.length - 1 && lead.stage !== ARCHIVE_STAGE;
    const statusLabel = isComplete ? 'Complete' : 'In Progress';
    return [
      lead.name || '',
      lead.phone || '',
      lead.contact || '',
      getVerticalLabel(lead.verticalId),
      getProductLabel(lead.productId),
      getFunnelLabel(lead.funnelId),
      lead.channel || '',
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
function escapeXml(value) {
  if (value === undefined || value === null) return '';
  return String(value).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
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

function showToast(msg, type) {
  const t = document.getElementById('toast');
  if (!t) return;
  t.textContent = msg;
  const isError = type === 'error';
  t.style.borderColor = isError ? 'var(--red)' : 'var(--green)';
  t.style.color = isError ? 'var(--red)' : 'var(--green)';
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2200);
}

const modal = document.getElementById('modal');
if (modal) {
  modal.addEventListener('click', function(e) {
    if (e.target === this) closeModal();
  });
}
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });
const contactInput = document.getElementById('inp-contact');
if (contactInput) contactInput.addEventListener('keydown', e => { if (e.key === 'Enter') addLead(); });
const filterTabs = document.getElementById('filter-tabs');
if (filterTabs) {
  filterTabs.addEventListener('click', function(e) {
    if (e.target.classList.contains('filter-tab')) {
      setFilter(e.target.dataset.filter);
    }
  });
}
const productSelect = document.getElementById('inp-product');
if (productSelect) {
  productSelect.addEventListener('change', function(e) {
    const value = e.target.value;
    updateModalFunnelOptions(value);
  });
}
const funnelSelect = document.getElementById('inp-funnel');
if (funnelSelect) {
  funnelSelect.addEventListener('change', updateModalChannelPreview);
}
const scriptVertical = document.getElementById('script-vertical');
if (scriptVertical) scriptVertical.addEventListener('change', e => handleScriptSelectionChange('vertical', e.target.value));
const scriptProduct = document.getElementById('script-product');
if (scriptProduct) scriptProduct.addEventListener('change', e => handleScriptSelectionChange('product', e.target.value));
const scriptFunnel = document.getElementById('script-funnel');
if (scriptFunnel) scriptFunnel.addEventListener('change', e => handleScriptSelectionChange('funnel', e.target.value));
const scriptStage = document.getElementById('script-stage');
if (scriptStage) scriptStage.addEventListener('change', e => handleScriptSelectionChange('stage', e.target.value));
const scriptStep = document.getElementById('script-step');
if (scriptStep) scriptStep.addEventListener('change', e => handleScriptSelectionChange('step', e.target.value));

populateFilterControls();
populateLeadFormControls();
fetchAssets();
fetchMetrics();
fetchLeads();
renderScriptEditor();
