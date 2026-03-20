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
  search: '',
};
const FOLLOWUP_THRESHOLD = 36 * 60 * 60 * 1000;
const NEW_LEAD_THRESHOLD = 30 * 60 * 60 * 1000;
const ACTION_DISPLAY_LIMIT = 3;
const NEXT_ACTION_LIMIT = 8;
const MAX_FOCUS_ITEMS = 6;
const COMMAND_FILTER = { vertical: 'all', stage: 'all', search: '' };
const PAGE_MAP = {
  war: 'war-room',
  command: 'command-panel',
  script: 'script-area',
};
let todayMode = false;
let currentCommandLeadId = null;
let activePageKey = 'war';

const LEAD_SYNC_WEBHOOK = 'https://automation.tharrosai.com/webhook/bc225c06-6ba3-49e7-9169-d5c49729591b';
const FETCH_LEADS_WEBHOOK = 'https://automation.tharrosai.com/webhook/127e7f4a-dd40-4f7c-b995-05efb0977b9a';
const DELETE_WEBHOOK_URL = 'https://tharros.app.n8n.cloud/webhook/1456e7f1-fbb4-4d32-975f-0b5a5f8132cd';
const ASSET_FETCH_URL = 'https://automation.tharrosai.com/webhook/907069ea-0044-4bda-b0c3-2529dd377014';
const ASSET_UPDATE_URL = 'https://automation.tharrosai.com/webhook/a90b8352-a749-499a-b711-742fc7e62814';
let leads = [];
let currentLeadId = null;
let assetsById = {};
let activeAssetSelection = {
  vertical: VERTICALS[0]?.id || '1',
  product: PRODUCTS[0]?.id || '1',
  funnel: FUNNELS_BY_PRODUCT[PRODUCTS[0]?.id]?.[0]?.id || FUNNELS[0]?.id || '1',
  stage: '1',
  step: '1',
};
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
  if (!preserveSelection) {
    leads = [];
    currentLeadId = null;
  }
  fetch(FETCH_LEADS_WEBHOOK, { method: 'GET', mode: 'cors', headers: { 'Accept': 'application/json' } })
    .then(res => {
      if (!res.ok) throw new Error('Network response was not ok');
      return res.json();
    })
    .then(data => {
      const rows = extractRows(data);
      leads = rows.map(normalizeLeadRow);
      mergePendingLeads();
      renderWarRoom();
      if (leads.length) {
        const focused = resolvePendingLead();
        if (!focused && !preserveSelection) {
          selectLead(leads[0].id);
        }
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


function getVisibleLeads() {
  const search = (FILTER_STATE.search || '').toLowerCase();
  return leads.filter(lead => {
    if (lead.stage === ARCHIVE_STAGE) return false;
    if (!lead.name) return false;
    if (!search) return true;
    const haystack = `${lead.name} ${lead.contact || ''} ${lead.phone || ''}`.toLowerCase();
    return haystack.includes(search);
  });
}

function updateHeaderStats() {
  const total = leads.length;
  const needAction = leads.filter(l => isActionDue(l) && l.stage < ARCHIVE_STAGE).length;
  const totalEl = document.getElementById('hdr-total');
  const actionEl = document.getElementById('hdr-action');
  if (totalEl) totalEl.textContent = total;
  if (actionEl) actionEl.textContent = needAction;
}

function renderWarRoom() {
  updateHeaderStats();
  renderFocusQueue();
  const prioritized = getPrioritizedLeads();
  renderActionEngine(prioritized);
  renderPipelinePulse();
  renderNextActionsList(prioritized);
  renderCommandSection(prioritized);
}

function renderFocusQueue() {
  const { needs, followUps, newLeads } = getFocusSections();
  const needsContainer = document.getElementById('focus-needs');
  const followContainer = document.getElementById('focus-follow');
  const newContainer = document.getElementById('focus-new');
  const needsCount = document.getElementById('needs-count');
  const followCount = document.getElementById('follow-count');
  const newCount = document.getElementById('new-count');
  if (needsCount) needsCount.textContent = needs.length;
  if (followCount) followCount.textContent = followUps.length;
  if (newCount) newCount.textContent = newLeads.length;
  if (needsContainer) needsContainer.innerHTML = buildFocusList(needs);
  if (followContainer) followContainer.innerHTML = buildFocusList(followUps);
  if (newContainer) newContainer.innerHTML = buildFocusList(newLeads);
}

function getFocusSections() {
  const needs = [];
  const followUps = [];
  const newLeads = [];
  const nowTs = now();
  const candidates = getVisibleLeads();
  candidates.forEach(lead => {
    if (isActionDue(lead)) {
      needs.push(lead);
      return;
    }
    if (lead.lastMessageAt && nowTs - lead.lastMessageAt > FOLLOWUP_THRESHOLD) {
      followUps.push(lead);
      return;
    }
    if (lead.createdAt && nowTs - lead.createdAt < NEW_LEAD_THRESHOLD) {
      newLeads.push(lead);
    }
  });
  return {
    needs: needs.slice(0, MAX_FOCUS_ITEMS),
    followUps: followUps.slice(0, MAX_FOCUS_ITEMS),
    newLeads: newLeads.slice(0, MAX_FOCUS_ITEMS),
  };
}

function buildFocusList(items) {
  if (!items.length) {
    return '<div class="focus-empty">No leads here</div>';
  }
  return items.map(lead => {
    const isActive = currentLeadId === lead.id;
    const lastTouch = fmtTime(lead.lastMessageAt || lead.createdAt);
    const stageLabel = STAGES[lead.stage]?.short || '';
    return `
      <button type="button" class="focus-item ${isActive ? 'active' : ''}" onclick="selectLead('${lead.id}')">
        <div class="focus-item-title">${escHTML(lead.name)}</div>
        <div class="focus-meta">
          <span>${stageLabel} · ${lastTouch}</span>
          <span class="focus-dot"></span>
        </div>
      </button>
    `;
  }).join('');
}

function getPrioritizedLeads() {
  const filtered = getVisibleLeads();
  const sorted = [...filtered].sort((a, b) => {
    const aUrgent = isActionDue(a);
    const bUrgent = isActionDue(b);
    if (aUrgent && !bUrgent) return -1;
    if (!aUrgent && bUrgent) return 1;
    const aTime = (a.lastMessageAt || a.createdAt) || 0;
    const bTime = (b.lastMessageAt || b.createdAt) || 0;
    return bTime - aTime;
  });
  if (todayMode) {
    const urgentOnly = sorted.filter(isActionDue);
    return urgentOnly.length ? urgentOnly : sorted;
  }
  return sorted;
}

function ensureCurrentLead(prioritized) {
  if (!prioritized.length) {
    currentLeadId = null;
    return;
  }
  const exists = leads.some(lead => lead.id === currentLeadId && lead.stage < ARCHIVE_STAGE);
  if (!currentLeadId || !exists) {
    currentLeadId = prioritized[0].id;
  }
}

function renderActionEngine(prioritized) {
  const container = document.getElementById('action-cards');
  if (!container) return;
  ensureCurrentLead(prioritized);
  if (!prioritized.length) {
    container.innerHTML = '<div class="action-card empty">No prioritized leads right now.</div>';
    return;
  }
  const cards = prioritized.slice(0, ACTION_DISPLAY_LIMIT).map(lead => buildActionCard(lead)).join('');
  container.innerHTML = cards;
}

function buildActionCard(lead) {
  const next = getNextAction(lead);
  const preview = next?.definition?.content ? next.definition.content.slice(0, 220) : 'Script not available for this combination.';
  const actionLabel = next?.definition ? `${next.definition.stepLabel} · ${next.definition.channel || 'Channel'}` : 'Waiting on response';
  const disabled = next && next.assetId ? '' : 'disabled';
  const stageName = STAGES[lead.stage]?.name || 'Stage';
  const cardClass = lead.id === currentLeadId ? 'active' : '';
  const markDoneButton = lead.stage < ARCHIVE_STAGE - 1 ? `<button class="btn btn-ghost" onclick="event.stopPropagation(); moveToStage(${Math.min(lead.stage + 1, ARCHIVE_STAGE - 1)})">Mark Done</button>` : '';
  return `
    <div class="action-card ${cardClass}" onclick="selectLead('${lead.id}')">
      <div class="action-card-header">
        <div>
          <div class="action-card-title">${escHTML(lead.name)}</div>
          <div class="action-card-stage">${stageName}</div>
        </div>
        <span class="focus-stage-tag">${STAGES[lead.stage]?.short || ''}</span>
      </div>
      <div class="action-card-meta">${actionLabel} · ${fmtTime(lead.lastMessageAt || lead.createdAt)}</div>
      <div class="action-script">${escHTML(preview)}</div>
      <div class="action-buttons">
        <button class="btn btn-primary" ${disabled} ${next && next.assetId ? `onclick="event.stopPropagation(); copyAndMark('${next.assetId}','${lead.id}')"` : ''}>Copy + Send</button>
        <button class="btn btn-ghost" onclick="event.stopPropagation(); skipWarLead('${lead.id}')">Skip</button>
        ${markDoneButton}
      </div>
    </div>
  `;
}

function buildNextStageButtons(lead) {
  const upcoming = STAGES.slice(lead.stage + 1, STAGES.length - 1);
  if (!upcoming.length) {
    return '<div class="stage-empty">No further stages</div>';
  }
  return upcoming.map((stage, idx) => {
    const stageIndex = lead.stage + 1 + idx;
    return `<button class="btn btn-ghost stage-btn" onclick="moveToStage(${stageIndex})">${escHTML(stage.short)}</button>`;
  }).join('');
}

function escapeTextarea(value = '') {
  return String(value || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function renderNextActionsList(prioritized = []) {
  const container = document.getElementById('next-actions');
  if (!container) return;
  if (!prioritized.length) {
    container.innerHTML = '<div class="next-actions-title">Next Actions</div><div class="next-actions-list"><div class="next-action-item">No next steps yet.</div></div>';
    return;
  }
  const list = prioritized.slice(0, NEXT_ACTION_LIMIT).map(lead => {
    const next = getNextAction(lead);
    const label = next?.definition ? `${next.definition.stepLabel} · ${next.definition.stageLabel}` : 'Waiting for reply';
    return `
      <div class="next-action-item">
        <span>${escHTML(lead.name)}</span>
        <span>${escHTML(label)}</span>
      </div>
    `;
  }).join('');
  container.innerHTML = `<div class="next-actions-title">Next Actions (${Math.min(prioritized.length, NEXT_ACTION_LIMIT)})</div><div class="next-actions-list">${list}</div>`;
}

function renderCommandSection(prioritized = []) {
  renderCommandFilters();
  renderCommandList(prioritized);
  renderCommandDetail();
}

function renderCommandFilters() {
  const verticalSelect = document.getElementById('command-vertical');
  const stageSelect = document.getElementById('command-stage');
  if (verticalSelect) {
    const options = ['<option value="all">All types</option>']
      .concat(VERTICALS.map(v => `<option value="${v.id}">${v.label}</option>`));
    verticalSelect.innerHTML = options.join('');
    verticalSelect.value = COMMAND_FILTER.vertical;
  }
  if (stageSelect) {
    const stageOptions = ['<option value="all">All stages</option>']
      .concat(STAGES.slice(0, -1).map((stage, idx) => `<option value="${idx}">${stage.short}</option>`));
    stageSelect.innerHTML = stageOptions.join('');
    stageSelect.value = COMMAND_FILTER.stage;
  }
  const searchInput = document.getElementById('command-search');
  if (searchInput) searchInput.value = COMMAND_FILTER.search;
}

function renderCommandList(prioritized) {
  const container = document.getElementById('command-lead-list');
  if (!container) return;
  const items = getCommandLeads(prioritized);
  if (!items.length) {
    container.innerHTML = '<div class="command-empty">No leads match the filters.</div>';
    return;
  }
  container.innerHTML = items.map(lead => {
    const stageLabel = STAGES[lead.stage]?.short || '';
    const lastTouch = fmtTime(lead.lastMessageAt || lead.createdAt);
    const isUrgent = isActionDue(lead);
    const active = currentCommandLeadId === lead.id;
    return `
      <button class="command-lead-item ${active ? 'active' : ''}" onclick="openCommandLead('${lead.id}')">
        <div class="command-lead-name">${escHTML(lead.name)}</div>
        <div class="command-lead-meta">
          <span class="command-lead-stage">${stageLabel}</span>
          <span>${escHTML(lead.verticalLabel)}</span>
          <span>${escHTML(lastTouch)}</span>
          <span class="command-lead-dot ${isUrgent ? 'urgent' : ''}"></span>
        </div>
      </button>
    `;
  }).join('');
}

function getCommandLeads(prioritized = []) {
  const search = (COMMAND_FILTER.search || '').toLowerCase();
  return (prioritized.length ? prioritized : leads)
    .filter(lead => lead.stage < ARCHIVE_STAGE)
    .filter(lead => {
      if (COMMAND_FILTER.vertical !== 'all' && lead.verticalId !== COMMAND_FILTER.vertical) return false;
      if (COMMAND_FILTER.stage !== 'all' && lead.stage !== Number(COMMAND_FILTER.stage)) return false;
      if (!search) return true;
      const haystack = `${lead.name} ${lead.contact || ''} ${lead.phone || ''}`.toLowerCase();
      return haystack.includes(search);
    });
}

function openCommandLead(leadId) {
  if (!leadId) return;
  currentCommandLeadId = leadId;
  selectLead(leadId);
}

function handleCommandSearch(event) {
  COMMAND_FILTER.search = (event.target.value || '').trim();
  renderCommandList();
}

function resetCommandFilters() {
  COMMAND_FILTER.vertical = 'all';
  COMMAND_FILTER.stage = 'all';
  COMMAND_FILTER.search = '';
  const searchInput = document.getElementById('command-search');
  if (searchInput) searchInput.value = '';
  renderCommandSection();
}

function renderCommandDetail() {
  const container = document.getElementById('command-detail-view');
  if (!container) return;
  const leadId = currentCommandLeadId || currentLeadId;
  const lead = leads.find(l => l.id === leadId && l.stage < ARCHIVE_STAGE);
  if (!lead) {
    container.innerHTML = '<div class="command-detail-empty">Select a lead to review the full tile.</div>';
    return;
  }
  const stageCount = STAGES.length - 1;
  const clampedStage = Math.min(lead.stage, stageCount - 1);
  const progressPercent = stageCount > 1 ? Math.round((clampedStage / (stageCount - 1)) * 100) : 0;
  const pipelineSteps = buildPipelineSteps(lead);
  const next = getNextAction(lead);
  const previewText = next?.definition?.content || 'Script not available for this stage yet.';
  const messageSequence = renderCommandMessages(lead);
  container.innerHTML = `
    <div class="command-detail-frame">
      <div class="command-detail-header">
        <div>
          <div class="detail-label">Full Lead Tile</div>
          <div class="detail-name">${escHTML(lead.name)}</div>
          <div class="detail-meta">${escHTML(lead.phone || 'No phone')} · ${escHTML(lead.contact || 'Unknown')}</div>
        </div>
        <div class="detail-stage-tag">${STAGES[lead.stage]?.short || 'Stage'}</div>
      </div>
      <div class="detail-pipeline">
        <div class="pipeline-line">
          <div class="pipeline-progress" style="width:${progressPercent}%"></div>
        </div>
        <div class="pipeline-steps">
          ${pipelineSteps}
        </div>
      </div>
      <div class="detail-next">
        <div class="detail-next-text">
          <div class="detail-next-label">${next?.definition ? `${escHTML(next.definition.stepLabel)} · ${escHTML(next.definition.channel || '')}` : 'Waiting on next action'}</div>
          <div class="detail-next-meta">${escHTML(fmtAbsTime(lead.lastMessageAt || lead.createdAt))}</div>
          <div class="detail-next-preview">${escHTML(previewText)}</div>
        </div>
        <div class="detail-next-actions">
          <button class="btn btn-primary" ${next?.assetId ? `onclick="copyAndMark('${next.assetId}','${lead.id}')"` : 'disabled'}>Copy + Send</button>
          <button class="btn btn-ghost" onclick="skipWarLead('${lead.id}')">Skip</button>
        </div>
      </div>
      <div class="detail-stage-actions">
        ${buildNextStageButtons(lead)}
      </div>
      <div class="command-messages">
        ${messageSequence}
      </div>
    </div>
  `;
}

function buildPipelineSteps(lead) {
  return STAGES.slice(0, -1).map((stage, idx) => {
    const passed = idx < lead.stage;
    const current = idx === lead.stage;
    const stateClasses = ['pipeline-step'];
    if (passed) stateClasses.push('done');
    if (current) stateClasses.push('current');
    const isInteractive = idx > lead.stage;
    return `
      <button type="button" class="${stateClasses.join(' ')}" ${isInteractive ? `onclick="moveToStage(${idx})"` : ''} ${idx <= lead.stage ? 'disabled' : ''}>
        <span class="pipeline-step-circle">${passed ? '✓' : ''}</span>
        <span class="pipeline-step-label">${escHTML(stage.short)}</span>
      </button>
    `;
  }).join('');
}

function renderCommandMessages(lead) {
  const assets = getStageAssets(lead);
  if (!assets.length) {
    return '<div class="command-messages-empty">No follow-up scripts available yet.</div>';
  }
  const sentSet = new Set(lead.sentMessages || []);
  return assets.map(assetId => {
    const definition = getAssetDefinition(assetId, lead);
    const content = definition.content || 'Script placeholder';
    const sent = sentSet.has(assetId);
    const label = definition.stepLabel || 'Message';
    const channel = definition.channel || 'Channel';
    return `
      <article class="command-message ${sent ? 'sent' : ''}">
        <div class="command-message-header">
          <span class="command-message-label">${escHTML(label)}</span>
          <span class="command-message-channel">${escHTML(channel)}</span>
          <span class="command-message-status">${sent ? 'Sent' : 'Pending'}</span>
        </div>
        <p class="command-message-body">${escHTML(content)}</p>
        <div class="command-message-actions">
          <button class="btn btn-primary" ${sent ? 'disabled' : ''} onclick="copyAndMark('${assetId}','${lead.id}')">Copy + Send</button>
          ${sent ? `<span class="command-message-sent">already sent</span>` : ''}
        </div>
      </article>
    `;
  }).join('');
}

function renderPipelinePulse() {
  const bottleneckEl = document.getElementById('pulse-bottleneck');
  const hotEl = document.getElementById('pulse-hot');
  const dropoffEl = document.getElementById('pulse-dropoff');
  const visible = getVisibleLeads();
  const total = visible.length || 1;
  const stageCounts = STAGES.slice(0, -1).map((stage, index) => visible.filter(lead => lead.stage === index).length);
  const maxIdx = stageCounts.reduce((acc, value, idx) => value > stageCounts[acc] ? idx : acc, 0);
  const bottleneckText = stageCounts[maxIdx] ? `${stageCounts[maxIdx]} leads stuck in ${STAGES[maxIdx].short}` : 'Pipeline is flowing';
  const hotDeals = visible.filter(lead => lead.stage >= STAGES.length - 2 && lead.stage < ARCHIVE_STAGE).slice(0, 3);
  const hotText = hotDeals.length ? hotDeals.map(lead => `${lead.name} · ${STAGES[lead.stage]?.short || ''}`).join('\n') : 'No hot deals at the moment';
  const dropList = stageCounts
    .map((count, idx) => ({ label: STAGES[idx]?.short || '', percent: Math.round((count / total) * 100) }))
    .filter(entry => entry.percent > 0)
    .slice(-3)
    .map(entry => `${entry.label}: ${entry.percent}%`)
    .join('\n') || 'No significant drop-offs';
  if (bottleneckEl) bottleneckEl.textContent = bottleneckText;
  if (hotEl) hotEl.textContent = hotText;
  if (dropoffEl) dropoffEl.textContent = dropList;
}

function skipWarLead(leadId) {
  const prioritized = getPrioritizedLeads();
  if (!prioritized.length) return;
  const idx = prioritized.findIndex(lead => lead.id === leadId);
  const nextLead = prioritized[(idx + 1) % prioritized.length];
  if (nextLead) {
    selectLead(nextLead.id);
  }
}

function toggleTodayMode() {
  todayMode = !todayMode;
  const toggle = document.getElementById('today-toggle');
  if (toggle) toggle.classList.toggle('active', todayMode);
  renderWarRoom();
}

function selectLead(id) {
  if (!id) return;
  currentLeadId = id;
  renderWarRoom();
}
function escHTML(str = '') {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>');
}

function markSent(leadId, assetId, options = {}) {
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
    .then(() => {
      renderWarRoom();
      if (!options.silent) {
        showToast(options.toastMessage || 'Marked as sent');
      }
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
      renderWarRoom();
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

function copyAndMark(assetId, leadId) {
  const lead = leads.find(l => l.id === leadId);
  if (!lead || !assetId) return;
  const asset = getAssetDefinition(assetId, lead);
  if (!asset || !asset.content) {
    showToast('Script not available', 'error');
    return;
  }
  navigator.clipboard.writeText(asset.content)
    .then(() => {
      markSent(leadId, assetId, { toastMessage: 'Copied + marked as sent' });
    })
    .catch(err => {
      console.error('copyAndMark error', err);
      showToast('Unable to copy script', 'error');
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
  renderWarRoom();
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
  renderWarRoom();
}

function openModal() {
  document.getElementById('modal').classList.add('open');
  document.body.classList.add('modal-open');
  setTimeout(() => document.getElementById('inp-name').focus(), 50);
}

function closeModal() {
  document.getElementById('modal').classList.remove('open');
  document.body.classList.remove('modal-open');
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
      fetchLeads({ preserveSelection: true });
    })
    .catch(err => {
      console.error('addLead error', err);
      showToast('Could not add lead', 'error');
      pendingLeadSelection = null;
    });
}

function showCommandPage() {
  switchPage('command');
}

function showScriptLibrary() {
  switchPage('script');
}

function showWarRoom() {
  switchPage('war');
}

function switchPage(pageKey) {
  if (!PAGE_MAP[pageKey]) return;
  Object.entries(PAGE_MAP).forEach(([key, sectionId]) => {
    const section = document.getElementById(sectionId);
    if (!section) return;
    section.classList.toggle('hidden', key !== pageKey);
  });
  activePageKey = pageKey;
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.page === pageKey);
  });
  if (pageKey === 'command') {
    currentCommandLeadId = currentCommandLeadId || currentLeadId;
  }
  if (pageKey === 'script') {
    renderScriptEditor();
  } else {
    renderWarRoom();
  }
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
      renderWarRoom();
    })
    .catch(err => {
      console.error('fetchAssets error', err);
      showToast('Unable to load script assets', 'error');
    });
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

const commandVerticalSelect = document.getElementById('command-vertical');
if (commandVerticalSelect) commandVerticalSelect.addEventListener('change', e => {
  COMMAND_FILTER.vertical = e.target.value;
  renderCommandList();
  renderCommandDetail();
});
const commandStageSelect = document.getElementById('command-stage');
if (commandStageSelect) commandStageSelect.addEventListener('change', e => {
  COMMAND_FILTER.stage = e.target.value;
  renderCommandList();
  renderCommandDetail();
});
const commandSearch = document.getElementById('command-search');
if (commandSearch) commandSearch.addEventListener('input', handleCommandSearch);

populateLeadFormControls();
fetchAssets();
fetchLeads();
renderScriptEditor();
showWarRoom();
