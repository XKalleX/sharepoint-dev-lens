import { getCurrentTab, sendToTab }  from './modules/messaging.js';
import { copyToClipboard, escapeHtml, formatDate, fieldTypeCss, getBaseTemplateName } from './modules/utils.js';
import { listSnippets, fieldSnippets, renderSnippets } from './modules/snippets.js';
import { toJson, toMarkdown, toCsv, download } from './modules/export.js';
import { t, toggleLang, getLang } from './modules/i18n.js';
import { icon } from './modules/icons.js';

// ── Application state ────────────────────────────────────────────────────────
const state = {
  tab: 'site',
  tabId: null,
  siteData: null,
  listDetail: null,
  selectedListId: null,
  exportFormat: 'markdown',
  listFilter: '',
  listShowHidden: false,
  fieldFilter: '',
  fieldShowHidden: false,
  loading: false,
  error: null,
};

// ── DOM refs ─────────────────────────────────────────────────────────────────
const content    = document.getElementById('tab-content');
const tabBar     = document.getElementById('tab-bar');
const headerSite = document.getElementById('header-site-name');
const spBadge    = document.getElementById('sp-badge');
const warnBadge  = document.getElementById('warn-badge');

// ── i18n + tab labels ─────────────────────────────────────────────────────────
function applyI18n() {
  document.getElementById('refresh-btn').innerHTML = icon('refresh', 15);
  document.getElementById('refresh-btn').title = t('refresh');
  document.getElementById('lang-btn').textContent = t('langToggle');

  document.getElementById('tab-site').innerHTML   = `${icon('home',14)}<span>${t('tabSite')}</span>`;
  document.getElementById('tab-lists').innerHTML  = `${icon('list',14)}<span>${t('tabLists')}</span>`;
  document.getElementById('tab-fields').innerHTML = `${icon('search',14)}<span>${t('tabFields')}</span>`;
  document.getElementById('tab-links').innerHTML  = `${icon('link',14)}<span>${t('tabLinks')}</span>`;
  document.getElementById('tab-export').innerHTML = `${icon('download',14)}<span>${t('tabExport')}</span>`;
}

// ── Init ─────────────────────────────────────────────────────────────────────
async function init() {
  applyI18n();
  state.loading = true; state.error = null;
  render();

  const tab = await getCurrentTab();
  if (!tab?.id) {
    state.loading = false;
    state.error = { msg: t('noActiveTab'), hint: null };
    render(); return;
  }
  state.tabId = tab.id;

  const ping = await sendToTab(tab.id, { type: 'PING' });
  if (!ping.success || !ping.data.isSharePoint) {
    state.loading = false;
    state.error = ping.success
      ? { msg: t('notSharePoint'), hint: null }
      : { msg: t('cannotConnect') + ' ' + (ping.error || ''), hint: t('connectionHint') };
    render(); return;
  }

  headerSite.textContent = ping.data.siteUrl.replace('https://', '');
  spBadge.classList.remove('hidden');

  const resp = await sendToTab(tab.id, { type: 'GET_SITE_DATA' });
  state.loading = false;
  if (resp.success) {
    state.siteData = resp.data;
    updateWarnBadge();
    enableTabs();
  } else {
    state.error = { msg: resp.error, hint: t('connectionHint') };
  }
  render();
}

function updateWarnBadge() {
  const n = state.siteData?.warnings?.length ?? 0;
  if (n > 0) {
    warnBadge.textContent = String(n);
    warnBadge.classList.remove('hidden');
  } else {
    warnBadge.classList.add('hidden');
  }
}

function enableTabs() {
  tabBar.querySelectorAll('.tab-btn').forEach(b => b.disabled = false);
}

document.getElementById('refresh-btn').addEventListener('click', init);
document.getElementById('lang-btn').addEventListener('click', () => {
  toggleLang();
  applyI18n();
  render();
  if (state.listDetail) bindFieldsEvents();
});

// ── Tab switching ─────────────────────────────────────────────────────────────
tabBar.addEventListener('click', e => {
  const btn = e.target.closest('.tab-btn');
  if (!btn || btn.disabled) return;
  const tab = btn.dataset.tab;
  if (tab === state.tab) return;
  state.tab = tab;
  tabBar.querySelectorAll('.tab-btn').forEach(b =>
    b.classList.toggle('active', b.dataset.tab === tab));
  if (tab === 'fields' && state.selectedListId && !state.listDetail) {
    loadListDetail(state.selectedListId);
  } else {
    render();
  }
});

// ── Data loading ──────────────────────────────────────────────────────────────
async function loadListDetail(listId) {
  state.loading = true; state.error = null; state.listDetail = null;
  render();
  const resp = await sendToTab(state.tabId, { type: 'GET_LIST_DETAIL', listId });
  state.loading = false;
  if (resp.success) { state.listDetail = resp.data; } else {
    state.error = { msg: resp.error, hint: null };
  }
  render();
}

// ── Master render ─────────────────────────────────────────────────────────────
function render() {
  if (state.loading) { content.innerHTML = renderLoading(); return; }
  if (state.error)   { content.innerHTML = renderError(state.error); bindErrorBtn(); return; }

  switch (state.tab) {
    case 'site':   content.innerHTML = renderSiteTab();   break;
    case 'lists':  content.innerHTML = renderListsTab();  bindListsEvents(); break;
    case 'fields': content.innerHTML = renderFieldsTab(); bindFieldsEvents(); break;
    case 'links':  content.innerHTML = renderLinksTab();  bindLinksEvents(); break;
    case 'export': content.innerHTML = renderExportTab(); bindExportEvents(); break;
  }
  content.addEventListener('click', handleCopyClick, { once: true });
}

// ── Copy handler (delegated) ──────────────────────────────────────────────────
function handleCopyClick(e) {
  content.addEventListener('click', handleCopyClick, { once: true });
  const btn = e.target.closest('.copy-btn');
  if (!btn) return;
  const text = btn.dataset.copy ?? '';
  copyToClipboard(text).then(ok => {
    if (!ok) return;
    btn.classList.add('copied');
    const orig = btn.innerHTML;
    btn.innerHTML = `${icon('check', 11)} ${t('copied')}`;
    setTimeout(() => { btn.classList.remove('copied'); btn.innerHTML = orig; }, 1500);
  });
}

// ── Loading / error states ────────────────────────────────────────────────────
function renderLoading(msg) {
  return `<div class="state-loading"><div class="spinner"></div><p>${msg ?? t('connecting')}</p></div>`;
}

function renderError(err) {
  const msg  = typeof err === 'string' ? err : err.msg;
  const hint = typeof err === 'string' ? null : err.hint;
  return `<div class="state-error">
    <strong>${icon('error', 14)} ${t('error')}</strong>
    ${escapeHtml(msg)}
    ${hint ? `<div class="error-hint">${icon('info', 12)} ${escapeHtml(hint)}</div>` : ''}
    <div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:10px">
      <button class="retry-btn" id="retry-btn">${icon('refresh', 12)} ${t('retry')}</button>
      <button class="reload-btn" id="reload-btn">${icon('externalLink', 12)} ${t('reloadPage')}</button>
    </div>
  </div>`;
}

function bindErrorBtn() {
  document.getElementById('retry-btn')?.addEventListener('click', init);
  document.getElementById('reload-btn')?.addEventListener('click', async () => {
    const tab = await getCurrentTab();
    if (tab?.id) chrome.tabs.reload(tab.id);
  });
}

// ── Governance helpers ────────────────────────────────────────────────────────
const SEVERITY_ORDER = { error: 0, warning: 1, info: 2 };

function sortWarnings(warnings) {
  return [...warnings].sort((a, b) => {
    const d = (SEVERITY_ORDER[a.severity] ?? 3) - (SEVERITY_ORDER[b.severity] ?? 3);
    if (d !== 0) return d;
    return (a.category ?? '').localeCompare(b.category ?? '');
  });
}

function calcHealthScore(warnings) {
  if (!warnings.length) return 100;
  const pen = { error: 20, warning: 8, info: 2 };
  return Math.max(0, 100 - warnings.reduce((s, w) => s + (pen[w.severity] ?? 0), 0));
}

function renderWarnings(warnings) {
  if (!warnings?.length) {
    return `<div class="no-warnings">${icon('check', 14)} ${t('noIssues')}</div>`;
  }
  const sorted = sortWarnings(warnings);
  const score  = calcHealthScore(sorted);
  const scoreColor = score >= 80 ? 'var(--sp-success)' : score >= 55 ? 'var(--sp-warn)' : 'var(--sp-error)';
  return `
  <div class="health-bar-wrap">
    <span class="health-label">${t('healthScore')}</span>
    <span class="health-score-num" style="color:${scoreColor}">${score}</span>
    <div class="health-bar-track"><div class="health-bar-fill" style="width:${score}%;background:${scoreColor}"></div></div>
  </div>
  <div class="warning-list">
    ${sorted.map(w => {
      const ico = w.severity === 'error' ? 'error' : w.severity === 'warning' ? 'warning' : 'info';
      const catKey = `cat${(w.category||'').charAt(0).toUpperCase()}${(w.category||'').slice(1)}`;
      const catLabel = t(catKey) || w.category;
      return `<div class="warning-item ${w.severity}">
        <div class="warning-title">
          ${icon(ico, 13)}
          <span>${escapeHtml(w.title)}</span>
          <span class="warning-cat">${escapeHtml(catLabel)}</span>
        </div>
        <div class="warning-desc">${escapeHtml(w.description)}</div>
      </div>`;
    }).join('')}
  </div>`;
}

// ── Site tab ──────────────────────────────────────────────────────────────────
function renderSiteTab() {
  if (!state.siteData) return `<div class="state-loading"><p>${t('noSiteData')}</p></div>`;
  const { web, lists, warnings, siteUrl } = state.siteData;
  const visible = lists.filter(l => !l.Hidden).length;
  const uniqueP = lists.filter(l => l.HasUniqueRoleAssignments).length;
  const large   = lists.filter(l => l.ItemCount > 5000).length;
  const hidden  = lists.filter(l => l.Hidden).length;
  const base    = `${siteUrl}/_layouts/15`;

  return `
  <div class="pad">
    <div class="section-title">${icon('home',12)} ${t('siteInfo')}</div>
    <div class="card" style="margin-bottom:10px">
      <div style="padding:8px 10px">
        <table class="info-table">
          <tr><td>${t('siteTitle')}</td><td>${escapeHtml(web.Title)}</td></tr>
          <tr><td>${t('siteUrl')}</td>
              <td class="mono">${escapeHtml(siteUrl)}
                <button class="copy-btn" data-copy="${escapeHtml(siteUrl)}" style="margin-left:4px">${icon('copy',11)} Copy</button>
              </td></tr>
          <tr><td>${t('siteWebId')}</td>
              <td class="mono">${escapeHtml(web.Id)}
                <button class="copy-btn" data-copy="${escapeHtml(web.Id)}" style="margin-left:4px">${icon('copy',11)} Copy</button>
              </td></tr>
          <tr><td>${t('siteServerPath')}</td><td class="mono">${escapeHtml(web.ServerRelativeUrl)}</td></tr>
          <tr><td>${t('siteTemplate')}</td><td>${escapeHtml(web.WebTemplate)}#${web.Configuration}</td></tr>
          <tr><td>${t('siteLanguage')}</td><td>${web.Language}</td></tr>
          <tr><td>${t('siteHubSite')}</td><td>${web.IsHubSite ? `${t('yes')} &nbsp;<span class="mono">${escapeHtml(web.HubSiteId)}</span>` : t('no')}</td></tr>
          ${web.AssociatedOwnerGroup ? `<tr><td>${t('siteOwners')}</td><td>${escapeHtml(web.AssociatedOwnerGroup.Title)}</td></tr>` : ''}
          ${web.AssociatedMemberGroup ? `<tr><td>${t('siteMembers')}</td><td>${escapeHtml(web.AssociatedMemberGroup.Title)}</td></tr>` : ''}
          <tr><td>${t('siteLastModified')}</td><td>${formatDate(web.LastItemModifiedDate)}</td></tr>
        </table>
      </div>
    </div>

    <div class="section-title">${icon('table',12)} ${t('siteQuickSummary')}</div>
    <div class="stats-grid" style="margin-bottom:10px">
      <div class="stat-card"><div class="stat-num stat-blue">${visible}</div><div class="stat-label">${t('siteVisibleLists')}</div></div>
      <div class="stat-card"><div class="stat-num ${hidden ? 'stat-warn':'stat-ok'}">${hidden}</div><div class="stat-label">${t('siteHiddenLists')}</div></div>
      <div class="stat-card"><div class="stat-num ${uniqueP ? 'stat-warn':'stat-ok'}">${uniqueP}</div><div class="stat-label">${t('siteUniquePerms')}</div></div>
      <div class="stat-card"><div class="stat-num ${large ? 'stat-error':'stat-ok'}">${large}</div><div class="stat-label">${t('siteLargeLists')}</div></div>
    </div>

    <div class="section-title">${icon('code',12)} ${t('siteRestEndpoints')}</div>
    <div class="card" style="margin-bottom:10px">
      <div style="padding:6px 10px">
        ${[
          ['/_api/web',             `${siteUrl}/_api/web`],
          ['/_api/web/lists',       `${siteUrl}/_api/web/lists`],
          ['/_api/web/siteusers',   `${siteUrl}/_api/web/siteusers`],
          ['/_api/web/roleassignments', `${siteUrl}/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings`],
        ].map(([label, href]) => `
          <div style="display:flex;align-items:center;gap:6px;padding:4px 0;border-bottom:1px solid var(--sp-border)">
            <a class="mono" href="${escapeHtml(href)}" target="_blank" rel="noreferrer" style="flex:1;font-size:11px">${escapeHtml(label)}</a>
            <button class="copy-btn" data-copy="${escapeHtml(href)}">${icon('copy',11)} Copy</button>
            <a href="${escapeHtml(href)}" target="_blank" rel="noreferrer" class="copy-btn" style="text-decoration:none">${icon('externalLink',11)}</a>
          </div>`).join('')}
      </div>
    </div>

    <div class="section-title">${icon('settings',12)} ${t('siteAdminLinks')}</div>
    <div class="link-list" style="margin-bottom:10px">
      ${[
        [t('siteSettings'),    `${base}/settings.aspx`],
        [t('siteContents'),    `${base}/viewlsts.aspx`],
        [t('sitePermissions'), `${base}/user.aspx`],
        [t('siteRecycleBin'),  `${base}/recyclebin.aspx`],
        [t('siteAppMgmt'),     `${base}/appinv.aspx`],
        [t('siteFeatures'),    `${base}/ManageFeatures.aspx`],
      ].map(([lbl, href]) =>
        `<a class="link-item" href="${escapeHtml(href)}" target="_blank" rel="noreferrer">
           <span>${escapeHtml(lbl)}</span>${icon('externalLink', 12)}
         </a>`
      ).join('')}
    </div>

    <div class="section-title">${icon('warning',12)} ${t('governanceWarnings')}
      ${warnings.length ? `<span class="badge badge-warn" style="margin-left:4px">${warnings.length}</span>` : ''}
    </div>
    <div style="margin-bottom:10px">${renderWarnings(warnings)}</div>
  </div>`;
}

// ── Lists tab ─────────────────────────────────────────────────────────────────
function renderListsTab() {
  const { lists, siteUrl } = state.siteData;
  const filtered = lists.filter(l => {
    if (!state.listShowHidden && l.Hidden) return false;
    if (state.listFilter && !l.Title.toLowerCase().includes(state.listFilter.toLowerCase())) return false;
    return true;
  });
  const label = filtered.length === 1 ? t('list') : t('lists');

  return `
  <div class="toolbar">
    <input type="search" id="list-search" placeholder="${t('filterLists')}" value="${escapeHtml(state.listFilter)}">
    <label><input type="checkbox" id="show-hidden-lists" ${state.listShowHidden ? 'checked' : ''}> ${t('showHidden')}</label>
  </div>
  <div class="count-bar"><span>${filtered.length} ${label}</span></div>
  <div class="card-list">
    ${filtered.length === 0
      ? `<p class="state-empty">${t('noListsMatch')}</p>`
      : filtered.map(l => renderListCard(l, siteUrl)).join('')}
  </div>`;
}

function renderListCard(l, siteUrl) {
  const restUrl = `${siteUrl}/_api/web/lists(guid'${l.Id}')`;
  const tplName = getBaseTemplateName(l.BaseTemplate);
  return `
  <div class="card">
    <div class="card-header">
      <div class="card-title">
        ${escapeHtml(l.Title)}
        ${l.Hidden ? `<span class="badge badge-default" style="margin-left:4px">${t('hiddenBadge')}</span>` : ''}
        ${l.HasUniqueRoleAssignments ? `<span class="badge badge-warn" style="margin-left:2px">${t('uniquePermsBadge')}</span>` : ''}
        ${l.ItemCount > 5000 ? '<span class="badge badge-error" style="margin-left:2px">&gt;5k</span>' : ''}
      </div>
      <button class="btn btn-secondary" style="font-size:11px;padding:3px 8px" data-inspect="${escapeHtml(l.Id)}">
        ${t('fieldsBtnLabel')}
      </button>
    </div>
    <div class="card-body">
      <div class="card-meta">
        <span>${t('template')}:</span><span class="val">${escapeHtml(tplName)}</span>
        <span>${t('items')}:</span><span class="val">${l.ItemCount.toLocaleString()}</span>
        <span>${t('versioning')}:</span><span class="${l.EnableVersioning ? 'ok' : 'err'}">${l.EnableVersioning ? t('on') : t('off')}</span>
        <span>${t('approval')}:</span><span class="val">${l.EnableModeration ? t('on') : t('off')}</span>
      </div>
      <div class="copy-group">
        <button class="copy-btn" data-copy="${escapeHtml(l.Id)}">${icon('copy',11)} GUID</button>
        <button class="copy-btn" data-copy="${escapeHtml(l.RootFolder.ServerRelativeUrl)}">${icon('copy',11)} Path</button>
        <button class="copy-btn" data-copy="${escapeHtml(restUrl)}">${icon('copy',11)} REST URL</button>
        <button class="copy-btn" data-copy="/_api/web/lists(guid'${escapeHtml(l.Id)}')">${icon('copy',11)} REST Path</button>
      </div>
    </div>
  </div>`;
}

function bindListsEvents() {
  document.getElementById('list-search')?.addEventListener('input', e => {
    state.listFilter = e.target.value;
    const { lists, siteUrl } = state.siteData;
    const f = lists.filter(l => {
      if (!state.listShowHidden && l.Hidden) return false;
      if (state.listFilter && !l.Title.toLowerCase().includes(state.listFilter.toLowerCase())) return false;
      return true;
    });
    const label = f.length === 1 ? t('list') : t('lists');
    document.querySelector('.count-bar span').textContent = `${f.length} ${label}`;
    document.querySelector('.card-list').innerHTML = f.length === 0
      ? `<p class="state-empty">${t('noListsMatch')}</p>`
      : f.map(l => renderListCard(l, siteUrl)).join('');
    bindInspectButtons();
  });

  document.getElementById('show-hidden-lists')?.addEventListener('change', e => {
    state.listShowHidden = e.target.checked;
    content.innerHTML = renderListsTab(); bindListsEvents();
  });

  bindInspectButtons();
}

function bindInspectButtons() {
  content.querySelectorAll('[data-inspect]').forEach(btn => {
    btn.addEventListener('click', () => {
      state.selectedListId = btn.dataset.inspect;
      state.listDetail = null;
      state.tab = 'fields';
      tabBar.querySelectorAll('.tab-btn').forEach(b =>
        b.classList.toggle('active', b.dataset.tab === 'fields'));
      loadListDetail(state.selectedListId);
    });
  });
}

// ── Fields tab ────────────────────────────────────────────────────────────────
function renderFieldsTab() {
  if (!state.selectedListId) {
    return `<div class="state-empty">
      ${icon('search', 32)}
      <p>${t('selectListFirst')}</p>
    </div>`;
  }
  if (!state.listDetail) return renderLoading(t('loadingFields'));

  const { list, fields, contentTypes, warnings } = state.listDetail;
  const visible = fields.filter(f => {
    if (!state.fieldShowHidden && f.Hidden) return false;
    if (state.fieldFilter &&
        !f.Title.toLowerCase().includes(state.fieldFilter.toLowerCase()) &&
        !f.InternalName.toLowerCase().includes(state.fieldFilter.toLowerCase())) return false;
    return true;
  });

  const snippets = renderSnippets(listSnippets(state.siteData.siteUrl, list));
  const fLabel = visible.length === 1 ? t('field') : t('fields');
  const ctVisible = contentTypes.filter(c => !c.Hidden).length;

  return `
  <div class="context-strip">
    <strong>${escapeHtml(list.Title)}</strong>
    <span class="mono">${escapeHtml(list.Id)}</span>
  </div>

  ${warnings.length ? `<div class="pad" style="padding-bottom:0">${renderWarnings(warnings)}</div>` : ''}

  <details style="border-bottom:1px solid var(--sp-border)">
    <summary style="padding:6px 10px;cursor:pointer;font-size:11px;color:var(--sp-blue);background:var(--sp-surface);user-select:none">
      ${icon('code',12)} ${t('listSnippets')} (${listSnippets(state.siteData.siteUrl, list).length})
    </summary>
    <div class="pad" style="padding-top:6px">${snippets}</div>
  </details>

  <div class="toolbar">
    <input type="search" id="field-search" placeholder="${t('filterFields')}" value="${escapeHtml(state.fieldFilter)}">
    <label><input type="checkbox" id="show-hidden-fields" ${state.fieldShowHidden ? 'checked' : ''}> ${t('showHidden')}</label>
  </div>
  <div class="count-bar">
    <span>${visible.length} ${fLabel}</span>
    <span>·</span>
    <span>${ctVisible} ${t('contentTypes')}</span>
    <button id="copy-csv-btn">${icon('copy',11)} ${t('copyCsv')}</button>
  </div>

  <div class="card-list" id="fields-list">
    ${visible.map(f => renderFieldRow(f, list)).join('')}
    ${visible.length === 0 ? `<p class="state-empty">${t('noListsMatch')}</p>` : ''}
  </div>`;
}

function renderFieldRow(f, list) {
  const cssType = fieldTypeCss(f.TypeAsString);
  const sUrl    = state.siteData?.siteUrl ?? '';
  const snips   = renderSnippets(fieldSnippets(sUrl, list, f));

  // Action buttons for write operations
  const writeOps = [];
  if (f.Hidden) {
    writeOps.push(`<button class="btn btn-action write-visible" data-write-visible="${escapeHtml(f.Id)}">${icon('eye', 12)} ${t('makeVisible')}</button>`);
  }
  if (f.ReadOnlyField) {
    writeOps.push(`<button class="btn btn-action write-editable" data-write-editable="${escapeHtml(f.Id)}">${icon('lockOpen', 12)} ${t('makeEditable')}</button>`);
  }

  return `
  <div class="field-row">
    <button class="field-row-header" data-toggle>
      <div class="field-row-info">
        <div class="field-name">${escapeHtml(f.Title)}</div>
        <div class="field-internal">${escapeHtml(f.InternalName)}</div>
        <div class="field-badges">
          <span class="badge type-${cssType}">${escapeHtml(f.TypeAsString)}</span>
          ${f.Required          ? `<span class="badge badge-error">${t('requiredBadge')}</span>` : ''}
          ${f.Indexed           ? `<span class="badge badge-ok">${t('indexedBadge')}</span>` : ''}
          ${f.Hidden            ? `<span class="badge badge-default">${icon('eyeOff',10)} ${t('hiddenBadge')}</span>` : ''}
          ${f.ReadOnlyField     ? `<span class="badge badge-default">${icon('lock',10)} ${t('readonlyBadge')}</span>` : ''}
          ${f.EnforceUniqueValues ? `<span class="badge badge-warn">${t('uniqueBadge')}</span>` : ''}
        </div>
      </div>
      <span class="chevron" style="color:var(--sp-subtle);flex-shrink:0">${icon('chevronDown', 14)}</span>
    </button>
    <div class="field-detail" style="display:none">
      <table class="info-table" style="margin-bottom:6px">
        <tr><td>${t('fieldId')}</td><td class="mono">${escapeHtml(f.Id)}</td></tr>
        <tr><td>${t('staticName')}</td><td class="mono">${escapeHtml(f.StaticName)}</td></tr>
        ${f.MaxLength  ? `<tr><td>${t('maxLength')}</td><td>${f.MaxLength}</td></tr>` : ''}
        ${f.DefaultValue ? `<tr><td>${t('defaultVal')}</td><td>${escapeHtml(f.DefaultValue)}</td></tr>` : ''}
        ${f.LookupList ? `<tr><td>${t('lookupList')}</td><td class="mono">${escapeHtml(f.LookupList)}</td></tr>` : ''}
        ${f.LookupField ? `<tr><td>${t('lookupField')}</td><td class="mono">${escapeHtml(f.LookupField)}</td></tr>` : ''}
        ${f.Choices?.results?.length ? `<tr><td>${t('choices')}</td><td>${escapeHtml(f.Choices.results.join(', '))}</td></tr>` : ''}
      </table>
      <div class="copy-group" style="margin-bottom:8px">
        <button class="copy-btn" data-copy="${escapeHtml(f.InternalName)}">${icon('copy',11)} Internal Name</button>
        <button class="copy-btn" data-copy="${escapeHtml(f.Id)}">${icon('copy',11)} Field ID</button>
        <button class="copy-btn" data-copy="${escapeHtml(f.StaticName)}">${icon('copy',11)} Static Name</button>
      </div>
      ${snips}
      ${writeOps.length ? `
      <div class="field-actions">
        <div class="field-actions-label">${icon('settings',10)} Dev-Aktionen</div>
        ${writeOps.join('')}
      </div>` : ''}
    </div>
  </div>`;
}

function bindFieldsEvents() {
  document.getElementById('field-search')?.addEventListener('input', e => {
    state.fieldFilter = e.target.value;
    const { fields, list } = state.listDetail;
    const v = fields.filter(f => {
      if (!state.fieldShowHidden && f.Hidden) return false;
      if (state.fieldFilter &&
          !f.Title.toLowerCase().includes(state.fieldFilter.toLowerCase()) &&
          !f.InternalName.toLowerCase().includes(state.fieldFilter.toLowerCase())) return false;
      return true;
    });
    const fLabel = v.length === 1 ? t('field') : t('fields');
    document.querySelector('.count-bar span').textContent = `${v.length} ${fLabel}`;
    document.getElementById('fields-list').innerHTML =
      v.map(f => renderFieldRow(f, list)).join('') ||
      `<p class="state-empty">${t('noListsMatch')}</p>`;
    bindExpandButtons();
    bindWriteButtons();
  });

  document.getElementById('show-hidden-fields')?.addEventListener('change', e => {
    state.fieldShowHidden = e.target.checked;
    content.innerHTML = renderFieldsTab(); bindFieldsEvents();
  });

  document.getElementById('copy-csv-btn')?.addEventListener('click', () => {
    const { fields } = state.listDetail;
    const csv = ['InternalName,Title,Type,Required,Indexed,Hidden,ReadOnly'].concat(
      fields.map(f =>
        `${f.InternalName},"${(f.Title||'').replace(/"/g,'""')}",${f.TypeAsString},${f.Required},${f.Indexed},${f.Hidden},${f.ReadOnlyField}`)
    ).join('\n');
    copyToClipboard(csv);
  });

  bindExpandButtons();
  bindWriteButtons();
}

function bindExpandButtons() {
  content.querySelectorAll('[data-toggle]').forEach(btn => {
    btn.addEventListener('click', () => {
      const detail  = btn.nextElementSibling;
      const chevron = btn.querySelector('.chevron');
      const open    = detail.style.display === 'none';
      detail.style.display = open ? 'block' : 'none';
      if (chevron) chevron.innerHTML = open ? icon('chevronUp', 14) : icon('chevronDown', 14);
    });
  });
}

function bindWriteButtons() {
  content.querySelectorAll('[data-write-visible]').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      showWriteModal({
        type: 'visible',
        fieldId: btn.dataset.writeVisible,
        msg:     t('writeOpMakeVisibleMsg'),
      });
    });
  });
  content.querySelectorAll('[data-write-editable]').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      showWriteModal({
        type: 'editable',
        fieldId: btn.dataset.writeEditable,
        msg:     t('writeOpMakeEditableMsg'),
      });
    });
  });
}

// ── Write operation modal ──────────────────────────────────────────────────────
let pendingWriteOp = null;

const overlay     = document.getElementById('write-overlay');
const modalHeader = document.getElementById('modal-header');
const modalMsg    = document.getElementById('modal-msg');
const modalWarn   = document.getElementById('modal-warn');
const modalCancel = document.getElementById('modal-cancel');
const modalConfirm= document.getElementById('modal-confirm');

function showWriteModal(op) {
  pendingWriteOp = op;
  modalHeader.innerHTML = `${icon('settings', 15)} ${t('writeOpConfirmTitle')}`;
  modalMsg.innerHTML    = op.msg;
  modalWarn.innerHTML   = `${icon('warning', 13)} ${t('writeOpWarning')}`;
  modalWarn.className   = 'write-modal-warn';
  modalCancel.textContent  = t('cancel');
  modalConfirm.textContent = t('confirm');
  modalConfirm.disabled    = false;
  modalCancel.disabled     = false;
  overlay.classList.add('open');
}

function hideWriteModal() {
  pendingWriteOp = null;
  overlay.classList.remove('open');
}

modalCancel.addEventListener('click', hideWriteModal);
overlay.addEventListener('click', e => { if (e.target === overlay) hideWriteModal(); });

modalConfirm.addEventListener('click', async () => {
  if (!pendingWriteOp) return;
  const op = pendingWriteOp;

  modalConfirm.disabled = true;
  modalCancel.disabled  = true;
  modalConfirm.textContent = t('updating');

  const props = {};
  if (op.type === 'visible')  props.Hidden        = false;
  if (op.type === 'editable') props.ReadOnlyField  = false;

  const resp = await sendToTab(state.tabId, {
    type:     'UPDATE_FIELD',
    listId:   state.selectedListId,
    fieldId:  op.fieldId,
    properties: props,
  });

  if (resp.success) {
    hideWriteModal();
    // Update field in state so the UI reflects the change immediately
    if (state.listDetail) {
      const field = state.listDetail.fields.find(f => f.Id === op.fieldId);
      if (field) {
        if (op.type === 'visible')  field.Hidden        = false;
        if (op.type === 'editable') field.ReadOnlyField = false;
      }
    }
    content.innerHTML = renderFieldsTab();
    bindFieldsEvents();
  } else {
    modalConfirm.disabled    = false;
    modalCancel.disabled     = false;
    modalConfirm.textContent = t('confirm');
    modalWarn.innerHTML   = `${icon('error', 13)} ${t('changeFailed')}: ${escapeHtml(resp.error || '')}`;
    modalWarn.className   = 'write-modal-warn is-error';
  }
});

// ── Links tab ──────────────────────────────────────────────────────────────────
function getLinkGroups() {
  const siteUrl = state.siteData?.siteUrl ?? '';
  return [
    { category: t('siteAdministration'), links: [
      [t('siteSettings'),    `${siteUrl}/_layouts/15/settings.aspx`],
      [t('siteContents'),    `${siteUrl}/_layouts/15/viewlsts.aspx`],
      [t('sitePermissions'), `${siteUrl}/_layouts/15/user.aspx`],
      [t('siteRecycleBin'),  `${siteUrl}/_layouts/15/recyclebin.aspx`],
      [t('siteAppMgmt'),     `${siteUrl}/_layouts/15/appinv.aspx`],
      [t('siteFeatures'),    `${siteUrl}/_layouts/15/ManageFeatures.aspx`],
    ]},
    { category: t('restApi'), links: [
      ['/_api/web',                `${siteUrl}/_api/web`],
      ['/_api/web/lists',          `${siteUrl}/_api/web/lists`],
      ['/_api/web/siteusers',      `${siteUrl}/_api/web/siteusers`],
      ['/_api/web/roleassignments',`${siteUrl}/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings`],
      ['/_api/web/features',       `${siteUrl}/_api/web/features`],
      ['/_api/contextinfo',        `${siteUrl}/_api/contextinfo`],
    ]},
    { category: t('galleries'), links: [
      [t('webPartGallery'),    `${siteUrl}/_catalogs/wp/Forms/AllItems.aspx`],
      [t('masterPageGallery'), `${siteUrl}/_catalogs/masterpage/Forms/AllItems.aspx`],
      [t('siteAssets'),        `${siteUrl}/SiteAssets/Forms/AllItems.aspx`],
    ]},
    { category: t('external'), links: [
      [t('powerAutomate'), 'https://make.powerautomate.com'],
      [t('m365Admin'),     'https://admin.microsoft.com'],
      [t('spAdmin'),       'https://admin.microsoft.com/SharePoint'],
    ]},
  ];
}

function renderLinksTab() {
  const siteUrl = state.siteData?.siteUrl ?? '';
  return `
  <div class="toolbar">
    <input type="search" id="link-search" placeholder="${t('searchLinks')}">
  </div>
  <div id="links-content" class="pad">
    <div class="section-title" style="margin-bottom:6px">${icon('externalLink',12)} ${t('openCustomPath')}</div>
    <div style="display:flex;gap:6px;margin-bottom:14px">
      <input type="text" id="custom-path" placeholder="${t('customPathPlaceholder')}"
             style="flex:1;padding:5px 8px;font-size:12px;border:1px solid var(--sp-border);border-radius:4px;font-family:monospace;background:var(--sp-bg)">
      <button class="btn btn-primary" id="open-custom-btn">${t('openBtn')}</button>
    </div>
    ${getLinkGroups().map(g => `
      <div class="section-title" style="margin-top:10px;margin-bottom:4px">${escapeHtml(g.category)}</div>
      <div class="link-list">
        ${g.links.map(([lbl, href]) =>
          `<a class="link-item" href="${escapeHtml(href)}" target="_blank" rel="noreferrer">
             <span>${escapeHtml(lbl)}</span>
             ${icon('externalLink', 12)}
           </a>`
        ).join('')}
      </div>`).join('')}
  </div>`;
}

function bindLinksEvents() {
  const siteUrl = state.siteData?.siteUrl ?? '';
  document.getElementById('open-custom-btn')?.addEventListener('click', () => {
    const path = document.getElementById('custom-path')?.value.trim();
    if (!path) return;
    const href = path.startsWith('http') ? path : `${siteUrl}${path.startsWith('/') ? '' : '/'}${path}`;
    window.open(href, '_blank', 'noreferrer');
  });
  document.getElementById('custom-path')?.addEventListener('keydown', e => {
    if (e.key === 'Enter') document.getElementById('open-custom-btn')?.click();
  });
  document.getElementById('link-search')?.addEventListener('input', e => {
    const q = e.target.value.toLowerCase();
    content.querySelectorAll('.link-item').forEach(a => {
      a.style.display = (!q || a.textContent.toLowerCase().includes(q) || a.href?.toLowerCase().includes(q)) ? '' : 'none';
    });
  });
}

// ── Export tab ────────────────────────────────────────────────────────────────
function renderExportTab() {
  if (!state.siteData) return `<div class="state-empty">${t('noDataToExport')}</div>`;
  const fmt = state.exportFormat;
  const content_text = fmt === 'json' ? toJson(state.siteData)
    : fmt === 'markdown' ? toMarkdown(state.siteData)
    : toCsv(state.siteData);
  const preview  = content_text.slice(0, 10000) + (content_text.length > 10000 ? '\n\n… (truncated)' : '');
  const siteName = state.siteData.web.Title.replace(/[^a-zA-Z0-9_-]/g, '_');
  const date     = new Date().toISOString().slice(0, 10);
  const ext      = fmt === 'json' ? 'json' : fmt === 'markdown' ? 'md' : 'csv';
  const mime     = fmt === 'json' ? 'application/json' : fmt === 'markdown' ? 'text/markdown' : 'text/csv';
  const filename = `sp-dev-lens_${siteName}_${date}.${ext}`;

  return `
  <div class="pad" style="border-bottom:1px solid var(--sp-border)">
    <div class="section-title" style="margin-bottom:6px">${t('format')}</div>
    <div class="format-selector">
      ${['json','markdown','csv'].map(f =>
        `<button class="format-btn ${f === fmt ? 'active' : ''}" data-format="${f}">${f.toUpperCase()}</button>`
      ).join('')}
    </div>
  </div>
  <div style="display:flex;gap:6px;padding:8px 10px;border-bottom:1px solid var(--sp-border);background:var(--sp-bg)">
    <button class="btn btn-primary btn-full" id="download-btn">${icon('download',12)} ${t('downloadBtn')} ${ext.toUpperCase()}</button>
    <button class="btn btn-secondary" id="copy-export-btn">${icon('copy',12)} ${t('copy')}</button>
  </div>
  <div style="padding:4px 10px 2px;border-bottom:1px solid var(--sp-border);font-size:11px;color:var(--sp-subtle)">
    ${escapeHtml(state.siteData.web.Title)} · ${state.siteData.lists.length} ${t('lists')}
    · ${state.siteData.warnings.length} ${t('warnings')}
    · ${formatDate(state.siteData.fetchedAt)}
  </div>
  <div class="count-bar"><span>${t('preview')} (${content_text.length.toLocaleString()} ${t('chars')})</span></div>
  <pre class="export-preview">${escapeHtml(preview)}</pre>`;
}

function bindExportEvents() {
  content.querySelectorAll('.format-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      state.exportFormat = btn.dataset.format;
      content.innerHTML = renderExportTab(); bindExportEvents();
    });
  });
  document.getElementById('download-btn')?.addEventListener('click', () => {
    const fmt  = state.exportFormat;
    const text = fmt === 'json' ? toJson(state.siteData)
      : fmt === 'markdown' ? toMarkdown(state.siteData) : toCsv(state.siteData);
    const site = state.siteData.web.Title.replace(/[^a-zA-Z0-9_-]/g, '_');
    const date = new Date().toISOString().slice(0, 10);
    const ext  = fmt === 'json' ? 'json' : fmt === 'markdown' ? 'md' : 'csv';
    const mime = fmt === 'json' ? 'application/json' : fmt === 'markdown' ? 'text/markdown' : 'text/csv';
    download(text, `sp-dev-lens_${site}_${date}.${ext}`, mime);
  });
  document.getElementById('copy-export-btn')?.addEventListener('click', async () => {
    const fmt  = state.exportFormat;
    const text = fmt === 'json' ? toJson(state.siteData)
      : fmt === 'markdown' ? toMarkdown(state.siteData) : toCsv(state.siteData);
    const ok = await copyToClipboard(text);
    if (ok) {
      const btn = document.getElementById('copy-export-btn');
      btn.innerHTML = `${icon('check',12)} ${t('copied')}`;
      btn.classList.add('copied');
      setTimeout(() => {
        btn.innerHTML = `${icon('copy',12)} ${t('copy')}`;
        btn.classList.remove('copied');
      }, 2000);
    }
  });
}

// ── Boot ──────────────────────────────────────────────────────────────────────
init();
