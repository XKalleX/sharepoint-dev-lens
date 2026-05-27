import { getCurrentTab, sendToTab }  from './modules/messaging.js';
import { copyToClipboard, escapeHtml, formatDate, fieldTypeCss, getBaseTemplateName } from './modules/utils.js';
import { listSnippets, fieldSnippets, renderSnippets } from './modules/snippets.js';
import { toJson, toMarkdown, toCsv, download } from './modules/export.js';

// ── Application state ────────────────────────────────────────────────────────
const state = {
  tab: 'site',
  tabId: null,
  siteData: null,      // { web, lists, warnings, fetchedAt, siteUrl }
  listDetail: null,    // { list, fields, contentTypes, warnings }
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

// ── Init ─────────────────────────────────────────────────────────────────────
async function init() {
  state.loading = true; state.error = null;
  render();

  const tab = await getCurrentTab();
  if (!tab?.id) { state.loading = false; state.error = 'No active tab found.'; render(); return; }
  state.tabId = tab.id;

  const ping = await sendToTab(tab.id, { type: 'PING' });
  if (!ping.success || !ping.data.isSharePoint) {
    state.loading = false;
    state.error = ping.success
      ? 'This page is not a SharePoint site.'
      : `Cannot connect to page. ${ping.error}`;
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
    state.error = resp.error;
  }
  render();
}

function updateWarnBadge() {
  const n = state.siteData?.warnings?.length ?? 0;
  if (n > 0) {
    warnBadge.textContent = `${n} ⚠`;
    warnBadge.classList.remove('hidden');
  } else {
    warnBadge.classList.add('hidden');
  }
}

function enableTabs() {
  tabBar.querySelectorAll('.tab-btn').forEach(b => b.disabled = false);
}

document.getElementById('refresh-btn').addEventListener('click', init);

// ── Tab switching ─────────────────────────────────────────────────────────────
tabBar.addEventListener('click', e => {
  const btn = e.target.closest('.tab-btn');
  if (!btn || btn.disabled) return;
  const tab = btn.dataset.tab;
  if (tab === state.tab) return;
  state.tab = tab;
  tabBar.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));

  // Load field detail if switching to fields and a list is selected
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
  if (resp.success) { state.listDetail = resp.data; } else { state.error = resp.error; }
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

  // Global copy-button handler (event delegation on content area)
  content.addEventListener('click', handleCopyClick, { once: true });
}

// ── Copy handler (delegated) ──────────────────────────────────────────────────
function handleCopyClick(e) {
  // Re-attach for next render cycle
  content.addEventListener('click', handleCopyClick, { once: true });

  const btn = e.target.closest('.copy-btn');
  if (!btn) return;
  const text = btn.dataset.copy ?? '';
  copyToClipboard(text).then(ok => {
    if (!ok) return;
    btn.classList.add('copied');
    const orig = btn.innerHTML;
    btn.innerHTML = '✓ Copied';
    setTimeout(() => { btn.classList.remove('copied'); btn.innerHTML = orig; }, 1500);
  });
}

// ── Loading / error states ────────────────────────────────────────────────────
function renderLoading(msg = 'Connecting to SharePoint…') {
  return `<div class="state-loading"><div class="spinner"></div><p>${msg}</p></div>`;
}

function renderError(msg) {
  return `<div class="state-error">
    <strong>Error</strong>${escapeHtml(msg)}
    <br><button class="retry-btn" id="retry-btn">Retry</button>
  </div>`;
}

function bindErrorBtn() {
  document.getElementById('retry-btn')?.addEventListener('click', init);
}

// ── Site tab ──────────────────────────────────────────────────────────────────
function renderSiteTab() {
  if (!state.siteData) return `<div class="state-loading"><p>No site data.</p></div>`;
  const { web, lists, warnings, siteUrl } = state.siteData;
  const visible = lists.filter(l => !l.Hidden).length;
  const uniqueP  = lists.filter(l => l.HasUniqueRoleAssignments).length;
  const large    = lists.filter(l => l.ItemCount > 5000).length;
  const hidden   = lists.filter(l => l.Hidden).length;
  const base     = `${siteUrl}/_layouts/15`;

  return `
  <div class="pad">
    <div class="section-title">Site Information</div>
    <div class="card" style="margin-bottom:10px">
      <div style="padding:8px 10px">
        <table class="info-table">
          <tr><td>Title</td><td>${escapeHtml(web.Title)}</td></tr>
          <tr><td>URL</td>
              <td class="mono">${escapeHtml(siteUrl)}
                <button class="copy-btn" data-copy="${escapeHtml(siteUrl)}" style="margin-left:4px">⎘</button>
              </td></tr>
          <tr><td>Web ID</td>
              <td class="mono">${escapeHtml(web.Id)}
                <button class="copy-btn" data-copy="${escapeHtml(web.Id)}" style="margin-left:4px">⎘</button>
              </td></tr>
          <tr><td>Server Path</td><td class="mono">${escapeHtml(web.ServerRelativeUrl)}</td></tr>
          <tr><td>Template</td><td>${escapeHtml(web.WebTemplate)}#${web.Configuration}</td></tr>
          <tr><td>Language</td><td>${web.Language}</td></tr>
          <tr><td>Hub Site</td><td>${web.IsHubSite ? `Yes &nbsp;<span class="mono">${escapeHtml(web.HubSiteId)}</span>` : 'No'}</td></tr>
          ${web.AssociatedOwnerGroup ? `<tr><td>Owners</td><td>${escapeHtml(web.AssociatedOwnerGroup.Title)}</td></tr>` : ''}
          ${web.AssociatedMemberGroup ? `<tr><td>Members</td><td>${escapeHtml(web.AssociatedMemberGroup.Title)}</td></tr>` : ''}
          <tr><td>Last modified</td><td>${formatDate(web.LastItemModifiedDate)}</td></tr>
        </table>
      </div>
    </div>

    <div class="section-title">Quick Summary</div>
    <div class="stats-grid" style="margin-bottom:10px">
      <div class="stat-card"><div class="stat-num stat-blue">${visible}</div><div class="stat-label">Visible Lists</div></div>
      <div class="stat-card"><div class="stat-num ${hidden ? 'stat-warn':'stat-ok'}">${hidden}</div><div class="stat-label">Hidden Lists</div></div>
      <div class="stat-card"><div class="stat-num ${uniqueP ? 'stat-warn':'stat-ok'}">${uniqueP}</div><div class="stat-label">Unique Perms</div></div>
      <div class="stat-card"><div class="stat-num ${large ? 'stat-error':'stat-ok'}">${large}</div><div class="stat-label">Large (&gt;5k)</div></div>
    </div>

    <div class="section-title">REST Endpoints</div>
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
            <button class="copy-btn" data-copy="${escapeHtml(href)}">⎘ Copy</button>
          </div>`).join('')}
      </div>
    </div>

    <div class="section-title">Admin Links</div>
    <div class="link-list" style="margin-bottom:10px">
      ${[
        ['Site Settings',    `${base}/settings.aspx`],
        ['Site Contents',    `${base}/viewlsts.aspx`],
        ['Site Permissions', `${base}/user.aspx`],
        ['Recycle Bin',      `${base}/recyclebin.aspx`],
        ['App Management',   `${base}/appinv.aspx`],
        ['Site Features',    `${base}/ManageFeatures.aspx`],
      ].map(([lbl,href]) =>
        `<a class="link-item" href="${escapeHtml(href)}" target="_blank" rel="noreferrer">
           <span>${escapeHtml(lbl)}</span><span>↗</span>
         </a>`
      ).join('')}
    </div>

    ${warnings.length ? `
    <div class="section-title">Governance Warnings <span class="badge badge-warn">${warnings.length}</span></div>
    ${renderWarnings(warnings)}` : `
    <div class="no-warnings">✓ No governance issues detected.</div>`}
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

  return `
  <div class="toolbar">
    <input type="search" id="list-search" placeholder="Filter lists…" value="${escapeHtml(state.listFilter)}">
    <label><input type="checkbox" id="show-hidden-lists" ${state.listShowHidden ? 'checked' : ''}> Hidden</label>
  </div>
  <div class="count-bar"><span>${filtered.length} list${filtered.length !== 1 ? 's' : ''}</span></div>
  <div class="card-list">
    ${filtered.length === 0 ? '<p class="state-empty">No lists match your filter.</p>' :
      filtered.map(l => renderListCard(l, siteUrl)).join('')}
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
        ${l.Hidden ? '<span class="badge badge-default" style="margin-left:4px">hidden</span>' : ''}
        ${l.HasUniqueRoleAssignments ? '<span class="badge badge-warn" style="margin-left:2px">unique perms</span>' : ''}
        ${l.ItemCount > 5000 ? '<span class="badge badge-error" style="margin-left:2px">&gt;5k</span>' : ''}
      </div>
      <button class="btn btn-secondary" style="font-size:11px;padding:3px 8px" data-inspect="${escapeHtml(l.Id)}">
        Fields →
      </button>
    </div>
    <div class="card-body">
      <div class="card-meta">
        <span>Template:</span><span class="val">${escapeHtml(tplName)}</span>
        <span>Items:</span><span class="val">${l.ItemCount.toLocaleString()}</span>
        <span>Versioning:</span><span class="${l.EnableVersioning ? 'ok' : 'err'}">${l.EnableVersioning ? 'On' : 'Off'}</span>
        <span>Approval:</span><span class="${l.EnableModeration ? 'val' : 'val'}">${l.EnableModeration ? 'On' : 'Off'}</span>
      </div>
      <div class="copy-group">
        <button class="copy-btn" data-copy="${escapeHtml(l.Id)}">⎘ GUID</button>
        <button class="copy-btn" data-copy="${escapeHtml(l.RootFolder.ServerRelativeUrl)}">⎘ Path</button>
        <button class="copy-btn" data-copy="${escapeHtml(restUrl)}">⎘ REST URL</button>
        <button class="copy-btn" data-copy="/_api/web/lists(guid'${escapeHtml(l.Id)}')">⎘ REST Path</button>
      </div>
    </div>
  </div>`;
}

function bindListsEvents() {
  document.getElementById('list-search')?.addEventListener('input', e => {
    state.listFilter = e.target.value;
    document.querySelector('.card-list').innerHTML =
      (() => {
        const { lists, siteUrl } = state.siteData;
        const f = lists.filter(l => {
          if (!state.listShowHidden && l.Hidden) return false;
          if (state.listFilter && !l.Title.toLowerCase().includes(state.listFilter.toLowerCase())) return false;
          return true;
        });
        document.querySelector('.count-bar span').textContent = `${f.length} list${f.length !== 1 ? 's' : ''}`;
        return f.length === 0 ? '<p class="state-empty">No lists match your filter.</p>' :
          f.map(l => renderListCard(l, siteUrl)).join('');
      })();
  });

  document.getElementById('show-hidden-lists')?.addEventListener('change', e => {
    state.listShowHidden = e.target.checked;
    content.innerHTML = renderListsTab(); bindListsEvents();
  });

  content.querySelectorAll('[data-inspect]').forEach(btn => {
    btn.addEventListener('click', () => {
      const listId = btn.dataset.inspect;
      state.selectedListId = listId;
      state.listDetail = null;
      state.tab = 'fields';
      tabBar.querySelectorAll('.tab-btn').forEach(b =>
        b.classList.toggle('active', b.dataset.tab === 'fields'));
      loadListDetail(listId);
    });
  });
}

// ── Fields tab ────────────────────────────────────────────────────────────────
function renderFieldsTab() {
  if (!state.selectedListId) {
    return `<div class="state-empty">
      <div style="font-size:32px">📋</div>
      <p>Select a list in the <strong>Lists</strong> tab to inspect its fields.</p>
    </div>`;
  }
  if (!state.listDetail) return renderLoading('Loading fields…');

  const { list, fields, contentTypes, warnings } = state.listDetail;
  const visible = fields.filter(f => {
    if (!state.fieldShowHidden && f.Hidden) return false;
    if (state.fieldFilter && !f.Title.toLowerCase().includes(state.fieldFilter.toLowerCase()) &&
        !f.InternalName.toLowerCase().includes(state.fieldFilter.toLowerCase())) return false;
    return true;
  });

  const snippets = renderSnippets(listSnippets(state.siteData.siteUrl, list));

  return `
  <div class="context-strip">
    <strong>${escapeHtml(list.Title)}</strong>
    <span class="mono">${escapeHtml(list.Id)}</span>
  </div>

  ${warnings.length ? `<div class="pad" style="padding-bottom:0">${renderWarnings(warnings)}</div>` : ''}

  <details style="border-bottom:1px solid var(--sp-border)">
    <summary style="padding:6px 10px;cursor:pointer;font-size:11px;color:var(--sp-blue);background:var(--sp-surface)">
      ▶ List snippets (${listSnippets(state.siteData.siteUrl, list).length})
    </summary>
    <div class="pad" style="padding-top:6px">${snippets}</div>
  </details>

  <div class="toolbar">
    <input type="search" id="field-search" placeholder="Filter fields…" value="${escapeHtml(state.fieldFilter)}">
    <label><input type="checkbox" id="show-hidden-fields" ${state.fieldShowHidden ? 'checked' : ''}> Hidden</label>
  </div>
  <div class="count-bar">
    <span>${visible.length} field${visible.length !== 1 ? 's' : ''}</span>
    <span>·</span>
    <span>${contentTypes.filter(c => !c.Hidden).length} content type${contentTypes.length !== 1 ? 's' : ''}</span>
    <button id="copy-csv-btn">Copy CSV</button>
  </div>

  <div class="card-list" id="fields-list">
    ${visible.map(f => renderFieldRow(f, list)).join('')}
    ${visible.length === 0 ? '<p class="state-empty">No fields match your filter.</p>' : ''}
  </div>`;
}

function renderFieldRow(f, list) {
  const cssType = fieldTypeCss(f.TypeAsString);
  const sUrl = state.siteData?.siteUrl ?? '';
  const snips = renderSnippets(fieldSnippets(sUrl, list, f));
  return `
  <div class="field-row">
    <button class="field-row-header" data-toggle>
      <div class="field-row-info">
        <div class="field-name">${escapeHtml(f.Title)}</div>
        <div class="field-internal">${escapeHtml(f.InternalName)}</div>
        <div class="field-badges">
          <span class="badge type-${cssType}">${escapeHtml(f.TypeAsString)}</span>
          ${f.Required ? '<span class="badge badge-error">required</span>' : ''}
          ${f.Indexed ? '<span class="badge badge-ok">indexed</span>' : ''}
          ${f.Hidden ? '<span class="badge badge-default">hidden</span>' : ''}
          ${f.ReadOnlyField ? '<span class="badge badge-default">readonly</span>' : ''}
          ${f.EnforceUniqueValues ? '<span class="badge badge-warn">unique</span>' : ''}
        </div>
      </div>
      <span style="color:var(--sp-subtle);font-size:11px;flex-shrink:0">▼</span>
    </button>
    <div class="field-detail" style="display:none">
      <table class="info-table" style="margin-bottom:6px">
        <tr><td>Field ID</td><td class="mono">${escapeHtml(f.Id)}</td></tr>
        <tr><td>Static Name</td><td class="mono">${escapeHtml(f.StaticName)}</td></tr>
        ${f.MaxLength ? `<tr><td>Max Length</td><td>${f.MaxLength}</td></tr>` : ''}
        ${f.DefaultValue ? `<tr><td>Default</td><td>${escapeHtml(f.DefaultValue)}</td></tr>` : ''}
        ${f.LookupList ? `<tr><td>Lookup List</td><td class="mono">${escapeHtml(f.LookupList)}</td></tr>` : ''}
        ${f.LookupField ? `<tr><td>Lookup Field</td><td class="mono">${escapeHtml(f.LookupField)}</td></tr>` : ''}
        ${f.Choices?.results?.length ? `<tr><td>Choices</td><td>${escapeHtml(f.Choices.results.join(', '))}</td></tr>` : ''}
      </table>
      <div class="copy-group" style="margin-bottom:8px">
        <button class="copy-btn" data-copy="${escapeHtml(f.InternalName)}">⎘ Internal Name</button>
        <button class="copy-btn" data-copy="${escapeHtml(f.Id)}">⎘ Field ID</button>
        <button class="copy-btn" data-copy="${escapeHtml(f.StaticName)}">⎘ Static Name</button>
      </div>
      ${snips}
    </div>
  </div>`;
}

function bindFieldsEvents() {
  document.getElementById('field-search')?.addEventListener('input', e => {
    state.fieldFilter = e.target.value;
    document.getElementById('fields-list').innerHTML = (() => {
      const { fields, list } = state.listDetail;
      const v = fields.filter(f => {
        if (!state.fieldShowHidden && f.Hidden) return false;
        if (state.fieldFilter &&
            !f.Title.toLowerCase().includes(state.fieldFilter.toLowerCase()) &&
            !f.InternalName.toLowerCase().includes(state.fieldFilter.toLowerCase())) return false;
        return true;
      });
      document.querySelector('.count-bar span').textContent =
        `${v.length} field${v.length !== 1 ? 's' : ''}`;
      return v.map(f => renderFieldRow(f, list)).join('') ||
        '<p class="state-empty">No fields match your filter.</p>';
    })();
    bindExpandButtons();
  });

  document.getElementById('show-hidden-fields')?.addEventListener('change', e => {
    state.fieldShowHidden = e.target.checked;
    content.innerHTML = renderFieldsTab(); bindFieldsEvents();
  });

  document.getElementById('copy-csv-btn')?.addEventListener('click', () => {
    const { fields } = state.listDetail;
    const csv = ['InternalName,Title,Type,Required,Indexed'].concat(
      fields.map(f => `${f.InternalName},"${f.Title.replace(/"/g,'""')}",${f.TypeAsString},${f.Required},${f.Indexed}`)
    ).join('\n');
    copyToClipboard(csv);
  });

  bindExpandButtons();
}

function bindExpandButtons() {
  content.querySelectorAll('[data-toggle]').forEach(btn => {
    btn.addEventListener('click', () => {
      const detail = btn.nextElementSibling;
      const arrow  = btn.querySelector('span:last-child');
      const open   = detail.style.display === 'none';
      detail.style.display = open ? 'block' : 'none';
      if (arrow) arrow.textContent = open ? '▲' : '▼';
    });
  });
}

// ── Links tab ──────────────────────────────────────────────────────────────────
const LINK_GROUPS = [
  { category: 'Site Administration', links: [
    ['Site Settings',    '/_layouts/15/settings.aspx'],
    ['Site Contents',    '/_layouts/15/viewlsts.aspx'],
    ['Site Permissions', '/_layouts/15/user.aspx'],
    ['Recycle Bin',      '/_layouts/15/recyclebin.aspx'],
    ['App Management',   '/_layouts/15/appinv.aspx'],
    ['Site Features',    '/_layouts/15/ManageFeatures.aspx'],
  ]},
  { category: 'REST API', links: [
    ['/_api/web',                '/_api/web'],
    ['/_api/web/lists',          '/_api/web/lists'],
    ['/_api/web/siteusers',      '/_api/web/siteusers'],
    ['/_api/web/roleassignments','/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings'],
    ['/_api/web/features',       '/_api/web/features'],
    ['/_api/contextinfo',        '/_api/contextinfo'],
  ]},
  { category: 'Galleries', links: [
    ['Web Part Gallery',   '/_catalogs/wp/Forms/AllItems.aspx'],
    ['Master Page Gallery','/_catalogs/masterpage/Forms/AllItems.aspx'],
    ['Site Assets',        '/SiteAssets/Forms/AllItems.aspx'],
  ]},
  { category: 'External', links: [
    ['Power Automate',      'https://make.powerautomate.com'],
    ['Microsoft 365 Admin', 'https://admin.microsoft.com'],
    ['SharePoint Admin',    'https://admin.microsoft.com/SharePoint'],
  ]},
];

function renderLinksTab() {
  const siteUrl = state.siteData?.siteUrl ?? '';
  return `
  <div class="toolbar">
    <input type="search" id="link-search" placeholder="Search links…">
  </div>
  <div id="links-content" class="pad">
    <div class="section-title" style="margin-bottom:6px">Open Custom Path</div>
    <div style="display:flex;gap:6px;margin-bottom:14px">
      <input type="text" id="custom-path" placeholder="/_api/web/lists  or  https://…"
             style="flex:1;padding:5px 8px;font-size:12px;border:1px solid var(--sp-border);border-radius:4px;font-family:monospace;background:var(--sp-bg)">
      <button class="btn btn-primary" id="open-custom-btn">Open ↗</button>
    </div>
    ${LINK_GROUPS.map(g => `
      <div class="section-title" style="margin-top:10px;margin-bottom:4px">${escapeHtml(g.category)}</div>
      <div class="link-list">
        ${g.links.map(([lbl, path]) => {
          const href = path.startsWith('http') ? path : `${siteUrl}${path}`;
          return `<a class="link-item" href="${escapeHtml(href)}" target="_blank" rel="noreferrer">
                    <span>${escapeHtml(lbl)}</span><span>↗</span>
                  </a>`;
        }).join('')}
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
    const links = content.querySelectorAll('.link-item');
    links.forEach(a => {
      const match = !q || a.textContent.toLowerCase().includes(q) || a.href.toLowerCase().includes(q);
      a.style.display = match ? '' : 'none';
    });
  });
}

// ── Export tab ────────────────────────────────────────────────────────────────
function renderExportTab() {
  if (!state.siteData) return `<div class="state-empty">No data to export. Load a site first.</div>`;
  const fmt = state.exportFormat;
  const content_text = fmt === 'json' ? toJson(state.siteData)
    : fmt === 'markdown' ? toMarkdown(state.siteData)
    : toCsv(state.siteData);
  const preview = content_text.slice(0, 10000) +
    (content_text.length > 10000 ? '\n\n… (truncated in preview)' : '');
  const site = state.siteData.web.Title.replace(/[^a-zA-Z0-9_-]/g, '_');
  const date = new Date().toISOString().slice(0, 10);
  const ext  = fmt === 'json' ? 'json' : fmt === 'markdown' ? 'md' : 'csv';
  const filename = `sp-dev-lens_${site}_${date}.${ext}`;

  return `
  <div class="pad" style="border-bottom:1px solid var(--sp-border)">
    <div class="section-title" style="margin-bottom:6px">Format</div>
    <div class="format-selector">
      ${['json','markdown','csv'].map(f =>
        `<button class="format-btn ${f === fmt ? 'active' : ''}" data-format="${f}">${f.toUpperCase()}</button>`
      ).join('')}
    </div>
  </div>
  <div style="display:flex;gap:6px;padding:8px 10px;border-bottom:1px solid var(--sp-border);background:var(--sp-bg)">
    <button class="btn btn-primary" style="flex:1" id="download-btn">⬇ Download ${ext.toUpperCase()}</button>
    <button class="btn btn-secondary" id="copy-export-btn">⎘ Copy</button>
  </div>
  <div style="padding:4px 10px 2px;border-bottom:1px solid var(--sp-border);font-size:11px;color:var(--sp-subtle)">
    ${escapeHtml(state.siteData.web.Title)} · ${state.siteData.lists.length} lists
    · ${state.siteData.warnings.length} warnings
    · ${formatDate(state.siteData.fetchedAt)}
  </div>
  <div class="count-bar"><span>Preview (${content_text.length.toLocaleString()} chars)</span></div>
  <pre class="export-preview" id="export-preview">${escapeHtml(preview)}</pre>`;
}

function bindExportEvents() {
  content.querySelectorAll('.format-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      state.exportFormat = btn.dataset.format;
      content.innerHTML = renderExportTab(); bindExportEvents();
    });
  });

  document.getElementById('download-btn')?.addEventListener('click', () => {
    const fmt = state.exportFormat;
    const text = fmt === 'json' ? toJson(state.siteData)
      : fmt === 'markdown' ? toMarkdown(state.siteData)
      : toCsv(state.siteData);
    const site = state.siteData.web.Title.replace(/[^a-zA-Z0-9_-]/g, '_');
    const date = new Date().toISOString().slice(0, 10);
    const ext  = fmt === 'json' ? 'json' : fmt === 'markdown' ? 'md' : 'csv';
    const mime = fmt === 'json' ? 'application/json' : fmt === 'markdown' ? 'text/markdown' : 'text/csv';
    download(text, `sp-dev-lens_${site}_${date}.${ext}`, mime);
  });

  document.getElementById('copy-export-btn')?.addEventListener('click', async () => {
    const fmt = state.exportFormat;
    const text = fmt === 'json' ? toJson(state.siteData)
      : fmt === 'markdown' ? toMarkdown(state.siteData) : toCsv(state.siteData);
    const ok = await copyToClipboard(text);
    if (ok) {
      const btn = document.getElementById('copy-export-btn');
      btn.textContent = '✓ Copied'; btn.classList.add('copied');
      setTimeout(() => { btn.textContent = '⎘ Copy'; btn.classList.remove('copied'); }, 2000);
    }
  });
}

// ── Shared: warnings ──────────────────────────────────────────────────────────
function renderWarnings(warnings) {
  if (!warnings.length) return '<div class="no-warnings">✓ No issues detected.</div>';
  return `<div class="warning-list">${warnings.map(w => `
    <div class="warning-item ${w.severity}">
      <div class="warning-title">[${w.category}] ${escapeHtml(w.title)}</div>
      <div class="warning-desc">${escapeHtml(w.description)}</div>
    </div>`).join('')}</div>`;
}

// ── Boot ──────────────────────────────────────────────────────────────────────
init();
