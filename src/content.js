/**
 * SharePoint Dev Lens – Content Script
 * Runs on SharePoint pages. Handles all REST API calls using the
 * user's authenticated browser session (same-origin, no OAuth needed).
 */
(function () {
  'use strict';

  // ── REST helpers ────────────────────────────────────────────────────────────

  const REST_HEADERS = {
    Accept: 'application/json;odata=verbose',
    'Content-Type': 'application/json;odata=verbose',
  };

  async function restGet(url) {
    const res = await fetch(url, { headers: REST_HEADERS, credentials: 'include' });
    if (!res.ok) {
      const text = await res.text().catch(() => '');
      throw new Error(`REST ${res.status} ${res.statusText}: ${text.slice(0, 200)}`);
    }
    const json = await res.json();
    // OData verbose wraps results in .d; minimal/json uses .value
    if (json && json.d !== undefined) return json.d;
    if (json && json.value !== undefined) return json.value;
    return json;
  }

  async function fetchWebInfo(siteUrl) {
    return restGet(
      `${siteUrl}/_api/web?$select=Id,Title,Url,ServerRelativeUrl,Language,` +
      `WebTemplate,Configuration,Created,LastItemModifiedDate,IsHubSite,HubSiteId` +
      `&$expand=AssociatedMemberGroup,AssociatedOwnerGroup`
    );
  }

  async function fetchLists(siteUrl) {
    const result = await restGet(
      `${siteUrl}/_api/web/lists?$select=Id,Title,` +
      `RootFolder/ServerRelativeUrl,BaseTemplate,Hidden,ItemCount,` +
      `EnableVersioning,EnableModeration,EnableAttachments,NoCrawl,` +
      `HasUniqueRoleAssignments,DefaultViewUrl,Description` +
      `&$expand=RootFolder&$orderby=Title`
    );
    return result.results || result;
  }

  async function fetchListFields(siteUrl, listId) {
    const result = await restGet(
      `${siteUrl}/_api/web/lists(guid'${listId}')/fields?` +
      `$select=Id,Title,InternalName,StaticName,TypeAsString,Required,Hidden,` +
      `ReadOnlyField,Indexed,EnforceUniqueValues,LookupList,LookupField,` +
      `Choices,DefaultValue,MaxLength&$orderby=Title`
    );
    return result.results || result;
  }

  async function fetchListContentTypes(siteUrl, listId) {
    const result = await restGet(
      `${siteUrl}/_api/web/lists(guid'${listId}')/ContentTypes?` +
      `$select=Id,Name,Description,Group,Hidden`
    );
    return result.results || result;
  }

  async function fetchRoleAssignments(siteUrl, listId) {
    const base = listId
      ? `${siteUrl}/_api/web/lists(guid'${listId}')/roleassignments`
      : `${siteUrl}/_api/web/roleassignments`;
    const result = await restGet(`${base}?$expand=Member,RoleDefinitionBindings`);
    return result.results || result;
  }

  // ── Governance rules ─────────────────────────────────────────────────────────

  const BASE_TEMPLATE_NAMES = {
    100: 'Custom List', 101: 'Document Library', 102: 'Survey',
    103: 'Links', 104: 'Announcements', 105: 'Contacts', 106: 'Events',
    107: 'Tasks', 108: 'Discussion Board', 109: 'Picture Library',
    115: 'XML Form Library', 119: 'Wiki Page Library', 130: 'Blog Posts',
    140: 'User Information', 301: 'Blog', 499: 'Site Assets',
    851: 'Asset Library', 1100: 'Issue Tracking', 1300: 'Promoted Links',
    2002: 'Site Pages', 2100: 'Project Tasks', 10102: 'Record Library',
  };

  function getBaseTemplateName(t) {
    return BASE_TEMPLATE_NAMES[t] || `Template ${t}`;
  }

  function analyzeList(list) {
    const warnings = [];
    const n = list.Title;

    if (list.ItemCount > 5000) {
      warnings.push({
        id: `large-${list.Id}`, severity: 'warning', category: 'performance',
        title: 'List exceeds view threshold',
        description: `"${n}" has ${list.ItemCount.toLocaleString()} items (>5,000). Views without indexed columns will fail.`,
      });
    }
    if (list.HasUniqueRoleAssignments) {
      warnings.push({
        id: `perms-${list.Id}`, severity: 'info', category: 'security',
        title: 'Unique permissions',
        description: `"${n}" has broken permission inheritance.`,
      });
    }
    if (!list.EnableVersioning && list.BaseTemplate !== 102) {
      warnings.push({
        id: `ver-${list.Id}`, severity: 'warning', category: 'governance',
        title: 'Versioning disabled',
        description: `"${n}" has versioning disabled. Accidental changes cannot be recovered.`,
      });
    }
    if (list.Hidden) {
      warnings.push({
        id: `hid-${list.Id}`, severity: 'info', category: 'configuration',
        title: 'Hidden list',
        description: `"${n}" is hidden from the UI.`,
      });
    }
    return warnings;
  }

  function analyzeFields(list, fields) {
    const warnings = [];
    const n = list.Title;
    const lookups = fields.filter(f => f.TypeAsString === 'Lookup' || f.TypeAsString === 'LookupMulti');

    if (lookups.length > 8) {
      warnings.push({
        id: `lookups-${list.Id}`, severity: 'warning', category: 'performance',
        title: 'Many lookup columns',
        description: `"${n}" has ${lookups.length} lookup columns. Queries may be slow.`,
      });
    }
    if (list.ItemCount > 5000) {
      const nonIndexed = lookups.filter(f => !f.Indexed);
      if (nonIndexed.length > 0) {
        warnings.push({
          id: `idx-${list.Id}`, severity: 'error', category: 'performance',
          title: 'Non-indexed lookups on large list',
          description: `"${n}" is large and has non-indexed lookups: ${nonIndexed.map(f => f.Title).join(', ')}.`,
        });
      }
    }
    return warnings;
  }

  // ── SharePoint detection ─────────────────────────────────────────────────────

  function resolveSiteUrl() {
    const { origin, pathname } = window.location;
    const parts = pathname.split('/').filter(Boolean);
    const managed = ['sites', 'teams', 'personal'];
    const idx = parts.findIndex(p => managed.includes(p));
    if (idx !== -1 && parts[idx + 1]) {
      return `${origin}/${parts.slice(0, idx + 2).join('/')}`;
    }
    return origin;
  }

  function detectSharePoint() {
    if (!window.location.hostname.endsWith('.sharepoint.com')) return false;
    return (
      !!document.getElementById('ms-designer-ribbon') ||
      !!document.getElementById('SPPageChrome') ||
      !!document.getElementById('O365_NavHeader') ||
      !!document.querySelector('[data-sp-webpartmanager]') ||
      !!(window._spPageContextInfo)
    );
  }

  // ── Request handler ──────────────────────────────────────────────────────────

  async function handleRequest(req) {
    const siteUrl = resolveSiteUrl();

    switch (req.type) {
      case 'PING':
        return {
          success: true,
          data: { isSharePoint: detectSharePoint(), siteUrl, pageTitle: document.title },
        };

      case 'GET_SITE_DATA': {
        const [web, lists] = await Promise.all([
          fetchWebInfo(siteUrl),
          fetchLists(siteUrl),
        ]);
        const warnings = lists.flatMap(l => analyzeList(l));
        return { success: true, data: { web, lists, warnings, fetchedAt: new Date().toISOString(), siteUrl } };
      }

      case 'GET_LIST_DETAIL': {
        const [allLists, fields, contentTypes] = await Promise.all([
          fetchLists(siteUrl),
          fetchListFields(siteUrl, req.listId),
          fetchListContentTypes(siteUrl, req.listId),
        ]);
        const list = allLists.find(l => l.Id === req.listId);
        if (!list) return { success: false, error: `List ${req.listId} not found` };
        const warnings = [...analyzeList(list), ...analyzeFields(list, fields)];
        return { success: true, data: { list, fields, contentTypes, warnings } };
      }

      case 'GET_PERMISSIONS': {
        const assignments = await fetchRoleAssignments(siteUrl, req.listId);
        return { success: true, data: assignments };
      }

      default:
        return { success: false, error: 'Unknown request type' };
    }
  }

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    handleRequest(message)
      .then(sendResponse)
      .catch(err => sendResponse({ success: false, error: err.message || String(err) }));
    return true; // keep channel open for async response
  });

})();
