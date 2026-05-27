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
    if (json && json.d !== undefined) return json.d;
    if (json && json.value !== undefined) return json.value;
    return json;
  }

  async function getFormDigest(siteUrl) {
    const res = await fetch(`${siteUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: REST_HEADERS,
      credentials: 'include',
    });
    if (!res.ok) throw new Error(`contextinfo ${res.status}`);
    const json = await res.json();
    const d = json?.d ?? json;
    return d?.GetContextWebInformation?.FormDigestValue ?? d?.FormDigestValue;
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

  // ── Governance analysis ──────────────────────────────────────────────────────

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

    if (list.ItemCount > 20000) {
      warnings.push({
        id: `xlarge-${list.Id}`, severity: 'error', category: 'performance',
        title: 'Liste sehr groß (>20.000 Elemente)',
        description: `"${n}" hat ${list.ItemCount.toLocaleString()} Elemente. Abfragen ohne indizierte Spalten schlagen fehl und API-Abrufe sind stark verlangsamt.`,
        quickfix: null,
      });
    } else if (list.ItemCount > 5000) {
      warnings.push({
        id: `large-${list.Id}`, severity: 'warning', category: 'performance',
        title: 'Liste überschreitet Ansichtsschwellenwert (>5.000)',
        description: `"${n}" hat ${list.ItemCount.toLocaleString()} Elemente. Ansichten ohne indizierte Spalten schlagen fehl.`,
        quickfix: null,
      });
    }

    if (list.HasUniqueRoleAssignments) {
      warnings.push({
        id: `perms-${list.Id}`, severity: 'info', category: 'security',
        title: 'Gebrochene Berechtigungsvererbung',
        description: `"${n}" hat eigene Berechtigungen (Vererbung gebrochen). Berechtigungsänderungen am übergeordneten Element werden nicht vererbt.`,
        quickfix: null,
      });
    }

    if (!list.EnableVersioning && list.BaseTemplate !== 102) {
      const iDocLib = list.BaseTemplate === 101;
      warnings.push({
        id: `ver-${list.Id}`, severity: iDocLib ? 'warning' : 'info', category: 'governance',
        title: iDocLib ? 'Dokumentbibliothek ohne Versionierung' : 'Versionierung deaktiviert',
        description: `"${n}" hat keine Versionierung. Versehentliche Änderungen und Löschungen können nicht wiederhergestellt werden.`,
        quickfix: null,
      });
    }

    if (list.Hidden) {
      warnings.push({
        id: `hid-${list.Id}`, severity: 'info', category: 'configuration',
        title: 'Versteckte Liste',
        description: `"${n}" ist in der Benutzeroberfläche ausgeblendet.`,
        quickfix: null,
      });
    }

    if (list.NoCrawl) {
      warnings.push({
        id: `crawl-${list.Id}`, severity: 'info', category: 'configuration',
        title: 'Von Suche ausgeschlossen',
        description: `"${n}" ist explizit von der SharePoint-Suche ausgeschlossen (NoCrawl = true).`,
        quickfix: null,
      });
    }

    return warnings;
  }

  function analyzeFields(list, fields) {
    const warnings = [];
    const n = list.Title;
    const lookups = fields.filter(f => f.TypeAsString === 'Lookup' || f.TypeAsString === 'LookupMulti');
    const calcFields = fields.filter(f => f.TypeAsString === 'Calculated');

    if (lookups.length > 8) {
      warnings.push({
        id: `lookups-${list.Id}`, severity: 'warning', category: 'performance',
        title: 'Viele Nachschlagespalten',
        description: `"${n}" hat ${lookups.length} Nachschlagespalten. Listenabfragen können stark verlangsamt sein. Empfohlen: max. 8.`,
        quickfix: null,
      });
    }

    if (list.ItemCount > 5000) {
      const nonIndexed = lookups.filter(f => !f.Indexed);
      if (nonIndexed.length > 0) {
        warnings.push({
          id: `idx-${list.Id}`, severity: 'error', category: 'performance',
          title: 'Nicht indizierte Nachschlagespalten auf großer Liste',
          description: `"${n}" ist groß und hat nicht indizierte Nachschlagespalten: ${nonIndexed.map(f => f.InternalName).join(', ')}. Abfragen über diese Spalten schlagen fehl.`,
          quickfix: null,
        });
      }
    }

    if (calcFields.length > 0) {
      warnings.push({
        id: `calc-${list.Id}`, severity: 'info', category: 'performance',
        title: 'Berechnete Spalten vorhanden',
        description: `"${n}" hat ${calcFields.length} berechnete Spalte(n): ${calcFields.map(f => f.Title).join(', ')}. Diese können bei Massenoperationen Probleme verursachen und verlangsamen Schreibzugriffe.`,
        quickfix: null,
      });
    }

    const hiddenNonSystem = fields.filter(f =>
      f.Hidden && !f.ReadOnlyField &&
      !['ID', 'ContentType', 'Modified', 'Created', 'Author', 'Editor',
        'Attachments', 'Edit', 'LinkTitleNoMenu', 'LinkTitle', 'DocIcon',
        'ServerUrl', 'EncodedAbsUrl', 'BaseName', 'FileSizeDisplay',
        '_UIVersionString', 'ItemChildCount', 'FolderChildCount'].includes(f.InternalName)
    );
    if (hiddenNonSystem.length > 0) {
      warnings.push({
        id: `hidf-${list.Id}`, severity: 'info', category: 'configuration',
        title: 'Versteckte benutzerdefinierte Spalten',
        description: `"${n}" hat ${hiddenNonSystem.length} versteckte benutzerdefinierte Spalte(n): ${hiddenNonSystem.slice(0, 5).map(f => f.InternalName).join(', ')}${hiddenNonSystem.length > 5 ? '…' : ''}. Diese können mit dem Felder-Tab sichtbar gemacht werden.`,
        quickfix: null,
      });
    }

    const readonlyNonSystem = fields.filter(f =>
      f.ReadOnlyField && !f.Hidden &&
      !['ID', 'ContentType', 'Modified', 'Created', 'Author', 'Editor',
        'Attachments', 'Edit', 'LinkTitleNoMenu', 'LinkTitle', 'DocIcon',
        'ServerUrl', 'EncodedAbsUrl', 'BaseName', 'FileSizeDisplay',
        '_UIVersionString', 'ItemChildCount', 'FolderChildCount'].includes(f.InternalName)
    );
    if (readonlyNonSystem.length > 0) {
      warnings.push({
        id: `rof-${list.Id}`, severity: 'info', category: 'configuration',
        title: 'Schreibgeschützte benutzerdefinierte Spalten',
        description: `"${n}" hat ${readonlyNonSystem.length} schreibgeschützte Spalte(n): ${readonlyNonSystem.slice(0, 5).map(f => f.InternalName).join(', ')}${readonlyNonSystem.length > 5 ? '…' : ''}. Diese können im Felder-Tab schreibbar gemacht werden.`,
        quickfix: null,
      });
    }

    return warnings;
  }

  function analyzeSite(web, lists) {
    const warnings = [];
    const uniquePermsCount = lists.filter(l => l.HasUniqueRoleAssignments).length;

    if (lists.length > 300) {
      warnings.push({
        id: 'many-lists', severity: 'warning', category: 'governance',
        title: 'Sehr viele Listen',
        description: `Diese Site hat ${lists.length} Listen/Bibliotheken. Microsoft empfiehlt max. 2.000, aber die Leistung sinkt ab ca. 300. Prüfe, ob Listen zusammengeführt oder auf andere Sites verteilt werden können.`,
        quickfix: null,
      });
    }

    if (uniquePermsCount > 5) {
      warnings.push({
        id: 'many-unique-perms', severity: 'warning', category: 'security',
        title: 'Viele gebrochene Berechtigungsvererbungen',
        description: `${uniquePermsCount} Listen/Bibliotheken haben eigene Berechtigungen. Dies erschwert die Berechtigungsverwaltung erheblich.`,
        quickfix: null,
      });
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
        const siteWarnings = analyzeSite(web, lists);
        const listWarnings = lists.flatMap(l => analyzeList(l));
        const warnings = [...siteWarnings, ...listWarnings];
        return { success: true, data: { web, lists, warnings, fetchedAt: new Date().toISOString(), siteUrl } };
      }

      case 'GET_LIST_DETAIL': {
        const [allLists, fields, contentTypes] = await Promise.all([
          fetchLists(siteUrl),
          fetchListFields(siteUrl, req.listId),
          fetchListContentTypes(siteUrl, req.listId),
        ]);
        const list = allLists.find(l => l.Id === req.listId);
        if (!list) return { success: false, error: `Liste ${req.listId} nicht gefunden` };
        const warnings = [...analyzeList(list), ...analyzeFields(list, fields)];
        return { success: true, data: { list, fields, contentTypes, warnings } };
      }

      case 'GET_PERMISSIONS': {
        const assignments = await fetchRoleAssignments(siteUrl, req.listId);
        return { success: true, data: assignments };
      }

      case 'UPDATE_FIELD': {
        const digest = await getFormDigest(siteUrl);
        if (!digest) return { success: false, error: 'Form Digest konnte nicht abgerufen werden.' };

        const fieldUrl = `${siteUrl}/_api/web/lists(guid'${req.listId}')/fields(guid'${req.fieldId}')`;
        const body = { __metadata: { type: 'SP.Field' }, ...req.properties };

        const res = await fetch(fieldUrl, {
          method: 'POST',
          headers: {
            ...REST_HEADERS,
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'X-RequestDigest': digest,
          },
          body: JSON.stringify(body),
          credentials: 'include',
        });

        if (!res.ok) {
          const text = await res.text().catch(() => '');
          throw new Error(`MERGE ${res.status} ${res.statusText}: ${text.slice(0, 300)}`);
        }
        return { success: true };
      }

      default:
        return { success: false, error: 'Unbekannter Anfrage-Typ' };
    }
  }

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    handleRequest(message)
      .then(sendResponse)
      .catch(err => sendResponse({ success: false, error: err.message || String(err) }));
    return true;
  });

})();
