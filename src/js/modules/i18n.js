// Runtime i18n – defaults to DE if browser language starts with 'de'
let lang = (localStorage.getItem('sdl-lang') || navigator.language || 'de').startsWith('de') ? 'de' : 'en';

const STRINGS = {
  de: {
    // Header
    refresh: 'Aktualisieren',
    langToggle: 'EN',

    // Tabs
    tabSite: 'Site',
    tabLists: 'Listen',
    tabFields: 'Felder',
    tabLinks: 'Links',
    tabExport: 'Export',

    // Site tab
    siteInfo: 'Seiteninformationen',
    siteTitle: 'Titel',
    siteUrl: 'URL',
    siteWebId: 'Web-ID',
    siteServerPath: 'Serverpfad',
    siteTemplate: 'Vorlage',
    siteLanguage: 'Sprache',
    siteHubSite: 'Hub-Site',
    siteOwners: 'Eigentümer',
    siteMembers: 'Mitglieder',
    siteLastModified: 'Zuletzt geändert',
    siteQuickSummary: 'Schnellübersicht',
    siteVisibleLists: 'Sichtbare Listen',
    siteHiddenLists: 'Versteckte Listen',
    siteUniquePerms: 'Eigene Berechtigungen',
    siteLargeLists: 'Große Listen (>5k)',
    siteRestEndpoints: 'REST-Endpunkte',
    siteAdminLinks: 'Admin-Links',
    siteSettings: 'Site-Einstellungen',
    siteContents: 'Websiteinhalte',
    sitePermissions: 'Berechtigungen',
    siteRecycleBin: 'Papierkorb',
    siteAppMgmt: 'App-Verwaltung',
    siteFeatures: 'Site-Features',

    // Lists tab
    filterLists: 'Listen filtern…',
    showHidden: 'Versteckte',
    noListsMatch: 'Keine Listen entsprechen dem Filter.',
    template: 'Vorlage',
    items: 'Elemente',
    versioning: 'Versionierung',
    approval: 'Genehmigung',
    on: 'An',
    off: 'Aus',
    yes: 'Ja',
    no: 'Nein',
    fieldsBtnLabel: 'Felder →',
    lists: 'Listen',
    list: 'Liste',

    // Fields tab
    selectListFirst: 'Wähle eine Liste im <strong>Listen</strong>-Tab aus, um deren Felder zu sehen.',
    filterFields: 'Felder filtern…',
    listSnippets: 'Listen-Snippets',
    fieldId: 'Feld-ID',
    staticName: 'Statischer Name',
    maxLength: 'Max. Länge',
    defaultVal: 'Standardwert',
    lookupList: 'Nachschlageliste',
    lookupField: 'Nachschlagefeld',
    choices: 'Optionen',
    copyCsv: 'CSV kopieren',
    contentTypes: 'Inhaltstypen',
    fields: 'Felder',
    field: 'Feld',

    // Badges
    hiddenBadge: 'versteckt',
    uniquePermsBadge: 'eigene Berechtigungen',
    requiredBadge: 'Pflichtfeld',
    indexedBadge: 'indiziert',
    readonlyBadge: 'schreibgeschützt',
    uniqueBadge: 'eindeutig',

    // Write operations
    makeVisible: 'Sichtbar machen',
    makeEditable: 'Schreibbar machen',
    writeOpConfirmTitle: 'Spalteneinstellung ändern',
    writeOpMakeVisibleMsg: 'Diese Spalte wird für Benutzer in Formularen und Ansichten sichtbar gemacht <code>(Hidden = false)</code>.',
    writeOpMakeEditableMsg: 'Diese Spalte wird auf schreibbar gesetzt <code>(ReadOnlyField = false)</code>. Systemeigene Spalten sollten nicht verändert werden.',
    writeOpWarning: 'Diese Änderung wirkt sich direkt auf die SharePoint-Spaltenstruktur aus. Dieser Vorgang kann nicht automatisch rückgängig gemacht werden.',
    cancel: 'Abbrechen',
    confirm: 'Bestätigen',
    updating: 'Wird geändert…',
    changeSuccess: 'Erfolgreich geändert',
    changeFailed: 'Fehler beim Ändern',

    // Links tab
    openCustomPath: 'Benutzerdefinierten Pfad öffnen',
    customPathPlaceholder: '/_api/web/lists  oder  https://…',
    openBtn: 'Öffnen ↗',
    searchLinks: 'Links suchen…',
    siteAdministration: 'Site-Administration',
    restApi: 'REST API',
    galleries: 'Galerien',
    external: 'Extern',
    webPartGallery: 'Webpart-Galerie',
    masterPageGallery: 'Gestaltungsvorlagenkatalog',
    siteAssets: 'Site-Assets',
    powerAutomate: 'Power Automate',
    m365Admin: 'Microsoft 365 Admin',
    spAdmin: 'SharePoint Admin',

    // Export tab
    format: 'Format',
    downloadBtn: 'Herunterladen',
    copy: 'Kopieren',
    copied: 'Kopiert',
    preview: 'Vorschau',
    chars: 'Zeichen',
    noDataToExport: 'Keine Daten zum Exportieren. Lade zuerst eine Site.',
    warnings: 'Warnungen',

    // Governance
    governanceWarnings: 'Governance-Warnungen',
    noIssues: 'Keine Governance-Probleme gefunden.',
    healthScore: 'Bewertung',
    severityError: 'Kritisch',
    severityWarning: 'Warnung',
    severityInfo: 'Info',
    catPerformance: 'Leistung',
    catSecurity: 'Sicherheit',
    catGovernance: 'Governance',
    catConfiguration: 'Konfiguration',

    // Status / Error
    noActiveTab: 'Kein aktives Tab gefunden.',
    notSharePoint: 'Diese Seite ist keine SharePoint-Site.',
    cannotConnect: 'Verbindung zur Seite nicht möglich.',
    connectionHint: 'Tipp: Lade die SharePoint-Seite neu und versuche es erneut.',
    reloadPage: 'Seite neu laden',
    retry: 'Erneut versuchen',
    connecting: 'Verbinde mit SharePoint…',
    loadingFields: 'Felder werden geladen…',
    noSiteData: 'Keine Seitendaten vorhanden.',
    error: 'Fehler',
    snippets: 'Snippets',

    // Popup
    popupSubtitle: 'Entwickler-Inspektor',
    detectingPage: 'Seite wird erkannt…',
    openPanel: 'Inspektor-Panel öffnen →',
    quickLinks: 'Schnelllinks',
    notAffiliated: 'Nicht mit Microsoft verbunden · v1.0.0',
    noSharePointDetected: 'Keine SharePoint-Seite erkannt.',
    cannotConnectPopup: 'Verbindung nicht möglich. Seite neu laden.',
    noTabFound: 'Kein aktives Tab gefunden.',
  },

  en: {
    // Header
    refresh: 'Refresh',
    langToggle: 'DE',

    // Tabs
    tabSite: 'Site',
    tabLists: 'Lists',
    tabFields: 'Fields',
    tabLinks: 'Links',
    tabExport: 'Export',

    // Site tab
    siteInfo: 'Site Information',
    siteTitle: 'Title',
    siteUrl: 'URL',
    siteWebId: 'Web ID',
    siteServerPath: 'Server Path',
    siteTemplate: 'Template',
    siteLanguage: 'Language',
    siteHubSite: 'Hub Site',
    siteOwners: 'Owners',
    siteMembers: 'Members',
    siteLastModified: 'Last Modified',
    siteQuickSummary: 'Quick Summary',
    siteVisibleLists: 'Visible Lists',
    siteHiddenLists: 'Hidden Lists',
    siteUniquePerms: 'Unique Perms',
    siteLargeLists: 'Large Lists (>5k)',
    siteRestEndpoints: 'REST Endpoints',
    siteAdminLinks: 'Admin Links',
    siteSettings: 'Site Settings',
    siteContents: 'Site Contents',
    sitePermissions: 'Site Permissions',
    siteRecycleBin: 'Recycle Bin',
    siteAppMgmt: 'App Management',
    siteFeatures: 'Site Features',

    // Lists tab
    filterLists: 'Filter lists…',
    showHidden: 'Hidden',
    noListsMatch: 'No lists match your filter.',
    template: 'Template',
    items: 'Items',
    versioning: 'Versioning',
    approval: 'Approval',
    on: 'On',
    off: 'Off',
    yes: 'Yes',
    no: 'No',
    fieldsBtnLabel: 'Fields →',
    lists: 'Lists',
    list: 'List',

    // Fields tab
    selectListFirst: 'Select a list in the <strong>Lists</strong> tab to inspect its fields.',
    filterFields: 'Filter fields…',
    listSnippets: 'List Snippets',
    fieldId: 'Field ID',
    staticName: 'Static Name',
    maxLength: 'Max Length',
    defaultVal: 'Default',
    lookupList: 'Lookup List',
    lookupField: 'Lookup Field',
    choices: 'Choices',
    copyCsv: 'Copy CSV',
    contentTypes: 'Content Types',
    fields: 'Fields',
    field: 'Field',

    // Badges
    hiddenBadge: 'hidden',
    uniquePermsBadge: 'unique perms',
    requiredBadge: 'required',
    indexedBadge: 'indexed',
    readonlyBadge: 'readonly',
    uniqueBadge: 'unique',

    // Write operations
    makeVisible: 'Make Visible',
    makeEditable: 'Make Editable',
    writeOpConfirmTitle: 'Change Column Setting',
    writeOpMakeVisibleMsg: 'This column will be made visible in forms and views <code>(Hidden = false)</code>.',
    writeOpMakeEditableMsg: 'This column will be set to editable <code>(ReadOnlyField = false)</code>. Avoid changing system columns.',
    writeOpWarning: 'This change directly affects the SharePoint column schema and cannot be automatically undone.',
    cancel: 'Cancel',
    confirm: 'Confirm',
    updating: 'Updating…',
    changeSuccess: 'Changed successfully',
    changeFailed: 'Failed to change',

    // Links tab
    openCustomPath: 'Open Custom Path',
    customPathPlaceholder: '/_api/web/lists  or  https://…',
    openBtn: 'Open ↗',
    searchLinks: 'Search links…',
    siteAdministration: 'Site Administration',
    restApi: 'REST API',
    galleries: 'Galleries',
    external: 'External',
    webPartGallery: 'Web Part Gallery',
    masterPageGallery: 'Master Page Gallery',
    siteAssets: 'Site Assets',
    powerAutomate: 'Power Automate',
    m365Admin: 'Microsoft 365 Admin',
    spAdmin: 'SharePoint Admin',

    // Export tab
    format: 'Format',
    downloadBtn: 'Download',
    copy: 'Copy',
    copied: 'Copied',
    preview: 'Preview',
    chars: 'chars',
    noDataToExport: 'No data to export. Load a site first.',
    warnings: 'Warnings',

    // Governance
    governanceWarnings: 'Governance Warnings',
    noIssues: 'No governance issues detected.',
    healthScore: 'Score',
    severityError: 'Critical',
    severityWarning: 'Warning',
    severityInfo: 'Info',
    catPerformance: 'Performance',
    catSecurity: 'Security',
    catGovernance: 'Governance',
    catConfiguration: 'Configuration',

    // Status / Error
    noActiveTab: 'No active tab found.',
    notSharePoint: 'This page is not a SharePoint site.',
    cannotConnect: 'Cannot connect to page.',
    connectionHint: 'Tip: Reload the SharePoint page and try again.',
    reloadPage: 'Reload page',
    retry: 'Retry',
    connecting: 'Connecting to SharePoint…',
    loadingFields: 'Loading fields…',
    noSiteData: 'No site data available.',
    error: 'Error',
    snippets: 'Snippets',

    // Popup
    popupSubtitle: 'Developer Inspector',
    detectingPage: 'Detecting page…',
    openPanel: 'Open Inspector Panel →',
    quickLinks: 'Quick Links',
    notAffiliated: 'Not affiliated with Microsoft · v1.0.0',
    noSharePointDetected: 'No SharePoint page detected.',
    cannotConnectPopup: 'Cannot connect. Reload the page.',
    noTabFound: 'No active tab found.',
  },
};

export function t(key) {
  return STRINGS[lang][key] ?? STRINGS.en[key] ?? key;
}

export function toggleLang() {
  lang = lang === 'de' ? 'en' : 'de';
  localStorage.setItem('sdl-lang', lang);
  return lang;
}

export function getLang() { return lang; }
