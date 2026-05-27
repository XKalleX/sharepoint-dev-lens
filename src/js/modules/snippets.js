import { escapeHtml } from './utils.js';

export function listSnippets(siteUrl, list) {
  const g = list.Id;
  return [
    {
      label: 'REST – Get items',
      code: `GET ${siteUrl}/_api/web/lists(guid'${g}')/items\nAccept: application/json;odata=verbose`,
    },
    {
      label: 'REST – List metadata',
      code: `GET ${siteUrl}/_api/web/lists(guid'${g}')\nAccept: application/json;odata=verbose`,
    },
    {
      label: 'Power Automate – HTTP URI',
      code: `${siteUrl}/_api/web/lists(guid'${g}')/items`,
    },
    {
      label: 'PnPjs – Get items',
      code: `const items = await sp.web.lists\n  .getById("${g}")\n  .items();`,
    },
    {
      label: 'Graph – List items',
      code: `GET https://graph.microsoft.com/v1.0/sites/{site-id}/lists/${g}/items`,
    },
  ];
}

export function fieldSnippets(siteUrl, list, field) {
  return [
    {
      label: 'REST – Filter by field',
      code: `GET ${siteUrl}/_api/web/lists(guid'${list.Id}')/items\n    ?$filter=${field.InternalName} eq 'value'`,
    },
    {
      label: 'REST – Select field',
      code: `GET ${siteUrl}/_api/web/lists(guid'${list.Id}')/items\n    ?$select=${field.InternalName}`,
    },
    {
      label: 'PnPjs – Filter',
      code: `const items = await sp.web.lists\n  .getById("${list.Id}")\n  .items\n  .filter(\`${field.InternalName} eq 'value'\`)();`,
    },
    {
      label: 'Power Automate – OData filter',
      code: `${field.InternalName} eq 'value'`,
    },
  ];
}

export function renderSnippets(snippets) {
  return snippets.map(s => `
    <div class="snippet-block">
      <div class="snippet-header">
        <span>${escapeHtml(s.label)}</span>
        <button class="copy-btn" data-copy="${escapeHtml(s.code)}">⎘ Copy</button>
      </div>
      <pre class="snippet-code">${escapeHtml(s.code)}</pre>
    </div>`).join('');
}
