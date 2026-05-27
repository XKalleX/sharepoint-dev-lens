export async function copyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    const el = document.createElement('textarea');
    el.value = text;
    el.style.cssText = 'position:fixed;opacity:0';
    document.body.appendChild(el);
    el.select();
    const ok = document.execCommand('copy');
    document.body.removeChild(el);
    return ok;
  }
}

export function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function formatDate(iso) {
  if (!iso) return '–';
  try { return new Date(iso).toLocaleString(); } catch { return iso; }
}

/** Map SharePoint field TypeAsString to a CSS class suffix */
export function fieldTypeCss(type) {
  if (!type) return 'default';
  const t = type.toLowerCase();
  if (t === 'text' || t === 'note') return 'text';
  if (t === 'number' || t === 'currency' || t === 'datetime') return 'number';
  if (t.startsWith('lookup')) return 'lookup';
  if (t.startsWith('choice') || t === 'multichoice') return 'choice';
  if (t.startsWith('user')) return 'user';
  if (t === 'calculated') return 'calc';
  return 'default';
}

const BASE_TEMPLATE_NAMES = {
  100: 'Custom List', 101: 'Document Library', 102: 'Survey',
  103: 'Links', 104: 'Announcements', 105: 'Contacts', 106: 'Events',
  107: 'Tasks', 108: 'Discussion Board', 109: 'Picture Library',
  115: 'XML Form Library', 119: 'Wiki Page Library', 130: 'Blog Posts',
  140: 'User Information', 301: 'Blog', 499: 'Site Assets',
  851: 'Asset Library', 1100: 'Issue Tracking', 1300: 'Promoted Links',
  2002: 'Site Pages', 2100: 'Project Tasks', 10102: 'Record Library',
};
export function getBaseTemplateName(t) {
  return BASE_TEMPLATE_NAMES[t] || `Template ${t}`;
}
