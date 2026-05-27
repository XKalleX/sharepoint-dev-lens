import { getCurrentTab, sendToTab } from './modules/messaging.js';

const dot   = document.getElementById('status-dot');
const title = document.getElementById('status-title');
const url   = document.getElementById('status-url');
const btn   = document.getElementById('open-panel-btn');
const links = document.getElementById('quick-links');
const list  = document.getElementById('link-list');

async function init() {
  const tab = await getCurrentTab();
  if (!tab?.id) { return setStatus('err', 'No active tab found.', ''); }

  const resp = await sendToTab(tab.id, { type: 'PING' });
  if (!resp.success) { return setStatus('err', 'Cannot connect to page.', tab.url ?? ''); }

  const { isSharePoint, siteUrl, pageTitle } = resp.data;

  if (!isSharePoint) {
    return setStatus('warn', 'No SharePoint page detected.', tab.url ?? '');
  }

  setStatus('ok', pageTitle || siteUrl, siteUrl);
  btn.disabled = false;
  btn.textContent = 'Open Inspector Panel →';

  // Quick links
  links.classList.remove('hidden');
  list.innerHTML = [
    ['Site Settings',  `${siteUrl}/_layouts/15/settings.aspx`],
    ['Site Contents',  `${siteUrl}/_layouts/15/viewlsts.aspx`],
    ['REST /_api/web', `${siteUrl}/_api/web`],
  ].map(([label, href]) =>
    `<a class="link-item" href="${href}" target="_blank" rel="noreferrer">
       <span>${label}</span><span>↗</span>
     </a>`
  ).join('');
}

function setStatus(state, t, u) {
  dot.className = `status-dot ${state}`;
  title.textContent = t;
  url.textContent = u;
}

btn.addEventListener('click', async () => {
  const tab = await getCurrentTab();
  if (tab?.id) {
    await chrome.sidePanel.open({ tabId: tab.id });
    window.close();
  }
});

init();
