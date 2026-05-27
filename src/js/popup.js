import { getCurrentTab, sendToTab } from './modules/messaging.js';
import { t } from './modules/i18n.js';
import { icon } from './modules/icons.js';

const dot        = document.getElementById('status-dot');
const titleEl    = document.getElementById('status-title');
const urlEl      = document.getElementById('status-url');
const btn        = document.getElementById('open-panel-btn');
const links      = document.getElementById('quick-links');
const list       = document.getElementById('link-list');
const hintEl     = document.getElementById('reload-hint');
const subtitleEl = document.getElementById('popup-subtitle');
const footerEl   = document.getElementById('popup-footer');
const iconEl     = document.getElementById('popup-icon');

function applyI18n() {
  subtitleEl.textContent = t('popupSubtitle');
  footerEl.textContent   = t('notAffiliated');
  titleEl.textContent    = t('detectingPage');
  btn.textContent        = t('openPanel');
  document.getElementById('quick-links-title').textContent = t('quickLinks');
  iconEl.innerHTML = icon('search', 26);
}

async function init() {
  applyI18n();

  const tab = await getCurrentTab();
  if (!tab?.id) {
    return setStatus('err', t('noTabFound'), '');
  }

  const resp = await sendToTab(tab.id, { type: 'PING' });

  if (!resp.success) {
    setStatus('err', t('cannotConnectPopup'), tab.url ?? '');
    hintEl.innerHTML = `${icon('info', 12)} ${t('connectionHint')}`;
    hintEl.classList.remove('hidden');
    return;
  }

  const { isSharePoint, siteUrl, pageTitle } = resp.data;

  if (!isSharePoint) {
    return setStatus('warn', t('noSharePointDetected'), tab.url ?? '');
  }

  setStatus('ok', pageTitle || siteUrl, siteUrl);
  btn.disabled = false;
  btn.textContent = t('openPanel');

  // Quick links
  links.classList.remove('hidden');
  list.innerHTML = [
    [t('siteSettings'),  `${siteUrl}/_layouts/15/settings.aspx`],
    [t('siteContents'),  `${siteUrl}/_layouts/15/viewlsts.aspx`],
    ['REST /_api/web',   `${siteUrl}/_api/web`],
  ].map(([label, href]) =>
    `<a class="link-item" href="${href}" target="_blank" rel="noreferrer">
       <span>${label}</span>${icon('externalLink', 12)}
     </a>`
  ).join('');
}

function setStatus(dotState, msg, u) {
  dot.className       = `status-dot ${dotState}`;
  titleEl.textContent = msg;
  urlEl.textContent   = u;
}

btn.addEventListener('click', async () => {
  const tab = await getCurrentTab();
  if (tab?.id) {
    await chrome.sidePanel.open({ tabId: tab.id });
    window.close();
  }
});

init();
