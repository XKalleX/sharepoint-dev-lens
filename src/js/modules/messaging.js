export async function getCurrentTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab ?? null;
}

export async function sendToTab(tabId, message) {
  try {
    return await chrome.tabs.sendMessage(tabId, message);
  } catch (err) {
    const msg = err.message || '';
    // Content script not loaded – try to inject it programmatically
    if (msg.includes('Receiving end does not exist') || msg.includes('Could not establish connection')) {
      try {
        await chrome.scripting.executeScript({ target: { tabId }, files: ['content.js'] });
        // Wait briefly for the script to register its message listener
        await new Promise(r => setTimeout(r, 400));
        return await chrome.tabs.sendMessage(tabId, message);
      } catch (injectErr) {
        return { success: false, error: injectErr.message || 'Content script konnte nicht geladen werden.' };
      }
    }
    return { success: false, error: msg || 'Verbindung zur Seite fehlgeschlagen.' };
  }
}
