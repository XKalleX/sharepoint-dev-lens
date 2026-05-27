# Privacy Policy – SharePoint Dev Lens

_Last updated: 2026-05-27_

## Summary

SharePoint Dev Lens does not collect, store or transmit any user data or SharePoint content to external servers.

## Data Processing

### What the extension reads
SharePoint Dev Lens reads SharePoint site data — list names, GUIDs, field metadata, and site configuration — from the SharePoint REST API (`/_api/`) of the site you are actively visiting. This data is only used to display information in the extension's UI.

### Where data goes
**Nowhere outside your browser.** All processing happens locally. No data is sent to any external server operated by the extension developer or any third party.

### Storage
The extension may store UI preferences (e.g. selected tab, filter values) in `chrome.storage.local`. This data never leaves your device.

### Authentication
The extension does not implement its own authentication. It uses the browser's existing authenticated session (cookies) for requests to the SharePoint domain you are visiting. No tokens are stored by the extension.

### Telemetry
None. The extension has no analytics, crash reporting or usage tracking.

## Permissions Used

| Permission | Reason |
|------------|--------|
| `activeTab` | Read the current tab URL to detect SharePoint |
| `storage` | Store UI preferences locally |
| `scripting` | Inject content script into SharePoint pages |
| `sidePanel` | Open the inspector side panel |
| `host_permissions: *.sharepoint.com` | Allow REST API calls on SharePoint pages |

## Contact

For questions about privacy, open an issue at:  
https://github.com/wagnerp/sharepoint-dev-lens/issues
