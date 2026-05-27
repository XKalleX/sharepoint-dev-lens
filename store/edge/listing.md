# Edge Add-ons Partner Center Listing – SharePoint Dev Lens

## Name
SharePoint Dev Lens

## Short Description (max 150 chars)
Inspect SharePoint lists, GUIDs, fields, permissions and developer URLs directly in Microsoft Edge.

## Description
(Same content as Chrome listing — adapt for Edge Partner Center UI)

SharePoint Dev Lens is a browser extension for SharePoint developers, consultants and Microsoft 365 administrators.

Features:
• All list/library GUIDs with one-click copy
• Field InternalNames, types, required and indexed status
• REST, PnPjs and Power Automate snippet generator
• Site metadata: Web ID, template, LCID, Hub Site status
• Developer quick links to /_layouts/15/ and REST endpoints
• Export as JSON, Markdown or CSV
• Governance warnings: large lists, missing indexes, unique permissions

Privacy: No data collection. Works entirely in your browser using your existing SharePoint session. No external servers, no telemetry.

Not affiliated with or endorsed by Microsoft Corporation.

## Category
Developer Tools / Productivity

## Privacy Policy URL
https://github.com/wagnerp/sharepoint-dev-lens/blob/main/docs/PRIVACY.md

## Testing Notes (for Edge reviewer)
To test this extension:
1. Navigate to any SharePoint Online site (e.g. https://contoso.sharepoint.com/sites/test)
2. Click the extension icon in the toolbar
3. The popup will detect the SharePoint site and show a button to open the inspector panel
4. Click "Open Inspector Panel" to open the side panel
5. The Lists tab shows all lists/libraries with their GUIDs
6. Click "Fields →" on any list to inspect its fields

The extension only activates on *.sharepoint.com URLs (declared in host_permissions).
It uses only the user's existing authenticated browser session — no login prompts.
