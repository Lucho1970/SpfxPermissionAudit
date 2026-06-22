# Permission Audit

SPFx web part for auditing SharePoint site, list/library, and list item permissions. It can expand SharePoint groups and Entra directory groups, show results in grouped or flat views, search loaded results by user/group, show detail HTML for selected rows, and export a flattened CSV.

## Project Stack

- SharePoint Framework 1.23
- React 17
- Fluent UI React v8
- PnPjs v4 for SharePoint calls
- Microsoft Graph client for directory group expansion
- Heft build pipeline

## Setup

Use Node.js 22.x, matching the `package.json` engine range.

```powershell
npm install
```

## Common Commands

```powershell
npm run start
npm run build
npm run clean
```

`npm run build` runs the production Heft test/build flow and packages the SPFx solution.

## Current Features

- Audit current site permissions.
- Optional group expansion for SharePoint groups and Entra-backed groups.
- Handles Entra group claim suffixes such as `_m` and `_o`.
- Optional list/library scan for unique permissions.
- Optional list item scan for unique permissions, independent of list-level unique permissions.
- Optional hidden-list inclusion.
- Progressive UI updates while list/item scans complete.
- Grouped/flat DetailsList toggle with tenant-and-user scoped local storage preference.
- Details panel for selected site/list/item/principal rows.
- User/group access search over the loaded audit result.
- CSV export of all loaded audit data in flattened format.

## Important Code Areas

- Web part entry point: `src/webparts/permissionAudit/PermissionAuditWebPart.ts`
- Main UI: `src/webparts/permissionAudit/AuditPage/AuditPage.tsx`
- Models: `src/webparts/permissionAudit/models/IPermissionAuditItem.ts`
- SharePoint service: `src/webparts/permissionAudit/services/SharePointPermissionAuditService.ts`
- Graph service: `src/webparts/permissionAudit/services/GraphPermissionAuditService.ts`
- Common controls: `src/common/controls`
- Localized strings: `src/webparts/permissionAudit/loc`

## Localization Rule

All user-facing text should go through the SPFx `mystrings` localization pattern:

- Add typings to `src/webparts/permissionAudit/loc/mystrings.d.ts`
- Add English values to `src/webparts/permissionAudit/loc/en-us.js`
- Import as `PermissionAuditWebPartStrings`

## Notes For Future Work

- Search currently runs against the already-loaded audit tree. To search group memberships, run the audit with `Expand groups` selected. To search list/item unique permissions, run with the relevant list or item audit options selected.
- CSV exports are generated user output and should not be committed.
- The details display accepts sanitized HTML produced by the web part code.
- SharePoint calls should remain centralized in `SharePointPermissionAuditService`.
- Graph calls should remain centralized in `GraphPermissionAuditService`.
- Keep reusable controls under `src/common/controls`, one control per folder.
