# Codex Handoff Notes

This repository is an SPFx permission audit web part. Read this file first when continuing work in a fresh Codex session.

## Build And Verify

- Working directory: repository root.
- Install: `npm install`
- Build/package: `npm run build`
- Dev server: `npm run start`
- Node requirement: `>=22.14.0 <23.0.0`

Always run `npm run build` before committing meaningful changes.

## Architecture

- `PermissionAuditWebPart.ts` creates the PnPjs SPFI instance, Graph service, SharePoint service, theme variables, and the tenant/user-scoped grouped-view preference key.
- `AuditPage/AuditPage.tsx` owns the UI state, DetailsList projection, CSV export, grouped/flat toggle, principal search, and details panel HTML generation.
- `SharePointPermissionAuditService.ts` is the only place SharePoint calls should be added.
- `GraphPermissionAuditService.ts` is the only place Graph calls should be added.
- `IPermissionAuditItem.ts` is the common tree item model used by site/list/item/principal nodes.
- `src/common/controls` contains reusable controls available to any component.

## Product Behavior To Preserve

- All UI text is localized through `PermissionAuditWebPartStrings`.
- Theme colors should come from SPFx/Fluent theme variables, not fixed palettes.
- Grouped DetailsList view and flat DetailsList view are projections of the same loaded audit result.
- List item unique-permission scanning must not require list-level unique permissions.
- Hidden list scanning is controlled by `Include hidden lists`.
- Long scans update progressively through the `onAuditUpdated` callback and keep the spinner active until all work completes.
- CSV export must be ungrouped and include all fields, including fields not visible in the DetailsList.

## Git Hygiene

- Generated audit CSV files under `assets` are user output and should not be committed.
- Build output folders are ignored: `dist`, `lib`, `lib-commonjs`, `release`, `temp`, etc.
