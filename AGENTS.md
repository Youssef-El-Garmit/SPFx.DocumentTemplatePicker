# AGENTS.md

## Cursor Cloud specific instructions

This is a **SharePoint Framework (SPFx) 1.21.1** web part project ("Document Template Picker"). It is a pure client-side component — no backend, no database, no Docker.

### Node.js requirement

Requires Node.js `>=22.14.0 <23.0.0` (enforced via `engines` in `package.json`). The VM ships with a compatible Node 22.x via nvm.

### Key commands

| Task | Command |
|------|---------|
| Install deps | `npm install` |
| Lint | `npx eslint src/ --ext .ts,.tsx` |
| Build (debug) | `npx gulp bundle` |
| Build (production) | `npx gulp bundle --ship` |
| Package solution | `npx gulp package-solution --ship` |
| Run tests | `npx gulp test` |
| Dev server | `npx gulp serve --nobrowser` |
| Trust dev cert | `npx gulp trust-dev-cert` |

### Dev server notes

- `gulp serve --nobrowser` starts a local HTTPS server on `https://localhost:4321` with LiveReload on port 35729.
- On Linux, `gulp trust-dev-cert` does not auto-trust the certificate in system stores. Browsers will show a certificate warning that must be bypassed manually.
- The dev server serves the compiled web part bundle and manifests at `https://localhost:4321/temp/build/manifests.js`.
- Full end-to-end testing requires a Microsoft 365 / SharePoint Online tenant. The local workbench (`/temp/workbench.html`) has been deprecated in SPFx 1.21; use the hosted workbench (`https://{tenantDomain}/_layouts/workbench.aspx?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/build/manifests.js`).

### Lint behavior

ESLint is configured with `@microsoft/eslint-config-spfx`. The existing codebase has 31 lint warnings (mostly `no-explicit-any` and `no-void`) and 0 errors. The `gulp bundle` and `gulp test` tasks also run lint as a subtask; lint warnings do not fail the build.
