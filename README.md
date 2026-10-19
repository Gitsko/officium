# Officium

Outlook add-in (new Outlook, Outlook on the web) that summarizes calendar
appointments per customer and period. Plan: `tasks/plan.md`, tasks: `tasks/todo.md`.

## Development (WSL2 + Outlook on Windows)

The dev server runs in WSL2; Outlook runs on Windows and reaches it via
`https://localhost:3000` (WSL localhost forwarding).

1. Install dependencies: `npm install`
2. Create the dev certificate (asks for `sudo` once, adds the CA to the WSL trust store):
   `npx office-addin-dev-certs install`
3. Trust the same CA on Windows (CurrentUser root store, Windows shows a confirmation dialog):
   `bash scripts/wsl-trust-dev-ca.sh`
4. Start the dev server: `npm run dev-server`
5. Check from Windows: open `https://localhost:3000/taskpane.html` in Edge - no certificate warning.
6. Sideload: `npm start` (signs in to Microsoft 365 and installs `manifest.json`), or see
   [Sideload add-ins with the unified manifest](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest).
   Stop and uninstall with `npm stop`.

New Outlook has no item-less calendar surface for add-ins: open an appointment
(or a mail) and use the **Officium** button in the ribbon or the **Apps** menu.

## Checks

`npm test`, `npm run typecheck`, `npm run lint`, `npm run validate`, `npm run build`
