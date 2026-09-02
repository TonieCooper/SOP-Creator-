# SOP Generator

A browser-based desktop-procedure builder with screenshot support, JSON project save/load, local autosave, Word `.docx` export, and print/PDF output.

## Important business-data note
Do not publish confidential, internal, customer, financial, or regulated information to a public repository or public GitHub Pages site. Confirm your organization's approved source-code hosting and third-party service requirements before business use.

## Run locally
1. Install Node.js 20 or later.
2. In this folder run `npm install`.
3. Run `npm run dev`.
4. Open the local address shown by Vite.

## Deploy with GitHub Pages
1. Create a repository and upload every file and folder in this project.
2. Use `main` as the default branch.
3. In repository **Settings > Pages**, set **Source** to **GitHub Actions**.
4. Push a commit to `main`, or run the workflow manually from **Actions**.

The included workflow builds the Vite app and deploys `dist`.

## Data handling
All SOP form data and screenshots stay in the browser unless the user downloads a project or document. Autosave uses localStorage. There is no server, sign-in, analytics, API key, or external database.

## AI attribution
Initial application code was generated with Microsoft Copilot and should be reviewed, tested, and approved before production use.
