# Vehicle Intake Form

Static form for customers to submit vehicles they're shipping. Submissions are
collected in a Google Sheet.

**Live:** https://carform-sigma.vercel.app

## How it works

`index.html` → Google Apps Script web app → Google Sheet, tab `Submissions`,
one row per VIN.

## Files

- `index.html` — the form. Single self-contained file: CSS, JS and the logo are
  all inlined, so it can be hosted anywhere static.
- `apps-script/Code.gs` — the submission endpoint. Lives in the Sheet's Apps
  Script editor, deployed as a web app (Execute as: Me, Access: Anyone).
- `KNL Logo.png` — source logo, not used at runtime.

## The sheet

Columns A–I are written by the form:

`Timestamp | Customer ID | Company | VIN | Brand | Model | Engine Type | Engine Capacity | Year`

Columns J–L are yours: `Status | Notes | Date received`. Submissions never
touch them.

Customers are identified by an anonymous ID stored in their browser, plus the
company name they type. Same customer on a different device = a new ID.

## Deploying

Push to `main` — Vercel redeploys automatically.

To change the endpoint, edit `Code.gs` in the Apps Script editor and deploy a
**new version of the existing deployment** (a *new deployment* changes the URL,
which you'd then have to paste into `ENDPOINT` in `index.html`).

## Notes

- The link is public; anyone who has it can submit.
- The EN/中文 toggle changes the interface only. Engine type is always stored in
  English.
- Download the sheet as `.xlsx` via File → Download when you need a file to work
  with; the Sheet is the master copy.
