# Telegram Expense Tracker

A Telegram Mini App for logging personal spending and income into a Google
Spreadsheet you own. No server, no database, no sign-up: the app is three static
files, the backend is a Google Apps Script bound to your own spreadsheet, and
your data never leaves your Google account.

Available in English, Russian and Vietnamese.

## How it works

```
  Telegram client
        │  opens the Mini App
        ▼
  index.html + app.js + styles.css          static hosting (GitLab/GitHub Pages)
        │  GET  ?action=addRecord&…         plain query params, no CORS preflight
        ▼
  Google Apps Script web app                apps-script/Code.gs
        │
        ▼
  Your Google Spreadsheet                   Records, Categories, Accounts, …
```

The front-end keeps categories, settings and UI state in `localStorage` (every
key namespaced by `APP_ID`, so several bots can share one browser origin) and
reconciles them with the spreadsheet on start-up. Records are written
optimistically: the form clears immediately and the write finishes in the
background, which keeps entering several purchases in a row fast.

## Features

- Outflow and inflow records with category, subcategory, account, transaction
  type and note — every optional field can be switched off.
- Categories, subcategories, accounts and transaction types are editable in the
  app: add, rename, reorder by drag, hide. Changes propagate to the spreadsheet.
- Monthly and yearly summaries: totals, balance, savings rate, share per
  category, comparison with the previous period and a six-period trend chart.
- History of the latest records, outflows or inflows, grouped by day and paged
  thirty at a time. A record can be edited in the record form or deleted from
  the spreadsheet, deletion asking for confirmation first.
- Duplicate guard for same-day records, so a double tap does not log twice.
- Follows the Telegram theme (light and dark) and the client's back button.

## Setup

**1. Create the spreadsheet.** Any empty Google Spreadsheet will do — the script
creates the sheets it needs on first write.

**2. Deploy the Apps Script.**

1. <https://script.google.com/home> → *New project*
2. Paste the contents of [`apps-script/Code.gs`](apps-script/Code.gs)
3. *Deploy* → *New deployment* → *Web app*, executing as **Me**, access
   **Anyone**
4. Copy the `/exec` URL

The `/exec` URL is the only credential this app has: anyone holding it can write
to the spreadsheet. Treat it accordingly.

**3. Configure the front-end.** Copy `config.js.example` to `config.js` and fill
in the URL:

```js
const GAS_URL = 'https://script.google.com/macros/s/…/exec';
const APP_ID  = 'rasshody'; // localStorage prefix — unique per bot deployment
```

`config.js` is gitignored; it belongs to a deployment, not to the source.

**4. Publish the app** on any static host (GitLab Pages, GitHub Pages, …) and
point a Telegram bot at it: [@BotFather](https://t.me/BotFather) →
*Bot Settings* → *Menu Button* → the published URL.

**5. Connect the spreadsheet** on first launch: the app asks for a name, which
optional fields to use, and the spreadsheet URL or ID.

### Updating the Apps Script

Pasting new code into the editor is not enough — the `/exec` URL keeps serving
the last *deployed* version. After editing: *Deploy* → *Manage deployments* →
edit the active deployment → *New version*.

To check which version is live:

```
<your /exec URL>?action=diag&spreadsheetId=<spreadsheet id>
```

It reports `SCRIPT_VERSION`, the configured columns, the sheet headers and how
the last few rows resolve to a date — enough to tell a stale deployment apart
from a spreadsheet whose columns are not what the script expects.

## Spreadsheet layout

`RECORD_COLUMNS` in `apps-script/Code.gs` defines the columns of a **newly
created** `Records` sheet: reorder the list to reorder them, comment entries in
and out to drop the optional ones.

A sheet that already exists keeps the order it was created with. Reads and
writes resolve each field from the header row, matching it against
`COLUMN_DISPLAY_NAMES`, so one script can serve several bots whose spreadsheets
are laid out differently, and columns of your own added to the right are left
alone.

When the headers cannot be matched — they are in another language, or you
renamed them — it falls back to the `RECORD_COLUMNS` order, and then that order
*must* match the sheet. Keep the block matching the spreadsheet of that
deployment and leave it alone when pasting an updated script, or translate
`COLUMN_DISPLAY_NAMES` to your headers to have them matched instead.
`?action=diag` reports which mapping is in effect.

`date` holds the purchase date and drives every summary; `timestamp` only
records when the row was entered.

Editing and deleting from the History tab address a row by its number, which
shifts whenever rows are added or removed elsewhere. Each such call therefore
carries the signature the row had when it was read — its timestamp, amount,
category and type — and the script refuses the write when the row no longer
matches, so the app reloads the list instead of touching the wrong record.

`Categories`, `Accounts`, `Transaction Types`, `Subcategories` and `Users` are
created and maintained by the script.

## Project layout

```
index.html            markup and inline handlers
app.js                application logic
styles.css            styling, driven by Telegram's theme variables
config.js             deployment config — gitignored, see config.js.example
lang/{en,ru,vi}.json  interface strings
apps-script/Code.gs   Google Apps Script backend
test/                 self-checks, run with node
```

No build step and no package manager: what is in the repository is what the
browser loads. Bootstrap and the Telegram SDK come from a CDN.

## Adding a language

1. Copy `lang/en.json` to `lang/<code>.json` and translate the values.
2. Add a button to the language picker in `index.html` and to the list in
   `updateLangButtons()` in `app.js`.

Strings are looked up by key with `t('key')`; a missing key falls back to the
key itself, so a partial translation degrades instead of breaking.

## Tests

```sh
node test/date-resolution.test.js   # summaries group by purchase date
node test/column-mapping.test.js    # fields resolve to the right columns
node test/record-editing.test.js    # an edit never lands on the wrong row
node test/wiring.test.js            # markup and app.js still agree
```

Plain `node`, no framework. The wiring check is what catches a renamed handler
or a missing element after refactoring the front-end.

## Licence

MIT — see [LICENSE](LICENSE).
