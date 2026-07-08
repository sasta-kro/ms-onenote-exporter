# MS OneNote Scraper

Export Microsoft OneNote notebooks that your work/school account can already
read into local files.

The app saves every page as `.html`. HTML is the format Microsoft Graph returns
and usually keeps more OneNote structure than plain text. If Pandoc is installed,
the app can also create `.md`, `.txt`, or `.rtf` copies.

This is not an organization-wide backup tool. It only exports notebooks visible
to the signed-in Microsoft account.

## What It Does

- Uses Microsoft device-code login in the terminal.
- Lists notebooks from your own OneNote area or a resolved SharePoint site.
- Supports Teams/Class Notebook links through a low-permission site ID helper.
- Walks notebooks, section groups, sections, and pages.
- Downloads pages as local HTML files.
- Writes one `manifest.json` inside each exported notebook folder.
- Optionally converts cleaned OneNote HTML pages with Pandoc.

## Quick Start

From this project directory:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
cp .env.example .env
```

Edit `.env` and set your Microsoft Entra application/client ID:

```text
ONENOTE_CLIENT_ID=paste-your-application-client-id-here
ONENOTE_TENANT_ID=organizations
```

Then try listing notebooks from your own OneNote area:

```bash
python main.py --list
```

On first run, Microsoft will ask you to open a device-login URL and enter the
code printed in the terminal. The code comes from this app's terminal output,
not from Teams or OneNote.

## Teams/Class Notebook Flow

Most class notebooks in Teams are stored in SharePoint. `/me` may show no
notebooks even when you can see the notebook in Teams.

Use this flow:

1. In Teams, open the class notebook.
2. Use the globe/open-in-browser button.
3. Copy the browser URL.
4. Run:

```bash
python main.py --site-url "PASTE_TEAMS_OR_ONENOTE_BROWSER_LINK"
```

The app prints something like this:

```text
>>>>> Detected notebook storage site (for checking only)
This is the Teams/SharePoint site that stores the notebook file. You usually do not need to open it.
╭────────────────────────────────────────────────────╮
│ https://yourtenant.sharepoint.com/sites/Section_... │
╰────────────────────────────────────────────────────╯

>>>>> Step 1 (SITE_GUID): open this in your signed-in browser

╭─────────────────────────────────────────────────────────────────╮
│ https://yourtenant.sharepoint.com/sites/Section_.../_api/site/id │
╰─────────────────────────────────────────────────────────────────╯

>>>>> Paste SITE_GUID page text
╭────────────────────────────────────────────────────────────╮
│ After pasting the XML text, press ENTER twice to continue. │
╰────────────────────────────────────────────────────────────╯
>

>>>>> Step 2 (WEB_GUID): open this in your signed-in browser

╭────────────────────────────────────────────────────────────────╮
│ https://yourtenant.sharepoint.com/sites/Section_.../_api/web/id │
╰────────────────────────────────────────────────────────────────╯

>>>>> Paste WEB_GUID page text
╭────────────────────────────────────────────────────────────╮
│ After pasting the XML text, press ENTER twice to continue. │
╰────────────────────────────────────────────────────────────╯
>

>>>>> Reusable command to see Notebooks in the link

    python main.py --site-id "yourtenant.sharepoint.com,SITE_GUID,WEB_GUID" --list

>>>>> Available notebooks
╭────────────────────────────╮
│ 1. 2026-1 BAD 542 Notebook │
╰────────────────────────────╯

>>>>> Auto-downloading the only notebook as HTML
╭─────────────────────────╮
│ 2026-1 BAD 542 Notebook │
╰─────────────────────────╯

>>>>> Optional Markdown/RTF commands

    python main.py --site-id "yourtenant.sharepoint.com,SITE_GUID,WEB_GUID" --notebook "2026-1 BAD 542 Notebook" --formats md
    python main.py --site-id "yourtenant.sharepoint.com,SITE_GUID,WEB_GUID" --notebook "2026-1 BAD 542 Notebook" --formats rtf
```

Open the Step 1 and Step 2 links in the browser where you are already signed
into your school account. Each page shows one GUID. You can paste the whole
page text, including the browser's "This XML file does not appear..." message,
or just the `<d:Id>...</d:Id>` XML line. The app extracts the GUID for you.

After both values are pasted, the app lists notebooks automatically. If exactly
one notebook is available, it downloads that notebook as HTML immediately. It
also prints optional Markdown and RTF commands for later.

If you only want to list notebooks without downloading, add `--list`:

```bash
python main.py --site-url "PASTE_TEAMS_OR_ONENOTE_BROWSER_LINK" --list
```

The app also prints a reusable `--site-id` command so you can skip the paste
step next time. If the app is running in a non-interactive shell, it prints the
manual helper links and command instead of prompting.

Reusable site ID example:

```bash
python main.py --site-id "assumptionuniversity.sharepoint.com,80a26a44-cf5b-42b2-bf61-c3a021fa18c7,5dbbcfdd-641d-42ed-b89a-2cb2451897ef" --list
```

When a site has more than one notebook, copy the exact command printed under
`To download one notebook`:

```text
>>>>> To download one notebook

    python main.py --site-id "..." --notebook "2026-1 BAD 542 Notebook"
```

Run that command to export the chosen notebook.

## Output

By default, exports go into `onenote_export/`:

```text
onenote_export/
  2026-1 BAD 542 Notebook/
    manifest.json
    _Content Library_03 Reverse Proxy/
      [CW] Reverse Proxy-c70597d6e4.html
      Reverse Proxy-c70597d6e4.html
```

Each notebook gets its own folder. Section groups and sections become folders
inside the notebook folder. The short suffix in each page filename comes from
the OneNote page ID and helps avoid filename collisions.

Use `--out` to choose another output root:

```bash
python main.py --out ~/OneNoteExport --notebook "2026-1 BAD 542 Notebook"
```

## Common Commands

List notebooks from your own OneNote area:

```bash
python main.py --list
```

Export one notebook:

```bash
python main.py --notebook "2026-1 BAD 542 Notebook"
```

Export to a specific folder:

```bash
python main.py --out ~/OneNoteExport --notebook "2026-1 BAD 542 Notebook"
```

Export HTML plus Markdown and TXT:

```bash
python main.py --formats md,txt --notebook "2026-1 BAD 542 Notebook"
```

Converted `.md`, `.txt`, and `.rtf` files are cleaned before Pandoc runs:
OneNote layout wrappers are stripped, raw `<span>` noise is removed, and strange
OneNote placeholder characters are turned into line breaks. Image URLs are
omitted from converted text formats by default.

Keep image links in converted files:

```bash
python main.py --formats md,txt --include-image-links --notebook "2026-1 BAD 542 Notebook"
```

List notebooks from a resolved SharePoint site:

```bash
python main.py --site-id "yourtenant.sharepoint.com,siteCollectionGuid,webGuid" --list
```

Export from a resolved SharePoint site:

```bash
python main.py --site-id "yourtenant.sharepoint.com,siteCollectionGuid,webGuid" --notebook "Notebook Name"
```

## Portable `.env` Setup

The app reads `.env` automatically. CLI flags override `.env`, and real shell
environment variables override `.env` too.

Useful `.env` values:

```text
ONENOTE_CLIENT_ID=paste-your-application-client-id-here
ONENOTE_TENANT_ID=organizations
ONENOTE_SITE_ID=yourtenant.sharepoint.com,siteCollectionGuid,webGuid
ONENOTE_OUT=onenote_export
ONENOTE_NOTEBOOK=2026-1 BAD 542 Notebook
ONENOTE_FORMATS=
```

Once `ONENOTE_SITE_ID` and `ONENOTE_NOTEBOOK` are set, you can usually run:

```bash
python main.py
```

## Requirements

Python dependencies live in `requirements.txt`:

```bash
python -m pip install -r requirements.txt
```

Packages:

- `msal`: Microsoft Authentication Library for device-code login.
- `requests`: HTTP client for Microsoft Graph calls.

Optional system dependency for Markdown/TXT/RTF conversion:

```bash
brew install pandoc
```

You do not need Pandoc if you only want `.html`.

Some OneNote pages are actually screenshots or pasted images. Markdown and TXT
cannot extract text from those images; they need OCR, which this app does not
perform yet.

## Microsoft App Setup

The app expects:

- A Microsoft Entra app registration.
- Public client/device-code flow enabled.
- Delegated Microsoft Graph permission: `Notes.Read.All`.
- The Application/client ID from the app registration.

You do not need a client secret.

## Token Cache

The app stores Microsoft login tokens in:

```text
.msal_token_cache.json
```

That file is ignored by git. Treat it as private. Delete it if you need to force
a fresh Microsoft login:

```bash
rm .msal_token_cache.json
```

## Troubleshooting

If `--list` shows no notebooks, make sure you are signed in with the same
account that can see the notebook in OneNote or Teams.

If a Teams notebook is missing from `/me`, use the Teams/Class Notebook flow with
`--site-url`.

If Microsoft says admin approval is required, your tenant may block user consent
for the requested permission. The normal setup only needs `Notes.Read.All`.

If Microsoft returns `AADSTS50059`, set this in `.env`:

```text
ONENOTE_TENANT_ID=organizations
```

If `.html` files export but `.md`, `.txt`, or `.rtf` files do not appear,
install Pandoc or leave `ONENOTE_FORMATS` blank for HTML-only export.

## Development

Run tests:

```bash
python -m unittest discover -s tests
```
