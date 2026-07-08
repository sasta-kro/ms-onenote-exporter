# MS OneNote Scraper

## What This Tool Does

MS OneNote Scraper exports Microsoft OneNote notebooks into local files.

It is made for class notebooks and Teams notebooks that are already readable in
a school or work Microsoft account. The common problem is simple: OneNote can be
viewed in Teams, but there is no clean bulk download button. This tool fills
that gap.

The main output is HTML. HTML keeps more of the original OneNote page structure
than plain text. If Pandoc is installed, the same pages can also be converted to
Markdown, TXT, or RTF.

This is not an admin backup tool. It does not download every notebook in an
organization. It only exports notebooks visible to the signed-in Microsoft
account.

## What It Can Export

- Teams/Class Notebook pages stored in SharePoint.
- Notebooks visible from the signed-in OneNote account.
- Sections, section groups, and pages.
- HTML files for every page.
- Optional Markdown, TXT, and RTF copies.
- A `manifest.json` file for each exported notebook.

## Quick Start

Run these commands inside this project folder:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
cp .env.example .env
```

Open `.env` and set the Microsoft Entra application/client ID:

```text
ONENOTE_CLIENT_ID=paste-application-client-id-here
ONENOTE_TENANT_ID=organizations
```

To find the client ID, open [Microsoft Entra admin center](https://entra.microsoft.com/),
then go to `Identity -> Applications -> App registrations`. Open the app
registration for this exporter and copy `Application (client) ID` from the
Overview page.

If no app registration exists yet, create one first. Microsoft has the official
guide here: [Register an application with the Microsoft identity platform](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app).

The client ID is not a student ID, email address, tenant ID, or SharePoint ID.
Do not paste `Directory (tenant) ID` into `ONENOTE_CLIENT_ID`.

After `.env` is ready, list notebooks visible from the signed-in OneNote
account:

```bash
python main.py --list
```

On the first run, Microsoft prints a browser login link and a device code in the
terminal. Open the link, paste the code shown in the terminal, and finish the
Microsoft login page. The code is not in Teams or OneNote.

## Example Flow: Export a Teams Class Notebook

Most Teams class notebooks are stored in SharePoint. In that case, `python
main.py --list` may show nothing even when the notebook is visible in Teams.
That is normal.

Flow:

1. Open the class notebook in Teams.
2. Press the globe/open-in-browser button.
3. Copy the browser URL.
4. Run this command:

```bash
python main.py --site-url "PASTE_TEAMS_OR_ONENOTE_BROWSER_LINK"
```

The app then prints two helper links. Open both links in a browser already
signed in with the same school or work account.

Each helper page shows XML text with one GUID. Paste the whole page text back
into the terminal. The app extracts the GUID automatically.

Example terminal shape:

```text
>>>>> Step 1 (SITE_GUID): open this in the signed-in browser

╭───────────────────────────────────────────────────────────────╮
│ https://school.sharepoint.com/sites/Section_.../_api/site/id  │
╰───────────────────────────────────────────────────────────────╯

>>>>> Paste SITE_GUID page text
╭────────────────────────────────────────────────────────────╮
│ After pasting the XML text, press ENTER twice to continue. │
╰────────────────────────────────────────────────────────────╯
>
```

The pasted text can be the full browser page text:

```xml
This XML file does not appear to have any style information associated with it.
<d:Id m:type="Edm.Guid">80a26a44-cf5b-42b2-bf61-c3a021fa18c7</d:Id>
```

or only the XML line:

```xml
<d:Id m:type="Edm.Guid">80a26a44-cf5b-42b2-bf61-c3a021fa18c7</d:Id>
```

After both GUID values are pasted, the app lists notebooks in that SharePoint
site. If only one notebook exists, it downloads that notebook automatically.

To export Markdown instead of only HTML, put `--formats md` anywhere in the
command:

```bash
python main.py --formats md --site-url "PASTE_TEAMS_OR_ONENOTE_BROWSER_LINK"
```

For TXT or RTF:

```bash
python main.py --formats txt --site-url "PASTE_TEAMS_OR_ONENOTE_BROWSER_LINK"
python main.py --formats rtf --site-url "PASTE_TEAMS_OR_ONENOTE_BROWSER_LINK"
```

To list notebooks without downloading, add `--list`:

```bash
python main.py --site-url "PASTE_TEAMS_OR_ONENOTE_BROWSER_LINK" --list
```

## Output Folder

Exports go into `onenote_export/` by default:

```text
onenote_export/
  2026-1 BAD 542 Notebook/
    manifest.json
    _Content Library_03 Reverse Proxy/
      [CW] Reverse Proxy-c70597d6e4.html
      Reverse Proxy-c70597d6e4.html
```

Each notebook gets one folder. Section groups and sections become folders
inside the notebook folder. Each page filename has a short ID suffix to avoid
name collisions.

To choose another output folder:

```bash
python main.py --out ~/OneNoteExport --notebook "2026-1 BAD 542 Notebook"
```

## Common Commands

List notebooks from the signed-in OneNote account:

```bash
python main.py --list
```

Export one notebook from the signed-in OneNote account:

```bash
python main.py --notebook "2026-1 BAD 542 Notebook"
```

Export from a Teams/Class Notebook browser link:

```bash
python main.py --site-url "PASTE_TEAMS_OR_ONENOTE_BROWSER_LINK"
```

Export HTML plus Markdown and TXT:

```bash
python main.py --formats md,txt --notebook "2026-1 BAD 542 Notebook"
```

Keep image links in converted Markdown/TXT/RTF files:

```bash
python main.py --formats md,txt --include-image-links --notebook "2026-1 BAD 542 Notebook"
```

List notebooks from a resolved SharePoint site ID:

```bash
python main.py --site-id "school.sharepoint.com,siteCollectionGuid,webGuid" --list
```

Export from a resolved SharePoint site ID:

```bash
python main.py --site-id "school.sharepoint.com,siteCollectionGuid,webGuid" --notebook "Notebook Name"
```

## Formats

HTML is always exported.

Accepted `--formats` values:

- `html`: accepted, but no extra file is made because HTML already exists.
- `md`: Markdown copy.
- `txt`: plain text copy.
- `rtf`: rich text copy.
- `md,txt`: comma-separated values also work.

Markdown, TXT, and RTF conversion needs Pandoc:

```bash
brew install pandoc
```

Converted files are cleaned before Pandoc runs. The cleaner removes common
OneNote layout wrappers, raw span noise, and strange placeholder characters.
Image URLs are omitted from converted text formats by default.

Some OneNote pages are screenshots or pasted images. Markdown and TXT cannot
extract text from images. OCR is not included.

## Portable `.env` Setup

The app reads `.env` automatically. CLI flags override `.env`, and real shell
environment variables override `.env` too.

Common `.env` values:

```text
ONENOTE_CLIENT_ID=paste-application-client-id-here
ONENOTE_TENANT_ID=organizations
ONENOTE_SITE_ID=school.sharepoint.com,siteCollectionGuid,webGuid
ONENOTE_OUT=onenote_export
ONENOTE_FORMATS=
ONENOTE_NOTEBOOK=2026-1 BAD 542 Notebook
ONENOTE_TOKEN_CACHE=.msal_token_cache.json
```

### `ONENOTE_CLIENT_ID`

Required. This comes from the Microsoft Entra app registration.

Path:

```text
Microsoft Entra admin center
Identity -> Applications -> App registrations
App Overview -> Application (client) ID
```

Portal: [https://entra.microsoft.com/](https://entra.microsoft.com/)

Official guide: [Register an application with the Microsoft identity platform](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app)

### `ONENOTE_TENANT_ID`

Default:

```text
ONENOTE_TENANT_ID=organizations
```

`organizations` means Microsoft work or school accounts. This is the normal
choice for university and company notebooks.

Other valid values:

- A Directory tenant GUID.
- A tenant domain, for example `school.edu`.
- `common`, for personal Microsoft accounts and work/school accounts.
- `consumers`, for personal Microsoft accounts only.

### `ONENOTE_SITE_ID`

Optional, but useful after a Teams/Class Notebook link has already been
resolved once.

To get it:

```bash
python main.py --site-url "PASTE_TEAMS_OR_ONENOTE_BROWSER_LINK"
```

After the two GUID helper pages are pasted, the app prints a resolved site ID.
Save that value in `.env` to skip the helper flow later.

Shape:

```text
school.sharepoint.com,siteCollectionGuid,webGuid
```

### Other `.env` Values

`ONENOTE_OUT` sets the export folder.

`ONENOTE_FORMATS` sets extra formats, for example `md`, `txt`, `rtf`, or
`md,txt`.

`ONENOTE_NOTEBOOK` filters notebook names. It can make `python main.py` export
one known notebook without typing the name each time.

`ONENOTE_TOKEN_CACHE` sets the Microsoft login token cache file. Keep this file
private.

## Requirements

Python dependencies live in `requirements.txt`:

```bash
python -m pip install -r requirements.txt
```

Packages:

- `msal`: Microsoft Authentication Library for device-code login.
- `requests`: HTTP client for Microsoft Graph calls.

Optional system dependency:

```bash
brew install pandoc
```

Pandoc is only needed for Markdown, TXT, and RTF conversion.

## Microsoft App Setup

The Microsoft Entra app registration needs:

- Public client/device-code flow enabled.
- Delegated Microsoft Graph permission: `Notes.Read.All`.
- No client secret.

The app uses device-code login. It does not need a password inside `.env`.

## Token Cache

The app stores Microsoft login tokens in:

```text
.msal_token_cache.json
```

That file is ignored by git. Treat it as private.

Delete it to force a fresh Microsoft login:

```bash
rm .msal_token_cache.json
```

## Troubleshooting

If `--list` shows no notebooks, the notebook may be stored under a Teams or
SharePoint site. Try the Teams/Class Notebook flow with `--site-url`.

If Microsoft returns `AADSTS50059`, keep this value in `.env`:

```text
ONENOTE_TENANT_ID=organizations
```

If Microsoft says admin approval is required, the Microsoft tenant may block
user consent for the requested permission. The normal setup only asks for
`Notes.Read.All`.

If `.html` files export but `.md`, `.txt`, or `.rtf` files do not appear,
install Pandoc or leave `ONENOTE_FORMATS` blank for HTML-only export.

If the same GUID is pasted for both SharePoint helper steps, open the second
helper link again. Step 1 and Step 2 must return different GUID values.

## Development

Run tests:

```bash
python -m unittest discover -s tests
```
