# MS OneNote Scraper

MS OneNote Scraper downloads readable Microsoft OneNote notebooks into local
files.

It is mainly made for AU students who can see shared class notebooks in Teams, but
cannot find a clean bulk download button.

## Main Use Case

Shared OneNote books can be opened in Teams. The hard part is saving all pages
locally.

This tool exports those pages into a normal folder on the computer.

## Output Files

Every page is saved as HTML.

HTML is the safest default because it keeps more of the original OneNote page
layout than plain text.

Optional extra formats:
- Markdown
- TXT

Markdown and TXT work without Pandoc. If Pandoc is installed, the app uses it
for cleaner conversion. Otherwise, the built-in text converter is used.

## Scope

What it can read:

- Teams/Class Notebook pages stored in SharePoint.
- Notebooks visible from the signed-in OneNote account.
- Sections, section groups, and pages.

What it creates:

- One folder per notebook.
- HTML files for every page.
- Optional Markdown and TXT copies.
- One `manifest.json` file per notebook.

This is not an admin backup tool. It does not download every notebook in an
organization. It only exports notebooks visible to the signed-in Microsoft
account.

___

## Quick Start

Run these commands inside this project folder:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
cp .env.example .env
```

For Assumption University students, `.env.example` already contains the shared
Microsoft app settings. No Entra app registration setup is needed.

After setup, list notebooks visible from the signed-in OneNote account:

```bash
python main.py --list
```

On the first run, Microsoft prints a browser login link and a device code in the
terminal. Open the link, paste the code shown in the terminal, and finish the
Microsoft login page. The code is not in Teams or OneNote.

___

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
>>>>> Step 1 (SITE_GUID): open this in the signed-in browser and copy the whole text.

╭───────────────────────────────────────────────────────────────╮
│ https://school.sharepoint.com/sites/Section_.../_api/site/id  │
╰───────────────────────────────────────────────────────────────╯

>>>>> Paste SITE_GUID page text
╭─────────────────────────────────────────────────────────────────────────────────────╮
│ After copying the whole XML text and pasting it here, press Return/Enter to continue. │
╰─────────────────────────────────────────────────────────────────────────────────────╯
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

## How It Works

The app uses Microsoft device-code login. The terminal prints a Microsoft login
URL and a short code. The browser handles the real Microsoft sign-in. The app
receives an access token after login succeeds.

For normal OneNote notebooks, the app asks Microsoft Graph for notebooks under
the signed-in account.

For Teams/Class Notebook links, the notebook usually lives in a SharePoint site.
The pasted Teams browser link contains the SharePoint site address, but
Microsoft Graph needs a resolved site ID. The app helps collect two SharePoint
GUID values from `_api/site/id` and `_api/web/id`, then builds the Graph site
ID from those values.

After the notebook is found, the app walks through notebooks, section groups,
sections, and pages. Each page is downloaded as HTML. Optional formats are made
from cleaned HTML. Pandoc is used when installed. If Pandoc is missing, the app
falls back to its built-in Markdown/TXT converter.

## Tech Used

- Python 3.
- `msal` for Microsoft device-code login and token caching.
- `requests` for Microsoft Graph and SharePoint-related HTTP calls.
- Microsoft Graph OneNote APIs for notebook, section, and page export.
- SharePoint `_api/site/id` and `_api/web/id` helper endpoints for Teams
  notebook site resolution.
- Optional Pandoc support for cleaner Markdown and TXT conversion.
- `unittest` for the local test suite.

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

Keep image links in converted Markdown/TXT files:

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
- `md,txt`: comma-separated values also work.

Markdown and TXT conversion works without Pandoc. Pandoc is optional:

```bash
brew install pandoc
```

Converted files are cleaned first. The cleaner removes common OneNote layout
wrappers, raw span noise, and strange placeholder characters. Image URLs are
omitted from converted text formats by default.

Some OneNote pages are screenshots or pasted images. Markdown and TXT cannot
extract text from images. OCR is not included.

## Optional `.env` Settings

The app reads `.env` automatically. CLI flags override `.env`, and real shell
environment variables override `.env` too.

For Assumption University students, the default auth values in `.env.example`
are already set. The values below are optional settings for repeated use.

```text
ONENOTE_SITE_ID=school.sharepoint.com,siteCollectionGuid,webGuid
ONENOTE_OUT=onenote_export
ONENOTE_FORMATS=
ONENOTE_NOTEBOOK=2026-1 BAD 542 Notebook
ONENOTE_TOKEN_CACHE=.msal_token_cache.json
```

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

`ONENOTE_FORMATS` sets extra formats, for example `md`, `txt`, or `md,txt`.

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

Pandoc is not required. It is only used when installed for cleaner Markdown and
TXT conversion.

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

If Microsoft returns `AADSTS50059`, Microsoft did not receive the school
tenant value during login. Keep this Assumption University value in `.env`:

```text
ONENOTE_TENANT_ID=c1f3dc23-b7f8-48d3-9b5d-2b12f158f01f
```

Then delete `.msal_token_cache.json` and run the command again.

If Microsoft says admin approval is required, the Microsoft tenant may block
user consent for the requested permission. The normal setup only asks for
`Notes.Read.All`.

If `.html` files export but `.md` or `.txt` files do not appear, check
`ONENOTE_FORMATS` or the `--formats` value.

If the same GUID is pasted for both SharePoint helper steps, open the second
helper link again. Step 1 and Step 2 must return different GUID values.

## For Use Outside Assumption University

The preset `ONENOTE_CLIENT_ID` is meant for Assumption University students. For
another school, company, or personal Microsoft tenant, create a separate
Microsoft Entra app registration and replace `ONENOTE_CLIENT_ID` in `.env`.

The app registration needs:

- Public client/device-code flow enabled.
- Delegated Microsoft Graph permission: `Notes.Read.All`.
- No client secret.

Microsoft guide:
[Register an application with the Microsoft identity platform](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app)

## Development

Run tests:

```bash
python -m unittest discover -s tests
```
