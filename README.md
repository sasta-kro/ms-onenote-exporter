# MS OneNote Scraper

Export OneNote notebooks that your Microsoft work/school account can already
access into local files.

The exporter always saves each page as `.html`. HTML is the source format
returned by Microsoft Graph and usually preserves more OneNote structure than
plain text. Optionally, you can also convert the HTML pages to Markdown, TXT, or
RTF with Pandoc.

## What This Does

- Signs in with your Microsoft account using a device-code flow.
- Lists notebooks visible through Microsoft Graph OneNote endpoints.
- Walks notebooks, sections, nested section groups, and pages.
- Downloads page content as local `.html` files.
- Writes a `manifest.json` with exported page metadata.
- Optionally runs Pandoc to create `.md`, `.txt`, or `.rtf` copies.

This is not an org-wide backup tool. It only exports notebooks that the signed-in
user can already access.

## Requirements

Python dependencies are documented in `requirements.txt`:

```bash
python -m pip install -r requirements.txt
```

Those packages are:

- `msal`: Microsoft Authentication Library. It handles the browser/device-code
  login and gets a Microsoft Graph access token. The script never sees your
  password.
- `requests`: HTTP client used to call Microsoft Graph with that token.

Optional system dependency:

```bash
brew install pandoc
```

`pandoc` is only needed if you pass `--formats md`, `--formats txt`,
`--formats rtf`, or a comma-separated combination like `--formats md,txt`.
If you only want `.html`, you do not need Pandoc and you do not need Homebrew.

Homebrew is mentioned because it is the common macOS way to install command-line
tools like Pandoc. You can inspect it first with:

```bash
brew info pandoc
brew home pandoc
```

## Microsoft App Setup

You said the web/auth side is already done. For reference, this script expects:

- A Microsoft Entra app registration.
- Public client/device-code flow enabled.
- Microsoft Graph delegated permission: `Notes.Read.All`.
- The Application/client ID from the app registration.

You do not need a client secret for this script.

## Setup

From this project directory:

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Create your local `.env` file:

```bash
cp .env.example .env
```

Then edit `.env` and paste your Entra Application/client ID:

```text
ONENOTE_CLIENT_ID=paste-your-application-client-id-here
ONENOTE_TENANT_ID=organizations
```

`organizations` is a good default for work/school Microsoft accounts. You can
replace it with your tenant ID or tenant domain if needed.

The script loads `.env` automatically before reading CLI defaults. Real shell
environment variables still win over `.env` values, and explicit CLI arguments
win over both.

## PyCharm Run Button

Use [main.py](main.py) as the PyCharm run target.

Why this exists: PyCharm's run button does not automatically inherit `export ...`
commands you typed in a separate terminal. The project therefore supports a
local `.env` file and a small `main.py` entrypoint so the run button can work
without custom shell setup.

Make sure PyCharm uses this project's virtualenv, not Miniforge/Conda base:

```text
PyCharm > Settings > Project > Python Interpreter
```

Select:

```text
<project folder>/.venv/bin/python
```

If PyCharm runs something like `/Users/.../miniforge3/bin/python3`, it is using
the wrong interpreter. Dependencies installed into `.venv` will not be visible
there.

For a portable setup on another machine:

```bash
git pull
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
cp .env.example .env
```

Then fill in `.env` and run `main.py` from PyCharm or the terminal.

## Usage

List notebooks first:

```bash
python main.py --list
```

On first run, the terminal will print a Microsoft device-login URL and a
`[DEVICE CODE]` line. Open the URL, enter the code from the terminal into the
Microsoft page, and sign in with the Microsoft account that can access the
OneNote notebooks. The code is not in Teams or OneNote.

Export everything visible to your account as HTML:

```bash
python main.py --out ~/OneNoteExport
```

Export only notebooks whose name contains some text:

```bash
python main.py --out ~/OneNoteExport --notebook "CSX4107"
```

Export HTML and also convert to Markdown and TXT:

```bash
python main.py --out ~/OneNoteExport --formats md,txt
```

Export from another Microsoft Graph OneNote root:

```bash
python main.py --location "/groups/GROUP_ID" --out ~/OneNoteExport
python main.py --location "/sites/SITE_ID" --out ~/OneNoteExport
```

You can also put stable defaults in `.env` instead of typing flags every time:

```text
ONENOTE_OUT=~/OneNoteExport
ONENOTE_NOTEBOOK=CSX4107
ONENOTE_FORMATS=md,txt
```

## Output

Output is organized like this:

```text
OneNoteExport/
  Notebook Name/
    Section Name/
      Page title-abc123.html
      Page title-abc123.md
      Page title-abc123.txt
  manifest.json
```

The short suffix in each filename comes from the OneNote page ID. It helps avoid
collisions when multiple pages have the same title.

## Token Cache

The script stores an MSAL token cache at:

```text
.msal_token_cache.json
```

That file is ignored by git. Treat it as private because it can contain reusable
login tokens. If you change permissions in Entra or want to force a fresh login,
delete the cache:

```bash
rm .msal_token_cache.json
```

## Troubleshooting

If `--list` says no notebooks were found, first try without `--notebook` and
confirm the notebook is visible to the same account in OneNote.

If Microsoft says admin approval is required, your org blocks user consent for
the requested Graph permission. Ask IT to approve delegated `Notes.Read.All` for
your app registration.

If Microsoft returns `AADSTS50059` or says no tenant-identifying information was
found, check `.env`:

```text
ONENOTE_TENANT_ID=organizations
```

You can also use the Directory/tenant ID from the app registration Overview page.
Do not leave `ONENOTE_TENANT_ID` blank in PyCharm's run configuration.

If `.html` files export but `.md`/`.txt`/`.rtf` files do not appear, install
Pandoc or run with `--formats ""` for HTML-only export.

If a class notebook is missing from `/me`, it may live under a Microsoft 365
group or SharePoint site. Use `--location "/groups/GROUP_ID"` or
`--location "/sites/SITE_ID"` once you know the correct ID.

## Development

Run tests with:

```bash
python -m unittest discover -s tests
```
