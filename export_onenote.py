#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
from collections.abc import Callable, Iterator
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import quote, unquote, urlparse

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_SCOPES = ["Notes.Read.All"]
SITE_URL_SCOPES = ["Notes.Read.All", "Sites.Read.All"]
PANDOC_TARGETS = {
    "md": "gfm",
    "txt": "plain",
    "rtf": "rtf",
}


class GraphError(RuntimeError):
    """Raised when Microsoft Graph returns an unrecoverable error."""


class MissingDependencyError(RuntimeError):
    def __init__(self, package: str) -> None:
        self.package = package
        super().__init__(f"Missing dependency '{package}'.")


def log_error(message: str) -> None:
    print(f"[ERROR] {message}")


def log_info(message: str) -> None:
    print(f"[INFO] {message}")


def log_action(message: str) -> None:
    print(f"[ACTION] {message}")


def log_device_code(code: str) -> None:
    print(f"[DEVICE CODE] {code}")


def log_recommendation(message: str) -> None:
    print(f"[RECOMMENDATION] {message}")


def env_value(name: str, default: str | None = None) -> str | None:
    value = os.getenv(name)
    if value is None or not value.strip():
        return default
    return value.strip()


def local_venv_python() -> Path:
    return Path(__file__).resolve().parent / ".venv" / "bin" / "python"


def log_missing_dependency(error: MissingDependencyError) -> None:
    venv_python = local_venv_python()
    log_error(f"Missing dependency '{error.package}' in the active Python interpreter.")
    log_info(f"Active Python: {sys.executable}")
    log_info(f"Project venv Python: {venv_python}")
    log_recommendation(f"In PyCharm, set the Project Interpreter to: {venv_python}")
    log_recommendation(
        "Then rerun main.py. Your dependencies are installed in the project .venv, "
        "not Miniforge base."
    )


def log_runtime_error(error: RuntimeError) -> None:
    message = str(error)
    if "AADSTS50059" in message:
        log_error("Microsoft login did not receive tenant-identifying information.")
        log_recommendation(
            "Set ONENOTE_TENANT_ID=organizations in .env, or use your Directory (tenant) ID."
        )
        log_info("Original error: AADSTS50059")
        return
    print(message, file=sys.stderr)


def log_device_flow(flow: dict[str, Any]) -> None:
    code = flow["user_code"]
    url = flow.get("verification_uri") or flow.get("verification_url")
    if not url:
        url = "https://login.microsoft.com/device"

    log_action(f"Open this URL in your browser: {url}")
    log_device_code(code)
    log_action("Paste the device code above into the Microsoft page, then click Next.")
    log_info("The code is printed here in the terminal. It is not in Teams or OneNote.")

    expires_in = flow.get("expires_in")
    if isinstance(expires_in, int) and expires_in > 0:
        minutes = max(1, round(expires_in / 60))
        log_info(f"Code expires in about {minutes} minutes.")


def safe_name(value: str | None, fallback: str = "untitled", limit: int = 140) -> str:
    value = value or fallback
    value = re.sub(r"[\\/:*?\"<>|#%{}$!@+`=]", "_", value)
    value = re.sub(r"\s+", " ", value).strip()
    return (value[:limit] or fallback).rstrip(". ")


def parse_formats(value: str) -> list[str]:
    if not value.strip():
        return []
    formats = [item.strip().lower() for item in value.split(",") if item.strip()]
    unsupported = sorted(set(formats) - set(PANDOC_TARGETS))
    if unsupported:
        raise argparse.ArgumentTypeError(
            f"unsupported format(s): {', '.join(unsupported)}. Use md, txt, rtf, or ''."
        )
    return formats


@dataclass(frozen=True)
class SharePointSiteIdHelperUrls:
    site_root: str
    site_id_url: str
    web_id_url: str
    site_id_template: str


def sharepoint_url_to_site_id_helper_urls(url: str) -> SharePointSiteIdHelperUrls:
    parsed = urlparse(url.strip())
    scheme = parsed.scheme or "https"
    host = parsed.netloc.lower()
    path_parts = [unquote(part) for part in parsed.path.split("/") if part]

    if not host or len(path_parts) < 2 or path_parts[0].lower() not in {"sites", "teams"}:
        raise ValueError(
            "Could not infer a SharePoint site from that URL. Expected a URL containing "
            "/sites/<name>/... or /teams/<name>/..."
        )

    site_kind = path_parts[0].lower()
    site_name = quote(path_parts[1], safe="")
    site_root = f"{scheme}://{host}/{site_kind}/{site_name}"
    return SharePointSiteIdHelperUrls(
        site_root=site_root,
        site_id_url=f"{site_root}/_api/site/id",
        web_id_url=f"{site_root}/_api/web/id",
        site_id_template=f"{host},SITE_GUID,WEB_GUID",
    )


def shell_double_quote(value: str) -> str:
    escaped = (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("$", "\\$")
        .replace("`", "\\`")
        .replace("\n", " ")
    )
    return f'"{escaped}"'


def print_site_id_helper(site_url: str, *, list_notebooks: bool = False, notebook: str | None = None) -> None:
    helper = sharepoint_url_to_site_id_helper_urls(site_url)
    if notebook:
        next_flag = f"--notebook {shell_double_quote(notebook)}"
    elif list_notebooks:
        next_flag = "--list"
    else:
        next_flag = "--list"

    print("")
    print("== SharePoint site ==")
    print(helper.site_root)
    print("")
    print("== Step 1: open this in your signed-in browser ==")
    print(helper.site_id_url)
    print("")
    print("== Step 2: open this in your signed-in browser ==")
    print(helper.web_id_url)
    print("")
    print("== Step 3: copy the two GUID values, then run ==")
    print(f"python main.py --site-id {shell_double_quote(helper.site_id_template)} {next_flag}")
    print("")


def sharepoint_url_to_site_lookup_path(url: str) -> str:
    parsed = urlparse(url.strip())
    host = parsed.netloc.lower()
    path_parts = [unquote(part) for part in parsed.path.split("/") if part]

    if not host or len(path_parts) < 2 or path_parts[0].lower() not in {"sites", "teams"}:
        raise ValueError(
            "Could not infer a SharePoint site from that URL. Expected a URL containing "
            "/sites/<name>/... or /teams/<name>/..."
        )

    site_kind = path_parts[0]
    site_name = path_parts[1]
    return f"/sites/{host}:/{site_kind}/{quote(site_name, safe='')}:"


def resolve_sharepoint_site_location(client: Any, site_url: str) -> str:
    lookup_path = sharepoint_url_to_site_lookup_path(site_url)
    site = client.json(lookup_path, params={"$select": "id,displayName,webUrl"})
    site_id = site.get("id")
    if not site_id:
        raise GraphError(f"Could not resolve SharePoint site ID for: {site_url}")
    log_info(f"Resolved SharePoint site: {site.get('displayName') or site.get('webUrl') or site_id}")
    return f"/sites/{site_id}"


def site_id_to_site_location(site_id: str) -> str:
    value = site_id.strip()
    if value.startswith("/sites/"):
        value = value[len("/sites/") :]
    if ":/" in value or value.endswith(":"):
        raise ValueError(
            "--site-id expects a resolved Graph site ID like "
            "hostname,siteCollectionGuid,webGuid. Use --site-url for SharePoint URLs."
        )
    parts = [part.strip() for part in value.split(",")]
    if len(parts) != 3 or not all(parts):
        raise ValueError(
            "--site-id expects a resolved Graph site ID like "
            "hostname,siteCollectionGuid,webGuid."
        )
    return f"/sites/{value}"


def unquote_env_value(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        return value[1:-1]
    return value


def load_dotenv(path: Path = Path(".env"), *, override: bool = False) -> bool:
    if not path.exists():
        return False

    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        key = key.strip()
        if not key or not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", key):
            continue
        if override or key not in os.environ:
            os.environ[key] = unquote_env_value(value)
    return True


def parse_args(argv: list[str] | None = None, env_file: Path | None = Path(".env")) -> argparse.Namespace:
    if env_file is not None:
        load_dotenv(env_file)

    parser = argparse.ArgumentParser(
        description="Export OneNote pages visible to your Microsoft account as local HTML files."
    )
    parser.add_argument("--client-id", default=env_value("ONENOTE_CLIENT_ID"))
    parser.add_argument(
        "--tenant-id",
        default=env_value("ONENOTE_TENANT_ID", "organizations"),
        help="Microsoft tenant ID, domain, or 'organizations'.",
    )
    parser.add_argument(
        "--location",
        default=env_value("ONENOTE_LOCATION", "/me"),
        help="Graph location root: /me, /users/{id}, /groups/{id}, or /sites/{id}.",
    )
    parser.add_argument(
        "--site-url",
        default=env_value("ONENOTE_SITE_URL"),
        help="Teams/SharePoint URL for a class notebook site. Overrides --location.",
    )
    parser.add_argument(
        "--site-id",
        default=env_value("ONENOTE_SITE_ID"),
        help="Resolved Graph site ID: hostname,siteCollectionGuid,webGuid. Overrides --site-url and --location.",
    )
    parser.add_argument(
        "--resolve-site-url-with-graph",
        action="store_true",
        help="Advanced: resolve --site-url through Microsoft Graph. Requires Sites.Read.All admin consent.",
    )
    parser.add_argument(
        "--out",
        default=env_value("ONENOTE_OUT", "onenote_export"),
        help="Output directory.",
    )
    parser.add_argument(
        "--notebook",
        default=env_value("ONENOTE_NOTEBOOK"),
        help="Only export notebooks whose name contains this text.",
    )
    parser.add_argument(
        "--formats",
        default=env_value("ONENOTE_FORMATS", ""),
        help="Optional comma-separated conversions: md,txt,rtf. Empty means HTML only.",
    )
    parser.add_argument("--list", action="store_true", help="List notebooks and exit.")
    parser.add_argument(
        "--cache",
        default=env_value("ONENOTE_TOKEN_CACHE", ".msal_token_cache.json"),
        help="MSAL token cache path. Keep this private.",
    )
    parser.add_argument(
        "--include-ids",
        action="store_true",
        help="Ask OneNote to include object IDs in exported HTML.",
    )
    return parser.parse_args(argv)


def normalize_location(location: str) -> str:
    location = location.strip().rstrip("/")
    if not location:
        return "/me"
    if not location.startswith("/"):
        location = f"/{location}"
    return location


def graph_url(url: str) -> str:
    if url.startswith("http://") or url.startswith("https://"):
        return url
    return f"{GRAPH_ROOT}{url}"


@dataclass
class GraphClient:
    token: str
    get_json: Callable[[str, str, dict[str, str] | None], dict[str, Any]] | None = None
    get_bytes: Callable[[str, str, dict[str, str] | None], bytes] | None = None

    def request(
        self,
        method: str,
        url: str,
        *,
        headers: dict[str, str] | None = None,
        params: dict[str, str] | None = None,
    ) -> Any:
        try:
            import requests
        except ImportError as exc:
            raise MissingDependencyError("requests") from exc

        request_headers = dict(headers or {})
        request_headers["Authorization"] = f"Bearer {self.token}"
        request_url = graph_url(url)

        for attempt in range(4):
            response = requests.request(
                method,
                request_url,
                headers=request_headers,
                params=params,
                timeout=60,
            )
            if response.status_code == 429:
                wait = int(response.headers.get("Retry-After", "10"))
                print(f"Microsoft Graph throttled this request. Waiting {wait}s...")
                time.sleep(wait)
                continue
            if response.status_code >= 500 and attempt < 3:
                time.sleep(2**attempt)
                continue
            if response.status_code >= 400:
                raise GraphError(
                    f"Microsoft Graph error {response.status_code} for {request_url}\n"
                    f"{response.text[:3000]}"
                )
            return response

        raise GraphError(f"Microsoft Graph request failed after retries: {request_url}")

    def json(self, url: str, params: dict[str, str] | None = None) -> dict[str, Any]:
        if self.get_json is not None:
            return self.get_json(self.token, url, params)
        response = self.request(
            "GET",
            url,
            headers={"Accept": "application/json"},
            params=params,
        )
        return response.json()

    def bytes(self, url: str, params: dict[str, str] | None = None) -> bytes:
        if self.get_bytes is not None:
            return self.get_bytes(self.token, url, params)
        response = self.request(
            "GET",
            url,
            headers={"Accept": "text/html"},
            params=params,
        )
        return response.content

    def paginate(
        self, url: str, params: dict[str, str] | None = None
    ) -> Iterator[dict[str, Any]]:
        while url:
            data = self.json(url, params=params)
            yield from data.get("value", [])
            url = data.get("@odata.nextLink")
            params = None

    def list_notebooks(self, location: str) -> list[dict[str, Any]]:
        params = {
            "$select": "id,displayName,sectionsUrl,sectionGroupsUrl,isShared,userRole",
            "$top": "100",
        }
        return list(self.paginate(f"{normalize_location(location)}/onenote/notebooks", params))

    def list_pages(self, section: dict[str, Any]) -> list[dict[str, Any]]:
        pages_url = section.get("pagesUrl")
        if not pages_url:
            return []
        params = {
            "$select": "id,title,contentUrl,lastModifiedDateTime",
            "$top": "100",
        }
        return list(self.paginate(pages_url, params))


def get_token(
    *,
    client_id: str,
    tenant_id: str,
    scopes: list[str],
    cache_path: Path,
) -> str:
    try:
        import msal
    except ImportError as exc:
        raise MissingDependencyError("msal") from exc

    cache = msal.SerializableTokenCache()
    if cache_path.exists():
        cache.deserialize(cache_path.read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache,
    )

    result: dict[str, Any] | None = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Could not start Microsoft device login: {flow}")
        log_device_flow(flow)
        result = app.acquire_token_by_device_flow(flow)

    if cache.has_state_changed:
        cache_path.write_text(cache.serialize(), encoding="utf-8")

    if not result or "access_token" not in result:
        raise RuntimeError(f"Could not get Microsoft Graph access token: {result}")

    return result["access_token"]


def iter_sections(
    client: GraphClient, container: dict[str, Any], prefix: str = ""
) -> Iterator[tuple[str, dict[str, Any]]]:
    sections_url = container.get("sectionsUrl")
    if sections_url:
        section_params = {"$select": "id,displayName,pagesUrl", "$top": "100"}
        for section in client.paginate(sections_url, params=section_params):
            section_name = section.get("displayName") or "Untitled section"
            full_name = f"{prefix}/{section_name}" if prefix else section_name
            yield full_name, section

    groups_url = container.get("sectionGroupsUrl")
    if groups_url:
        group_params = {
            "$select": "id,displayName,sectionsUrl,sectionGroupsUrl",
            "$top": "100",
        }
        for group in client.paginate(groups_url, params=group_params):
            group_name = group.get("displayName") or "Untitled group"
            new_prefix = f"{prefix}/{group_name}" if prefix else group_name
            yield from iter_sections(client, group, new_prefix)


def convert_with_pandoc(html_path: Path, output_base: Path, formats: list[str]) -> None:
    if not formats:
        return

    pandoc = shutil.which("pandoc")
    if not pandoc:
        print(
            "pandoc was not found. HTML was saved, but requested conversions were skipped."
        )
        return

    for fmt in formats:
        out_path = output_base.with_suffix(f".{fmt}")
        result = subprocess.run(
            [pandoc, str(html_path), "-f", "html", "-t", PANDOC_TARGETS[fmt], "-o", str(out_path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            print(f"pandoc failed for {html_path.name} -> {fmt}")
            print(result.stderr[:1000])


def page_content_url(location: str, page: dict[str, Any], include_ids: bool) -> tuple[str, dict[str, str] | None]:
    content_url = page.get("contentUrl")
    params = {"includeIDs": "true"} if include_ids else None
    if content_url:
        return content_url, params
    page_id = quote(page["id"], safe="")
    return f"{normalize_location(location)}/onenote/pages/{page_id}/content", params


def export_page(
    client: GraphClient,
    *,
    location: str,
    page: dict[str, Any],
    output_dir: Path,
    formats: list[str],
    include_ids: bool,
) -> dict[str, str]:
    title = page.get("title") or "Untitled page"
    page_id = page["id"]
    short_id = re.sub(r"\W+", "", page_id)[-10:] or "page"
    output_base = output_dir / f"{safe_name(title)}-{short_id}"
    html_path = output_base.with_suffix(".html")

    url, params = page_content_url(location, page, include_ids)
    html_path.write_bytes(client.bytes(url, params=params))
    convert_with_pandoc(html_path, output_base, formats)

    return {
        "title": title,
        "id": page_id,
        "html": str(html_path),
        "lastModifiedDateTime": page.get("lastModifiedDateTime", ""),
    }


def export_notebooks(
    client: GraphClient,
    *,
    location: str,
    output_dir: Path,
    notebook_filter: str | None,
    formats: list[str],
    include_ids: bool,
) -> int:
    all_notebooks = client.list_notebooks(location)
    notebooks = all_notebooks
    if notebook_filter:
        notebooks = [
            notebook
            for notebook in notebooks
            if notebook_filter.lower() in (notebook.get("displayName") or "").lower()
        ]

    if not notebooks:
        if notebook_filter and all_notebooks:
            log_error(f"No notebooks matched filter: {notebook_filter}")
            log_info(f"Notebooks visible at {normalize_location(location)}:")
            for notebook in all_notebooks[:20]:
                print(f"  - {notebook.get('displayName') or 'Untitled notebook'}")
            if len(all_notebooks) > 20:
                print(f"  ... and {len(all_notebooks) - 20} more")
            log_recommendation("Copy one of the names above exactly, or run with --list.")
            log_recommendation(
                "If your class notebook is not listed, it may live under a Microsoft 365 group "
                "or SharePoint site."
            )
        else:
            log_error(f"No notebooks found at {normalize_location(location)}.")
            log_recommendation(
                "Try --location /groups/GROUP_ID or --location /sites/SITE_ID if this is a class notebook."
            )
        return 0

    output_dir.mkdir(parents=True, exist_ok=True)
    total_pages = 0
    notebook_outputs: list[tuple[str, Path, Path]] = []

    for notebook in notebooks:
        notebook_name = notebook.get("displayName") or "Untitled notebook"
        print(f"\nNotebook: {notebook_name}")
        notebook_dir = output_dir / safe_name(notebook_name)
        notebook_dir.mkdir(parents=True, exist_ok=True)
        notebook_manifest: list[dict[str, str]] = []

        for section_path, section in iter_sections(client, notebook):
            section_dir = notebook_dir / safe_name(section_path)
            section_dir.mkdir(parents=True, exist_ok=True)
            pages = client.list_pages(section)
            print(f"  Section: {section_path} ({len(pages)} pages)")

            for page in pages:
                print(f"    Exporting: {page.get('title') or 'Untitled page'}")
                record = export_page(
                    client,
                    location=location,
                    page=page,
                    output_dir=section_dir,
                    formats=formats,
                    include_ids=include_ids,
                )
                record["notebook"] = notebook_name
                record["section"] = section_path
                notebook_manifest.append(record)
                total_pages += 1

        manifest_path = notebook_dir / "manifest.json"
        manifest_path.write_text(json.dumps(notebook_manifest, indent=2), encoding="utf-8")
        notebook_outputs.append((notebook_name, notebook_dir, manifest_path))

    print(f"\nExported {total_pages} page(s).")
    print(f"Output root: {output_dir}")
    for notebook_name, notebook_dir, manifest_path in notebook_outputs:
        print(f"Notebook output: {notebook_name}")
        print(f"  Folder: {notebook_dir}")
        print(f"  Manifest: {manifest_path}")
    return total_pages


def print_notebooks(client: GraphClient, location: str, export_command_base: str | None = None) -> None:
    notebooks = client.list_notebooks(location)
    if not notebooks:
        print("No notebooks found.")
        return
    for index, notebook in enumerate(notebooks, start=1):
        name = notebook.get("displayName") or "Untitled notebook"
        shared = notebook.get("isShared")
        role = notebook.get("userRole")
        print(f"{index}. {name} | shared={shared} | role={role}")
    if export_command_base:
        first_name = notebooks[0].get("displayName") or "Untitled notebook"
        print("")
        print("== To download one notebook ==")
        print(f"{export_command_base} --notebook {shell_double_quote(first_name)}")


def main(
    argv: list[str] | None = None,
    *,
    token_provider: Callable[..., str] = get_token,
    client_factory: Callable[[str], GraphClient] = GraphClient,
) -> int:
    args = parse_args(argv)
    if args.site_url and not args.site_id and not args.resolve_site_url_with_graph:
        try:
            print_site_id_helper(args.site_url, list_notebooks=args.list, notebook=args.notebook)
            return 0
        except ValueError as exc:
            print(str(exc), file=sys.stderr)
            return 1

    if not args.client_id:
        log_error("Missing Microsoft Entra application/client ID.")
        log_recommendation("Set ONENOTE_CLIENT_ID or pass --client-id.")
        return 2

    try:
        formats = parse_formats(args.formats)
        scopes = (
            SITE_URL_SCOPES
            if args.site_url and not args.site_id and args.resolve_site_url_with_graph
            else DEFAULT_SCOPES
        )
        token = token_provider(
            client_id=args.client_id,
            tenant_id=args.tenant_id,
            scopes=scopes,
            cache_path=Path(args.cache),
        )
        client = client_factory(token)
        if args.site_id:
            location = site_id_to_site_location(args.site_id)
            site_id_value = location[len("/sites/") :]
            export_command_base = f"python main.py --site-id {shell_double_quote(site_id_value)}"
        elif args.site_url:
            location = resolve_sharepoint_site_location(client, args.site_url)
            export_command_base = (
                f"python main.py --site-url {shell_double_quote(args.site_url)} "
                "--resolve-site-url-with-graph"
            )
        else:
            location = normalize_location(args.location)
            if location == "/me":
                export_command_base = "python main.py"
            else:
                export_command_base = f"python main.py --location {shell_double_quote(location)}"

        if args.list:
            print_notebooks(client, location, export_command_base=export_command_base)
            return 0

        export_notebooks(
            client,
            location=location,
            output_dir=Path(args.out).expanduser().resolve(),
            notebook_filter=args.notebook,
            formats=formats,
            include_ids=args.include_ids,
        )
        return 0
    except MissingDependencyError as exc:
        log_missing_dependency(exc)
        return 1
    except (GraphError, ValueError, argparse.ArgumentTypeError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except RuntimeError as exc:
        log_runtime_error(exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
