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
from urllib.parse import quote

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_SCOPES = ["Notes.Read.All"]
PANDOC_TARGETS = {
    "md": "gfm",
    "txt": "plain",
    "rtf": "rtf",
}


class GraphError(RuntimeError):
    """Raised when Microsoft Graph returns an unrecoverable error."""


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


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export OneNote pages visible to your Microsoft account as local HTML files."
    )
    parser.add_argument("--client-id", default=os.getenv("ONENOTE_CLIENT_ID"))
    parser.add_argument(
        "--tenant-id",
        default=os.getenv("ONENOTE_TENANT_ID", "organizations"),
        help="Microsoft tenant ID, domain, or 'organizations'.",
    )
    parser.add_argument(
        "--location",
        default="/me",
        help="Graph location root: /me, /users/{id}, /groups/{id}, or /sites/{id}.",
    )
    parser.add_argument("--out", default="onenote_export", help="Output directory.")
    parser.add_argument("--notebook", help="Only export notebooks whose name contains this text.")
    parser.add_argument(
        "--formats",
        default="",
        help="Optional comma-separated conversions: md,txt,rtf. Empty means HTML only.",
    )
    parser.add_argument("--list", action="store_true", help="List notebooks and exit.")
    parser.add_argument(
        "--cache",
        default=".msal_token_cache.json",
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
            raise RuntimeError(
                "Missing dependency 'requests'. Install dependencies with: "
                "python -m pip install -r requirements.txt"
            ) from exc

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
        raise RuntimeError(
            "Missing dependency 'msal'. Install dependencies with: "
            "python -m pip install -r requirements.txt"
        ) from exc

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
        print(flow["message"])
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
    notebooks = client.list_notebooks(location)
    if notebook_filter:
        notebooks = [
            notebook
            for notebook in notebooks
            if notebook_filter.lower() in (notebook.get("displayName") or "").lower()
        ]

    if not notebooks:
        print("No notebooks found.")
        if notebook_filter:
            print("Try running with --list or without --notebook to see available names.")
        return 0

    output_dir.mkdir(parents=True, exist_ok=True)
    total_pages = 0
    manifest: list[dict[str, str]] = []

    for notebook in notebooks:
        notebook_name = notebook.get("displayName") or "Untitled notebook"
        print(f"\nNotebook: {notebook_name}")
        notebook_dir = output_dir / safe_name(notebook_name)
        notebook_dir.mkdir(parents=True, exist_ok=True)

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
                manifest.append(record)
                total_pages += 1

    manifest_path = output_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(f"\nExported {total_pages} page(s).")
    print(f"Output: {output_dir}")
    print(f"Manifest: {manifest_path}")
    return total_pages


def print_notebooks(client: GraphClient, location: str) -> None:
    notebooks = client.list_notebooks(location)
    if not notebooks:
        print("No notebooks found.")
        return
    for index, notebook in enumerate(notebooks, start=1):
        name = notebook.get("displayName") or "Untitled notebook"
        shared = notebook.get("isShared")
        role = notebook.get("userRole")
        print(f"{index}. {name} | shared={shared} | role={role}")


def main(
    argv: list[str] | None = None,
    *,
    token_provider: Callable[..., str] = get_token,
    client_factory: Callable[[str], GraphClient] = GraphClient,
) -> int:
    args = parse_args(argv)
    if not args.client_id:
        print("Missing Microsoft Entra application/client ID.")
        print("Set ONENOTE_CLIENT_ID or pass --client-id.")
        return 2

    try:
        formats = parse_formats(args.formats)
        location = normalize_location(args.location)
        token = token_provider(
            client_id=args.client_id,
            tenant_id=args.tenant_id,
            scopes=DEFAULT_SCOPES,
            cache_path=Path(args.cache),
        )
        client = client_factory(token)

        if args.list:
            print_notebooks(client, location)
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
    except (GraphError, RuntimeError, argparse.ArgumentTypeError) as exc:
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
