from __future__ import annotations

import argparse
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import export_onenote


class SafeNameTests(unittest.TestCase):
    def test_safe_name_replaces_filesystem_hostile_characters(self) -> None:
        value = 'Week 1: intro / "setup" <draft>? #100%'

        result = export_onenote.safe_name(value)

        self.assertEqual(result, "Week 1_ intro _ _setup_ _draft__ _100_")

    def test_safe_name_uses_fallback_for_blank_values(self) -> None:
        self.assertEqual(export_onenote.safe_name("   ", fallback="Untitled"), "Untitled")
        self.assertEqual(export_onenote.safe_name(None, fallback="Untitled"), "Untitled")

    def test_safe_name_limits_long_names(self) -> None:
        self.assertEqual(len(export_onenote.safe_name("x" * 200)), 140)


class PaginationTests(unittest.TestCase):
    def test_paginate_yields_all_values_and_clears_params_after_first_page(self) -> None:
        calls: list[tuple[str, dict[str, str] | None]] = []

        def fake_get_json(token: str, url: str, params: dict[str, str] | None = None) -> dict:
            calls.append((url, params))
            if url == "/first":
                return {
                    "value": [{"id": "1"}],
                    "@odata.nextLink": "https://example.test/next",
                }
            return {"value": [{"id": "2"}]}

        client = export_onenote.GraphClient(token="token", get_json=fake_get_json)

        result = list(client.paginate("/first", params={"$top": "100"}))

        self.assertEqual(result, [{"id": "1"}, {"id": "2"}])
        self.assertEqual(
            calls,
            [
                ("/first", {"$top": "100"}),
                ("https://example.test/next", None),
            ],
        )


class SectionTraversalTests(unittest.TestCase):
    def test_iter_sections_walks_nested_section_groups(self) -> None:
        graph_data = {
            "sections://notebook": {
                "value": [{"displayName": "Root", "pagesUrl": "pages://root"}]
            },
            "groups://notebook": {
                "value": [
                    {
                        "displayName": "Group A",
                        "sectionsUrl": "sections://group-a",
                        "sectionGroupsUrl": "groups://group-a",
                    }
                ]
            },
            "sections://group-a": {
                "value": [{"displayName": "Nested", "pagesUrl": "pages://nested"}]
            },
            "groups://group-a": {"value": []},
        }

        def fake_get_json(token: str, url: str, params: dict[str, str] | None = None) -> dict:
            return graph_data[url]

        client = export_onenote.GraphClient(token="token", get_json=fake_get_json)
        notebook = {
            "sectionsUrl": "sections://notebook",
            "sectionGroupsUrl": "groups://notebook",
        }

        result = list(export_onenote.iter_sections(client, notebook))

        self.assertEqual(
            [(path, section["pagesUrl"]) for path, section in result],
            [("Root", "pages://root"), ("Group A/Nested", "pages://nested")],
        )


class CliTests(unittest.TestCase):
    def test_load_dotenv_reads_simple_project_env_without_overriding_existing_values(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            env_path = Path(tmpdir) / ".env"
            env_path.write_text(
                "\n".join(
                    [
                        "ONENOTE_CLIENT_ID=from-file",
                        "ONENOTE_TENANT_ID='quoted-tenant'",
                        "ONENOTE_FORMATS=\"md,txt\"",
                        "IGNORED_LINE",
                        "# comment",
                    ]
                ),
                encoding="utf-8",
            )

            with patch.dict(os.environ, {"ONENOTE_CLIENT_ID": "already-set"}, clear=False):
                loaded = export_onenote.load_dotenv(env_path)

                self.assertTrue(loaded)
                self.assertEqual(os.environ["ONENOTE_CLIENT_ID"], "already-set")
                self.assertEqual(os.environ["ONENOTE_TENANT_ID"], "quoted-tenant")
                self.assertEqual(os.environ["ONENOTE_FORMATS"], "md,txt")

    def test_parse_args_uses_dotenv_values_for_pycharm_style_run(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            env_path = Path(tmpdir) / ".env"
            env_path.write_text(
                "\n".join(
                    [
                        "ONENOTE_CLIENT_ID=abc",
                        "ONENOTE_OUT=~/OneNoteExport",
                        "ONENOTE_NOTEBOOK=CSX4107",
                        "ONENOTE_FORMATS=md,txt",
                    ]
                ),
                encoding="utf-8",
            )

            with patch.dict(os.environ, {}, clear=True):
                args = export_onenote.parse_args([], env_file=env_path)

        self.assertEqual(args.client_id, "abc")
        self.assertEqual(args.out, "~/OneNoteExport")
        self.assertEqual(args.notebook, "CSX4107")
        self.assertEqual(args.formats, "md,txt")

    def test_parse_args_defaults_to_html_only_and_organizations_tenant(self) -> None:
        with patch.dict(os.environ, {}, clear=True):
            args = export_onenote.parse_args(["--client-id", "abc"], env_file=None)

        self.assertEqual(args.client_id, "abc")
        self.assertEqual(args.tenant_id, "organizations")
        self.assertEqual(args.formats, "")
        self.assertEqual(args.location, "/me")

    def test_main_prints_tagged_error_when_client_id_is_missing(self) -> None:
        with patch.dict(os.environ, {}, clear=True), patch("builtins.print") as print_mock:
            exit_code = export_onenote.main([], token_provider=Mock())

        self.assertEqual(exit_code, 2)
        self.assertEqual(
            [call.args[0] for call in print_mock.call_args_list],
            [
                "[ERROR] Missing Microsoft Entra application/client ID.",
                "[RECOMMENDATION] Set ONENOTE_CLIENT_ID or pass --client-id.",
            ],
        )

    def test_main_lists_notebooks_without_exporting_pages(self) -> None:
        token_provider = Mock(return_value="token")
        client = Mock()
        client.list_notebooks.return_value = [
            {"displayName": "Course Notes", "isShared": True, "userRole": "Reader"}
        ]

        with patch("builtins.print") as print_mock:
            exit_code = export_onenote.main(
                ["--client-id", "abc", "--list"],
                token_provider=token_provider,
                client_factory=lambda token: client,
            )

        self.assertEqual(exit_code, 0)
        token_provider.assert_called_once_with(
            client_id="abc",
            tenant_id="organizations",
            scopes=["Notes.Read.All"],
            cache_path=Path(".msal_token_cache.json"),
        )
        client.list_notebooks.assert_called_once_with("/me")
        self.assertTrue(
            any("Course Notes" in str(call.args[0]) for call in print_mock.call_args_list)
        )

    def test_parse_args_rejects_unknown_formats(self) -> None:
        with self.assertRaises(argparse.ArgumentTypeError):
            export_onenote.parse_formats("pdf")


if __name__ == "__main__":
    unittest.main()
