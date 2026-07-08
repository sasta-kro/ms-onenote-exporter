from __future__ import annotations

import argparse
import io
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


class SharePointUrlTests(unittest.TestCase):
    def test_sharepoint_url_to_site_id_helper_urls_supports_doc_links(self) -> None:
        url = (
            "https://school.sharepoint.com/sites/Section_123/_layouts/15/Doc.aspx"
            "?sourcedoc={abc}&action=view"
        )

        result = export_onenote.sharepoint_url_to_site_id_helper_urls(url)

        self.assertEqual(
            result.site_root,
            "https://school.sharepoint.com/sites/Section_123",
        )
        self.assertEqual(
            result.site_id_url,
            "https://school.sharepoint.com/sites/Section_123/_api/site/id",
        )
        self.assertEqual(
            result.web_id_url,
            "https://school.sharepoint.com/sites/Section_123/_api/web/id",
        )
        self.assertEqual(
            result.site_id_template,
            "school.sharepoint.com,SITE_GUID,WEB_GUID",
        )

    def test_site_id_to_site_location_accepts_graph_site_id(self) -> None:
        result = export_onenote.site_id_to_site_location(
            "school.sharepoint.com,site-guid,web-guid"
        )

        self.assertEqual(result, "/sites/school.sharepoint.com,site-guid,web-guid")

    def test_site_id_to_site_location_accepts_prefixed_location(self) -> None:
        result = export_onenote.site_id_to_site_location(
            "/sites/school.sharepoint.com,site-guid,web-guid"
        )

        self.assertEqual(result, "/sites/school.sharepoint.com,site-guid,web-guid")

    def test_site_id_to_site_location_rejects_path_lookup_shape(self) -> None:
        with self.assertRaises(ValueError):
            export_onenote.site_id_to_site_location(
                "school.sharepoint.com:/sites/GDD542:"
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


class ExportNotebookTests(unittest.TestCase):
    def test_export_page_preserves_titles_with_dot_words(self) -> None:
        client = Mock()
        client.bytes.return_value = b"<html><body>Env setup</body></html>"
        page = {
            "id": "page-fe6a099c87",
            "title": "[LAB] 3: Configuration Management (.env) (Local Machine)",
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            record = export_onenote.export_page(
                client,
                location="/me",
                page=page,
                output_dir=Path(tmpdir),
                formats=[],
            )

            html_path = Path(record["html"])

        self.assertEqual(
            html_path.name,
            "[LAB] 3_ Configuration Management (.env) (Local Machine)-fe6a099c87.html",
        )

    def test_clean_onenote_html_for_text_removes_layout_spans_images_and_placeholders(self) -> None:
        html = """
        <html><body data-absolute-enabled="true">
          <div style="position:absolute;left:48px;top:115px;width:720px">
            <p>Use <span lang="en-US">.env</span> files.</p>
            <p>PORT=3000\ufffcDB_HOST=127.0.0.1</p>
            <img src="https://graph.microsoft.com/v1.0/resource/$value" />
          </div>
        </body></html>
        """

        result = export_onenote.clean_onenote_html_for_text(html, omit_images=True)

        self.assertIn(".env", result)
        self.assertIn("PORT=3000<br />\nDB_HOST=127.0.0.1", result)
        self.assertNotIn("<span", result)
        self.assertNotIn("<div", result)
        self.assertNotIn("\ufffc", result)
        self.assertNotIn("graph.microsoft.com", result)

    def test_clean_onenote_html_for_text_can_keep_image_links(self) -> None:
        html = '<html><body><div><img src="https://graph.microsoft.com/v1.0/resource/$value" /></div></body></html>'

        result = export_onenote.clean_onenote_html_for_text(html, omit_images=False)

        self.assertIn('<img src="https://graph.microsoft.com/v1.0/resource/$value">', result)

    def test_convert_with_pandoc_uses_cleaned_html_for_markdown_by_default(self) -> None:
        html = """
        <html><body>
          <div style="position:absolute;left:48px;top:115px;width:720px">
            <p>Hello <span lang="en-US">world</span>\ufffcAgain</p>
            <img src="https://graph.microsoft.com/v1.0/resource/$value" />
          </div>
        </body></html>
        """

        with tempfile.TemporaryDirectory() as tmpdir:
            output_base = Path(tmpdir) / "sample.with.dot"
            html_path = Path(f"{output_base}.html")
            html_path.write_text(html, encoding="utf-8")

            export_onenote.convert_with_pandoc(html_path, output_base, ["md"])

            markdown = Path(f"{output_base}.md").read_text(encoding="utf-8")

        self.assertIn("Hello world\nAgain", markdown)
        self.assertNotIn("\\\n", markdown)
        self.assertNotIn("<div", markdown)
        self.assertNotIn("<span", markdown)
        self.assertNotIn("graph.microsoft.com", markdown)

    def test_export_notebooks_writes_manifest_inside_each_notebook_dir(self) -> None:
        client = Mock()
        client.list_notebooks.return_value = [
            {
                "displayName": "Course Notes",
                "sectionsUrl": "sections-url",
            }
        ]
        client.paginate.return_value = [
            {
                "displayName": "Week 1",
                "pagesUrl": "pages-url",
            }
        ]
        client.list_pages.return_value = [
            {
                "id": "page-1234567890",
                "title": "Intro",
                "lastModifiedDateTime": "2026-07-06T00:00:00Z",
            }
        ]
        client.bytes.return_value = b"<html><body>Intro</body></html>"

        with tempfile.TemporaryDirectory() as tmpdir, patch("builtins.print"):
            output_dir = Path(tmpdir) / "out"

            count = export_onenote.export_notebooks(
                client,
                location="/me",
                output_dir=output_dir,
                notebook_filter="Course",
                formats=[],
            )

            notebook_dir = output_dir / "Course Notes"
            manifest_path = notebook_dir / "manifest.json"

            self.assertEqual(count, 1)
            self.assertTrue((notebook_dir / "Week 1" / "Intro-1234567890.html").exists())
            self.assertTrue(manifest_path.exists())
            self.assertFalse((output_dir / "manifest.json").exists())

    def test_export_notebooks_shows_available_names_when_filter_matches_nothing(self) -> None:
        client = Mock()
        client.list_notebooks.return_value = [
            {"displayName": "2026-1 GDD542 Notebook"},
            {"displayName": "Personal Notes"},
        ]

        with patch("builtins.print") as print_mock:
            count = export_onenote.export_notebooks(
                client,
                location="/me",
                output_dir=Path("out"),
                notebook_filter="2026-1 GDD 542 Notebook",
                formats=[],
            )

        self.assertEqual(count, 0)
        self.assertEqual(
            [call.args[0] for call in print_mock.call_args_list],
            [
                "[ERROR] No notebooks matched filter: 2026-1 GDD 542 Notebook",
                "[INFO] Notebooks visible at /me:",
                "  - 2026-1 GDD542 Notebook",
                "  - Personal Notes",
                "[RECOMMENDATION] Copy one of the names above exactly, or run with --list.",
                "[RECOMMENDATION] If your class notebook is not listed, it may live under a Microsoft 365 group or SharePoint site.",
            ],
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

    def test_parse_args_accepts_site_url(self) -> None:
        args = export_onenote.parse_args(
            [
                "--client-id",
                "abc",
                "--site-url",
                "https://school.sharepoint.com/sites/GDD542/Shared%20Documents",
            ],
            env_file=None,
        )

        self.assertEqual(args.site_url, "https://school.sharepoint.com/sites/GDD542/Shared%20Documents")

    def test_parse_args_accepts_site_id(self) -> None:
        args = export_onenote.parse_args(
            [
                "--client-id",
                "abc",
                "--site-id",
                "school.sharepoint.com,site-guid,web-guid",
            ],
            env_file=None,
        )

        self.assertEqual(args.site_id, "school.sharepoint.com,site-guid,web-guid")

    def test_parse_args_accepts_include_image_links(self) -> None:
        args = export_onenote.parse_args(
            ["--client-id", "abc", "--include-image-links"],
            env_file=None,
        )

        self.assertTrue(args.include_image_links)

    def test_parse_args_rejects_removed_flags(self) -> None:
        removed_flags = [
            ["--location", "/groups/abc"],
            ["--resolve-site-url-with-graph"],
            ["--include-ids"],
        ]

        for flag_args in removed_flags:
            with (
                self.subTest(flag_args=flag_args),
                patch("sys.stderr", io.StringIO()),
                self.assertRaises(SystemExit),
            ):
                export_onenote.parse_args(["--client-id", "abc", *flag_args], env_file=None)

    def test_parse_args_treats_blank_env_values_as_missing(self) -> None:
        with patch.dict(
            os.environ,
            {
                "ONENOTE_CLIENT_ID": "",
                "ONENOTE_TENANT_ID": "",
                "ONENOTE_OUT": "",
                "ONENOTE_FORMATS": "",
            },
            clear=True,
        ):
            args = export_onenote.parse_args([], env_file=None)

        self.assertIsNone(args.client_id)
        self.assertEqual(args.tenant_id, "organizations")
        self.assertEqual(args.out, "onenote_export")
        self.assertEqual(args.formats, "")

    def test_main_prints_tagged_error_when_client_id_is_missing(self) -> None:
        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch("builtins.print") as print_mock,
        ):
            exit_code = export_onenote.main([], token_provider=Mock())

        self.assertEqual(exit_code, 2)
        self.assertEqual(
            [call.args[0] for call in print_mock.call_args_list],
            [
                "[ERROR] Missing Microsoft Entra application/client ID.",
                "[RECOMMENDATION] Set ONENOTE_CLIENT_ID or pass --client-id.",
            ],
        )

    def test_main_explains_missing_dependency_interpreter_mismatch(self) -> None:
        token_provider = Mock(side_effect=export_onenote.MissingDependencyError("msal"))

        with (
            patch("builtins.print") as print_mock,
            patch.object(export_onenote.sys, "executable", "/opt/miniforge3/bin/python3"),
            patch.object(export_onenote, "local_venv_python", return_value=Path("/project/.venv/bin/python")),
        ):
            exit_code = export_onenote.main(
                ["--client-id", "abc"],
                token_provider=token_provider,
            )

        self.assertEqual(exit_code, 1)
        self.assertEqual(
            [call.args[0] for call in print_mock.call_args_list],
            [
                "[ERROR] Missing dependency 'msal' in the active Python interpreter.",
                "[INFO] Active Python: /opt/miniforge3/bin/python3",
                "[INFO] Project venv Python: /project/.venv/bin/python",
                "[RECOMMENDATION] Run with the project venv Python: /project/.venv/bin/python main.py",
                "[RECOMMENDATION] Or activate the venv before running commands: source .venv/bin/activate",
            ],
        )

    def test_main_explains_tenant_missing_auth_error(self) -> None:
        auth_error = RuntimeError(
            "Could not start Microsoft device login: "
            "{'error': 'invalid_request', 'error_description': 'AADSTS50059: "
            "No tenant-identifying information found'}"
        )
        token_provider = Mock(side_effect=auth_error)

        with patch("builtins.print") as print_mock:
            exit_code = export_onenote.main(
                ["--client-id", "abc"],
                token_provider=token_provider,
            )

        self.assertEqual(exit_code, 1)
        self.assertEqual(
            [call.args[0] for call in print_mock.call_args_list],
            [
                "[ERROR] Microsoft login did not receive tenant-identifying information.",
                "[RECOMMENDATION] Set ONENOTE_TENANT_ID=organizations in .env, or use your Directory (tenant) ID.",
                "[INFO] Original error: AADSTS50059",
            ],
        )

    def test_log_device_flow_prints_code_and_url_explicitly(self) -> None:
        flow = {
            "user_code": "SRNMMXBNA",
            "verification_uri": "https://login.microsoft.com/device",
            "expires_in": 1800,
        }

        with patch("builtins.print") as print_mock:
            export_onenote.log_device_flow(flow)

        self.assertEqual(
            [call.args[0] for call in print_mock.call_args_list],
            [
                "[ACTION] Open this URL in your browser: https://login.microsoft.com/device",
                "[DEVICE CODE] SRNMMXBNA",
                "[ACTION] Paste the device code above into the Microsoft page, then click Next.",
                "[INFO] The code is printed here in the terminal. It is not in Teams or OneNote.",
                "[INFO] Code expires in about 30 minutes.",
            ],
        )

    def test_main_lists_notebooks_without_exporting_pages(self) -> None:
        token_provider = Mock(return_value="token")
        client = Mock()
        client.list_notebooks.return_value = [
            {"displayName": "Course Notes", "isShared": True, "userRole": "Reader"}
        ]

        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch("builtins.print") as print_mock,
        ):
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

    def test_main_prints_site_id_helper_for_site_url_by_default(self) -> None:
        token_provider = Mock(return_value="token")
        client = Mock()

        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch("builtins.print") as print_mock,
        ):
            exit_code = export_onenote.main(
                [
                    "--client-id",
                    "abc",
                    "--site-url",
                    "https://school.sharepoint.com/teams/2026-GDD-542/Shared%20Documents",
                    "--list",
                ],
                token_provider=token_provider,
                client_factory=lambda token: client,
            )

        self.assertEqual(exit_code, 0)
        token_provider.assert_not_called()
        client.list_notebooks.assert_not_called()
        self.assertEqual(
            [call.args[0] for call in print_mock.call_args_list],
            [
                "",
                "== Detected notebook storage site (for checking only) ==",
                "This is the Teams/SharePoint site that stores the notebook file. You usually do not need to open it.",
                "https://school.sharepoint.com/teams/2026-GDD-542",
                "",
                "== Step 1 (SITE_GUID): open this in your signed-in browser ==",
                "https://school.sharepoint.com/teams/2026-GDD-542/_api/site/id",
                "",
                "== Step 2 (WEB_GUID): open this in your signed-in browser ==",
                "https://school.sharepoint.com/teams/2026-GDD-542/_api/web/id",
                "",
                "== Step 3: copy the two GUID values, then run ==",
                'python main.py --site-id "school.sharepoint.com,SITE_GUID,WEB_GUID" --list',
                "",
            ],
        )

    def test_main_uses_site_id_without_sites_read_scope(self) -> None:
        token_provider = Mock(return_value="token")
        client = Mock()
        client.list_notebooks.return_value = [
            {"displayName": "2026-1 CSX4107(541) Notebook", "isShared": False, "userRole": "Owner"}
        ]

        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch("builtins.print") as print_mock,
        ):
            exit_code = export_onenote.main(
                [
                    "--client-id",
                    "abc",
                    "--site-id",
                    "school.sharepoint.com,site-guid,web-guid",
                    "--list",
                ],
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
        client.json.assert_not_called()
        client.list_notebooks.assert_called_once_with(
            "/sites/school.sharepoint.com,site-guid,web-guid"
        )
        self.assertEqual(
            [call.args[0] for call in print_mock.call_args_list],
            [
                "1. 2026-1 CSX4107(541) Notebook | shared=False | role=Owner",
                "",
                "== To download one notebook ==",
                'python main.py --site-id "school.sharepoint.com,site-guid,web-guid" --notebook "2026-1 CSX4107(541) Notebook"',
            ],
        )

    def test_main_rejects_invalid_site_url_cleanly(self) -> None:
        stderr = io.StringIO()
        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch("sys.stderr", stderr),
        ):
            exit_code = export_onenote.main(
                ["--client-id", "abc", "--site-url", "https://school.sharepoint.com/_layouts/15/start.aspx"],
                token_provider=Mock(return_value="token"),
            )

        self.assertEqual(exit_code, 1)
        self.assertEqual(
            stderr.getvalue(),
            "Could not infer a SharePoint site from that URL. Expected a URL containing "
            "/sites/<name>/... or /teams/<name>/...\n"
        )

    def test_parse_args_rejects_unknown_formats(self) -> None:
        with self.assertRaises(argparse.ArgumentTypeError):
            export_onenote.parse_formats("pdf")


if __name__ == "__main__":
    unittest.main()
