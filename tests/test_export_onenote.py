from __future__ import annotations

import argparse
import io
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import export_onenote


class FakeInteractiveStdin(io.StringIO):
    def isatty(self) -> bool:
        return True


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


class CliHeadingTests(unittest.TestCase):
    def test_section_heading_accepts_full_named_template(self) -> None:
        with patch.object(export_onenote, "SECTION_HEADING_DECORATION_TEXT", ">>> {text}"):
            result = export_onenote.section_heading("Step 1")

        self.assertEqual(result, ">>> Step 1")

    def test_section_heading_accepts_plain_decoration_token(self) -> None:
        with patch.object(export_onenote, "SECTION_HEADING_DECORATION_TEXT", "##"):
            result = export_onenote.section_heading("Step 1")

        self.assertEqual(result, "## Step 1 ##")

    def test_copy_block_indents_copyable_values(self) -> None:
        result = export_onenote.copy_block('python main.py --site-id "abc" --list')

        self.assertEqual(result, '    python main.py --site-id "abc" --list')

    def test_info_box_formats_info_values_with_rounded_unicode_box(self) -> None:
        result = export_onenote.info_box(["1. Course Notebook", "2. Lab Notebook"])

        self.assertEqual(
            result,
            "\n".join(
                [
                    "╭────────────────────╮",
                    "│ 1. Course Notebook │",
                    "│ 2. Lab Notebook    │",
                    "╰────────────────────╯",
                ]
            ),
        )

    def test_error_box_formats_error_values_with_title_in_border(self) -> None:
        result = export_onenote.error_box(["[ERROR] SITE_GUID was not found in that paste."])

        self.assertEqual(
            result,
            "\n".join(
                [
                    "╭───────────[ERROR]──────────────────────────────╮",
                    "│ [ERROR] SITE_GUID was not found in that paste. │",
                    "╰────────────────────────────────────────────────╯",
                ]
            ),
        )

    def test_read_pasted_guid_accepts_one_enter_after_guid_line(self) -> None:
        stdin = io.StringIO(
            '<d:Id m:type="Edm.Guid">80a26a44-cf5b-42b2-bf61-c3a021fa18c7</d:Id>\n'
        )

        with patch("builtins.print") as print_mock:
            result = export_onenote.read_pasted_guid("SITE_GUID", stdin)

        self.assertEqual(result, "80a26a44-cf5b-42b2-bf61-c3a021fa18c7")
        self.assertIn(
            export_onenote.info_box(
                [
                    "After copying the whole XML text and pasting it here, press Return/Enter to continue.",
                ]
            ),
            [call.args[0] for call in print_mock.call_args_list],
        )

    def test_read_pasted_guid_accepts_one_enter_after_browser_text(self) -> None:
        stdin = io.StringIO(
            "This XML file does not appear to have any style information associated with it.\n"
            '<d:Id m:type="Edm.Guid">5dbbcfdd-641d-42ed-b89a-2cb2451897ef</d:Id>\n'
        )

        with patch("builtins.print"):
            result = export_onenote.read_pasted_guid("WEB_GUID", stdin)

        self.assertEqual(result, "5dbbcfdd-641d-42ed-b89a-2cb2451897ef")


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
    def test_site_url_next_flag_prefers_notebook_then_list(self) -> None:
        self.assertEqual(
            export_onenote.site_url_next_flag(notebook="Course Notes"),
            '--notebook "Course Notes"',
        )
        self.assertEqual(
            export_onenote.site_url_next_flag(notebook=None),
            "--list",
        )

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

    def test_sharepoint_url_to_site_id_helper_urls_supports_sharepoint_redirect_links(self) -> None:
        url = (
            "https://school.sharepoint.com/:o:/r/sites/Section_123/_layouts/15/Doc.aspx"
            "?sourcedoc={abc}&action=edit"
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

    def test_site_id_to_site_location_rejects_duplicate_site_and_web_guids(self) -> None:
        with self.assertRaisesRegex(ValueError, "SITE_GUID and WEB_GUID are identical"):
            export_onenote.site_id_to_site_location(
                "school.sharepoint.com,80a26a44-cf5b-42b2-bf61-c3a021fa18c7,80a26a44-cf5b-42b2-bf61-c3a021fa18c7"
            )

    def test_extract_sharepoint_guid_accepts_full_browser_xml_text(self) -> None:
        pasted = """
        This XML file does not appear to have any style information associated with it.
        <d:Id xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices"
          m:type="Edm.Guid">80a26a44-cf5b-42b2-bf61-c3a021fa18c7</d:Id>
        """

        result = export_onenote.extract_sharepoint_guid(pasted, "SITE_GUID")

        self.assertEqual(result, "80a26a44-cf5b-42b2-bf61-c3a021fa18c7")

    def test_extract_sharepoint_guid_accepts_xml_only(self) -> None:
        pasted = (
            '<d:Id xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices" '
            'm:type="Edm.Guid">5dbbcfdd-641d-42ed-b89a-2cb2451897ef</d:Id>'
        )

        result = export_onenote.extract_sharepoint_guid(pasted, "WEB_GUID")

        self.assertEqual(result, "5dbbcfdd-641d-42ed-b89a-2cb2451897ef")

    def test_extract_sharepoint_guid_explains_missing_guid(self) -> None:
        with self.assertRaisesRegex(
            ValueError,
            r"\[ERROR\] WEB_GUID was not found in that paste\.",
        ):
            export_onenote.extract_sharepoint_guid("not the XML page", "WEB_GUID")

    def test_extract_sharepoint_guid_explains_collapsed_xml(self) -> None:
        collapsed_xml = (
            'This XML file does not appear to have any style information associated with it.\n'
            '<m:error xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata">\n'
            "...\n"
            "</m:error>"
        )

        with self.assertRaisesRegex(
            ValueError,
            r"Click the triangle/arrow next to the XML line in the browser to expand it",
        ):
            export_onenote.extract_sharepoint_guid(collapsed_xml, "SITE_GUID")

    def test_extract_sharepoint_guid_explains_unauthorized_xml(self) -> None:
        unauthorized_xml = (
            'This XML file does not appear to have any style information associated with it.\n'
            '<m:error xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata">\n'
            "<m:code>-2147024891, System.UnauthorizedAccessException</m:code>\n"
            '<m:message xml:lang="en-US">Attempted to perform an unauthorized operation.</m:message>\n'
            "</m:error>"
        )

        with self.assertRaisesRegex(
            ValueError,
            r"Open the link in a browser signed in with the Assumption University Microsoft account",
        ):
            export_onenote.extract_sharepoint_guid(unauthorized_xml, "SITE_GUID")


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
    def test_filter_notebooks_matches_display_name_case_insensitively(self) -> None:
        notebooks = [
            {"displayName": "2026-1 BAD 542 Notebook"},
            {"displayName": "Personal Notes"},
        ]

        self.assertEqual(
            export_onenote.filter_notebooks(notebooks, "bad 542"),
            [{"displayName": "2026-1 BAD 542 Notebook"}],
        )
        self.assertEqual(export_onenote.filter_notebooks(notebooks, None), notebooks)

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

    def test_convert_with_pandoc_removes_cleaned_html_when_pandoc_raises(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            output_base = Path(tmpdir) / "sample"
            html_path = Path(f"{output_base}.html")
            html_path.write_text("<html><body>Hello</body></html>", encoding="utf-8")
            cleaned_path = Path(f"{output_base}.cleaned.html")

            with (
                patch.object(export_onenote.shutil, "which", return_value="/usr/bin/pandoc"),
                patch.object(export_onenote.subprocess, "run", side_effect=RuntimeError("boom")),
                self.assertRaises(RuntimeError),
            ):
                export_onenote.convert_with_pandoc(html_path, output_base, ["md"])

            self.assertFalse(cleaned_path.exists())

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
                export_onenote.error_box(
                    ["[ERROR] No notebooks matched filter: 2026-1 GDD 542 Notebook"]
                ),
                "[INFO] Notebooks visible at /me:",
                "  - 2026-1 GDD542 Notebook",
                "  - Personal Notes",
                "[RECOMMENDATION] Copy one of the names above exactly, or run with --list.",
                "[RECOMMENDATION] If your class notebook is not listed, it may live under a Microsoft 365 group or SharePoint site.",
            ],
        )


class CliTests(unittest.TestCase):
    def test_env_example_documents_portable_env_keys_only(self) -> None:
        env_example = Path(".env.example").read_text(encoding="utf-8")
        env_keys = [
            line.split("=", 1)[0]
            for line in env_example.splitlines()
            if line.startswith("ONENOTE_") and "=" in line
        ]

        self.assertEqual(
            env_keys,
            [
                "ONENOTE_CLIENT_ID",
                "ONENOTE_TENANT_ID",
                "ONENOTE_SITE_ID",
                "ONENOTE_OUT",
                "ONENOTE_FORMATS",
                "ONENOTE_NOTEBOOK",
                "ONENOTE_TOKEN_CACHE",
            ],
        )
        self.assertNotIn("ONENOTE_SITE_URL=", env_example)

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

    def test_resolve_tenant_id_uses_au_tenant_for_preset_app(self) -> None:
        self.assertEqual(
            export_onenote.resolve_tenant_id(
                export_onenote.ASSUMPTION_UNIVERSITY_CLIENT_ID,
                "organizations",
            ),
            export_onenote.ASSUMPTION_UNIVERSITY_TENANT_ID,
        )
        self.assertEqual(
            export_onenote.resolve_tenant_id("abc", "organizations"),
            "organizations",
        )

    def test_parse_formats_treats_html_as_implicit_noop(self) -> None:
        self.assertEqual(export_onenote.parse_formats("html"), [])
        self.assertEqual(export_onenote.parse_formats("html,md,rtf"), ["md", "rtf"])

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

    def test_build_export_context_uses_site_id_or_me_location(self) -> None:
        site_args = argparse.Namespace(site_id="school.sharepoint.com,site-guid,web-guid")
        me_args = argparse.Namespace(site_id=None)

        site_context = export_onenote.build_export_context(site_args)
        me_context = export_onenote.build_export_context(me_args)

        self.assertEqual(
            site_context.location,
            "/sites/school.sharepoint.com,site-guid,web-guid",
        )
        self.assertEqual(
            site_context.export_command_base,
            'python main.py --site-id "school.sharepoint.com,site-guid,web-guid"',
        )
        self.assertEqual(me_context.location, "/me")
        self.assertEqual(me_context.export_command_base, "python main.py")

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
                export_onenote.error_box(["[ERROR] Missing Microsoft Entra application/client ID."]),
                "[RECOMMENDATION] Set ONENOTE_CLIENT_ID or pass --client-id.",
            ],
        )

    def test_main_explains_missing_dependency_interpreter_mismatch(self) -> None:
        token_provider = Mock(side_effect=export_onenote.MissingDependencyError("msal"))

        with (
            patch("builtins.print") as print_mock,
            patch.object(export_onenote.sys, "executable", "/opt/miniforge3/bin/python3"),
        ):
            exit_code = export_onenote.main(
                ["--client-id", "abc"],
                token_provider=token_provider,
            )

        venv_python = Path(export_onenote.__file__).resolve().parent / ".venv" / "bin" / "python"
        self.assertEqual(exit_code, 1)
        self.assertEqual(
            [call.args[0] for call in print_mock.call_args_list],
            [
                export_onenote.error_box(
                    ["[ERROR] Missing dependency 'msal' in the active Python interpreter."]
                ),
                "[INFO] Active Python: /opt/miniforge3/bin/python3",
                f"[INFO] Project venv Python: {venv_python}",
                f"[RECOMMENDATION] Run with the project venv Python: {venv_python} main.py",
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
                export_onenote.error_box(
                    ["[ERROR] Microsoft login did not receive tenant-identifying information."]
                ),
                "[RECOMMENDATION] For the Assumption University preset app, set "
                f"ONENOTE_TENANT_ID={export_onenote.ASSUMPTION_UNIVERSITY_TENANT_ID} in .env.",
                "[RECOMMENDATION] After changing .env, delete .msal_token_cache.json and run the command again.",
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
                export_onenote.info_box(
                    [
                        "This is a one-time setup process for a new user.",
                        "Microsoft handles the authentication.",
                    ]
                ),
                "[ACTION] Open this URL in your browser: https://login.microsoft.com/device",
                "[DEVICE CODE]",
                "",
                export_onenote.copy_block("SRNMMXBNA"),
                "",
                "[ACTION] Paste the device code above into the Microsoft page, then click Next.",
                "[INFO] The code is printed here in the terminal. It is not in Teams or OneNote.",
                "[INFO] Code expires in about 30 minutes.",
                "[INFO] After login finishes in the browser, this program will continue automatically.",
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

    def test_main_resolves_au_preset_app_to_au_tenant(self) -> None:
        token_provider = Mock(return_value="token")
        client = Mock()
        client.list_notebooks.return_value = []

        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch("builtins.print"),
        ):
            exit_code = export_onenote.main(
                [
                    "--client-id",
                    export_onenote.ASSUMPTION_UNIVERSITY_CLIENT_ID,
                    "--tenant-id",
                    "organizations",
                    "--list",
                ],
                token_provider=token_provider,
                client_factory=lambda token: client,
            )

        self.assertEqual(exit_code, 0)
        token_provider.assert_called_once_with(
            client_id=export_onenote.ASSUMPTION_UNIVERSITY_CLIENT_ID,
            tenant_id=export_onenote.ASSUMPTION_UNIVERSITY_TENANT_ID,
            scopes=["Notes.Read.All"],
            cache_path=Path(".msal_token_cache.json"),
        )

    def test_main_prints_site_id_helper_for_site_url_when_not_interactive(self) -> None:
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
                export_onenote.section_heading(
                    "Step 1 (SITE_GUID): open this in your signed-in browser and copy the whole text."
                ),
                "",
                export_onenote.info_box(["https://school.sharepoint.com/teams/2026-GDD-542/_api/site/id"]),
                "",
                export_onenote.section_heading(
                    "Step 2 (WEB_GUID): open this in your signed-in browser and copy the whole text."
                ),
                "",
                export_onenote.info_box(["https://school.sharepoint.com/teams/2026-GDD-542/_api/web/id"]),
                "",
                export_onenote.section_heading("Step 3: copy the two GUID values, then run"),
                "",
                export_onenote.copy_block('python main.py --site-id "school.sharepoint.com,SITE_GUID,WEB_GUID" --list'),
                "",
            ],
        )

    def test_main_prompts_for_site_url_guids_and_lists_notebooks_when_interactive(self) -> None:
        print_events: list[str] = []

        def record_token_provider(**_: object) -> str:
            print_events.append("token_provider")
            return "token"

        token_provider = Mock(side_effect=record_token_provider)

        def record_print(*args: object, **_: object) -> None:
            print_events.append(str(args[0]) if args else "")

        client = Mock()
        client.list_notebooks.return_value = [
            {"displayName": "2026-1 BAD 542 Notebook", "isShared": False, "userRole": "Owner"}
        ]
        stdin = FakeInteractiveStdin(
            "\n".join(
                [
                    "This XML file does not appear to have any style information associated with it.",
                    '<d:Id m:type="Edm.Guid">80a26a44-cf5b-42b2-bf61-c3a021fa18c7</d:Id>',
                    "",
                    '<d:Id m:type="Edm.Guid">5dbbcfdd-641d-42ed-b89a-2cb2451897ef</d:Id>',
                    "",
                ]
            )
        )

        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch("builtins.print", side_effect=record_print) as print_mock,
            patch("sys.stdin", stdin),
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
        token_provider.assert_called_once_with(
            client_id="abc",
            tenant_id="organizations",
            scopes=["Notes.Read.All"],
            cache_path=Path(".msal_token_cache.json"),
        )
        client.list_notebooks.assert_called_once_with(
            "/sites/school.sharepoint.com,80a26a44-cf5b-42b2-bf61-c3a021fa18c7,5dbbcfdd-641d-42ed-b89a-2cb2451897ef"
        )
        self.assertIn(
            export_onenote.info_box(["https://school.sharepoint.com/teams/2026-GDD-542/_api/site/id"]),
            [call.args[0] for call in print_mock.call_args_list],
        )
        first_helper_prompt_index = print_events.index(
            export_onenote.section_heading(
                "Step 1 (SITE_GUID): open this in your signed-in browser and copy the whole text."
            )
        )
        token_provider_index = print_events.index("token_provider")
        self.assertLess(token_provider_index, first_helper_prompt_index)
        self.assertIn(
            export_onenote.info_box(["https://school.sharepoint.com/teams/2026-GDD-542/_api/web/id"]),
            [call.args[0] for call in print_mock.call_args_list],
        )
        resolved_site_id = (
            "school.sharepoint.com,"
            "80a26a44-cf5b-42b2-bf61-c3a021fa18c7,"
            "5dbbcfdd-641d-42ed-b89a-2cb2451897ef"
        )
        self.assertIn(
            export_onenote.info_box([resolved_site_id]),
            [call.args[0] for call in print_mock.call_args_list],
        )
        self.assertIn(
            export_onenote.copy_block(
                'python main.py --site-id "school.sharepoint.com,80a26a44-cf5b-42b2-bf61-c3a021fa18c7,5dbbcfdd-641d-42ed-b89a-2cb2451897ef" --list'
            ),
            [call.args[0] for call in print_mock.call_args_list],
        )

    def test_main_auto_exports_single_site_url_notebook_as_html_when_not_list_only(self) -> None:
        token_provider = Mock(return_value="token")
        client = Mock()
        client.list_notebooks.return_value = [
            {"displayName": "2026-1 BAD 542 Notebook", "isShared": False, "userRole": "Owner"}
        ]
        stdin = FakeInteractiveStdin(
            "\n".join(
                [
                    '<d:Id m:type="Edm.Guid">80a26a44-cf5b-42b2-bf61-c3a021fa18c7</d:Id>',
                    "",
                    '<d:Id m:type="Edm.Guid">5dbbcfdd-641d-42ed-b89a-2cb2451897ef</d:Id>',
                    "",
                ]
            )
        )

        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch.object(export_onenote, "export_notebooks", return_value=1) as export_mock,
            patch("builtins.print") as print_mock,
            patch("sys.stdin", stdin),
        ):
            exit_code = export_onenote.main(
                [
                    "--client-id",
                    "abc",
                    "--site-url",
                    "https://school.sharepoint.com/teams/2026-GDD-542/Shared%20Documents",
                ],
                token_provider=token_provider,
                client_factory=lambda token: client,
            )

        self.assertEqual(exit_code, 0)
        export_mock.assert_called_once_with(
            client,
            location="/sites/school.sharepoint.com,80a26a44-cf5b-42b2-bf61-c3a021fa18c7,5dbbcfdd-641d-42ed-b89a-2cb2451897ef",
            output_dir=Path("onenote_export").resolve(),
            notebook_filter="2026-1 BAD 542 Notebook",
            formats=[],
            include_image_links=False,
        )
        self.assertIn(
            export_onenote.section_heading("Auto-downloading the only notebook"),
            [call.args[0] for call in print_mock.call_args_list],
        )
        self.assertIn(
            export_onenote.copy_block(
                'python main.py --site-id "school.sharepoint.com,80a26a44-cf5b-42b2-bf61-c3a021fa18c7,5dbbcfdd-641d-42ed-b89a-2cb2451897ef" --notebook "2026-1 BAD 542 Notebook" --formats md'
            ),
            [call.args[0] for call in print_mock.call_args_list],
        )
        self.assertIn(
            export_onenote.copy_block(
                'python main.py --site-id "school.sharepoint.com,80a26a44-cf5b-42b2-bf61-c3a021fa18c7,5dbbcfdd-641d-42ed-b89a-2cb2451897ef" --notebook "2026-1 BAD 542 Notebook" --formats rtf'
            ),
            [call.args[0] for call in print_mock.call_args_list],
        )
        self.assertIn(
            export_onenote.copy_block(
                'python main.py --site-id "school.sharepoint.com,80a26a44-cf5b-42b2-bf61-c3a021fa18c7,5dbbcfdd-641d-42ed-b89a-2cb2451897ef" --notebook "2026-1 BAD 542 Notebook" --formats txt'
            ),
            [call.args[0] for call in print_mock.call_args_list],
        )

    def test_main_auto_export_honors_formats_before_site_url(self) -> None:
        token_provider = Mock(return_value="token")
        client = Mock()
        client.list_notebooks.return_value = [
            {"displayName": "2026-1 BAD 542 Notebook", "isShared": False, "userRole": "Owner"}
        ]
        stdin = FakeInteractiveStdin(
            "\n".join(
                [
                    '<d:Id m:type="Edm.Guid">80a26a44-cf5b-42b2-bf61-c3a021fa18c7</d:Id>',
                    "",
                    '<d:Id m:type="Edm.Guid">5dbbcfdd-641d-42ed-b89a-2cb2451897ef</d:Id>',
                    "",
                ]
            )
        )

        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch.object(export_onenote, "export_notebooks", return_value=1) as export_mock,
            patch("builtins.print"),
            patch("sys.stdin", stdin),
        ):
            exit_code = export_onenote.main(
                [
                    "--client-id",
                    "abc",
                    "--formats",
                    "md",
                    "--site-url",
                    "https://school.sharepoint.com/teams/2026-GDD-542/Shared%20Documents",
                ],
                token_provider=token_provider,
                client_factory=lambda token: client,
            )

        self.assertEqual(exit_code, 0)
        self.assertEqual(export_mock.call_args.kwargs["formats"], ["md"])

    def test_main_returns_clear_error_when_interactive_site_url_paste_has_no_guid(self) -> None:
        stdin = FakeInteractiveStdin("not the XML page\n\n")

        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch("builtins.print") as print_mock,
            patch("sys.stdin", stdin),
        ):
            exit_code = export_onenote.main(
                [
                    "--client-id",
                    "abc",
                    "--site-url",
                    "https://school.sharepoint.com/teams/2026-GDD-542/Shared%20Documents",
                    "--list",
                ],
                token_provider=Mock(return_value="token"),
            )

        self.assertEqual(exit_code, 1)
        self.assertIn(
            export_onenote.error_box(
                [
                    "[ERROR] SITE_GUID was not found in that paste.",
                    "[RECOMMENDATION] Paste the full SharePoint XML page text, including the long value inside <d:Id>...</d:Id>.",
                ]
            ),
            [call.args[0] for call in print_mock.call_args_list],
        )

    def test_main_explains_collapsed_interactive_site_url_xml(self) -> None:
        stdin = FakeInteractiveStdin(
            'This XML file does not appear to have any style information associated with it.\n'
            '<m:error xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata">\n'
            "...\n"
            "</m:error>\n\n"
        )

        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch("builtins.print") as print_mock,
            patch("sys.stdin", stdin),
        ):
            exit_code = export_onenote.main(
                [
                    "--client-id",
                    "abc",
                    "--site-url",
                    "https://school.sharepoint.com/teams/2026-GDD-542/Shared%20Documents",
                    "--list",
                ],
                token_provider=Mock(return_value="token"),
            )

        self.assertEqual(exit_code, 1)
        self.assertIn(
            export_onenote.error_box(
                [
                    "[ERROR] SITE_GUID was not found in that paste.",
                    "[RECOMMENDATION] The pasted XML looks collapsed. Click the triangle/arrow next to "
                    "the XML line in the browser to expand it, then copy and paste the expanded text.",
                ]
            ),
            [call.args[0] for call in print_mock.call_args_list],
        )

    def test_main_explains_unauthorized_interactive_site_url_xml(self) -> None:
        stdin = FakeInteractiveStdin(
            'This XML file does not appear to have any style information associated with it.\n'
            '<m:error xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata">\n'
            "<m:code>-2147024891, System.UnauthorizedAccessException</m:code>\n"
            '<m:message xml:lang="en-US">Attempted to perform an unauthorized operation.</m:message>\n'
            "</m:error>\n\n"
        )

        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch("builtins.print") as print_mock,
            patch("sys.stdin", stdin),
        ):
            exit_code = export_onenote.main(
                [
                    "--client-id",
                    "abc",
                    "--site-url",
                    "https://school.sharepoint.com/teams/2026-GDD-542/Shared%20Documents",
                    "--list",
                ],
                token_provider=Mock(return_value="token"),
            )

        self.assertEqual(exit_code, 1)
        self.assertIn(
            export_onenote.error_box(
                [
                    "[ERROR] SharePoint denied access to the SITE_GUID page.",
                    "[RECOMMENDATION] Open the link in a browser signed in with the Assumption "
                    "University Microsoft account, then copy the page text again.",
                ]
            ),
            [call.args[0] for call in print_mock.call_args_list],
        )

    def test_main_rejects_duplicate_interactive_site_and_web_guid_after_login(self) -> None:
        token_provider = Mock(return_value="token")
        same_guid = "80a26a44-cf5b-42b2-bf61-c3a021fa18c7"
        stdin = FakeInteractiveStdin(
            "\n".join(
                [
                    f'<d:Id m:type="Edm.Guid">{same_guid}</d:Id>',
                    "",
                    f'<d:Id m:type="Edm.Guid">{same_guid}</d:Id>',
                    "",
                ]
            )
        )

        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch("builtins.print") as print_mock,
            patch("sys.stdin", stdin),
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
                client_factory=Mock(),
        )

        self.assertEqual(exit_code, 1)
        token_provider.assert_called_once_with(
            client_id="abc",
            tenant_id="organizations",
            scopes=["Notes.Read.All"],
            cache_path=Path(".msal_token_cache.json"),
        )
        self.assertIn(
            export_onenote.error_box(
                [
                    "[ERROR] SITE_GUID and WEB_GUID are identical.",
                    "[RECOMMENDATION] The Step 1 SITE_GUID page was probably pasted twice. "
                    "Open the Step 2 WEB_GUID link, then paste that page instead.",
                ]
            ),
            [call.args[0] for call in print_mock.call_args_list],
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
        notebook_line = "1. 2026-1 CSX4107(541) Notebook"
        self.assertEqual(
            [call.args[0] for call in print_mock.call_args_list],
            [
                export_onenote.section_heading("Available notebooks"),
                export_onenote.info_box([notebook_line]),
                "",
                export_onenote.section_heading("To download one notebook"),
                "",
                export_onenote.copy_block(
                    'python main.py --site-id "school.sharepoint.com,site-guid,web-guid" --notebook "2026-1 CSX4107(541) Notebook"'
                ),
                "\n>>> Auto-download will start soon if only 1 notebook is found.",
            ],
        )

    def test_main_rejects_sharepoint_url_in_site_id_before_login(self) -> None:
        token_provider = Mock(return_value="token")
        stderr = io.StringIO()

        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(export_onenote, "load_dotenv", return_value=False),
            patch("sys.stderr", stderr),
        ):
            exit_code = export_onenote.main(
                [
                    "--client-id",
                    "abc",
                    "--site-id",
                    "https://school.sharepoint.com/sites/Section_123/_layouts/15/Doc.aspx?sourcedoc={abc}",
                ],
                token_provider=token_provider,
            )

        self.assertEqual(exit_code, 1)
        token_provider.assert_not_called()
        self.assertEqual(
            stderr.getvalue(),
            "--site-id expects a resolved Graph site ID like "
            "hostname,siteCollectionGuid,webGuid. Use --site-url for SharePoint URLs.\n",
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
