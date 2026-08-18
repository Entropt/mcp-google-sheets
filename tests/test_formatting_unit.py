import unittest
from types import SimpleNamespace

from mcp_google_sheets import formatting


class FakeRequest:
    def __init__(self, result):
        self.result = result

    def execute(self):
        return self.result


class FakeSpreadsheets:
    def __init__(self, metadata, get_result=None):
        self._metadata = metadata
        self._get_result = get_result or {}
        self.batch_updates = []
        self.get_calls = []

    def get(self, **kwargs):
        self.get_calls.append(kwargs)
        # _get_sheet_id asks for the sheet list; get_cell_formats asks for ranges.
        if "ranges" in kwargs:
            return FakeRequest(self._get_result)
        return FakeRequest(self._metadata)

    def batchUpdate(self, **kwargs):  # noqa: N802 - mirrors the Google client's name
        self.batch_updates.append(kwargs)
        return FakeRequest({"replies": [{}]})


class FakeSheetsService:
    def __init__(self, sheet_titles=("Sheet1",), get_result=None):
        metadata = {
            "sheets": [
                {"properties": {"title": title, "sheetId": 100 + index}}
                for index, title in enumerate(sheet_titles)
            ]
        }
        self._spreadsheets = FakeSpreadsheets(metadata, get_result)

    def spreadsheets(self):
        return self._spreadsheets


def make_ctx(sheets_service):
    lifespan = SimpleNamespace(sheets_service=sheets_service, drive_service=None, folder_id=None)
    return SimpleNamespace(request_context=SimpleNamespace(lifespan_context=lifespan))


def sent_request(service, index=0, batch=0):
    return service.spreadsheets().batch_updates[batch]["body"]["requests"][index]


class ColorParsingTests(unittest.TestCase):
    def test_six_digit_hex(self):
        self.assertEqual(
            formatting._parse_color("#ff0000"), {"red": 1.0, "green": 0.0, "blue": 0.0}
        )

    def test_hex_without_hash_and_three_digit_shorthand(self):
        self.assertEqual(formatting._parse_color("00ff00")["green"], 1.0)
        self.assertEqual(formatting._parse_color("#00f"), {"red": 0.0, "green": 0.0, "blue": 1.0})

    def test_named_color(self):
        self.assertEqual(formatting._parse_color("LightGreen"), formatting._parse_color("#d9ead3"))

    def test_dict_passthrough(self):
        self.assertEqual(
            formatting._parse_color({"red": 0.5, "green": 0.25, "blue": 0.0}),
            {"red": 0.5, "green": 0.25, "blue": 0.0},
        )

    def test_clear_sentinels_return_none(self):
        for sentinel in ("clear", "none", "Default", "RESET"):
            self.assertIsNone(formatting._parse_color(sentinel))

    def test_rejects_garbage_and_out_of_range(self):
        with self.assertRaises(ValueError):
            formatting._parse_color("chartreuse")
        with self.assertRaises(ValueError):
            formatting._parse_color("#12345")
        with self.assertRaises(ValueError):
            formatting._parse_color({"red": 2.0})

    def test_hex_round_trips_through_rgb_to_hex(self):
        for hex_color in ("#000000", "#ffffff", "#4a86e8", "#d9ead3"):
            self.assertEqual(formatting._rgb_to_hex(formatting._parse_color(hex_color)), hex_color)


class CellFormatBuildingTests(unittest.TestCase):
    def test_background_only_sets_one_field(self):
        cell_format, fields = formatting._build_cell_format(background_color="red")
        self.assertEqual(fields, ["userEnteredFormat.backgroundColorStyle"])
        self.assertEqual(
            cell_format["backgroundColorStyle"]["rgbColor"],
            {"red": 1.0, "green": 0.0, "blue": 0.0},
        )

    def test_clearing_names_the_field_but_omits_the_value(self):
        cell_format, fields = formatting._build_cell_format(background_color="clear")
        self.assertEqual(fields, ["userEnteredFormat.backgroundColorStyle"])
        self.assertNotIn("backgroundColorStyle", cell_format)

    def test_false_booleans_are_applied_not_skipped(self):
        cell_format, fields = formatting._build_cell_format(bold=False)
        self.assertEqual(fields, ["userEnteredFormat.textFormat.bold"])
        self.assertIs(cell_format["textFormat"]["bold"], False)

    def test_number_format_with_and_without_pattern(self):
        cell_format, fields = formatting._build_cell_format(number_format="PERCENT:0.0%")
        self.assertEqual(cell_format["numberFormat"], {"type": "PERCENT", "pattern": "0.0%"})
        self.assertEqual(fields, ["userEnteredFormat.numberFormat"])

        cell_format, _ = formatting._build_cell_format(number_format="text")
        self.assertEqual(cell_format["numberFormat"], {"type": "TEXT"})

    def test_enums_are_upper_cased_and_validated(self):
        cell_format, _ = formatting._build_cell_format(horizontal_alignment="center")
        self.assertEqual(cell_format["horizontalAlignment"], "CENTER")
        with self.assertRaises(ValueError):
            formatting._build_cell_format(vertical_alignment="sideways")
        with self.assertRaises(ValueError):
            formatting._build_cell_format(wrap_strategy="SQUISH")

    def test_font_size_must_be_positive(self):
        with self.assertRaises(ValueError):
            formatting._build_cell_format(font_size=0)

    def test_nothing_set_yields_no_fields(self):
        cell_format, fields = formatting._build_cell_format()
        self.assertEqual((cell_format, fields), ({}, []))


class RepeatCellRequestTests(unittest.TestCase):
    def test_range_becomes_a_half_open_grid_range(self):
        request = formatting._repeat_cell_request(7, "B2:C10", background_color="blue")
        self.assertEqual(
            request["repeatCell"]["range"],
            {
                "sheetId": 7,
                "startColumnIndex": 1,
                "endColumnIndex": 3,
                "startRowIndex": 1,
                "endRowIndex": 10,
            },
        )

    def test_sheet_prefix_and_dollar_signs_are_stripped(self):
        request = formatting._repeat_cell_request(7, "Sheet1!$A$1", background_color="blue")
        self.assertEqual(
            request["repeatCell"]["range"],
            {
                "sheetId": 7,
                "startColumnIndex": 0,
                "endColumnIndex": 1,
                "startRowIndex": 0,
                "endRowIndex": 1,
            },
        )

    def test_whole_column_range_omits_row_bounds(self):
        request = formatting._repeat_cell_request(7, "C:D", background_color="blue")
        grid_range = request["repeatCell"]["range"]
        self.assertEqual(grid_range["startColumnIndex"], 2)
        self.assertEqual(grid_range["endColumnIndex"], 4)
        self.assertNotIn("startRowIndex", grid_range)

    def test_empty_formatting_is_rejected(self):
        with self.assertRaises(ValueError):
            formatting._repeat_cell_request(7, "A1")


class FormatCellsTests(unittest.TestCase):
    def test_sets_background_and_reports_fields(self):
        service = FakeSheetsService()
        result = formatting.format_cells(
            "sid", "Sheet1", "A1:B2", background_color="#4a86e8", ctx=make_ctx(service)
        )

        self.assertTrue(result["success"])
        self.assertEqual(result["range"], "Sheet1!A1:B2")
        self.assertEqual(result["fields"], "userEnteredFormat.backgroundColorStyle")
        request = sent_request(service)["repeatCell"]
        self.assertEqual(request["range"]["sheetId"], 100)
        self.assertAlmostEqual(
            request["cell"]["userEnteredFormat"]["backgroundColorStyle"]["rgbColor"]["blue"],
            232 / 255.0,
        )

    def test_combined_properties_build_one_request(self):
        service = FakeSheetsService()
        result = formatting.format_cells(
            "sid",
            "Sheet1",
            "A1:D1",
            background_color="#4a86e8",
            text_color="white",
            bold=True,
            horizontal_alignment="center",
            ctx=make_ctx(service),
        )

        self.assertEqual(len(service.spreadsheets().batch_updates[0]["body"]["requests"]), 1)
        cell_format = sent_request(service)["repeatCell"]["cell"]["userEnteredFormat"]
        self.assertIs(cell_format["textFormat"]["bold"], True)
        self.assertEqual(cell_format["horizontalAlignment"], "CENTER")
        self.assertIn("userEnteredFormat.textFormat.foregroundColorStyle", result["fields"])

    def test_unknown_sheet_is_reported(self):
        service = FakeSheetsService(sheet_titles=("Other",))
        result = formatting.format_cells(
            "sid", "Sheet1", "A1", background_color="red", ctx=make_ctx(service)
        )
        self.assertIn("error", result)
        self.assertEqual(service.spreadsheets().batch_updates, [])

    def test_bad_colour_returns_error_without_calling_the_api(self):
        service = FakeSheetsService()
        result = formatting.format_cells(
            "sid", "Sheet1", "A1", background_color="chartreuse", ctx=make_ctx(service)
        )
        self.assertIn("error", result)
        self.assertEqual(service.spreadsheets().batch_updates, [])

    def test_no_properties_returns_error(self):
        service = FakeSheetsService()
        result = formatting.format_cells("sid", "Sheet1", "A1", ctx=make_ctx(service))
        self.assertIn("error", result)
        self.assertEqual(service.spreadsheets().batch_updates, [])


class BatchFormatCellsTests(unittest.TestCase):
    def test_several_ranges_go_out_in_one_call(self):
        service = FakeSheetsService()
        result = formatting.batch_format_cells(
            "sid",
            "Sheet1",
            {
                "A1:D1": {"background_color": "#4a86e8", "text_color": "white", "bold": True},
                "C2:C50": {"background_color": "lightgreen"},
            },
            ctx=make_ctx(service),
        )

        self.assertTrue(result["success"])
        self.assertEqual(len(service.spreadsheets().batch_updates), 1)
        requests = service.spreadsheets().batch_updates[0]["body"]["requests"]
        self.assertEqual(len(requests), 2)
        self.assertEqual(result["rangesFormatted"], ["A1:D1", "C2:C50"])

    def test_empty_formats_rejected(self):
        service = FakeSheetsService()
        result = formatting.batch_format_cells("sid", "Sheet1", {}, ctx=make_ctx(service))
        self.assertIn("error", result)

    def test_unknown_option_is_reported_and_nothing_is_sent(self):
        service = FakeSheetsService()
        result = formatting.batch_format_cells(
            "sid", "Sheet1", {"A1": {"backgroud_color": "red"}}, ctx=make_ctx(service)
        )
        self.assertIn("error", result)
        self.assertEqual(service.spreadsheets().batch_updates, [])

    def test_one_bad_range_aborts_the_whole_batch(self):
        service = FakeSheetsService()
        result = formatting.batch_format_cells(
            "sid",
            "Sheet1",
            {"A1": {"background_color": "red"}, "B2": {"background_color": "nonsense"}},
            ctx=make_ctx(service),
        )
        self.assertIn("error", result)
        self.assertEqual(service.spreadsheets().batch_updates, [])


class GetCellFormatsTests(unittest.TestCase):
    def test_reports_styled_cells_and_skips_defaults(self):
        grid = {
            "sheets": [
                {
                    "data": [
                        {
                            "startRow": 0,
                            "startColumn": 0,
                            "rowData": [
                                {
                                    "values": [
                                        {
                                            "userEnteredFormat": {
                                                "backgroundColorStyle": {
                                                    "rgbColor": {
                                                        "red": 1.0,
                                                        "green": 1.0,
                                                        "blue": 1.0,
                                                    }
                                                },
                                                "textFormat": {"bold": False},
                                            }
                                        },
                                        {
                                            "userEnteredFormat": {
                                                "backgroundColorStyle": {
                                                    "rgbColor": {
                                                        "red": 1.0,
                                                        "green": 0.0,
                                                        "blue": 0.0,
                                                    }
                                                },
                                                "textFormat": {"bold": True},
                                                "horizontalAlignment": "CENTER",
                                            }
                                        },
                                    ]
                                }
                            ],
                        }
                    ]
                }
            ]
        }
        service = FakeSheetsService(get_result=grid)

        result = formatting.get_cell_formats("sid", "Sheet1", "A1:B1", ctx=make_ctx(service))

        self.assertEqual(result["count"], 1)
        self.assertEqual(
            result["cells"][0],
            {
                "backgroundColor": "#ff0000",
                "bold": True,
                "horizontalAlignment": "CENTER",
                "cell": "B1",
            },
        )

    def test_block_offsets_shift_the_reported_cell(self):
        grid = {
            "sheets": [
                {
                    "data": [
                        {
                            "startRow": 4,
                            "startColumn": 2,
                            "rowData": [
                                {
                                    "values": [
                                        {
                                            "userEnteredFormat": {
                                                "backgroundColorStyle": {
                                                    "rgbColor": {
                                                        "red": 0.0,
                                                        "green": 0.0,
                                                        "blue": 1.0,
                                                    }
                                                }
                                            }
                                        }
                                    ]
                                }
                            ],
                        }
                    ]
                }
            ]
        }
        service = FakeSheetsService(get_result=grid)
        result = formatting.get_cell_formats("sid", "Sheet1", "C5", ctx=make_ctx(service))
        self.assertEqual(result["cells"][0]["cell"], "C5")

    def test_unstyled_sheet_reports_nothing(self):
        grid = {"sheets": [{"data": [{"rowData": [{"values": [{"userEnteredFormat": {}}]}]}]}]}
        service = FakeSheetsService(get_result=grid)
        result = formatting.get_cell_formats("sid", "Sheet1", ctx=make_ctx(service))
        self.assertEqual(result, {"sheet": "Sheet1", "count": 0, "cells": []})


if __name__ == "__main__":
    unittest.main()
