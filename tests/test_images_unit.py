import base64
import io
import tempfile
import unittest
import zipfile
from types import SimpleNamespace

from mcp_google_sheets import images

# 1x1 transparent PNG.
TINY_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="
)

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def build_xlsx(col: int = 2, row: int = 2, sheet_name: str = "Sheet1") -> bytes:
    """Build a minimal XLSX holding one anchored picture, as Drive exports one."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr(
            "xl/workbook.xml",
            f'<workbook xmlns="{MAIN_NS}" xmlns:r="{R_NS}">'
            f'<sheets><sheet name="{sheet_name}" sheetId="1" r:id="rId1"/></sheets>'
            "</workbook>",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            f'<Relationships xmlns="{REL_NS}">'
            '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/>'
            "</Relationships>",
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            f'<worksheet xmlns="{MAIN_NS}" xmlns:r="{R_NS}">'
            '<sheetData/><drawing r:id="rId9"/>'
            "</worksheet>",
        )
        archive.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            f'<Relationships xmlns="{REL_NS}">'
            '<Relationship Id="rId9" Target="../drawings/drawing1.xml"/>'
            "</Relationships>",
        )
        archive.writestr(
            "xl/drawings/drawing1.xml",
            f'<wsDr xmlns:xdr="{XDR_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">'
            "<xdr:twoCellAnchor>"
            f"<xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            '<xdr:pic><xdr:blipFill><a:blip r:embed="rId3"/></xdr:blipFill></xdr:pic>'
            "</xdr:twoCellAnchor></wsDr>",
        )
        archive.writestr(
            "xl/drawings/_rels/drawing1.xml.rels",
            f'<Relationships xmlns="{REL_NS}">'
            '<Relationship Id="rId3" Target="../media/image1.png"/>'
            "</Relationships>",
        )
        archive.writestr("xl/media/image1.png", TINY_PNG)
    return buffer.getvalue()


class FakeRequest:
    def __init__(self, result):
        self.result = result

    def execute(self):
        return self.result


class FakeValues:
    def __init__(self, get_result=None):
        self.get_result = get_result or {}
        self.updates = []

    def get(self, **kwargs):
        return FakeRequest(self.get_result)

    def update(self, **kwargs):
        self.updates.append(kwargs)
        return FakeRequest({"updatedRange": kwargs["range"]})


class FakeSpreadsheets:
    def __init__(self, values, metadata=None):
        self._values = values
        self._metadata = metadata or {}

    def values(self):
        return self._values

    def get(self, **kwargs):
        return FakeRequest(self._metadata)


class FakeSheetsService:
    def __init__(self, values, metadata=None):
        self._spreadsheets = FakeSpreadsheets(values, metadata)

    def spreadsheets(self):
        return self._spreadsheets


class FakeFiles:
    def __init__(self, export_bytes=b""):
        self.export_bytes = export_bytes
        self.created = []

    def create(self, **kwargs):
        self.created.append(kwargs)
        return FakeRequest({"id": "drive-file-1", "name": kwargs["body"]["name"]})

    def export(self, **kwargs):
        return FakeRequest(self.export_bytes)


class FakePermissions:
    def __init__(self):
        self.created = []

    def create(self, **kwargs):
        self.created.append(kwargs)
        return FakeRequest({"id": "perm-1"})


class FakeDriveService:
    def __init__(self, export_bytes=b""):
        self._files = FakeFiles(export_bytes)
        self._permissions = FakePermissions()

    def files(self):
        return self._files

    def permissions(self):
        return self._permissions


def make_ctx(sheets_service=None, drive_service=None, folder_id=None):
    lifespan = SimpleNamespace(
        sheets_service=sheets_service,
        drive_service=drive_service,
        folder_id=folder_id,
    )
    return SimpleNamespace(request_context=SimpleNamespace(lifespan_context=lifespan))


class HelperTests(unittest.TestCase):
    def test_sniff_mime_from_magic_bytes(self):
        self.assertEqual(images._sniff_mime(TINY_PNG), "image/png")
        self.assertEqual(images._sniff_mime(b"\xff\xd8\xff\xe0rest"), "image/jpeg")
        self.assertEqual(images._sniff_mime(b"RIFF\x00\x00\x00\x00WEBPVP8 "), "image/webp")

    def test_normalise_cell_accepts_and_rejects(self):
        self.assertEqual(images._normalise_cell("Sheet1!$b$2"), "B2")
        self.assertEqual(images._normalise_cell("aa10"), "AA10")
        with self.assertRaises(ValueError):
            images._normalise_cell("A1:B2")
        with self.assertRaises(ValueError):
            images._normalise_cell("nonsense")

    def test_image_formula_regex(self):
        match = images.IMAGE_FORMULA_RE.match('=IMAGE("https://example.com/a.png", 1)')
        self.assertIsNotNone(match)
        self.assertEqual(match.group(1), "https://example.com/a.png")
        self.assertIsNone(images.IMAGE_FORMULA_RE.match("=SUM(A1:A2)"))

    def test_resolve_source_data_uri(self):
        uri = "data:image/png;base64," + base64.b64encode(TINY_PNG).decode()
        data, mime, name = images._resolve_source(uri)
        self.assertEqual(data, TINY_PNG)
        self.assertEqual(mime, "image/png")
        self.assertTrue(name.startswith("image"))

    def test_resolve_source_rejects_non_image_base64(self):
        with self.assertRaises(ValueError):
            images._resolve_source(base64.b64encode(b"not an image at all").decode())

    def test_rels_resolve_parent_relative_targets(self):
        with zipfile.ZipFile(io.BytesIO(build_xlsx())) as archive:
            rels = images._rels_for(archive, "xl/worksheets/sheet1.xml")
        self.assertEqual(rels["rId9"], "xl/drawings/drawing1.xml")


class EmbeddedImageTests(unittest.TestCase):
    def test_embedded_images_reports_anchor_cell(self):
        found = images._embedded_images(build_xlsx(col=2, row=2), sheet=None)
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0]["sheet"], "Sheet1")
        self.assertEqual(found[0]["cell"], "C3")
        self.assertEqual(found[0]["mediaPath"], "xl/media/image1.png")
        self.assertEqual(found[0]["mimeType"], "image/png")
        self.assertEqual(found[0]["sizeBytes"], len(TINY_PNG))

    def test_embedded_images_filters_by_sheet_name(self):
        xlsx = build_xlsx(sheet_name="Data")
        self.assertEqual(len(images._embedded_images(xlsx, sheet="Data")), 1)
        self.assertEqual(len(images._embedded_images(xlsx, sheet="Other")), 0)


class FormulaImageTests(unittest.TestCase):
    def test_formula_images_finds_image_cells(self):
        values = FakeValues(
            get_result={
                "values": [
                    ["header", "x"],
                    ['=IMAGE("https://example.com/a.png", 1)', "plain"],
                    ["", '=image("https://example.com/b.jpg")'],
                ]
            }
        )
        found = images._formula_images(FakeSheetsService(values), "sid", "Sheet1")
        self.assertEqual(
            [(image["cell"], image["url"]) for image in found],
            [
                ("A2", "https://example.com/a.png"),
                ("B3", "https://example.com/b.jpg"),
            ],
        )


class UploadImageToCellTests(unittest.TestCase):
    def test_local_image_is_uploaded_and_formula_written(self):
        values = FakeValues()
        drive = FakeDriveService()
        ctx = make_ctx(FakeSheetsService(values), drive)
        uri = "data:image/png;base64," + base64.b64encode(TINY_PNG).decode()

        result = images.upload_image_to_cell("sid", "Sheet1", "b2", uri, ctx=ctx)

        self.assertTrue(result["success"])
        self.assertEqual(result["cell"], "Sheet1!B2")
        self.assertEqual(
            result["formula"],
            '=IMAGE("https://lh3.googleusercontent.com/d/drive-file-1", 1)',
        )
        self.assertEqual(values.updates[0]["valueInputOption"], "USER_ENTERED")
        self.assertEqual(len(drive.permissions().created), 1)

    def test_http_source_is_not_copied_to_drive(self):
        values = FakeValues()
        drive = FakeDriveService()
        ctx = make_ctx(FakeSheetsService(values), drive)

        result = images.upload_image_to_cell(
            "sid",
            "Sheet1",
            "A1",
            "https://example.com/a.png",
            mode=4,
            height=80,
            width=60,
            ctx=ctx,
        )

        self.assertEqual(result["formula"], '=IMAGE("https://example.com/a.png", 4, 80, 60)')
        self.assertEqual(drive.files().created, [])

    def test_mode_four_requires_dimensions(self):
        ctx = make_ctx(FakeSheetsService(FakeValues()), FakeDriveService())
        result = images.upload_image_to_cell(
            "sid", "Sheet1", "A1", "https://example.com/a.png", mode=4, ctx=ctx
        )
        self.assertIn("error", result)

    def test_invalid_cell_is_rejected(self):
        ctx = make_ctx(FakeSheetsService(FakeValues()), FakeDriveService())
        result = images.upload_image_to_cell(
            "sid", "Sheet1", "A1:B2", "https://example.com/a.png", ctx=ctx
        )
        self.assertIn("error", result)


class ListAndReadTests(unittest.TestCase):
    def test_list_combines_formula_and_embedded_images(self):
        values = FakeValues(get_result={"values": [['=IMAGE("https://example.com/a.png")']]})
        ctx = make_ctx(FakeSheetsService(values), FakeDriveService(build_xlsx()))

        result = images.list_sheet_images("sid", "Sheet1", ctx=ctx)

        self.assertEqual(result["count"], 2)
        self.assertEqual({image["source"] for image in result["images"]}, {"formula", "embedded"})
        self.assertNotIn("warnings", result)

    def test_read_embedded_image_by_cell_returns_bytes(self):
        values = FakeValues(get_result={"values": [[""]]})
        ctx = make_ctx(FakeSheetsService(values), FakeDriveService(build_xlsx()))

        result = images.read_sheet_image("sid", "Sheet1", cell="C3", ctx=ctx)

        self.assertEqual(result.data, TINY_PNG)

    def test_read_saves_to_requested_path(self):
        values = FakeValues(get_result={"values": [[""]]})
        ctx = make_ctx(FakeSheetsService(values), FakeDriveService(build_xlsx()))
        with tempfile.TemporaryDirectory() as directory:
            target = f"{directory}/out.png"
            result = images.read_sheet_image("sid", "Sheet1", cell="C3", save_to=target, ctx=ctx)
            self.assertEqual(result["savedTo"], target)
            with open(target, "rb") as handle:
                self.assertEqual(handle.read(), TINY_PNG)

    def test_read_reports_missing_image(self):
        values = FakeValues(get_result={"values": [[""]]})
        ctx = make_ctx(FakeSheetsService(values), FakeDriveService(build_xlsx()))

        result = images.read_sheet_image("sid", "Sheet1", cell="Z99", ctx=ctx)

        self.assertIn("error", result)

    def test_read_requires_cell_or_media_path(self):
        ctx = make_ctx(FakeSheetsService(FakeValues()), FakeDriveService())
        self.assertIn("error", images.read_sheet_image("sid", "Sheet1", ctx=ctx))


if __name__ == "__main__":
    unittest.main()
