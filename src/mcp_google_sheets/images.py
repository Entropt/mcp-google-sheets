#!/usr/bin/env python
"""Image support for the Google Sheets MCP server.

The Google Sheets REST API v4 has no image primitives at all: ``CellData`` has
no image field and ``batchUpdate`` has no add-image request. Only Apps Script
can create over-grid images or native in-cell (``CellImage``) values.

This module works around that with two REST-only paths:

* **Writing** -- the image is uploaded to Google Drive, shared with
  link-readers, and an ``=IMAGE("url")`` formula is written into the target
  cell. The result renders in the cell like a native in-cell image.
* **Reading** -- ``=IMAGE()`` formulas are read straight from the Sheets API,
  while genuine embedded images (over-grid pictures and in-cell images added
  through the UI) are recovered by exporting the spreadsheet to XLSX through
  the Drive API and unpacking ``xl/media`` plus the drawing anchors that say
  which cell each picture sits on.
"""

import base64
import io
import logging
import mimetypes
import os
import re
import tempfile
import zipfile
from typing import Any, Dict, List, Optional
from xml.etree import ElementTree

import httplib2
from googleapiclient.http import MediaIoBaseUpload
from mcp.server.fastmcp import Context, Image
from mcp.types import ToolAnnotations

from .server import _column_index_to_letter, tool

logger = logging.getLogger(__name__)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Drive's files.export endpoint refuses documents whose exported form exceeds
# 10 MB. Spreadsheets with more image payload than that cannot be read back.
EXPORT_SIZE_LIMIT_BYTES = 10 * 1024 * 1024

# Above this size an image is written to a temp file instead of being inlined
# as base64 into the model's context.
INLINE_IMAGE_LIMIT_BYTES = 750 * 1024

MAX_DOWNLOAD_BYTES = 64 * 1024 * 1024

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

# Leading bytes -> mime type, for sources that carry no filename.
MAGIC_BYTES = (
    (b"\x89PNG\r\n\x1a\n", "image/png"),
    (b"\xff\xd8\xff", "image/jpeg"),
    (b"GIF87a", "image/gif"),
    (b"GIF89a", "image/gif"),
    (b"BM", "image/bmp"),
)

IMAGE_FORMULA_RE = re.compile(r'^\s*=\s*IMAGE\s*\(\s*"([^"]+)"', re.IGNORECASE)


def _sniff_mime(data: bytes, hint: Optional[str] = None) -> str:
    """Best-effort mime type for raw image bytes."""
    for magic, mime in MAGIC_BYTES:
        if data.startswith(magic):
            return mime
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "image/webp"
    if data.lstrip()[:5] in (b"<svg ", b"<?xml"):
        return "image/svg+xml"
    if hint:
        guessed, _ = mimetypes.guess_type(hint)
        if guessed and guessed.startswith("image/"):
            return guessed
    return "application/octet-stream"


def _download(url: str) -> bytes:
    """Fetch an http(s) URL, refusing other schemes and oversized bodies.

    Uses httplib2 rather than urllib because it ships its own CA bundle;
    urllib relies on the platform trust store, which is absent in a bare
    Python install on macOS and fails every HTTPS fetch with
    CERTIFICATE_VERIFY_FAILED.
    """
    if not url.lower().startswith(("http://", "https://")):
        raise ValueError(f"Only http(s) URLs can be fetched, got: {url}")

    response, content = httplib2.Http(timeout=60).request(
        url, "GET", headers={"User-Agent": "mcp-google-sheets"}
    )
    if response.status >= 400:
        raise ValueError(f"Fetching {url} returned HTTP {response.status}")

    declared = response.get("content-length")
    if declared and declared.isdigit() and int(declared) > MAX_DOWNLOAD_BYTES:
        raise ValueError(f"Image at {url} exceeds the {MAX_DOWNLOAD_BYTES} byte limit")
    if len(content) > MAX_DOWNLOAD_BYTES:
        raise ValueError(f"Image at {url} exceeds the {MAX_DOWNLOAD_BYTES} byte limit")
    return content


def _resolve_source(source: str) -> tuple[bytes, str, str]:
    """Turn a user-supplied image reference into (bytes, mime, suggested name).

    Accepts a local file path, an http(s) URL, a ``data:`` URI, or a bare
    base64 payload.
    """
    if not source or not source.strip():
        raise ValueError("source must be a file path, http(s) URL, data URI, or base64 string")
    source = source.strip()

    if source.startswith("data:"):
        header, _, payload = source.partition(",")
        if not payload:
            raise ValueError("Malformed data URI: no payload after the comma")
        data = base64.b64decode(payload, validate=False)
        declared = header[5:].split(";")[0]
        mime = declared if declared.startswith("image/") else _sniff_mime(data)
        return data, mime, "image" + (mimetypes.guess_extension(mime) or "")

    if source.lower().startswith(("http://", "https://")):
        data = _download(source)
        mime = _sniff_mime(data, hint=source)
        return data, mime, os.path.basename(source.split("?")[0]) or "image"

    expanded = os.path.expanduser(source)
    if os.path.isfile(expanded):
        with open(expanded, "rb") as handle:
            data = handle.read()
        return data, _sniff_mime(data, hint=expanded), os.path.basename(expanded)

    # Last resort: treat the string as raw base64.
    try:
        data = base64.b64decode(source, validate=True)
    except Exception as exc:  # noqa: BLE001 - surfaced to the caller as a clear message
        raise ValueError(
            f"source is not an existing file, an http(s) URL, a data URI, or valid base64: {exc}"
        ) from exc
    mime = _sniff_mime(data)
    if mime == "application/octet-stream":
        raise ValueError("Decoded base64 payload is not a recognised image format")
    return data, mime, "image" + (mimetypes.guess_extension(mime) or "")


def _upload_to_drive(
    drive_service: Any,
    data: bytes,
    mime: str,
    name: str,
    folder_id: Optional[str],
    share_publicly: bool,
) -> Dict[str, str]:
    """Upload image bytes to Drive and return its id plus a hotlinkable URL."""
    body: Dict[str, Any] = {"name": name}
    if folder_id:
        body["parents"] = [folder_id]

    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=False)
    created = (
        drive_service.files()
        .create(body=body, media_body=media, fields="id,name,webViewLink")
        .execute()
    )
    file_id = created["id"]

    if share_publicly:
        drive_service.permissions().create(
            fileId=file_id, body={"role": "reader", "type": "anyone"}
        ).execute()

    return {
        "fileId": file_id,
        "name": created.get("name", name),
        "webViewLink": created.get("webViewLink", ""),
        # lh3 serves Drive-hosted image bytes directly; the older
        # drive.google.com/uc?export=view form is frequently rejected by the
        # IMAGE() fetcher.
        "imageUrl": f"https://lh3.googleusercontent.com/d/{file_id}",
    }


def _export_xlsx(drive_service: Any, spreadsheet_id: str) -> bytes:
    """Export a Google Sheet to XLSX bytes via the Drive API."""
    return drive_service.files().export(fileId=spreadsheet_id, mimeType=XLSX_MIME).execute()


def _rels_for(archive: zipfile.ZipFile, part_path: str) -> Dict[str, str]:
    """Return {relationship id: resolved archive path} for an OPC part."""
    directory, _, filename = part_path.rpartition("/")
    rels_path = f"{directory}/_rels/{filename}.rels"
    if rels_path not in archive.namelist():
        return {}

    resolved: Dict[str, str] = {}
    root = ElementTree.fromstring(archive.read(rels_path))
    for relationship in root.findall("rel:Relationship", NS):
        target = relationship.get("Target", "")
        rel_id = relationship.get("Id")
        if not rel_id or not target:
            continue
        if target.startswith("/"):
            resolved[rel_id] = target.lstrip("/")
            continue
        base = directory
        while target.startswith("../"):
            target = target[3:]
            base = base.rpartition("/")[0]
        resolved[rel_id] = f"{base}/{target}" if base else target
    return resolved


def _drawing_images(archive: zipfile.ZipFile, drawing_path: str) -> List[Dict[str, Any]]:
    """Extract (cell anchor, media part) pairs from one drawing part."""
    media_rels = _rels_for(archive, drawing_path)
    root = ElementTree.fromstring(archive.read(drawing_path))

    found: List[Dict[str, Any]] = []
    for anchor in root:
        blip = anchor.find(".//a:blip", NS)
        if blip is None:
            continue
        embed_id = blip.get(f"{{{NS['r']}}}embed")
        media_path = media_rels.get(embed_id or "")
        if not media_path:
            continue

        cell = None
        from_node = anchor.find("xdr:from", NS)
        if from_node is not None:
            col_node = from_node.find("xdr:col", NS)
            row_node = from_node.find("xdr:row", NS)
            if col_node is not None and row_node is not None:
                cell = (
                    f"{_column_index_to_letter(int(col_node.text or 0))}"
                    f"{int(row_node.text or 0) + 1}"
                )

        found.append(
            {
                "cell": cell,
                "mediaPath": media_path,
                "anchorType": anchor.tag.rpartition("}")[2],
            }
        )
    return found


def _embedded_images(xlsx_bytes: bytes, sheet: Optional[str]) -> List[Dict[str, Any]]:
    """List every picture embedded in an exported workbook, with cell anchors."""
    images: List[Dict[str, Any]] = []
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as archive:
        names = set(archive.namelist())
        workbook_rels = _rels_for(archive, "xl/workbook.xml")
        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))

        for sheet_node in workbook.findall("main:sheets/main:sheet", NS):
            sheet_name = sheet_node.get("name", "")
            if sheet and sheet_name != sheet:
                continue
            sheet_path = workbook_rels.get(sheet_node.get(f"{{{NS['r']}}}id", ""))
            if not sheet_path or sheet_path not in names:
                continue

            sheet_rels = _rels_for(archive, sheet_path)
            sheet_xml = ElementTree.fromstring(archive.read(sheet_path))
            for drawing_node in sheet_xml.findall("main:drawing", NS):
                drawing_path = sheet_rels.get(drawing_node.get(f"{{{NS['r']}}}id", ""))
                if not drawing_path or drawing_path not in names:
                    continue
                for image in _drawing_images(archive, drawing_path):
                    info = archive.getinfo(image["mediaPath"])
                    images.append(
                        {
                            "sheet": sheet_name,
                            "cell": image["cell"],
                            "source": "embedded",
                            "anchorType": image["anchorType"],
                            "mediaPath": image["mediaPath"],
                            "mimeType": _sniff_mime(
                                archive.read(image["mediaPath"])[:32], hint=image["mediaPath"]
                            ),
                            "sizeBytes": info.file_size,
                        }
                    )
    return images


def _formula_images(
    sheets_service: Any, spreadsheet_id: str, sheet: Optional[str]
) -> List[Dict[str, Any]]:
    """List cells holding an ``=IMAGE("...")`` formula."""
    if sheet:
        sheet_names = [sheet]
    else:
        metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_names = [s["properties"]["title"] for s in metadata.get("sheets", [])]

    images: List[Dict[str, Any]] = []
    for name in sheet_names:
        result = (
            sheets_service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=name, valueRenderOption="FORMULA")
            .execute()
        )
        for row_index, row in enumerate(result.get("values", []), start=1):
            for col_index, value in enumerate(row):
                if not isinstance(value, str):
                    continue
                match = IMAGE_FORMULA_RE.match(value)
                if not match:
                    continue
                images.append(
                    {
                        "sheet": name,
                        "cell": f"{_column_index_to_letter(col_index)}{row_index}",
                        "source": "formula",
                        "url": match.group(1),
                        "formula": value,
                    }
                )
    return images


def _normalise_cell(cell: str) -> str:
    """Validate an A1 single-cell reference and strip any sheet prefix."""
    reference = cell.split("!")[-1].replace("$", "").strip().upper()
    if not re.fullmatch(r"[A-Z]+[0-9]+", reference):
        raise ValueError(f"cell must be a single A1 reference like 'B2', got: {cell}")
    return reference


def _as_result(data: bytes, mime: str, save_to: Optional[str], meta: Dict[str, Any]):
    """Return image bytes inline when small, otherwise write them to disk."""
    if save_to:
        path = os.path.expanduser(save_to)
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        with open(path, "wb") as handle:
            handle.write(data)
        return {**meta, "savedTo": path, "sizeBytes": len(data), "mimeType": mime}

    if len(data) > INLINE_IMAGE_LIMIT_BYTES:
        suffix = mimetypes.guess_extension(mime) or ".bin"
        handle = tempfile.NamedTemporaryFile(prefix="gsheet-image-", suffix=suffix, delete=False)
        with handle:
            handle.write(data)
        return {
            **meta,
            "savedTo": handle.name,
            "sizeBytes": len(data),
            "mimeType": mime,
            "note": (
                f"Image is larger than {INLINE_IMAGE_LIMIT_BYTES} bytes, so it was written to "
                "disk instead of being inlined. Read the file to view it."
            ),
        }

    return Image(data=data, format=(mimetypes.guess_extension(mime) or ".bin").lstrip("."))


@tool(
    annotations=ToolAnnotations(
        title="Upload Image To Cell",
        destructiveHint=True,
    ),
)
def upload_image_to_cell(spreadsheet_id: str,
                         sheet: str,
                         cell: str,
                         source: str,
                         mode: int = 1,
                         height: Optional[int] = None,
                         width: Optional[int] = None,
                         drive_folder_id: Optional[str] = None,
                         share_publicly: bool = True,
                         ctx: Context = None) -> Dict[str, Any]:
    """
    Put an image into a spreadsheet cell.

    The image is uploaded to Google Drive, shared so Sheets can fetch it, and
    referenced from the cell with an =IMAGE() formula. This is the only way to
    place an image with the Sheets REST API, which has no native image request.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL)
        sheet: The name of the sheet
        cell: Single target cell in A1 notation (e.g. 'B2')
        source: Local file path, http(s) URL, data URI, or base64 image data.
            An http(s) URL that Sheets can already reach is used as-is and is
            not copied to Drive.
        mode: IMAGE() sizing mode - 1 fit to cell (default), 2 stretch,
            3 original size, 4 custom size (requires height and width)
        height: Image height in pixels, only used when mode is 4
        width: Image width in pixels, only used when mode is 4
        drive_folder_id: Optional Drive folder to upload into. Defaults to the
            server's configured folder.
        share_publicly: Share the uploaded Drive file with anyone who has the
            link. Sheets cannot render the image without this unless the file
            is already readable by every viewer of the spreadsheet.

    Returns:
        Details of the upload and the formula written into the cell
    """
    if mode not in (1, 2, 3, 4):
        return {"error": "mode must be 1 (fit), 2 (stretch), 3 (original), or 4 (custom)"}
    if mode == 4 and (not height or not width):
        return {"error": "mode 4 requires both height and width in pixels"}

    try:
        target_cell = _normalise_cell(cell)
    except ValueError as exc:
        return {"error": str(exc)}

    lifespan = ctx.request_context.lifespan_context
    sheets_service = lifespan.sheets_service
    drive_service = lifespan.drive_service

    upload: Optional[Dict[str, str]] = None
    stripped = source.strip()
    if stripped.lower().startswith(("http://", "https://")):
        image_url = stripped
    else:
        try:
            data, mime, name = _resolve_source(stripped)
        except ValueError as exc:
            return {"error": str(exc)}
        upload = _upload_to_drive(
            drive_service,
            data,
            mime,
            name,
            drive_folder_id or lifespan.folder_id,
            share_publicly,
        )
        image_url = upload["imageUrl"]

    escaped = image_url.replace('"', '""')
    if mode == 4:
        formula = f'=IMAGE("{escaped}", 4, {height}, {width})'
    else:
        formula = f'=IMAGE("{escaped}", {mode})'

    result = (
        sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet}!{target_cell}",
            valueInputOption="USER_ENTERED",
            body={"values": [[formula]]},
        )
        .execute()
    )

    return {
        "success": True,
        "cell": f"{sheet}!{target_cell}",
        "formula": formula,
        "imageUrl": image_url,
        "drive": upload,
        "updatedRange": result.get("updatedRange"),
    }


@tool(
    annotations=ToolAnnotations(
        title="List Sheet Images",
        readOnlyHint=True,
    ),
)
def list_sheet_images(spreadsheet_id: str,
                      sheet: Optional[str] = None,
                      ctx: Context = None) -> Dict[str, Any]:
    """
    List every image in a spreadsheet, from both storage mechanisms.

    Reports cells holding an =IMAGE() formula, plus pictures genuinely embedded
    in the file (over-grid images and in-cell images added through the Sheets
    UI). Embedded pictures are discovered by exporting the spreadsheet to XLSX
    through the Drive API, because the Sheets REST API does not expose them.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL)
        sheet: Optional sheet name to restrict the scan to

    Returns:
        A list of images with their sheet, anchor cell, source kind, and size
    """
    lifespan = ctx.request_context.lifespan_context

    images: List[Dict[str, Any]] = []
    warnings: List[str] = []

    try:
        images.extend(_formula_images(lifespan.sheets_service, spreadsheet_id, sheet))
    except Exception as exc:  # noqa: BLE001 - partial results beat a hard failure
        warnings.append(f"Could not scan =IMAGE() formulas: {exc}")

    try:
        xlsx_bytes = _export_xlsx(lifespan.drive_service, spreadsheet_id)
        images.extend(_embedded_images(xlsx_bytes, sheet))
    except Exception as exc:  # noqa: BLE001
        warnings.append(
            f"Could not read embedded images via XLSX export: {exc}. Exports above "
            f"{EXPORT_SIZE_LIMIT_BYTES} bytes are rejected by Drive."
        )

    response: Dict[str, Any] = {"count": len(images), "images": images}
    if warnings:
        response["warnings"] = warnings
    return response


@tool(
    annotations=ToolAnnotations(
        title="Read Image From Sheet",
        readOnlyHint=True,
    ),
)
def read_sheet_image(spreadsheet_id: str,
                     sheet: str,
                     cell: Optional[str] = None,
                     media_path: Optional[str] = None,
                     save_to: Optional[str] = None,
                     ctx: Context = None):
    """
    Read one image out of a spreadsheet and return its bytes.

    Give either a cell (the image anchored to or referenced from that cell) or
    a media_path from list_sheet_images. Images up to 750 KB are returned
    inline so they can be viewed directly; anything larger is written to a file
    and its path returned. Pass save_to to always write to a chosen path.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL)
        sheet: The name of the sheet holding the image
        cell: Cell in A1 notation holding an =IMAGE() formula or an anchored picture
        media_path: Archive path from list_sheet_images (e.g. 'xl/media/image1.png')
        save_to: Optional local path to write the image to

    Returns:
        The image itself, or metadata describing where it was saved
    """
    if not cell and not media_path:
        return {"error": "Provide either cell or media_path"}

    lifespan = ctx.request_context.lifespan_context
    target_cell = None

    if cell:
        try:
            target_cell = _normalise_cell(cell)
        except ValueError as exc:
            return {"error": str(exc)}

        formula_result = (
            lifespan.sheets_service.spreadsheets()
            .values()
            .get(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet}!{target_cell}",
                valueRenderOption="FORMULA",
            )
            .execute()
        )
        values = formula_result.get("values", [[]])
        raw = values[0][0] if values and values[0] else ""
        match = IMAGE_FORMULA_RE.match(raw) if isinstance(raw, str) else None
        if match:
            url = match.group(1)
            try:
                data = _download(url)
            except Exception as exc:  # noqa: BLE001
                return {"error": f"Could not fetch {url}: {exc}"}
            return _as_result(
                data,
                _sniff_mime(data, hint=url),
                save_to,
                {"sheet": sheet, "cell": target_cell, "source": "formula", "url": url},
            )

    # No formula in that cell, so look for a picture embedded in the file.
    try:
        xlsx_bytes = _export_xlsx(lifespan.drive_service, spreadsheet_id)
    except Exception as exc:  # noqa: BLE001
        return {"error": f"Could not export the spreadsheet to read embedded images: {exc}"}

    if not media_path:
        candidates = [
            image
            for image in _embedded_images(xlsx_bytes, sheet)
            if image["cell"] == target_cell
        ]
        if not candidates:
            return {
                "error": (
                    f"No =IMAGE() formula and no embedded picture found at {sheet}!{target_cell}. "
                    "Run list_sheet_images to see what this spreadsheet contains."
                )
            }
        media_path = candidates[0]["mediaPath"]

    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as archive:
        if media_path not in archive.namelist():
            return {"error": f"{media_path} is not present in the exported workbook"}
        data = archive.read(media_path)

    return _as_result(
        data,
        _sniff_mime(data, hint=media_path),
        save_to,
        {"sheet": sheet, "cell": target_cell, "source": "embedded", "mediaPath": media_path},
    )


@tool(
    annotations=ToolAnnotations(
        title="Upload Image To Drive",
        destructiveHint=True,
    ),
)
def upload_image_to_drive(source: str,
                          name: Optional[str] = None,
                          drive_folder_id: Optional[str] = None,
                          share_publicly: bool = True,
                          ctx: Context = None) -> Dict[str, Any]:
    """
    Upload an image to Google Drive and return a URL usable in =IMAGE().

    Useful when the same image goes into many cells, or when building the
    formula by hand. upload_image_to_cell does this automatically for a single
    cell.

    Args:
        source: Local file path, http(s) URL, data URI, or base64 image data
        name: Optional file name to store in Drive
        drive_folder_id: Optional Drive folder to upload into
        share_publicly: Share with anyone who has the link, required for Sheets
            to render the image

    Returns:
        The Drive file id and a hotlinkable image URL
    """
    lifespan = ctx.request_context.lifespan_context
    try:
        data, mime, default_name = _resolve_source(source)
    except ValueError as exc:
        return {"error": str(exc)}

    upload = _upload_to_drive(
        lifespan.drive_service,
        data,
        mime,
        name or default_name,
        drive_folder_id or lifespan.folder_id,
        share_publicly,
    )
    return {"success": True, "mimeType": mime, "sizeBytes": len(data), **upload}
