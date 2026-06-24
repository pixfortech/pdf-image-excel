"""Document extraction orchestration.

Implements the "smart" extraction strategy:

1.  Try digital text/table extraction with pdfplumber.
2.  Fall back to PyMuPDF for text when pdfplumber yields little/no text.
3.  Fall back to OCR (see :mod:`src.ocr`) for scanned PDFs and images.

The output is a uniform :class:`ExtractionResult` containing per-page text,
detected tables, the engine used, and any warnings.  Parsing into structured
records is handled separately by :mod:`src.parser`, keeping extraction free of
business assumptions.
"""
from __future__ import annotations

import io
import os
from dataclasses import dataclass, field
from typing import List, Optional

from . import ocr as ocr_module
from . import table_detector


IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".tif", ".gif"}


@dataclass
class ExtractionResult:
    pages_text: List[str] = field(default_factory=list)
    tables: List[dict] = field(default_factory=list)   # normalised tables
    engine: str = ""
    ocr_confidence: Optional[float] = None
    warnings: List[str] = field(default_factory=list)
    page_count: int = 0


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def extract(source, *, filename: str = "", mode: str = "auto", min_chars_per_page: int = 20) -> ExtractionResult:
    """Extract content from ``source``.

    ``source`` may be a path (str) or a bytes/file-like object.  ``filename`` is
    used to detect the file type when ``source`` is bytes.  ``mode`` is one of
    ``auto``, ``pdf_text``, ``ocr``, ``table``.
    """
    data, name = _read_source(source, filename)
    ext = os.path.splitext(name)[1].lower()

    if ext in IMAGE_EXTS:
        return _extract_image(data, name)

    # Treat everything else as a PDF.
    if mode == "ocr":
        return _extract_pdf_ocr(data, name)

    result = _extract_pdf_digital(data, name, want_tables=mode in ("auto", "table", "pdf_text"))

    has_text = sum(len((t or "").strip()) for t in result.pages_text) >= min_chars_per_page
    if mode == "auto" and not has_text:
        result.warnings.append("Little or no selectable text found; falling back to OCR.")
        ocr_result = _extract_pdf_ocr(data, name)
        if ocr_result.pages_text and any(t.strip() for t in ocr_result.pages_text):
            ocr_result.warnings = result.warnings + ocr_result.warnings
            return ocr_result
    return result


def _read_source(source, filename):
    if isinstance(source, (bytes, bytearray)):
        return bytes(source), filename or "upload.pdf"
    if hasattr(source, "read"):
        data = source.read()
        name = filename or getattr(source, "name", "upload.pdf")
        return data, name
    # Assume path.
    with open(source, "rb") as fh:
        return fh.read(), filename or os.path.basename(source)


# ---------------------------------------------------------------------------
# PDF: digital text + tables via pdfplumber, PyMuPDF fallback
# ---------------------------------------------------------------------------

def _extract_pdf_digital(data: bytes, name: str, want_tables: bool) -> ExtractionResult:
    result = ExtractionResult(engine="pdfplumber")
    try:
        import pdfplumber
    except Exception as exc:
        result.warnings.append(f"pdfplumber unavailable ({exc}); trying PyMuPDF.")
        return _extract_pdf_pymupdf(data, name)

    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            result.page_count = len(pdf.pages)
            for i, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                result.pages_text.append(text)
                if want_tables:
                    try:
                        for tbl in page.extract_tables() or []:
                            norm = table_detector.normalise_table(tbl, page=i)
                            norm = table_detector.drop_empty_rows(norm)
                            if norm["rows"]:
                                result.tables.append(norm)
                    except Exception as exc:  # pragma: no cover
                        result.warnings.append(f"Table extraction failed on page {i}: {exc}")
    except Exception as exc:
        result.warnings.append(f"pdfplumber failed ({exc}); trying PyMuPDF.")
        return _extract_pdf_pymupdf(data, name)

    if not any(t.strip() for t in result.pages_text):
        # No text at all – let the caller decide about OCR.
        result.warnings.append("pdfplumber found no text.")
    return result


def _extract_pdf_pymupdf(data: bytes, name: str) -> ExtractionResult:
    result = ExtractionResult(engine="pymupdf")
    try:
        import fitz  # PyMuPDF
    except Exception as exc:
        result.warnings.append(f"PyMuPDF unavailable ({exc}).")
        return result
    try:
        doc = fitz.open(stream=data, filetype="pdf")
        result.page_count = doc.page_count
        for page in doc:
            result.pages_text.append(page.get_text("text") or "")
    except Exception as exc:  # pragma: no cover
        result.warnings.append(f"PyMuPDF failed: {exc}")
    return result


# ---------------------------------------------------------------------------
# PDF: OCR per page
# ---------------------------------------------------------------------------

def _extract_pdf_ocr(data: bytes, name: str) -> ExtractionResult:
    result = ExtractionResult(engine="ocr")
    if not ocr_module.ocr_available():
        result.warnings.append("OCR requested but no OCR engine is installed.")
        return result
    try:
        import fitz
    except Exception as exc:
        result.warnings.append(f"PyMuPDF needed to rasterise PDF for OCR is unavailable ({exc}).")
        return result

    try:
        from PIL import Image
    except Exception as exc:
        result.warnings.append(f"Pillow needed for OCR is unavailable ({exc}).")
        return result

    confs: List[float] = []
    try:
        doc = fitz.open(stream=data, filetype="pdf")
        result.page_count = doc.page_count
        for page in doc:
            pix = page.get_pixmap(dpi=200)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            ocr_res = ocr_module.ocr_image(img)
            result.pages_text.append(ocr_res.text)
            result.warnings.extend(ocr_res.warnings)
            if ocr_res.confidence is not None:
                confs.append(ocr_res.confidence)
            # Build a whitespace table to help downstream parsing.
            tbl = table_detector.text_to_table(ocr_res.text, page=page.number + 1)
            tbl = table_detector.drop_empty_rows(tbl)
            if tbl["rows"]:
                tbl["ocr_confidence"] = ocr_res.confidence
                result.tables.append(tbl)
    except Exception as exc:  # pragma: no cover
        result.warnings.append(f"OCR extraction failed: {exc}")

    result.ocr_confidence = sum(confs) / len(confs) if confs else None
    return result


# ---------------------------------------------------------------------------
# Image files
# ---------------------------------------------------------------------------

def _extract_image(data: bytes, name: str) -> ExtractionResult:
    result = ExtractionResult(engine="ocr-image")
    if not ocr_module.ocr_available():
        result.warnings.append("Image OCR requested but no OCR engine is installed.")
        return result
    try:
        from PIL import Image
        img = Image.open(io.BytesIO(data))
    except Exception as exc:
        result.warnings.append(f"Could not open image: {exc}")
        return result

    ocr_res = ocr_module.ocr_image(img)
    result.pages_text.append(ocr_res.text)
    result.warnings.extend(ocr_res.warnings)
    result.ocr_confidence = ocr_res.confidence
    result.page_count = 1
    tbl = table_detector.text_to_table(ocr_res.text, page=1)
    tbl = table_detector.drop_empty_rows(tbl)
    if tbl["rows"]:
        tbl["ocr_confidence"] = ocr_res.confidence
        result.tables.append(tbl)
    return result
