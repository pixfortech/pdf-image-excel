"""Turn a PDF or image into positioned words.

Whatever the source, the output is the same: pages of words with their x/y
positions, plus any ruled table-row boxes pdfplumber can see (used to find
exact column boundaries of a boxed header row).  Keeping positions is what
lets the parser assign values to columns by layout instead of by token order.

Strategy: pdfplumber words -> PyMuPDF words (if pdfplumber finds little text)
-> OCR with word boxes (scanned PDFs and images).
"""
from __future__ import annotations

import io
import os
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
_MIN_WORDS_PER_PAGE = 5
_OCR_DPI = 300
_LOW_OCR_CONFIDENCE = 70.0


@dataclass
class Word:
    text: str
    x0: float
    x1: float
    top: float
    bottom: float

    @property
    def xc(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def yc(self) -> float:
        return (self.top + self.bottom) / 2

    @property
    def height(self) -> float:
        return max(self.bottom - self.top, 1.0)


Box = Tuple[float, float, float, float]  # x0, top, x1, bottom


@dataclass
class Page:
    number: int
    words: List[Word] = field(default_factory=list)
    ruled_rows: List[List[Box]] = field(default_factory=list)  # first row of each ruled table


@dataclass
class Document:
    pages: List[Page] = field(default_factory=list)
    engine: str = ""
    warnings: List[str] = field(default_factory=list)
    ocr_confidence: Optional[float] = None

    @property
    def word_count(self) -> int:
        return sum(len(p.words) for p in self.pages)


def extract(data: bytes, filename: str) -> Document:
    ext = os.path.splitext(filename or "")[1].lower()
    if ext in IMAGE_EXTS:
        return _ocr_image_bytes(data)

    doc = _pdfplumber_words(data)
    if _too_little_text(doc):
        fallback = _pymupdf_words(data)
        fallback.warnings = doc.warnings + fallback.warnings
        doc = fallback if fallback.word_count > doc.word_count else doc
    if _too_little_text(doc):
        ocr = _ocr_pdf(data)
        ocr.warnings = doc.warnings + ["Little or no selectable text; used OCR."] + ocr.warnings
        doc = ocr if ocr.word_count > doc.word_count else doc
    return doc


def _too_little_text(doc: Document) -> bool:
    return not doc.pages or doc.word_count < _MIN_WORDS_PER_PAGE * len(doc.pages)


# ---------------------------------------------------------------------------
# Digital PDFs
# ---------------------------------------------------------------------------

def _pdfplumber_words(data: bytes) -> Document:
    doc = Document(engine="pdfplumber")
    try:
        import pdfplumber
    except Exception as exc:  # pragma: no cover - dependency problem
        doc.warnings.append(f"pdfplumber unavailable: {exc}")
        return doc
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for i, pg in enumerate(pdf.pages, start=1):
                page = Page(number=i)
                for w in pg.extract_words():
                    page.words.append(Word(w["text"], w["x0"], w["x1"], w["top"], w["bottom"]))
                try:
                    for table in pg.find_tables():
                        cells = [c for c in table.rows[0].cells if c] if table.rows else []
                        if len(cells) >= 2:
                            page.ruled_rows.append([tuple(c) for c in cells])
                except Exception:  # ruled boxes are an optional hint only
                    pass
                doc.pages.append(page)
    except Exception as exc:
        doc.warnings.append(f"pdfplumber could not read the PDF: {exc}")
    return doc


def _pymupdf_words(data: bytes) -> Document:
    doc = Document(engine="pymupdf")
    try:
        import pymupdf
    except Exception as exc:  # pragma: no cover
        doc.warnings.append(f"PyMuPDF unavailable: {exc}")
        return doc
    try:
        with pymupdf.open(stream=data, filetype="pdf") as pdf:
            for i, pg in enumerate(pdf, start=1):
                page = Page(number=i)
                for x0, y0, x1, y1, text, *_ in pg.get_text("words"):
                    page.words.append(Word(text, x0, x1, y0, y1))
                doc.pages.append(page)
    except Exception as exc:
        doc.warnings.append(f"PyMuPDF could not read the PDF: {exc}")
    return doc


# ---------------------------------------------------------------------------
# OCR (scanned PDFs and images)
# ---------------------------------------------------------------------------

def _ocr_pdf(data: bytes) -> Document:
    doc = Document(engine="ocr")
    try:
        import pymupdf
        from PIL import Image
    except Exception as exc:  # pragma: no cover
        doc.warnings.append(f"OCR needs PyMuPDF and Pillow: {exc}")
        return doc
    confidences = []
    with pymupdf.open(stream=data, filetype="pdf") as pdf:
        for i, pg in enumerate(pdf, start=1):
            pix = pg.get_pixmap(dpi=_OCR_DPI)
            image = Image.open(io.BytesIO(pix.tobytes("png")))
            words, conf, warn = _ocr_words(image, scale=72.0 / _OCR_DPI)
            doc.pages.append(Page(number=i, words=words))
            doc.warnings.extend(warn)
            if conf is not None:
                confidences.append(conf)
            if warn and "unavailable" in warn[0]:
                break
    _finish_ocr(doc, confidences)
    return doc


def _ocr_image_bytes(data: bytes) -> Document:
    doc = Document(engine="ocr")
    try:
        from PIL import Image
        image = Image.open(io.BytesIO(data))
    except Exception as exc:
        doc.warnings.append(f"Could not open image: {exc}")
        return doc
    words, conf, warn = _ocr_words(image, scale=1.0)
    doc.pages.append(Page(number=1, words=words))
    doc.warnings.extend(warn)
    _finish_ocr(doc, [conf] if conf is not None else [])
    return doc


def _ocr_words(image, scale: float):
    """Run Tesseract and return (words with boxes, mean confidence, warnings)."""
    try:
        import pytesseract
        data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
    except Exception as exc:
        return [], None, [f"OCR engine unavailable (install Tesseract + pytesseract): {exc}"]
    words, confs = [], []
    for text, left, top, width, height, conf in zip(
            data["text"], data["left"], data["top"], data["width"], data["height"], data["conf"]):
        text = str(text).strip()
        if not text:
            continue
        words.append(Word(text, left * scale, (left + width) * scale,
                          top * scale, (top + height) * scale))
        try:
            if float(conf) >= 0:
                confs.append(float(conf))
        except (TypeError, ValueError):
            pass
    return words, (sum(confs) / len(confs) if confs else None), []


def _finish_ocr(doc: Document, confidences: list) -> None:
    if confidences:
        doc.ocr_confidence = sum(confidences) / len(confidences)
        if doc.ocr_confidence < _LOW_OCR_CONFIDENCE:
            doc.warnings.append(
                f"Low OCR confidence ({doc.ocr_confidence:.0f}%). Check every value in the preview.")
