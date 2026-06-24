"""OCR helpers for scanned PDFs and image files.

Uses pytesseract when available, with EasyOCR as an alternative.  Both are
optional dependencies; the module degrades gracefully and reports a clear
message instead of crashing when neither engine (or no image backend) is
installed.  OCR confidence is surfaced so the UI can warn the user.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional


@dataclass
class OCRResult:
    text: str
    confidence: Optional[float]   # 0-100, None if unknown
    engine: str
    warnings: List[str]


def _available_engines() -> List[str]:
    engines = []
    try:
        import pytesseract  # noqa: F401
        engines.append("pytesseract")
    except Exception:
        pass
    try:
        import easyocr  # noqa: F401
        engines.append("easyocr")
    except Exception:
        pass
    return engines


def ocr_available() -> bool:
    return bool(_available_engines())


def ocr_image(image, *, engine: str = "auto", lang: str = "eng") -> OCRResult:
    """Run OCR on a PIL image (or a path) and return text + confidence.

    ``image`` may be a path string or a PIL.Image.Image instance.
    """
    warnings: List[str] = []
    engines = _available_engines()
    if not engines:
        return OCRResult(
            text="",
            confidence=None,
            engine="none",
            warnings=["No OCR engine available. Install pytesseract or easyocr."],
        )

    chosen = engine if engine in engines else engines[0]

    pil_image = _load_image(image, warnings)
    if pil_image is None and chosen == "pytesseract":
        return OCRResult("", None, "none", warnings)

    if chosen == "pytesseract":
        return _ocr_pytesseract(pil_image, lang, warnings)
    return _ocr_easyocr(image, pil_image, warnings)


def _load_image(image, warnings):
    try:
        from PIL import Image
    except Exception:
        warnings.append("Pillow not installed; cannot load image.")
        return None
    if isinstance(image, str):
        try:
            return Image.open(image)
        except Exception as exc:  # pragma: no cover - depends on file
            warnings.append(f"Could not open image: {exc}")
            return None
    return image


def _ocr_pytesseract(pil_image, lang, warnings) -> OCRResult:
    import pytesseract
    try:
        data = pytesseract.image_to_data(pil_image, lang=lang, output_type=pytesseract.Output.DICT)
        words = []
        confs = []
        for txt, conf in zip(data.get("text", []), data.get("conf", [])):
            if str(txt).strip():
                words.append(txt)
                try:
                    c = float(conf)
                    if c >= 0:
                        confs.append(c)
                except (TypeError, ValueError):
                    pass
        # Reconstruct text preserving line breaks via image_to_string.
        text = pytesseract.image_to_string(pil_image, lang=lang)
        confidence = sum(confs) / len(confs) if confs else None
        if confidence is not None and confidence < 70:
            warnings.append(f"Low OCR confidence ({confidence:.1f}%). Please verify extracted values.")
        return OCRResult(text=text, confidence=confidence, engine="pytesseract", warnings=warnings)
    except Exception as exc:  # pragma: no cover - depends on tesseract binary
        warnings.append(f"pytesseract failed: {exc}")
        return OCRResult("", None, "pytesseract", warnings)


def _ocr_easyocr(image, pil_image, warnings) -> OCRResult:  # pragma: no cover - heavy dep
    import numpy as np
    import easyocr

    reader = easyocr.Reader(["en"], gpu=False)
    if pil_image is not None:
        arr = np.array(pil_image.convert("RGB"))
        results = reader.readtext(arr)
    else:
        results = reader.readtext(image)
    lines = [r[1] for r in results]
    confs = [float(r[2]) * 100 for r in results if len(r) > 2]
    confidence = sum(confs) / len(confs) if confs else None
    if confidence is not None and confidence < 70:
        warnings.append(f"Low OCR confidence ({confidence:.1f}%). Please verify extracted values.")
    return OCRResult(text="\n".join(lines), confidence=confidence, engine="easyocr", warnings=warnings)
