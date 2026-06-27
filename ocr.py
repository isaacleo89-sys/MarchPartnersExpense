"""
OCR helpers for extracting monetary amounts from receipt images.
Uses Tesseract (via pytesseract) with image pre-processing.
Falls back gracefully — the bot will ask the user if extraction fails.
"""

import io
import logging
import re

from PIL import Image, ImageEnhance, ImageFilter

logger = logging.getLogger(__name__)

# Try to import pytesseract; fail softly if tesseract not installed
try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False
    logger.warning("pytesseract not installed — OCR disabled.")


# ── Image pre-processing ──────────────────────────────────────────────────────

def _preprocess(image_bytes: bytes) -> Image.Image:
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")

    # Scale up small images
    if img.width < 1200:
        scale = 1200 / img.width
        img = img.resize(
            (int(img.width * scale), int(img.height * scale)), Image.LANCZOS
        )

    # Sharpen and boost contrast
    img = ImageEnhance.Sharpness(img).enhance(2.0)
    img = ImageEnhance.Contrast(img).enhance(1.5)

    return img


# ── Amount extraction ─────────────────────────────────────────────────────────

# Lines that usually precede the total on a receipt
_TOTAL_LABELS = re.compile(
    r"(?:grand\s+)?total|amount\s+(?:due|paid|payable)|balance\s+(?:due|payable)?|"
    r"subtotal|nett?\s+(?:total|amount)",
    re.IGNORECASE,
)

# SGD / generic numeric amount pattern (handles commas)
_AMOUNT_PATTERN = re.compile(r"S?\$?\s*([\d,]+\.\d{2})\b")


def extract_amount(image_bytes: bytes) -> float | None:
    """
    Attempt to extract the total payable amount from a receipt image.
    Returns a float (SGD) or None if extraction fails.
    """
    if not TESSERACT_AVAILABLE:
        return None

    try:
        img = _preprocess(image_bytes)
        text = pytesseract.image_to_string(img, config="--psm 6")
        logger.debug("OCR raw text:\n%s", text)
    except Exception as e:
        logger.warning("OCR failed: %s", e)
        return None

    lines = text.splitlines()

    # Pass 1: look for "Total" or similar label on the same or next line
    for i, line in enumerate(lines):
        if _TOTAL_LABELS.search(line):
            # Search this line and the next for an amount
            search_in = line + " " + (lines[i + 1] if i + 1 < len(lines) else "")
            m = _AMOUNT_PATTERN.search(search_in)
            if m:
                try:
                    return float(m.group(1).replace(",", ""))
                except ValueError:
                    pass

    # Pass 2: collect all currency-like numbers and return the largest
    #         (totals are usually the biggest number on a receipt)
    candidates: list[float] = []
    for m in _AMOUNT_PATTERN.finditer(text):
        try:
            candidates.append(float(m.group(1).replace(",", "")))
        except ValueError:
            pass

    if candidates:
        return max(candidates)

    return None
