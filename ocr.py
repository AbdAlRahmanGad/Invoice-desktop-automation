"""Turning a PDF or image file into positioned text.

The only module that knows about PaddleOCR and pypdfium2.
"""

from dataclasses import dataclass
from functools import lru_cache
from typing import Any

from layout import TextBox


#: Caps PaddleOCR's internal downscaling; must clear our largest render.
DET_LIMIT_SIDE_LEN = 2200

#: Width a PDF page renders to at scale 1; used to bring image inputs to a
#: comparable resolution so detector sensitivity is the same for both.
BASE_PAGE_WIDTH = 540


@dataclass(frozen=True)
class Page:
    """OCR result for a single page."""

    boxes: list[TextBox]
    width: int


def _disable_ir_optim() -> None:
    """Stop paddle from rewriting the OCR graph into oneDNN fused kernels.

    paddleocr 2.10 hardcodes `switch_ir_optim(True)` and ships PP-OCRv3/v4
    models exported under paddle 2.x; running those through paddle 3.3's CPU
    fusion passes produces a `fused_conv2d` the oneDNN context cannot fill
    ("OneDnnContext does not have the input Filter"). The unfused graph runs
    fine, so we turn the passes off at predictor construction -- the only hook
    paddleocr leaves us, since its own `ir_optim` argument is ignored.
    """
    from paddle import inference

    if getattr(inference.create_predictor, "_ir_optim_disabled", False):
        return

    original = inference.create_predictor

    def create_predictor(config: Any) -> Any:
        config.switch_ir_optim(False)
        return original(config)

    create_predictor._ir_optim_disabled = True
    inference.create_predictor = create_predictor


@lru_cache(maxsize=1)
def _engine() -> Any:
    """Build the PaddleOCR engine once; model weights load lazily on first use."""
    from paddleocr import PaddleOCR

    _disable_ir_optim()
    return PaddleOCR(lang="en", det_limit_side_len=DET_LIMIT_SIDE_LEN)


def render(path: str, scale: int) -> Any:
    """Render the first page of `path` to an RGB numpy array at `scale`.

    Args:
        path (str): Path to a PDF or image.
        scale (int): Render scale; a PDF page becomes BASE_PAGE_WIDTH * scale wide.

    Returns:
        Any: The page as a numpy array.
    """
    import numpy as np

    if not path.lower().endswith(".pdf"):
        from PIL import Image

        img = Image.open(path).convert("RGB")
        target = BASE_PAGE_WIDTH * scale
        if img.width < target:
            ratio = target / img.width
            img = img.resize((target, round(img.height * ratio)), Image.LANCZOS)
        return np.array(img)

    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(path)
    try:
        return pdf[0].render(scale=scale).to_numpy()
    finally:
        pdf.close()


def read_page(path: str, scale: int) -> Page:
    """OCR the first page of `path` at `scale`.

    Args:
        path (str): Path to a PDF or image.
        scale (int): Render scale (see `render`).

    Returns:
        Page: Recognised boxes plus the pixel width they are positioned in.
    """
    image = render(path, scale)
    result = _engine().ocr(image)
    boxes = [
        TextBox(
            text=text.strip(),
            x=sum(p[0] for p in box) / 4,
            y=sum(p[1] for p in box) / 4,
            score=score,
        )
        for page in result or []
        for box, (text, score) in page or []
    ]
    return Page(boxes=boxes, width=image.shape[1])
