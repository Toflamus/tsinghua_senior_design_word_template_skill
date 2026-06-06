"""Image safety helper: re-encode + downsize via Pillow.

Why: real-world Word documents often embed multi-megapixel scans (page-sized
TIFF/PNG at 16000+px). Two problems:
  1. python-docx's image header parser is strict — some exotic PNG variants
     raise UnrecognizedImageError when add_picture() tries to embed them.
  2. Embedding a 20 MB image multiplied across a thesis produces 60+ MB docx
     files that Word loads slowly.

`safe_image` re-encodes any input through Pillow as a clean PNG with the long
edge capped (default 1600 px ≈ 278 DPI at 14.6 cm display, plenty for print).
The output goes into a cache directory; calling twice on the same source
reuses the cached output if it's newer than the source.

Pillow is an optional dependency. If not installed, `safe_image` raises with
an install hint.
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional


_PILLOW_HINT = (
    "Pillow not installed. `safe_image` needs Pillow:\n"
    "  pip install Pillow"
)


def safe_image(
    src_path: Path | str,
    dst_dir: Optional[Path | str] = None,
    *,
    max_long_edge: int = 1600,
) -> Path:
    """Re-encode `src_path` as a clean RGB PNG capped at `max_long_edge` pixels
    on its long edge. Returns the cached output path.

    Args:
        src_path: Source image (any Pillow-readable format).
        dst_dir:  Output directory; defaults to a sibling `media_proc/`
                  alongside `src_path.parent`.
        max_long_edge: Cap the long edge; pass 0 to disable resizing.

    Cache behavior: if the output already exists and is newer than the source,
    the existing file is returned without re-encoding.
    """
    try:
        from PIL import Image
    except ImportError as e:
        raise RuntimeError(_PILLOW_HINT) from e
    # Pillow trips a safety check on huge images; we trust thesis content.
    Image.MAX_IMAGE_PIXELS = None  # type: ignore[attr-defined]

    src = Path(src_path)
    if not src.exists():
        raise FileNotFoundError(src)
    if dst_dir is None:
        dst_dir = src.parent.parent / "media_proc"
    dst_dir = Path(dst_dir)
    dst_dir.mkdir(parents=True, exist_ok=True)
    out = dst_dir / (src.stem + ".png")

    # cache: skip if output is newer than source
    if out.exists() and out.stat().st_mtime >= src.stat().st_mtime:
        return out

    im = Image.open(src)
    im.load()
    w, h = im.size
    if max_long_edge and max(w, h) > max_long_edge:
        scale = max_long_edge / max(w, h)
        im = im.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    im.save(out, "PNG", optimize=True)
    return out
