"""Shared path configuration — auto-detects Google Colab vs local machine."""

from __future__ import annotations

from pathlib import Path

# --- Google Colab (Drive) paths — data folders on Shared Drive ---
COLAB_DRIVE_BASE = (
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_APTIV_VERSIGENT/RMT_Road"
)
COLAB_INPUT_DIR = f"{COLAB_DRIVE_BASE}/input"
COLAB_OUTPUT_DIR = f"{COLAB_DRIVE_BASE}/output"
COLAB_PROCESSING_DIR = f"{COLAB_DRIVE_BASE}/processing"

# Active layout keys (see converters/__init__.py LAYOUT_LABELS for display names)
LAYOUTS: tuple[str, ...] = ("usual_rate", "new_grid")

# Legacy folder names still scanned on Drive/older clones (optional)
LEGACY_INPUT_FOLDERS: dict[str, tuple[str, ...]] = {
    "usual_rate": ("layout1", "layout2", "layout4"),
    "new_grid": ("layout3",),
}

LEGACY_PROCESSING_FOLDERS: dict[str, tuple[str, ...]] = {
    "usual_rate": ("layout1", "layout2", "layout4"),
    "new_grid": ("layout3",),
}


def _is_colab() -> bool:
    try:
        import google.colab  # noqa: F401

        return True
    except ImportError:
        return False


IS_COLAB = _is_colab()

_SCRIPT_ROOT = Path(__file__).resolve().parent
_DRIVE_INPUT = Path(COLAB_INPUT_DIR)

if IS_COLAB and _DRIVE_INPUT.is_dir():
    PROJECT_ROOT = _SCRIPT_ROOT
    INPUT_DIR = _DRIVE_INPUT
    OUTPUT_DIR = Path(COLAB_OUTPUT_DIR)
    PROCESSING_DIR = Path(COLAB_PROCESSING_DIR)
else:
    PROJECT_ROOT = _SCRIPT_ROOT
    INPUT_DIR = PROJECT_ROOT / "input"
    OUTPUT_DIR = PROJECT_ROOT / "output"
    PROCESSING_DIR = PROJECT_ROOT / "processing"


def normalize_layout(layout: str) -> str:
    """Map legacy layout1–layout4 names to usual_rate / new_grid."""
    from converters import LEGACY_LAYOUT_ALIASES

    return LEGACY_LAYOUT_ALIASES.get(layout, layout)


def input_dirs_for(layout: str) -> list[Path]:
    """
    Input folder for this layout.

    Uses ``input/usual_rate`` or ``input/new_grid`` when that folder exists.
    Legacy ``layout1``–``layout4`` are only scanned if the new folder is missing
    (unmigrated Drive clones).
    """
    key = normalize_layout(layout)
    primary = (INPUT_DIR / key).resolve()
    if primary.is_dir():
        return [primary]
    dirs: list[Path] = []
    for name in LEGACY_INPUT_FOLDERS.get(key, ()):
        p = (INPUT_DIR / name).resolve()
        if p.is_dir():
            dirs.append(p)
    if not dirs:
        primary.mkdir(parents=True, exist_ok=True)
        dirs.append(primary)
    return dirs


def processing_dirs_for(layout: str) -> list[Path]:
    """Same rule as input_dirs_for but under processing/."""
    key = normalize_layout(layout)
    primary = (PROCESSING_DIR / key).resolve()
    if primary.is_dir():
        return [primary]
    dirs: list[Path] = []
    for name in LEGACY_PROCESSING_FOLDERS.get(key, ()):
        p = (PROCESSING_DIR / name).resolve()
        if p.is_dir():
            dirs.append(p)
    return dirs or [primary]


def ensure_dirs() -> None:
    """Create processing/ and output/ trees (including per-layout subfolders)."""
    for base in (PROCESSING_DIR, OUTPUT_DIR):
        base.mkdir(parents=True, exist_ok=True)
        for layout in LAYOUTS:
            (base / layout).mkdir(parents=True, exist_ok=True)


def matrix_output_path(converted_path: Path) -> Path:
    """Mirror processing/<layout>/… under output/<layout>/… for matrix files."""
    stem = converted_path.stem.replace("_converted", "") + "_matrix.xlsx"
    try:
        rel = converted_path.resolve().relative_to(PROCESSING_DIR.resolve())
        if len(rel.parts) > 1:
            return OUTPUT_DIR / rel.parent / stem
    except ValueError:
        pass
    return OUTPUT_DIR / stem
