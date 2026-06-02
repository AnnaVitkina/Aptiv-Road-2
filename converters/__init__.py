from converters.new_grid import convert_file as convert_new_grid
from converters.usual_rate import convert_file as convert_usual_rate

CONVERTERS = {
    "usual_rate": convert_usual_rate,
    "new_grid": convert_new_grid,
}

# Legacy layout folder names → same converters
LEGACY_LAYOUT_ALIASES = {
    "layout1": "usual_rate",
    "layout2": "usual_rate",
    "layout4": "usual_rate",
    "layout3": "new_grid",
}

LAYOUT_LABELS = {
    "usual_rate": "Usual rate (lane books, WMT, simple pricelists)",
    "new_grid": "New grid (Rate Grid FTL / LTL / multi-stop)",
}
