"""python3 make.py <deck> [<deck> ...]   builds out/decks/<deck>.pptx from decks/<deck>.py (SPECS list)."""
import importlib
import sys
from pathlib import Path

import kit

OUT = Path(__file__).parent / "out" / "decks"

if __name__ == "__main__":
    for name in sys.argv[1:]:
        specs = importlib.import_module(f"decks.{name}").SPECS
        path = kit.build_deck(specs, OUT / f"{name}.pptx")
        print(path, len(specs), "slides")
