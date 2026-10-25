# Local deck editing

Edit the real workshop deck from Python, render it, look at it, then move slides into Google Slides.

```
deck/
  source/workshop-deck.pptx   the deck exported from Google Slides (input, never modified)
  icons/                      icons extracted from the deck itself (same look as the existing slides)
  build.py                    builds the opening slides into the deck
  out/                        generated: workshop-deck.pptx + render/ previews (git-ignored)
```

## Loop

```bash
cd deck
python3 build.py                                   # -> out/workshop-deck.pptx
S=<pptx skill dir>/scripts
python3 $S/office/validate.py out/workshop-deck.pptx --original source/workshop-deck.pptx
python3 $S/office/soffice.py --headless --convert-to pdf --outdir out/render out/workshop-deck.pptx
pdftoppm -jpeg -r 110 -f 2 -l 3 out/render/workshop-deck.pdf out/render/s   # look at the slides
```

Change wording or layout in `build.py`, re-run, re-render. The script finds the slides it owns
(by title) and replaces them, so it is safe to run again. Slides use the deck's own layouts
(`TitleDark`, `ContentLight`), title placeholders, Cambria/Calibri and palette, so new slides match.

## Getting slides into Google Slides

1. Open the deck in Google Slides, then File > Import slides, upload `out/workshop-deck.pptx`.
2. Tick only the new slides (2, 3 and 6), keep "Keep original theme", import.
3. Delete whichever of the old intro slides you do not want.

## Refreshing the source

If you change the deck in Google Slides, export File > Download > Microsoft PowerPoint, overwrite
`source/workshop-deck.pptx`, and re-run `build.py`.

## Notes

- Speaker notes (facilitation script for the warm-up, talk track for the intro) are in the notes pane of each new slide.
- No internet in some sandboxes: icons come from the deck itself, so no icon package is needed.
