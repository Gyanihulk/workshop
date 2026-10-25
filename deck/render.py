"""python3 render.py <deck> [per_sheet]  -> out/render/<deck>/sheet-N.jpg (2 columns), for visual QA."""
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw

SKILL = Path("/root/.claude/skills/synced/61f51ad4-e363-4958-9b90-894e0ee506ce_a211e624-7974-4b2a-9bd1-b173e277d1bf/pptx")
HERE = Path(__file__).parent


def main(name, per=6):
    pptx = HERE / "out" / "decks" / f"{name}.pptx"
    out = HERE / "out" / "render" / name
    out.mkdir(parents=True, exist_ok=True)
    for f in out.glob("*"):
        f.unlink()
    subprocess.run(["python3", str(SKILL / "scripts/office/soffice.py"), "--headless", "--convert-to", "pdf", "--outdir", str(out), str(pptx)], check=True, capture_output=True)
    subprocess.run(["pdftoppm", "-jpeg", "-r", "62", str(out / f"{name}.pdf"), str(out / "s")], check=True)
    pages = sorted(out.glob("s-*.jpg"))
    for k in range(0, len(pages), per):
        chunk = pages[k : k + per]
        ims = [Image.open(p) for p in chunk]
        w, h = ims[0].size
        rows = (len(ims) + 1) // 2
        sheet = Image.new("RGB", (w * 2 + 10, rows * (h + 10)), (60, 60, 70))
        d = ImageDraw.Draw(sheet)
        for i, im in enumerate(ims):
            x, y = (i % 2) * (w + 10), (i // 2) * (h + 10)
            sheet.paste(im, (x, y))
            d.rectangle([x, y, x + 34, y + 20], fill=(245, 158, 11))
            d.text((x + 6, y + 4), str(k + i + 1), fill=(0, 0, 0))
        sheet.save(out / f"sheet-{k // per + 1}.jpg", quality=84)
    print(len(pages), "pages ->", out)


if __name__ == "__main__":
    main(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 6)
