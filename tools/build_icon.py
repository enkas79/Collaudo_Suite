"""Genera l'icona dell'applicazione (PNG + ICO multi-risoluzione).

Uso: python tools/build_icon.py
Output: collaudo_suite/assets/app_icon.png e collaudo_suite/assets/app_icon.ico
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "collaudo_suite" / "assets"

# Palette della suite (vedi collaudo_suite/styles.py).
BLUE = (40, 120, 168, 255)
DARK = (44, 62, 80, 255)
WHITE = (255, 255, 255, 255)
LINE = (190, 202, 214, 255)
GREEN = (38, 140, 79, 255)

SIZE = 1024  # disegno ad alta risoluzione, poi ridimensionato


def draw_icon() -> Image.Image:
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    s = SIZE / 256  # coordinate progettate su griglia 256

    def box(x0: float, y0: float, x1: float, y1: float) -> tuple[float, float, float, float]:
        return (x0 * s, y0 * s, x1 * s, y1 * s)

    # Fondo: quadrato arrotondato blu con bordo scuro.
    d.rounded_rectangle(box(8, 8, 248, 248), radius=48 * s, fill=DARK)
    d.rounded_rectangle(box(16, 16, 240, 240), radius=42 * s, fill=BLUE)

    # Foglio checklist (clipboard).
    d.rounded_rectangle(box(60, 52, 196, 220), radius=14 * s, fill=WHITE)
    # Fermaglio superiore.
    d.rounded_rectangle(box(98, 38, 158, 66), radius=10 * s, fill=DARK)

    # Righe della checklist: casella + linea di testo.
    for i, y in enumerate((92, 128, 164)):
        d.rounded_rectangle(box(76, y, 96, y + 20), radius=4 * s, outline=DARK, width=int(4 * s))
        d.rounded_rectangle(box(106, y + 6, 180, y + 14), radius=4 * s, fill=LINE)
        if i < 2:
            d.line(
                [(79 * s, (y + 10) * s), (85 * s, (y + 16) * s), (95 * s, (y + 3) * s)],
                fill=GREEN,
                width=int(5 * s),
                joint="curve",
            )

    # Badge di collaudo superato in basso a destra.
    d.ellipse(box(150, 150, 236, 236), fill=WHITE)
    d.ellipse(box(158, 158, 228, 228), fill=GREEN)
    d.line(
        [(175 * s, 193 * s), (189 * s, 207 * s), (213 * s, 178 * s)],
        fill=WHITE,
        width=int(11 * s),
        joint="curve",
    )
    return img


def main() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    icon = draw_icon()
    icon.resize((256, 256), Image.Resampling.LANCZOS).save(ASSETS / "app_icon.png")
    icon.save(
        ASSETS / "app_icon.ico",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )
    print(f"Icona generata in {ASSETS}")


if __name__ == "__main__":
    main()
