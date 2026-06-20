from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw


OUTPUT = Path("assets/daily-report-app.ico")


def make_icon(size: int) -> Image.Image:
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    # Rounded green tile.
    margin = int(size * 0.08)
    draw.rounded_rectangle(
        [margin, margin, size - margin, size - margin],
        radius=int(size * 0.18),
        fill=(15, 118, 110, 255),
    )

    # White report sheet.
    sheet_margin = int(size * 0.22)
    sheet = [sheet_margin, int(size * 0.16), int(size * 0.78), int(size * 0.82)]
    draw.rounded_rectangle(sheet, radius=int(size * 0.04), fill=(255, 255, 255, 255))

    # Sheet header and lines.
    draw.rectangle(
        [sheet[0], sheet[1], sheet[2], sheet[1] + int(size * 0.12)],
        fill=(204, 251, 241, 255),
    )
    for i in range(3):
        y = sheet[1] + int(size * (0.24 + i * 0.12))
        draw.line([sheet[0] + int(size * 0.07), y, sheet[2] - int(size * 0.07), y], fill=(15, 118, 110, 255), width=max(1, size // 32))

    # Simple chart bars.
    bar_w = int(size * 0.08)
    base_y = int(size * 0.72)
    x0 = int(size * 0.36)
    heights = [int(size * 0.16), int(size * 0.26), int(size * 0.36)]
    for i, height in enumerate(heights):
        x = x0 + i * int(size * 0.13)
        draw.rounded_rectangle(
            [x, base_y - height, x + bar_w, base_y],
            radius=max(1, size // 64),
            fill=(15, 118, 110, 255),
        )

    return image


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    base = make_icon(256)
    base.save(OUTPUT, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print(f"created={OUTPUT.resolve()}")


if __name__ == "__main__":
    main()
