"""Renders the app icon to a PNG using only the standard library.

No imaging library is installed, so this writes the PNG container by hand and
rasterises the shapes directly. Rounded corners get analytic coverage-based
antialiasing; everything else is axis-aligned and needs none.
"""

import struct
import sys
import zlib

SIZE = 1024


def write_png(path, width, height, pixels):
    """Writes 8-bit RGBA pixels (a flat bytearray) as a PNG."""
    stride = width * 4
    raw = bytearray()
    for y in range(height):
        raw.append(0)  # filter type 0: none
        raw += pixels[y * stride:(y + 1) * stride]

    def chunk(tag, data):
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    out = b"\x89PNG\r\n\x1a\n"
    out += chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0))
    out += chunk(b"IDAT", zlib.compress(bytes(raw), 9))
    out += chunk(b"IEND", b"")
    with open(path, "wb") as handle:
        handle.write(out)


def rounded_rect_coverage(x, y, left, top, right, bottom, radius):
    """Returns 0..1 coverage of the pixel centre by a rounded rectangle."""
    if x < left - 1 or x > right + 1 or y < top - 1 or y > bottom + 1:
        return 0.0

    # Distance from the rounded boundary, negative inside.
    cx = min(max(x, left + radius), right - radius)
    cy = min(max(y, top + radius), bottom - radius)
    dx = x - cx
    dy = y - cy

    if dx == 0 and dy == 0:
        inside = min(x - left, right - x, y - top, bottom - y)
        return 1.0 if inside >= 0.5 else max(0.0, min(1.0, inside + 0.5))

    dist = (dx * dx + dy * dy) ** 0.5 - radius
    # One pixel of linear falloff centred on the edge.
    return max(0.0, min(1.0, 0.5 - dist))


def lerp(a, b, t):
    return tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def blend(dst, offset, colour, alpha):
    if alpha <= 0:
        return
    if alpha >= 1:
        dst[offset] = colour[0]
        dst[offset + 1] = colour[1]
        dst[offset + 2] = colour[2]
        dst[offset + 3] = 255
        return

    existing_a = dst[offset + 3] / 255
    out_a = alpha + existing_a * (1 - alpha)
    if out_a <= 0:
        return
    for i in range(3):
        src = colour[i] * alpha
        old = dst[offset + i] * existing_a * (1 - alpha)
        dst[offset + i] = round((src + old) / out_a)
    dst[offset + 3] = round(out_a * 255)


def render():
    # A teal-green that reads as "spreadsheet" without copying Excel's brand.
    top_colour = (18, 134, 112)
    bottom_colour = (9, 86, 72)
    sheet_colour = (255, 255, 255)
    header_colour = (158, 206, 193)

    pixels = bytearray(SIZE * SIZE * 4)

    radius = SIZE * 0.225
    margin = SIZE * 0.045
    left = margin
    top = margin
    right = SIZE - margin
    bottom = SIZE - margin

    # Sheet glyph geometry.
    sheet_w = SIZE * 0.56
    sheet_h = SIZE * 0.50
    sheet_left = (SIZE - sheet_w) / 2
    sheet_top = (SIZE - sheet_h) / 2
    sheet_right = sheet_left + sheet_w
    sheet_bottom = sheet_top + sheet_h
    sheet_radius = SIZE * 0.028

    header_h = sheet_h * 0.24
    line = max(2.0, SIZE * 0.012)
    col1 = sheet_left + sheet_w / 3
    col2 = sheet_left + sheet_w * 2 / 3
    row1 = sheet_top + header_h + (sheet_h - header_h) / 3
    row2 = sheet_top + header_h + (sheet_h - header_h) * 2 / 3

    for y in range(SIZE):
        py = y + 0.5
        gradient = lerp(top_colour, bottom_colour, py / SIZE)
        row_offset = y * SIZE * 4

        for x in range(SIZE):
            px = x + 0.5
            offset = row_offset + x * 4

            body = rounded_rect_coverage(px, py, left, top, right, bottom, radius)
            if body <= 0:
                continue
            blend(pixels, offset, gradient, body)

            sheet = rounded_rect_coverage(
                px, py, sheet_left, sheet_top, sheet_right, sheet_bottom, sheet_radius
            )
            if sheet <= 0:
                continue
            blend(pixels, offset, sheet_colour, sheet * body)

            # Cut the grid back out of the white sheet, in the background colour.
            in_header = py < sheet_top + header_h
            on_column = abs(px - col1) < line / 2 or abs(px - col2) < line / 2
            on_row = abs(py - row1) < line / 2 or abs(py - row2) < line / 2

            if in_header:
                # A light tint, not the background: filling the header with the
                # background colour makes it vanish into the surrounding green
                # and the sheet reads as if its top were cut off.
                blend(pixels, offset, header_colour, sheet * body)
                if on_column:
                    blend(pixels, offset, gradient, sheet * body)
            elif on_column or on_row:
                blend(pixels, offset, gradient, sheet * body)

    return pixels


if __name__ == "__main__":
    target = sys.argv[1]
    write_png(target, SIZE, SIZE, render())
    print(f"wrote {target}")
