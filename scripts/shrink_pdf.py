# --- shrink_pdf.py ---
# Creates a lower-resolution copy of a PDF by rasterizing pages with PyMuPDF.

import sys
import fitz  # PyMuPDF


def clamp(value, lo, hi):
    return max(lo, min(hi, value))


def shrink_pdf(input_path, output_path, percent):
    if percent >= 100:
        print("Percent is 100; no shrinking requested.")
        return

    scale = percent / 100.0
    src = fitz.open(input_path)
    dst = fitz.open()

    for page in src:
        rect = page.rect
        mat = fitz.Matrix(scale, scale)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        new_page = dst.new_page(width=rect.width, height=rect.height)
        new_page.insert_image(rect, pixmap=pix)

    dst.save(output_path, garbage=4, deflate=True, clean=True)
    dst.close()
    src.close()


def main():
    if len(sys.argv) < 4:
        print(
            "Usage: python shrink_pdf.py <input_pdf> <output_pdf> <percent>",
            file=sys.stderr,
        )
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    try:
        percent = int(float(sys.argv[3]))
    except ValueError:
        print("Percent must be a number.", file=sys.stderr)
        sys.exit(1)

    percent = clamp(percent, 5, 100)
    if percent >= 100:
        print("Percent is 100; skipping shrink.")
        return

    shrink_pdf(input_path, output_path, percent)
    print(f"Shrunk PDF created at {output_path} ({percent}%).")


if __name__ == "__main__":
    main()
