# --- detect_pdf_size.py ---
# Detects the paper size of PDF files in a project's PDF folder
# Returns the matching paper size name for the PlotDWGs.ps1 script

import sys
import os
import fitz  # PyMuPDF library


# Standard paper sizes in points (1 inch = 72 points)
# Using tolerance-based matching to account for slight variations
PAPER_SIZES = {
    "ANSI full bleed D (22.00 x 34.00 Inches)": (22 * 72, 34 * 72),   # 1584 x 2448
    "ARCH full bleed D (24.00 x 36.00 Inches)": (24 * 72, 36 * 72),   # 1728 x 2592
    "ARCH full bleed E1 (30.00 x 42.00 Inches)": (30 * 72, 42 * 72),  # 2160 x 3024
}

# Tolerance in points (0.5 inch = 36 points)
TOLERANCE = 36


def get_project_root(dwg_path):
    r"""
    Given a DWG file path, find the project root folder.
    Project structure: \\acies.lan\cachedrive\Projects\Nelson\BAC\2025\250368 ProjectName\Electrical\file.dwg
    Project root is the folder with the project number (e.g., "250368 ProjectName")
    """
    path = os.path.normpath(dwg_path)

    # Handle UNC paths (\\server\share\...)
    # os.path.normpath converts \\ to \ but we need to handle the split correctly
    is_unc = path.startswith('\\\\') or dwg_path.startswith('\\\\')

    # Split by backslash (works for both UNC and regular paths on Windows)
    parts = path.replace('/', '\\').split('\\')

    # Remove empty parts (from leading \\ in UNC paths)
    parts = [p for p in parts if p]

    # Look for the "Projects" folder in the path
    try:
        projects_idx = None
        for i, part in enumerate(parts):
            if part.lower() == "projects":
                projects_idx = i
                break

        if projects_idx is not None and projects_idx + 4 < len(parts):
            # Project root is 4 levels down from "Projects": Nelson/BAC/2025/ProjectFolder
            root_parts = parts[:projects_idx + 5]

            # Reconstruct the path
            if is_unc:
                project_root = '\\\\' + '\\'.join(root_parts)
            else:
                project_root = '\\'.join(root_parts)

            return project_root
    except (ValueError, IndexError):
        pass

    return None


def find_pdf_folder(project_root):
    """
    Find the PDF subfolder in the project root.
    """
    if not project_root:
        return None

    pdf_folder = os.path.join(project_root, "PDF")
    if os.path.isdir(pdf_folder):
        return pdf_folder

    return None


def get_most_recent_pdf(pdf_folder):
    """
    Get the most recent PDF file(s) from the PDF folder.
    Looks for the most recently modified PDF.
    """
    if not pdf_folder or not os.path.isdir(pdf_folder):
        return None

    pdf_files = []

    # Walk through all subdirectories to find PDFs
    for root, dirs, files in os.walk(pdf_folder):
        for file in files:
            if file.lower().endswith('.pdf'):
                full_path = os.path.join(root, file)
                try:
                    mtime = os.path.getmtime(full_path)
                    pdf_files.append((full_path, mtime))
                except OSError:
                    continue

    if not pdf_files:
        return None

    # Sort by modification time (most recent first)
    pdf_files.sort(key=lambda x: x[1], reverse=True)

    return pdf_files[0][0]


def detect_pdf_page_size(pdf_path):
    """
    Read a PDF file and return the page dimensions of the first page.
    Returns (width, height) in points.
    """
    try:
        with fitz.open(pdf_path) as pdf:
            if len(pdf) == 0:
                return None

            page = pdf[0]
            rect = page.rect
            width = rect.width
            height = rect.height

            # Ensure width is the smaller dimension (portrait normalization)
            # Actually, for engineering drawings, we want landscape (width > height)
            # So we normalize to ensure we're comparing correctly
            if width < height:
                width, height = height, width

            return (width, height)
    except Exception as e:
        print(f"Error reading PDF: {e}", file=sys.stderr)
        return None


def match_paper_size(dimensions):
    """
    Match detected dimensions to a standard paper size.
    Returns the paper size name or None if no match.
    """
    if not dimensions:
        return None

    detected_width, detected_height = dimensions

    for size_name, (std_width, std_height) in PAPER_SIZES.items():
        # Check both orientations
        if (abs(detected_width - std_width) <= TOLERANCE and
            abs(detected_height - std_height) <= TOLERANCE):
            return size_name
        if (abs(detected_width - std_height) <= TOLERANCE and
            abs(detected_height - std_width) <= TOLERANCE):
            return size_name

    return None


def detect_paper_size(dwg_path):
    """
    Main function to detect paper size from existing PDFs in the project.
    Returns the matching paper size name or empty string if not found.
    """
    project_root = get_project_root(dwg_path)
    if not project_root:
        return ""

    pdf_folder = find_pdf_folder(project_root)
    if not pdf_folder:
        return ""

    recent_pdf = get_most_recent_pdf(pdf_folder)
    if not recent_pdf:
        return ""

    dimensions = detect_pdf_page_size(recent_pdf)
    if not dimensions:
        return ""

    matched_size = match_paper_size(dimensions)
    return matched_size if matched_size else ""


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python detect_pdf_size.py <dwg_path>", file=sys.stderr)
        sys.exit(1)

    dwg_path = sys.argv[1]
    result = detect_paper_size(dwg_path)

    # Output the result (will be captured by PowerShell)
    print(result)
