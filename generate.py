#!/usr/bin/env python3
"""
Certificate Generator
Reads participant data from Excel and overlays it onto a certificate template.
"""

from __future__ import annotations

import re
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

EXCEL_PATH = "players_spreadsheet.xlsx"
TEMPLATE_PATH = "merit certificate.png"
OUTPUT_DIR = "output"
CALIBRATION_MODE = False
CERTIFICATE_NUMBER_START =2294

FONTS_DIR = Path("fonts")

# Body field text — black
CERTIFICATE_TEXT_COLOR = (0, 0, 0)

# Certificate serial number only — burgundy to match HOA/HSG/M/ on template
CERTIFICATE_NO_COLOR = (99, 8, 53)

FONT_CANDIDATES = {
    "bold": [
        FONTS_DIR / "Times New Roman Bold.ttf",
        "/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf",
        "/Library/Fonts/Times New Roman Bold.ttf",
    ],
    "regular": [
        FONTS_DIR / "Times New Roman.ttf",
        "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
        "/Library/Fonts/Times New Roman.ttf",
    ],
}

REQUIRED_COLUMNS = [
    "Full Name",
    "Father Name",
    "Mother Name",
    "DOB",
    "Aadhaar Number",
    "Position",
    "Sport Event",
    "Sport",
    "District",
]

# All field positions come from this configuration.
# Calibrated against merit certificate.png (3508x2482). Fine-tune as needed.
FIELD_CONFIG: Dict[str, Dict[str, Any]] = {
    "full_name": {
        "label": "FULL_NAME",
        "x": 635,
        "y": 1007,
        "max_width": 820,
        "font_size": 60,
        "min_font_size": 36,
        "align": "left",
        "font_weight": "bold",
        "calibration_color": (255, 0, 0),
    },
    "father_name": {
        "label": "FATHER_NAME",
        "x": 1912,
        "y": 1007,
        "max_width": 820,
        "font_size": 60,
        "min_font_size": 36,
        "align": "left",
        "font_weight": "bold",
        "calibration_color": (0, 128, 255),
    },
    "mother_name": {
        "label": "MOTHER_NAME",
        "x": 808,
        "y": 1138,
        "max_width": 820,
        "font_size": 60,
        "min_font_size": 36,
        "align": "left",
        "font_weight": "bold",
        "calibration_color": (0, 180, 0),
    },
    "dob": {
        "label": "DOB",
        "x": 2122,
        "y": 1138,
        "max_width": 520,
        "font_size": 60,
        "min_font_size": 36,
        "align": "left",
        "font_weight": "bold",
        "calibration_color": (255, 128, 0),
    },
    "aadhaar": {
        "label": "AADHAAR",
        "x": 753,
        "y": 1279,
        "max_width": 805,
        "font_size": 60,
        "min_font_size": 36,
        "align": "left",
        "font_weight": "bold",
        "calibration_color": (128, 0, 255),
    },
    "position": {
        "label": "POSITION",
        "x": 725,
        "y": 1410,
        "max_width": 370,
        "font_size": 60,
        "min_font_size": 36,
        "align": "left",
        "font_weight": "bold",
        "calibration_color": (255, 0, 128),
    },
    "sport_event": {
        "label": "SPORT_EVENT",
        "x":  1513,
        "y": 1410,
        "max_width": 1220,
        "font_size": 60,
        "min_font_size": 36,
        "align": "left",
        "font_weight": "bold",
        "calibration_color": (0, 128, 128),
    },
    "sport": {
        "label": "SPORT",
        "x": 745,
        "y": 1541,
        "max_width": 700,
        "font_size": 60,
        "min_font_size": 36,
        "align": "left",
        "font_weight": "bold",
        "calibration_color": (180, 180, 0),
    },
    "district": {
        "label": "DISTRICT",
        "x": 2048,
        "y": 1674,
        "max_width": 770,
        "font_size": 60,
        "min_font_size": 36,
        "align": "left",
        "font_weight": "bold",
        "calibration_color": (128, 64, 0),
    },
    "certificate_no": {
        "label": "CERTIFICATE_NO",
        "x": 3060,
        "y": 585,
        "max_width": 120,
        "font_size": 42,
        "min_font_size": 22,
        "align": "left",
        "font_weight": "bold",
        "color": CERTIFICATE_NO_COLOR,
        "calibration_color": (255, 0, 0),
        # Template already prints HOA/HSG/M/ — render only the numeric suffix.
        "prefix": "",
    },
}

CALIBRATION_BOX_HEIGHT = 48
CALIBRATION_LABEL_OFFSET = (-4, -22)


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------


class CertificateGeneratorError(Exception):
    """Base exception for certificate generation errors."""


class MissingFileError(CertificateGeneratorError):
    """Raised when a required file is missing."""


class MissingColumnsError(CertificateGeneratorError):
    """Raised when required Excel columns are missing."""


class InvalidDataError(CertificateGeneratorError):
    """Raised when participant data is invalid."""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def ensure_project_fonts() -> None:
    """Copy certificate fonts into fonts/ when missing locally."""
    FONTS_DIR.mkdir(parents=True, exist_ok=True)
    bundled_fonts = {
        "Times New Roman Bold.ttf": [
            "/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf",
            "/Library/Fonts/Times New Roman Bold.ttf",
        ],
        "Times New Roman.ttf": [
            "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
            "/Library/Fonts/Times New Roman.ttf",
        ],
    }
    for filename, sources in bundled_fonts.items():
        dest = FONTS_DIR / filename
        if dest.is_file():
            continue
        for source in sources:
            source_path = Path(source)
            if source_path.is_file():
                shutil.copy2(source_path, dest)
                break


def resolve_font_path(weight: str = "bold") -> Path:
    """Locate a usable TrueType font for the requested weight."""
    ensure_project_fonts()
    candidates = FONT_CANDIDATES.get(weight, FONT_CANDIDATES["bold"])
    for candidate in candidates:
        path = Path(candidate)
        if path.is_file():
            return path
    raise MissingFileError(
        f"Font file not found for weight '{weight}'. "
        f"Place Times New Roman Bold.ttf in {FONTS_DIR}/"
    )


def load_font(font_path: Path, size: int) -> ImageFont.FreeTypeFont:
    """Load a font at the given size."""
    try:
        return ImageFont.truetype(str(font_path), size=size)
    except OSError as exc:
        raise CertificateGeneratorError(f"Unable to load font '{font_path}': {exc}") from exc


def fit_text_to_width(
    text: str,
    font_path: Path,
    max_width: int,
    initial_size: int,
    min_size: int = 12,
) -> Tuple[ImageFont.FreeTypeFont, str]:
    """
    Reduce font size until text fits within max_width.
    Returns the chosen font and the (possibly truncated) text.
    """
    if not text:
        text = ""

    size = initial_size
    while size >= min_size:
        font = load_font(font_path, size)
        bbox = font.getbbox(text)
        width = bbox[2] - bbox[0]
        if width <= max_width:
            return font, text
        size -= 1

    font = load_font(font_path, min_size)
    if not text:
        return font, text

    # Truncate with ellipsis if still too wide at minimum size
    ellipsis = "..."
    truncated = text
    while truncated and (font.getbbox(truncated + ellipsis)[2] - font.getbbox(truncated + ellipsis)[0]) > max_width:
        truncated = truncated[:-1]
    return font, (truncated + ellipsis) if truncated != text else truncated


def format_dob(value: Any) -> str:
    """Format DOB as DD-MM-YYYY."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        raise InvalidDataError("DOB is missing.")

    if isinstance(value, datetime):
        return value.strftime("%d-%m-%Y")

    text = str(value).strip()
    if not text or text.lower() == "nan":
        raise InvalidDataError("DOB is missing.")

    # Already in DD-MM-YYYY
    if re.fullmatch(r"\d{2}-\d{2}-\d{4}", text):
        return text

    # Try common parse formats
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue

    raise InvalidDataError(f"Invalid DOB format: {text!r}. Expected DD-MM-YYYY.")


def format_aadhaar(value: Any) -> str:
    """Ensure Aadhaar is always treated as a plain string."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        raise InvalidDataError("Aadhaar Number is missing.")

    if isinstance(value, float):
        # Avoid scientific notation for numeric Excel values
        if value.is_integer():
            text = str(int(value))
        else:
            text = format(value, "f").rstrip("0").rstrip(".")
    else:
        text = str(value).strip()

    if not text or text.lower() == "nan":
        raise InvalidDataError("Aadhaar Number is missing.")

    # Normalize common Excel/scientific artifacts
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]

    return text


def sanitize_filename(name: str) -> str:
    """Convert participant name into a safe filename segment."""
    cleaned = re.sub(r"[^\w\s-]", "", name.strip())
    cleaned = re.sub(r"\s+", "_", cleaned)
    cleaned = cleaned.strip("_")
    if not cleaned:
        raise InvalidDataError("Full Name is blank or invalid.")
    return cleaned


def format_certificate_number(
    row_index: int,
    certificate_number_start: Optional[int] = None,
) -> str:
    """Generate zero-padded certificate number based on row order."""
    start = (
        certificate_number_start
        if certificate_number_start is not None
        else CERTIFICATE_NUMBER_START
    )
    return f"{start + row_index - 1:04d}"


def validate_excel_file(excel_path: str | Path) -> None:
    """Ensure the spreadsheet file exists."""
    if not Path(excel_path).is_file():
        raise MissingFileError(f"Excel file not found: {excel_path}")


def validate_template_file(template_path: str | Path) -> None:
    """Ensure the template image exists."""
    if not Path(template_path).is_file():
        raise MissingFileError(f"Template image not found: {template_path}")


def validate_required_files(
    excel_path: str | Path = EXCEL_PATH,
    template_path: str | Path = TEMPLATE_PATH,
) -> None:
    """Ensure spreadsheet and template exist."""
    validate_excel_file(excel_path)
    validate_template_file(template_path)


def validate_columns(columns: List[str]) -> None:
    """Ensure all required Excel columns are present."""
    missing = [col for col in REQUIRED_COLUMNS if col not in columns]
    if missing:
        raise MissingColumnsError(
            f"Missing required columns: {', '.join(missing)}"
        )


def row_to_field_values(row: pd.Series, certificate_number: str) -> Dict[str, str]:
    """Map a spreadsheet row to certificate field values."""
    full_name = str(row["Full Name"]).strip()
    if not full_name or full_name.lower() == "nan":
        raise InvalidDataError("Full Name is missing.")

    prefix = FIELD_CONFIG["certificate_no"].get("prefix", "")
    cert_display = f"{prefix}{certificate_number}" if prefix else certificate_number
    return {
        "full_name": full_name,
        "father_name": str(row["Father Name"]).strip(),
        "mother_name": str(row["Mother Name"]).strip(),
        "dob": format_dob(row["DOB"]),
        "aadhaar": format_aadhaar(row["Aadhaar Number"]),
        "position": str(row["Position"]).strip(),
        "sport_event": str(row["Sport Event"]).strip(),
        "sport": str(row["Sport"]).strip(),
        "district": str(row["District"]).strip(),
        "certificate_no": cert_display,
    }


# ---------------------------------------------------------------------------
# Core rendering
# ---------------------------------------------------------------------------


def draw_field(
    draw: ImageDraw.ImageDraw,
    field_key: str,
    text: str,
    calibration: bool = False,
) -> None:
    """Draw a single configured field onto the certificate."""
    config = FIELD_CONFIG[field_key]
    x = config["x"]
    y = config["y"]
    max_width = config["max_width"]
    font_size = config["font_size"]
    min_font_size = config.get("min_font_size", 12)
    align = config.get("align", "left")
    text_color = config.get("color", CERTIFICATE_TEXT_COLOR)
    font_weight = config.get("font_weight", "bold")
    field_font_path = resolve_font_path(font_weight)

    font, fitted_text = fit_text_to_width(
        text=text,
        font_path=field_font_path,
        max_width=max_width,
        initial_size=font_size,
        min_size=min_font_size,
    )

    # Use default (la) bbox — no anchor arg.
    bbox = font.getbbox(fitted_text)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]

    # Anchor every text by its ink BOTTOM so names with or without descenders
    # (e.g. "Deepshikha" vs "Himanshi") always sit on the same line.
    # config_y was calibrated as the draw-y for a name that reaches the full
    # font descent (ascent+descent pixels below draw-y), so:
    ascent, descent = font.getmetrics()
    underline_y = y + ascent + descent      # the line position on the template
    draw_y = underline_y - bbox[3]          # shift each text so its ink bottom = underline_y

    if align == "center":
        draw_x = x + (max_width - text_width) // 2
    elif align == "right":
        draw_x = x + max_width - text_width
    else:
        draw_x = x

    draw.text((draw_x, draw_y), fitted_text, font=font, fill=text_color)

    if calibration:
        box_color = config.get("calibration_color", (255, 0, 0))
        ink_top = draw_y + bbox[1]
        ink_bottom = draw_y + bbox[3]
        draw.rectangle(
            [x, ink_top - 6, x + max_width, ink_bottom + 6],
            outline=box_color,
            width=3,
        )
        label = config.get("label", field_key.upper())
        label_x = x + CALIBRATION_LABEL_OFFSET[0]
        label_y = ink_top + CALIBRATION_LABEL_OFFSET[1]
        label_font = load_font(resolve_font_path("regular"), 18)
        draw.text(
            (label_x, label_y),
            f"{label} ({x},{y})",
            font=label_font,
            fill=box_color,
        )


def draw_coordinate_guides(
    draw: ImageDraw.ImageDraw,
    image_size: Tuple[int, int],
    font_path: Path,
) -> None:
    """Draw light grid guides to help calibration."""
    width, height = image_size
    guide_color = (200, 200, 200)
    label_font = load_font(font_path, 16)

    for x in range(0, width, 250):
        draw.line([(x, 0), (x, height)], fill=guide_color, width=1)
        draw.text((x + 4, 4), str(x), font=label_font, fill=guide_color)

    for y in range(0, height, 250):
        draw.line([(0, y), (width, y)], fill=guide_color, width=1)
        draw.text((4, y + 4), str(y), font=label_font, fill=guide_color)


def generate_certificate(
    template: Image.Image,
    field_values: Dict[str, str],
    calibration: bool = False,
) -> Image.Image:
    """Render one certificate from template and field values."""
    # Preserve original resolution — copy, never resize
    certificate = template.copy()
    draw = ImageDraw.Draw(certificate)

    if calibration:
        draw_coordinate_guides(draw, certificate.size, resolve_font_path("regular"))

    for field_key in FIELD_CONFIG:
        draw_field(
            draw=draw,
            field_key=field_key,
            text=field_values[field_key],
            calibration=calibration,
        )

    return certificate


def create_calibration_preview(
    template: Image.Image,
    first_row: pd.Series,
    output_path: Path,
) -> None:
    """Generate a single calibration certificate with field boxes and guides."""
    certificate_number = format_certificate_number(1)
    field_values = row_to_field_values(first_row, certificate_number)
    preview = generate_certificate(
        template=template,
        field_values=field_values,
        calibration=True,
    )
    preview.save(output_path, format="PNG")
    print(f"Generated:\n{output_path.name}")


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------


def load_excel(path: str | Path = EXCEL_PATH) -> pd.DataFrame:
    """Load and validate participant data from Excel."""
    validate_excel_file(path)

    try:
        df = pd.read_excel(
            path,
            dtype={"Aadhaar Number": str},
            engine="openpyxl",
        )
    except FileNotFoundError as exc:
        raise MissingFileError(f"Excel file not found: {path}") from exc
    except Exception as exc:
        raise CertificateGeneratorError(f"Failed to read Excel file: {exc}") from exc

    validate_columns(list(df.columns))

    # Force Aadhaar to string and strip whitespace
    df["Aadhaar Number"] = df["Aadhaar Number"].apply(
        lambda v: "" if pd.isna(v) else str(v).strip()
    )

    if df.empty:
        raise InvalidDataError("Excel file contains no participant rows.")

    return df


# ---------------------------------------------------------------------------
# Batch generation
# ---------------------------------------------------------------------------


def generate_all_certificates(
    excel_path: str | Path,
    template_path: str | Path,
    output_dir: str | Path,
    *,
    certificate_number_start: Optional[int] = None,
    on_start: Optional[Any] = None,
    on_generating: Optional[Any] = None,
    on_generated: Optional[Any] = None,
    on_skipped: Optional[Any] = None,
) -> Tuple[int, List[Path], int]:
    """
    Generate one PNG certificate per valid Excel row.

    Returns (generated_count, list_of_output_paths, skipped_count).
    Invalid rows are skipped (same behavior as the CLI).
    """
    excel_path = Path(excel_path)
    template_path = Path(template_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    ensure_project_fonts()
    resolve_font_path("bold")
    df = load_excel(excel_path)
    total_rows = len(df)

    if on_start:
        on_start(total_rows)

    try:
        template = Image.open(template_path)
    except FileNotFoundError as exc:
        raise MissingFileError(f"Template image not found: {template_path}") from exc
    except Exception as exc:
        raise CertificateGeneratorError(
            f"Failed to open template image: {exc}"
        ) from exc

    generated_paths: List[Path] = []
    skipped_count = 0

    for row_index, (_, row) in enumerate(df.iterrows(), start=1):
        certificate_number = format_certificate_number(
            row_index,
            certificate_number_start=certificate_number_start,
        )

        try:
            field_values = row_to_field_values(row, certificate_number)
            full_name = field_values["full_name"]
            filename = f"{certificate_number}_{sanitize_filename(full_name)}.png"
            output_path = output_dir / filename

            if on_generating:
                on_generating(
                    {
                        "row": row_index,
                        "total_rows": total_rows,
                        "name": full_name,
                        "certificate_no": certificate_number,
                        "generated": len(generated_paths),
                        "skipped": skipped_count,
                        "percent": round((row_index - 1) / total_rows * 100, 1),
                    }
                )

            certificate = generate_certificate(
                template=template,
                field_values=field_values,
                calibration=False,
            )
            certificate.save(output_path, format="PNG")
            generated_paths.append(output_path)

            if on_generated:
                on_generated(
                    {
                        "row": row_index,
                        "total_rows": total_rows,
                        "name": full_name,
                        "certificate_no": certificate_number,
                        "filename": filename,
                        "generated": len(generated_paths),
                        "skipped": skipped_count,
                        "percent": round(row_index / total_rows * 100, 1),
                    }
                )
        except InvalidDataError as exc:
            skipped_count += 1
            if on_skipped:
                on_skipped(
                    {
                        "row": row_index,
                        "total_rows": total_rows,
                        "name": str(row.get("Full Name", "")).strip(),
                        "reason": str(exc),
                        "generated": len(generated_paths),
                        "skipped": skipped_count,
                        "percent": round(row_index / total_rows * 100, 1),
                    }
                )

    return len(generated_paths), generated_paths, skipped_count


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------


def main() -> int:
    """Run calibration preview or full certificate generation."""
    try:
        output_dir = Path(OUTPUT_DIR)
        output_dir.mkdir(parents=True, exist_ok=True)

        validate_required_files()
        ensure_project_fonts()
        resolve_font_path("bold")

        if CALIBRATION_MODE:
            df = load_excel()
            try:
                template = Image.open(TEMPLATE_PATH)
            except FileNotFoundError as exc:
                raise MissingFileError(
                    f"Template image not found: {TEMPLATE_PATH}"
                ) from exc
            except Exception as exc:
                raise CertificateGeneratorError(
                    f"Failed to open template image: {exc}"
                ) from exc

            preview_path = output_dir / "calibration_preview.png"
            create_calibration_preview(
                template=template,
                first_row=df.iloc[0],
                output_path=preview_path,
            )
            generated_count = 1
            print("\nCalibration preview created.")
            print("Adjust FIELD_CONFIG coordinates, then set CALIBRATION_MODE = False.")
        else:
            generated_count, _, _ = generate_all_certificates(
                excel_path=EXCEL_PATH,
                template_path=TEMPLATE_PATH,
                output_dir=output_dir,
                on_generated=lambda info: print(f"Generated:\n{info['filename']}"),
                on_skipped=lambda info: print(
                    f"Skipped row {info['row']}: {info['reason']}",
                    file=sys.stderr,
                ),
            )

        print("\n---")
        print(f"Certificates Generated: {generated_count}")
        print(f"Output Folder: {OUTPUT_DIR}")
        print("---------------------")
        return 0

    except CertificateGeneratorError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Unexpected error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
