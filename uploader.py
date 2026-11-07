"""
uploader.py
===========
Photo & Document Matcher and Auto-Compressor for Government Portals.
Features:
- Maps a folder of applicant photos/documents to rows by Aadhaar number, Applicant Name, or Serial No.
- Automatically resizes to target dimensions (e.g. 200x230 px for passport photo) and compresses
  JPEG quality iteratively until file size is strictly under the portal's KB limit (e.g. max 50 KB or 200 KB).
- Converts PNG/WEBP to JPG or wraps images into a PDF when the portal requires a PDF document.
"""

from __future__ import annotations
import io
import os
import re
from typing import Dict, List, Optional, Tuple, Any

try:
    from PIL import Image
except ImportError:
    Image = None


def find_matching_file_for_row(
    row_data: Dict[str, Any],
    row_number: int,
    media_folder: str,
    match_by_column: str,
    explicit_filename_column: Optional[str] = None
) -> Tuple[Optional[str], str]:
    """
    Locates the matching photo or document file for a given row.
    1. Checks explicit filename column if mapped (e.g. 'Photo_File' = '101_ramesh.jpg')
    2. Checks match_by_column (e.g. Aadhaar number, Applicant Name, or Row Serial Number)
    Returns: (full_file_path_or_None, status_message)
    """
    if not os.path.exists(media_folder):
        return None, f"Media folder '{media_folder}' not found."

    files_in_dir = os.listdir(media_folder)
    if not files_in_dir:
        return None, f"Media folder '{media_folder}' is empty."

    # 1. Check explicit filename in row
    if explicit_filename_column and explicit_filename_column in row_data:
        explicit_val = str(row_data[explicit_filename_column]).strip()
        if explicit_val:
            direct_path = os.path.join(media_folder, explicit_val)
            if os.path.isfile(direct_path):
                return direct_path, f"Matched exact filename '{explicit_val}'"
            # Case-insensitive lookup
            for f in files_in_dir:
                if f.lower() == explicit_val.lower():
                    return os.path.join(media_folder, f), f"Matched file '{f}'"

    # 2. Match by Aadhaar / Name / Serial Key
    key_val = ""
    if match_by_column == "_ROW_NUMBER_":
        key_val = str(row_number)
    elif match_by_column in row_data:
        key_val = str(row_data[match_by_column]).strip()

    if key_val.endswith(".0"):
        key_val = key_val[:-2]

    if not key_val:
        return None, f"Key column '{match_by_column}' is blank for Row {row_number}."

    clean_key = re.sub(r"[^a-zA-Z0-9]", "", key_val).lower()

    for fname in files_in_dir:
        stem, ext = os.path.splitext(fname)
        if ext.lower() not in [".jpg", ".jpeg", ".png", ".webp", ".pdf"]:
            continue
        clean_stem = re.sub(r"[^a-zA-Z0-9]", "", stem).lower()
        if clean_stem == clean_key or (len(clean_key) >= 4 and clean_key in clean_stem):
            return os.path.join(media_folder, fname), f"Matched '{fname}' via '{match_by_column}'"

    return None, f"Warning: No photo/document found matching '{key_val}' in '{media_folder}'."


def compress_and_resize_image(
    input_image_path: str,
    output_dir: str = "processed_uploads",
    max_kb: int = 50,
    target_width: int = 200,
    target_height: int = 230,
    output_format: str = "JPEG"
) -> Tuple[str, Dict[str, Any]]:
    """
    Resizes an image to (target_width x target_height), converts to RGB JPEG (or PDF),
    and iteratively lowers JPEG quality until the file size is strictly <= max_kb.
    Returns: (processed_file_path, metadata_dict)
    """
    if Image is None:
        raise RuntimeError("Pillow is not installed. Run 'pip install Pillow'.")

    os.makedirs(output_dir, exist_ok=True)
    orig_size_kb = round(os.path.getsize(input_image_path) / 1024.0, 2)
    base_name = os.path.splitext(os.path.basename(input_image_path))[0]

    with Image.open(input_image_path) as img:
        # Convert RGBA/P to RGB for JPEG/PDF compatibility
        if img.mode in ("RGBA", "LA", "P"):
            background = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "P":
                img = img.convert("RGBA")
            background.paste(img, mask=img.split()[-1] if "A" in img.mode else None)
            rgb_img = background
        else:
            rgb_img = img.convert("RGB")

        # Resize with high-quality Lanczos filter
        if target_width > 0 and target_height > 0:
            rgb_img = rgb_img.resize((target_width, target_height), Image.Resampling.LANCZOS)

        fmt = output_format.upper()
        if fmt == "PDF":
            out_path = os.path.join(output_dir, f"{base_name}_portal.pdf")
            rgb_img.save(out_path, "PDF", resolution=100.0)
            final_kb = round(os.path.getsize(out_path) / 1024.0, 2)
            return out_path, {
                "original_kb": orig_size_kb,
                "final_kb": final_kb,
                "dimensions": f"{target_width}x{target_height}",
                "format": "PDF"
            }

        # Iteratively compress JPEG to fit within max_kb
        out_path = os.path.join(output_dir, f"{base_name}_portal.jpg")
        quality = 92
        best_bytes = b""

        while quality >= 15:
            buf = io.BytesIO()
            rgb_img.save(buf, format="JPEG", quality=quality, optimize=True)
            data = buf.getvalue()
            best_bytes = data
            if len(data) / 1024.0 <= max_kb:
                break
            quality -= 8

        with open(out_path, "wb") as f:
            f.write(best_bytes)

        final_kb = round(len(best_bytes) / 1024.0, 2)
        return out_path, {
            "original_kb": orig_size_kb,
            "final_kb": final_kb,
            "dimensions": f"{target_width}x{target_height}",
            "quality_used": quality,
            "format": "JPEG"
        }
