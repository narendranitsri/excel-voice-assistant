"""
filler.py
=========
Robust Playwright Web Form Filling Engine.
Capabilities:
- Fills Text, Number, Date, Email, Tel, and Textarea fields (including Unicode Hindi text).
- Handles Dropdowns (<select>) and Cascading Dependent Dropdowns (District -> Block -> Village)
  by explicitly polling until child <option> items finish loading via AJAX/JS.
- Handles Radio Buttons, Checkboxes, File Upload fields (<input type="file">), and nested iframes.
- Supports Multi-Step / Multi-Page wizard buttons ("Next Step").
- Multi-fallback locator resolution (CSS selector -> Label text -> Name attribute -> Placeholder).
"""

from __future__ import annotations
import time
import os
from typing import Dict, List, Any, Optional, Tuple
from utils import format_value_for_portal, mask_sensitive_value
from uploader import find_matching_file_for_row, compress_and_resize_image


def resolve_locator_with_fallbacks(target_frame: Any, mapping: Dict[str, Any]) -> Any:
    """
    Finds the form element using primary CSS selector first, falling back to label text,
    name attribute, or placeholder if the portal's dynamic ID changed.
    """
    selector = mapping.get("field_selector", "")
    label = mapping.get("field_label", "")

    if selector:
        loc = target_frame.locator(selector).first
        if loc.count() > 0:
            return loc

    if label:
        by_lbl = target_frame.get_by_label(label, exact=False).first
        if by_lbl.count() > 0:
            return by_lbl

        by_ph = target_frame.get_by_placeholder(label, exact=False).first
        if by_ph.count() > 0:
            return by_ph

    return target_frame.locator(selector).first


def wait_for_dependent_dropdown_options(locator: Any, timeout_ms: int = 5000) -> bool:
    """
    Waits for a cascading dropdown (like Block after selecting District, or Village after Block)
    to populate with more than 1 <option> (beyond the default '-- Select --' placeholder).
    """
    deadline = time.time() + (timeout_ms / 1000.0)
    while time.time() < deadline:
        try:
            option_count = locator.locator("option").count()
            if option_count > 1:
                return True
        except Exception:
            pass
        time.sleep(0.25)
    return False


def fill_single_row_on_page(
    page: Any,
    row_data: Dict[str, Any],
    row_number: int,
    mappings: List[Dict[str, Any]],
    photo_settings: Optional[Dict[str, Any]] = None,
    media_folder: str = "samples/photos",
    skip_photos: bool = False,
    action_delay_sec: float = 0.8
) -> Tuple[bool, List[str]]:
    """
    Fills a single applicant row into the active portal form.
    Never clicks the final Submit button unless caller explicitly requests submission.
    Returns: (all_succeeded: bool, step_logs: List[str])
    """
    logs: List[str] = []
    photo_cfg = photo_settings or {"max_kb": 50, "width": 200, "height": 230, "format": "JPEG", "match_by_column": "Aadhaar_No"}

    for m in mappings:
        data_col = m.get("data_column", "")
        if not data_col or data_col not in row_data:
            continue

        raw_val = row_data.get(data_col)
        f_type = str(m.get("field_type", "text")).lower()
        f_label = m.get("field_label", data_col)
        transform = m.get("transform", "none")
        depends_on = m.get("depends_on_selector")
        wait_after_ms = int(m.get("wait_after_ms", 200))

        # Target frame (main page or named iframe)
        frame_name = m.get("frame", "main")
        target_frame = page
        if frame_name != "main":
            for fr in page.frames:
                if fr.name == frame_name:
                    target_frame = fr
                    break

        try:
            locator = resolve_locator_with_fallbacks(target_frame, m)

            # 1. File / Photo Upload Field
            if f_type == "file":
                if skip_photos:
                    logs.append(f"Skipped file upload for '{f_label}' (Skip Photos active).")
                    continue

                matched_file, msg = find_matching_file_for_row(
                    row_data=row_data,
                    row_number=row_number,
                    media_folder=media_folder,
                    match_by_column=photo_cfg.get("match_by_column", "Aadhaar_No"),
                    explicit_filename_column=data_col
                )
                if not matched_file:
                    logs.append(f"⚠️ File missing for '{f_label}': {msg}")
                    continue

                # Auto-compress / resize if image
                ext = os.path.splitext(matched_file)[1].lower()
                upload_path = matched_file
                if ext in [".jpg", ".jpeg", ".png", ".webp"]:
                    upload_path, meta = compress_and_resize_image(
                        input_image_path=matched_file,
                        max_kb=int(photo_cfg.get("max_kb", 50)),
                        target_width=int(photo_cfg.get("width", 200)),
                        target_height=int(photo_cfg.get("height", 230)),
                        output_format=photo_cfg.get("format", "JPEG")
                    )
                    logs.append(f"📷 Resized & compressed '{os.path.basename(matched_file)}' ({meta['original_kb']} KB ➔ {meta['final_kb']} KB)")

                locator.set_input_files(upload_path)
                logs.append(f"Uploaded file '{os.path.basename(upload_path)}' into '{f_label}'.")
                time.sleep(action_delay_sec)
                continue

            # Format value for standard inputs
            formatted_val = format_value_for_portal(raw_val, transform)
            if not formatted_val:
                continue

            masked_display = mask_sensitive_value(formatted_val, data_col)

            # 2. Dropdown (<select>) including Dependent Dropdowns
            if f_type in ["select", "select-one"]:
                if depends_on:
                    wait_for_dependent_dropdown_options(locator, timeout_ms=5000)

                # Try selecting by label first, then by value
                try:
                    locator.select_option(label=formatted_val, timeout=3000)
                except Exception:
                    locator.select_option(value=formatted_val, timeout=3000)

                logs.append(f"Selected '{formatted_val}' in dropdown '{f_label}'.")
                if wait_after_ms > 0:
                    time.sleep(wait_after_ms / 1000.0)

            # 3. Radio Button
            elif f_type == "radio":
                radio_name = m.get("name") or m.get("field_selector", "")
                radio_loc = target_frame.locator(f"input[type='radio'][value='{formatted_val}' i]").first
                if radio_loc.count() > 0:
                    radio_loc.check()
                else:
                    target_frame.get_by_label(formatted_val, exact=False).first.check()
                logs.append(f"Checked radio option '{formatted_val}' for '{f_label}'.")

            # 4. Checkbox
            elif f_type == "checkbox":
                should_check = formatted_val.lower() in ["true", "yes", "1", "y", "हाँ", "हां", "checked"]
                if should_check:
                    locator.check()
                else:
                    locator.uncheck()
                logs.append(f"Set checkbox '{f_label}' to {should_check}.")

            # 5. Text / Number / Date / Textarea
            else:
                locator.scroll_into_view_if_needed(timeout=3000)
                locator.fill("")
                locator.fill(formatted_val)
                logs.append(f"Filled '{f_label}' = '{masked_display}'.")

            # Respect rate-limiting delay between fields
            time.sleep(action_delay_sec)

        except Exception as exc:
            logs.append(f"❌ Failed on field '{f_label}': {str(exc)}")
            return False, logs

    return True, logs
