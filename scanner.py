"""
scanner.py
==========
"Teach Mode" Web Form Scanner.
Scans any live portal page (or local HTML file) including nested iframes and extracts all:
- Text / Number / Tel / Email / Date inputs
- Dropdowns (<select>) with their available option labels
- Radio buttons and Checkboxes
- Textareas
- File upload inputs (<input type="file">)
Resolves Hindi, English, and Hinglish labels via <label for="...">, aria-label, placeholder, or adjacent text.
"""

from __future__ import annotations
from typing import Dict, List, Any, Optional
from bs4 import BeautifulSoup


JS_DOM_SCANNER_SCRIPT = """
() => {
    const fields = [];
    const elements = document.querySelectorAll('input, select, textarea');
    
    function getLabelText(el) {
        // 1. Explicit <label for="id">
        if (el.id) {
            const lbl = document.querySelector(`label[for="${CSS.escape(el.id)}"]`);
            if (lbl && lbl.innerText.trim()) return lbl.innerText.trim();
        }
        // 2. Parent <label>
        const parentLbl = el.closest('label');
        if (parentLbl && parentLbl.innerText.trim()) return parentLbl.innerText.trim();
        // 3. aria-label or placeholder
        if (el.getAttribute('aria-label')) return el.getAttribute('aria-label').trim();
        // 4. Previous sibling or table cell <td>/<th> label
        const prev = el.previousElementSibling;
        if (prev && prev.innerText && prev.innerText.trim().length < 80) return prev.innerText.trim();
        const td = el.closest('td, .form-group, .field-row');
        if (td) {
            const lblInGroup = td.querySelector('label, .label, span.title');
            if (lblInGroup && lblInGroup.innerText.trim()) return lblInGroup.innerText.trim();
        }
        return el.placeholder || el.name || el.id || 'Unnamed Field';
    }

    elements.forEach((el, idx) => {
        const tag = el.tagName.toLowerCase();
        const type = (el.getAttribute('type') || (tag === 'select' ? 'select' : tag)).toLowerCase();
        if (['hidden', 'submit', 'button', 'reset', 'image'].includes(type)) return;
        if (el.offsetParent === null && type !== 'file') return; // Skip invisible fields

        let options = [];
        if (tag === 'select') {
            options = Array.from(el.options).map(o => o.text.trim()).filter(Boolean).slice(0, 25);
        }

        const id = el.id || '';
        const name = el.name || '';
        let selector = '';
        if (id) selector = `#${CSS.escape(id)}`;
        else if (name) selector = `${tag}[name="${name}"]`;
        else selector = `${tag}:nth-of-type(${idx + 1})`;

        fields.push({
            field_index: idx + 1,
            label: getLabelText(el).replace(/\\s+/g, ' '),
            tag: tag,
            type: type,
            id: id,
            name: name,
            placeholder: el.placeholder || '',
            required: el.required || el.getAttribute('aria-required') === 'true',
            options: options,
            selector: selector
        });
    });
    return fields;
}
"""


def scan_live_page_fields(page: Any) -> List[Dict[str, Any]]:
    """
    Executes DOM inspection on the active Playwright page (and any child iframes)
    to return a clean list of form fields for Teach Mode.
    """
    all_fields: List[Dict[str, Any]] = []

    # 1. Scan main frame
    try:
        main_fields = page.evaluate(JS_DOM_SCANNER_SCRIPT)
        for f in main_fields:
            f["frame"] = "main"
            all_fields.append(f)
    except Exception as e:
        print(f"Main frame scan warning: {e}")

    # 2. Scan child iframes (common in SSO / eMitra portals)
    try:
        for frame_idx, frame in enumerate(page.frames[1:], start=1):
            try:
                frame_fields = frame.evaluate(JS_DOM_SCANNER_SCRIPT)
                for f in frame_fields:
                    f["frame"] = frame.name or f"iframe_{frame_idx}"
                    all_fields.append(f)
            except Exception:
                pass
    except Exception:
        pass

    return all_fields


def scan_html_string_fields(html_content: str) -> List[Dict[str, Any]]:
    """
    Scans raw HTML string (for offline demo or uploaded HTML portal forms) without needing a live browser.
    Uses regex/DOM parsing fallback.
    """
    import re
    fields: List[Dict[str, Any]] = []

    # Find <label for="X">Text</label>
    label_map: Dict[str, str] = {}
    for m in re.finditer(r'<label[^>]*for=["\']([^"\']+)["\'][^>]*>(.*?)</label>', html_content, re.IGNORECASE | re.DOTALL):
        clean_lbl = re.sub(r"<[^>]+>", "", m.group(2)).strip()
        label_map[m.group(1)] = clean_lbl

    # Find <input>, <select>, <textarea>
    pattern = re.compile(r'<(input|select|textarea)\b([^>]*)>', re.IGNORECASE)
    idx = 0
    for m in pattern.finditer(html_content):
        tag = m.group(1).lower()
        attrs_str = m.group(2)

        def get_attr(attr_name: str) -> str:
            am = re.search(rf'\b{attr_name}=["\']([^"\']*)["\']', attrs_str, re.IGNORECASE)
            return am.group(1) if am else ""

        f_type = (get_attr("type") or (tag if tag != "input" else "text")).lower()
        if f_type in ["hidden", "submit", "button", "reset"]:
            continue

        idx += 1
        f_id = get_attr("id")
        f_name = get_attr("name")
        f_placeholder = get_attr("placeholder")
        is_req = bool(re.search(r'\brequired\b', attrs_str, re.IGNORECASE))

        label = label_map.get(f_id) or get_attr("data-label") or f_placeholder or f_name or f_id or f"Field_{idx}"
        selector = f"#{f_id}" if f_id else (f'{tag}[name="{f_name}"]' if f_name else f"{tag}")

        fields.append({
            "field_index": idx,
            "label": label.strip(),
            "tag": tag,
            "type": f_type,
            "id": f_id,
            "name": f_name,
            "placeholder": f_placeholder,
            "required": is_req,
            "options": [],
            "selector": selector,
            "frame": "main"
        })

    return fields
