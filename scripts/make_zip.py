"""
scripts/make_zip.py
Packages the complete "Excel <-> Website Automation Assistant" Python project
(all 13 modules + requirements.txt + README.md + templates/ + samples/)
into public/excel_website_automation_assistant.zip.
"""

import os
import shutil
import zipfile

CORE_FILES = [
    "app.py",
    "importer.py",
    "validator.py",
    "browser.py",
    "scanner.py",
    "llm.py",
    "uploader.py",
    "filler.py",
    "runner.py",
    "extractor.py",
    "reconcile.py",
    "report.py",
    "utils.py",
    "requirements.txt",
    "README.md",
]

FOLDERS_TO_INCLUDE = [
    "templates",
    "samples",
]


def build_zip():
    os.makedirs("public", exist_ok=True)
    zip_path = os.path.join("public", "excel_website_automation_assistant.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in CORE_FILES:
            if os.path.exists(fname):
                zf.write(fname, arcname=f"excel_website_automation_assistant/{fname}")

        for folder in FOLDERS_TO_INCLUDE:
            if os.path.exists(folder):
                for root, _, files in os.walk(folder):
                    for file in files:
                        full_path = os.path.join(root, file)
                        rel_path = os.path.relpath(full_path, ".")
                        zf.write(full_path, arcname=f"excel_website_automation_assistant/{rel_path}")

    # Also keep alias for backward compatibility in public/
    alias_path = os.path.join("public", "voice_controlled_excel_assistant.zip")
    shutil.copyfile(zip_path, alias_path)
    print(f"Created {zip_path} ({os.path.getsize(zip_path)} bytes)")


if __name__ == "__main__":
    build_zip()
