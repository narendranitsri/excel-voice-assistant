import zipfile
import os

files_to_pack = [
    "app.py",
    "actions.py",
    "llm_parser.py",
    "formulas.py",
    "quality.py",
    "charts.py",
    "workflows.py",
    "utils.py",
    "generate_samples.py",
    "requirements.txt",
    "README.md",
    "sample_orders.xlsx",
    "sample_prices.xlsx",
    "sample_inventory.xlsx"
]

os.makedirs("public", exist_ok=True)
zip_destinations = [
    "voice_controlled_excel_assistant.zip",
    "public/voice_controlled_excel_assistant.zip"
]

for zip_path in zip_destinations:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        # Add root files
        for fname in files_to_pack:
            if os.path.exists(fname):
                zipf.write(fname, arcname=fname)
                print(f"Added {fname} to {zip_path}")
            else:
                print(f"Warning: {fname} does not exist.")

        # Add samples folder
        if os.path.exists("samples"):
            for sname in os.listdir("samples"):
                spath = os.path.join("samples", sname)
                if os.path.isfile(spath):
                    zipf.write(spath, arcname=f"samples/{sname}")
                    print(f"Added samples/{sname} to {zip_path}")

print("Successfully generated ZIP files in root and in public/!")
