# watch_folder.py
# Drop this file into the same folder as main.py (your "Converter" folder).
# It watches the "input_excels" folder and automatically creates PDFs in "output_pdfs".
import os
import time
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Make paths relative to this file reliable
BASE_DIR = Path(__file__).resolve().parent
os.chdir(BASE_DIR)

# Import your existing converter logic
from main import process_excel, INPUT_FOLDER, OUTPUT_FOLDER

# Ensure output folder exists
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def is_excel(path: str) -> bool:
    p = path.lower()
    return p.endswith(".xlsx") or p.endswith(".xlsm")

class ExcelHandler(FileSystemEventHandler):
    # Handle new files copied into the input folder
    def on_created(self, event):
        if event.is_directory:
            return
        path = event.src_path
        if not is_excel(path):
            return

        # Wait until the file copy finishes (size stable)
        last_size = -1
        stable_count = 0
        for _ in range(120):  # up to ~2 minutes
            try:
                size = os.path.getsize(path)
                if size == last_size:
                    stable_count += 1
                    if stable_count >= 3:
                        break
                else:
                    stable_count = 0
                    last_size = size
            except FileNotFoundError:
                pass
            time.sleep(1)

        try:
            print(f"Processing: {path}")
            process_excel(path)
            print(f"Done: {path}")
            # Move processed file to an archive folder (optional but recommended)
            processed_dir = os.path.join(INPUT_FOLDER, "_processed")
            os.makedirs(processed_dir, exist_ok=True)
            dest = os.path.join(processed_dir, os.path.basename(path))
            # If a file with the same name exists in _processed, add a suffix
            base_name, ext = os.path.splitext(dest)
            idx = 1
            while os.path.exists(dest):
                dest = f"{base_name}_{idx}{ext}"
                idx += 1
            os.replace(path, dest)
        except Exception as e:
            print(f"ERROR processing {path}: {e}")
            failed_dir = os.path.join(INPUT_FOLDER, "_failed")
            os.makedirs(failed_dir, exist_ok=True)
            dest = os.path.join(failed_dir, os.path.basename(path))
            os.replace(path, dest)

if __name__ == "__main__":
    # Process any Excel files already present on startup
    if os.path.isdir(INPUT_FOLDER):
        for name in os.listdir(INPUT_FOLDER):
            full = os.path.join(INPUT_FOLDER, name)
            if os.path.isfile(full) and is_excel(full):
                try:
                    process_excel(full)
                except Exception as e:
                    print(f"Startup processing failed for {full}: {e}")
    else:
        os.makedirs(INPUT_FOLDER, exist_ok=True)

    observer = Observer()
    handler = ExcelHandler()
    observer.schedule(handler, INPUT_FOLDER, recursive=False)
    observer.start()
    print(f"Watching '{INPUT_FOLDER}' for new Excel files...")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
