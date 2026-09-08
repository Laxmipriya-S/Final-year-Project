import os
import shutil
import time
from datetime import datetime
import openpyxl

# CONFIG
DASHBOARD_FOLDER = "uploads"  # from app.config['UPLOAD_FOLDER']
ARCHIVE_FOLDER = "archived_files"
EXCEL_LOG = "archived_files_log.xlsx"
DAYS_TO_KEEP_UPLOADS = 7       # Number of days before moving files to archive
DAYS_TO_KEEP_ARCHIVE = 30      # Number of days before deleting archived files

# Ensure archive folder exists
os.makedirs(ARCHIVE_FOLDER, exist_ok=True)

def log_to_excel(file_name, original_path, archive_path):
    """Log archive action to Excel."""
    if not os.path.exists(EXCEL_LOG):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["File Name", "Original Path", "Archive Path", "Date Archived"])
        wb.save(EXCEL_LOG)

    wb = openpyxl.load_workbook(EXCEL_LOG)
    ws = wb.active
    ws.append([
        file_name,
        original_path,
        archive_path,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ])
    wb.save(EXCEL_LOG)

def cleanup_old_files():
    """Move old uploads to archive."""
    now = time.time()
    for file_name in os.listdir(DASHBOARD_FOLDER):
        file_path = os.path.join(DASHBOARD_FOLDER, file_name)
        if os.path.isfile(file_path):
            file_age_days = (now - os.path.getmtime(file_path)) / 86400
            if file_age_days > DAYS_TO_KEEP_UPLOADS:
                archive_path = os.path.join(ARCHIVE_FOLDER, file_name)
                shutil.move(file_path, archive_path)
                log_to_excel(file_name, file_path, archive_path)
                print(f"Archived: {file_name}")

def delete_old_archives():
    """Delete archived files older than DAYS_TO_KEEP_ARCHIVE."""
    now = time.time()
    for file_name in os.listdir(ARCHIVE_FOLDER):
        file_path = os.path.join(ARCHIVE_FOLDER, file_name)
        if os.path.isfile(file_path):
            file_age_days = (now - os.path.getmtime(file_path)) / 86400
            if file_age_days > DAYS_TO_KEEP_ARCHIVE:
                try:
                    os.remove(file_path)
                    print(f"Deleted old archive: {file_name}")
                except Exception as e:
                    print(f"Error deleting {file_name}: {e}")

if __name__ == "__main__":
    cleanup_old_files()
    delete_old_archives()