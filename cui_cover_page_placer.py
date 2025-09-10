import os
from pathlib import Path
import xlwings as xw

TEMPLATE_FILE = r"template.xlsx"           # template workbook (sheet to copy)
ROOT_DIR = r"path\to\your\root\directory"  # root containing subfolders with Excel files
TARGET_SHEET_TITLE = "CUI Cover Page"
CUIStr = "CUI"

def file_has_cui_sheet(book: xw.Book) -> bool:
    return any(CUIStr in sht.name for sht in book.sheets)

def move_leftmost_and_rename(sht: xw.Sheet, new_name: str):
    sht.name = new_name
    # Move leftmost: xlwings index is 1-based in .api; we can copy before first sheet
    # Easiest: copy placed before the first sheet OR reposition via .api.Move
    sht.api.Move(Before=sht.book.sheets[0].api)  # leftmost

def main():
    app = xw.App(visible=False)  # headless Excel
    try:
        # Open template and get the single template sheet (first sheet)
        tpl_book = app.books.open(TEMPLATE_FILE)
        tpl_sheet = tpl_book.sheets[0]

        for root, _, files in os.walk(ROOT_DIR):
            for name in files:
                if name.startswith("~$"):
                    continue
                if not name.lower().endswith((".xlsx", ".xlsm")):
                    continue

                path = str(Path(root) / name)
                try:
                    wb = app.books.open(path)
                except Exception as e:
                    print(f"❌ Could not open {path}: {e}")
                    continue

                try:
                    if file_has_cui_sheet(wb):
                        # Move the first CUI-like sheet to leftmost
                        for sht in wb.sheets:
                            if CUIStr in sht.name:
                                # If it's not already first, move it
                                sht.api.Move(Before=wb.sheets[0].api)
                                # Optional: normalize name exactly
                                sht.name = TARGET_SHEET_TITLE
                                print(f"↔️ Moved existing '{sht.name}' leftmost in {path}")
                                break
                    else:
                        # Copy template sheet BEFORE first sheet, then rename
                        tpl_sheet.api.Copy(Before=wb.sheets[0].api)
                        new_sht = wb.sheets[0]  # the copied sheet is now leftmost
                        move_leftmost_and_rename(new_sht, TARGET_SHEET_TITLE)
                        print(f"➕ Inserted '{TARGET_SHEET_TITLE}' (leftmost) into {path}")

                    wb.save()
                except Exception as e:
                    print(f"❌ Error processing {path}: {e}")
                finally:
                    wb.close()
        tpl_book.close()
    finally:
        app.quit()

if __name__ == "__main__":
    main()
