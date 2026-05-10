"""
Extract a PST file to an Excel workbook.

Run this on the Windows machine where the PST lives (it uses Outlook via COM,
so Microsoft Outlook must be installed and able to open the PST).

Output:
  - emails.xlsx              : one row per message (Folder, Date, From, To, CC,
                               Subject, Body, Attachments)
  - attachments/<msg_id>/... : every attachment, grouped by message

Setup (PowerShell or cmd):
  py -m pip install pywin32 openpyxl
  py pst_to_excel.py "C:\\Users\\User\\OneDrive - Walter Kidde Portable Equipment\\emails  backup - inbox.pst"

If the path has no .pst suffix on disk, pass the actual filename as-is.
"""

import os
import re
import sys
import uuid
from pathlib import Path

import win32com.client  # pywin32
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

COLUMNS = ["Folder", "Received", "From", "To", "CC", "Subject", "Body", "Attachments"]
EXCEL_CELL_LIMIT = 32_767  # Excel hard limit per cell


def sanitize(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .") or "unnamed"


def truncate_for_excel(value: str) -> str:
    if value is None:
        return ""
    if len(value) <= EXCEL_CELL_LIMIT:
        return value
    return value[: EXCEL_CELL_LIMIT - 20] + "...[TRUNCATED]"


def iter_messages(folder, path_parts):
    current_path = "/".join(path_parts + [folder.Name])
    try:
        items = folder.Items
    except Exception:
        items = None
    if items is not None:
        for i in range(1, items.Count + 1):
            try:
                yield current_path, items.Item(i)
            except Exception as exc:
                print(f"  ! could not read item {i} in {current_path}: {exc}")
    try:
        subfolders = folder.Folders
    except Exception:
        subfolders = None
    if subfolders is not None:
        for j in range(1, subfolders.Count + 1):
            yield from iter_messages(subfolders.Item(j), path_parts + [folder.Name])


def extract(pst_path: Path, out_dir: Path) -> None:
    if not pst_path.exists():
        sys.exit(f"PST not found: {pst_path}")

    out_dir.mkdir(parents=True, exist_ok=True)
    attach_root = out_dir / "attachments"
    attach_root.mkdir(exist_ok=True)
    xlsx_path = out_dir / "emails.xlsx"

    print(f"Opening Outlook and attaching {pst_path} ...")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    existing = {outlook.Stores.Item(i).FilePath for i in range(1, outlook.Stores.Count + 1)}
    if str(pst_path) not in existing:
        outlook.AddStoreEx(str(pst_path), 3)  # 3 = olStoreUnicode

    root = None
    for i in range(1, outlook.Stores.Count + 1):
        if outlook.Stores.Item(i).FilePath == str(pst_path):
            root = outlook.Stores.Item(i).GetRootFolder()
            break
    if root is None:
        sys.exit("Could not attach the PST to Outlook.")

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Emails")
    ws.append(COLUMNS)

    count = 0
    for folder_path, msg in iter_messages(root, []):
        count += 1
        if count % 100 == 0:
            print(f"  ... {count} messages")

        try:
            received = getattr(msg, "ReceivedTime", None) or getattr(msg, "SentOn", None)
            received_str = received.strftime("%Y-%m-%d %H:%M:%S") if received else ""
        except Exception:
            received_str = ""

        sender = getattr(msg, "SenderName", "") or ""
        sender_email = getattr(msg, "SenderEmailAddress", "") or ""
        from_field = f"{sender} <{sender_email}>" if sender_email else sender
        to_field = getattr(msg, "To", "") or ""
        cc_field = getattr(msg, "CC", "") or ""
        subject = getattr(msg, "Subject", "") or ""
        body = getattr(msg, "Body", "") or ""

        attach_names = []
        try:
            attachments = msg.Attachments
        except Exception:
            attachments = None
        if attachments is not None and attachments.Count > 0:
            msg_id = uuid.uuid4().hex[:12]
            msg_dir = attach_root / msg_id
            msg_dir.mkdir(parents=True, exist_ok=True)
            for k in range(1, attachments.Count + 1):
                att = attachments.Item(k)
                fname = sanitize(att.FileName or f"attachment_{k}")
                target = msg_dir / fname
                # avoid collisions
                n = 1
                while target.exists():
                    stem, dot, ext = fname.rpartition(".")
                    target = msg_dir / (f"{stem}_{n}.{ext}" if dot else f"{fname}_{n}")
                    n += 1
                try:
                    att.SaveAsFile(str(target))
                    attach_names.append(f"{msg_id}/{target.name}")
                except Exception as exc:
                    attach_names.append(f"[FAILED: {fname} - {exc}]")

        ws.append([
            folder_path,
            received_str,
            truncate_for_excel(from_field),
            truncate_for_excel(to_field),
            truncate_for_excel(cc_field),
            truncate_for_excel(subject),
            truncate_for_excel(body),
            truncate_for_excel("; ".join(attach_names)),
        ])

    # set reasonable widths via a second sheet pass not possible in write_only mode;
    # widths can be tweaked after open. Save and detach.
    wb.save(xlsx_path)
    try:
        outlook.RemoveStore(root)
    except Exception:
        pass

    print(f"\nDone. {count} messages written to {xlsx_path}")
    print(f"Attachments saved under {attach_root}")


def main() -> None:
    if len(sys.argv) < 2:
        sys.exit("Usage: py pst_to_excel.py <path-to-pst> [output-dir]")
    pst_path = Path(sys.argv[1]).resolve()
    out_dir = Path(sys.argv[2]).resolve() if len(sys.argv) > 2 else Path.cwd() / "pst_export"
    extract(pst_path, out_dir)


if __name__ == "__main__":
    main()
