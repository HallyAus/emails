"""
Extract a PST to Excel with proper SMTP addresses, reply tracking, and
per-email attachment folders organised by sender.

Run on the Windows machine where the PST lives (uses Outlook via COM).

Usage:
  py pst_to_excel.py <path-to-pst> [your-email] [output-dir]

Defaults:
  your-email  ""                      (used to compute Direction / reply flags)
  output-dir  <pst-folder>/pst_export (places output next to the PST so it
                                       syncs through OneDrive)
"""

import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import win32com.client
from openpyxl import Workbook

PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
PR_SENT_REPR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D02001F"
PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"

COLUMNS = [
    "Folder", "Direction", "Date", "From", "To", "CC", "BCC",
    "Subject", "Is Reply", "Got Reply", "I Replied",
    "Conversation", "Body", "Attachments",
]
EXCEL_CELL_LIMIT = 32_767


def sanitize(name):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name or "").strip(" .") or "unnamed"


def truncate(value):
    if value is None:
        return ""
    s = str(value)
    if len(s) <= EXCEL_CELL_LIMIT:
        return s
    return s[: EXCEL_CELL_LIMIT - 20] + "...[TRUNCATED]"


def safe_component(s, max_len=80):
    s = sanitize(s)
    return s[:max_len] or "unknown"


def get_prop(obj, prop):
    try:
        v = obj.PropertyAccessor.GetProperty(prop)
        return str(v) if v else ""
    except Exception:
        return ""


def resolve_exchange(addr_entry):
    try:
        if addr_entry and addr_entry.Type == "EX":
            ex = addr_entry.GetExchangeUser()
            if ex and ex.PrimarySmtpAddress:
                return ex.PrimarySmtpAddress
    except Exception:
        pass
    return ""


def sender_smtp(msg):
    for prop in (PR_SENDER_SMTP, PR_SENT_REPR_SMTP):
        v = get_prop(msg, prop)
        if "@" in v:
            return v.lower()
    raw = getattr(msg, "SenderEmailAddress", "") or ""
    if "@" in raw:
        return raw.lower()
    try:
        smtp = resolve_exchange(msg.Sender)
        if smtp:
            return smtp.lower()
    except Exception:
        pass
    return raw.lower()


def recipient_smtp(rcp):
    v = get_prop(rcp, PR_SMTP)
    if "@" in v:
        return v.lower()
    raw = getattr(rcp, "Address", "") or ""
    if "@" in raw:
        return raw.lower()
    try:
        smtp = resolve_exchange(rcp.AddressEntry)
        if smtp:
            return smtp.lower()
    except Exception:
        pass
    return raw.lower()


def extract_recipients(msg):
    to_, cc, bcc = [], [], []
    try:
        recipients = msg.Recipients
    except Exception:
        return "", "", ""
    for k in range(1, recipients.Count + 1):
        try:
            r = recipients.Item(k)
            name = getattr(r, "Name", "") or ""
            smtp = recipient_smtp(r)
            display = f"{name} <{smtp}>" if smtp and "@" in smtp else (name or smtp)
            t = getattr(r, "Type", 1)
            if t == 1:
                to_.append(display)
            elif t == 2:
                cc.append(display)
            elif t == 3:
                bcc.append(display)
        except Exception:
            continue
    return "; ".join(to_), "; ".join(cc), "; ".join(bcc)


def iter_messages(folder, path_parts):
    current = "/".join(path_parts + [folder.Name])
    try:
        items = folder.Items
    except Exception:
        items = None
    if items is not None:
        for i in range(1, items.Count + 1):
            try:
                yield current, items.Item(i)
            except Exception as exc:
                print(f"  ! skip item {i} in {current}: {exc}")
    try:
        subs = folder.Folders
    except Exception:
        subs = None
    if subs is not None:
        for j in range(1, subs.Count + 1):
            yield from iter_messages(subs.Item(j), path_parts + [folder.Name])


def extract(pst_path, user_email, out_dir):
    user_email = (user_email or "").lower().strip()
    out_dir.mkdir(parents=True, exist_ok=True)
    attach_root = out_dir / "attachments"
    attach_root.mkdir(exist_ok=True)
    xlsx_path = out_dir / "emails.xlsx"

    print(f"Opening Outlook and attaching {pst_path} ...")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    existing = {outlook.Stores.Item(i).FilePath for i in range(1, outlook.Stores.Count + 1)}
    if str(pst_path) not in existing:
        outlook.AddStoreEx(str(pst_path), 3)
    root = None
    for i in range(1, outlook.Stores.Count + 1):
        if outlook.Stores.Item(i).FilePath == str(pst_path):
            root = outlook.Stores.Item(i).GetRootFolder()
            break
    if root is None:
        sys.exit("Could not attach PST to Outlook.")

    print("Pass 1: reading messages and saving attachments...")
    records = []
    count = 0
    for folder_path, msg in iter_messages(root, []):
        try:
            mc = getattr(msg, "MessageClass", "") or ""
        except Exception:
            continue
        if not mc.startswith("IPM.Note"):
            continue  # skip calendar/contacts/tasks/drafts of non-mail kind
        count += 1
        if count % 200 == 0:
            print(f"  ... {count} messages")

        date_obj = None
        for prop in ("ReceivedTime", "SentOn", "CreationTime"):
            try:
                v = getattr(msg, prop, None)
                if v:
                    date_obj = v
                    break
            except Exception:
                continue
        # COM datetime → naive python datetime for comparison
        try:
            date_cmp = datetime(date_obj.year, date_obj.month, date_obj.day,
                                date_obj.hour, date_obj.minute, date_obj.second) if date_obj else None
        except Exception:
            date_cmp = None
        date_str = date_obj.strftime("%Y-%m-%d %H:%M:%S") if date_obj else ""

        s_name = getattr(msg, "SenderName", "") or ""
        s_smtp = sender_smtp(msg)
        from_field = f"{s_name} <{s_smtp}>" if s_smtp and "@" in s_smtp else (s_name or s_smtp)

        to_field, cc_field, bcc_field = extract_recipients(msg)
        subject = getattr(msg, "Subject", "") or ""
        body = getattr(msg, "Body", "") or ""

        try:
            conv_id = getattr(msg, "ConversationID", "") or ""
        except Exception:
            conv_id = ""
        try:
            conv_topic = getattr(msg, "ConversationTopic", "") or subject
        except Exception:
            conv_topic = subject

        is_reply = bool(re.match(r"^\s*(re|aw|sv|antw|r):", subject, re.I))
        direction = "Sent" if user_email and s_smtp == user_email else "Received"

        # Attachments: attachments/<sender>/<YYYY-MM-DD>_<subject>/<file>
        attach_names = []
        try:
            attachments = msg.Attachments
        except Exception:
            attachments = None
        if attachments is not None and attachments.Count > 0:
            sender_dir = safe_component(s_smtp if "@" in s_smtp else (s_name or "unknown_sender"))
            date_part = date_obj.strftime("%Y-%m-%d") if date_obj else "no_date"
            subj_part = safe_component(subject, 60) or "no_subject"
            target_dir = attach_root / sender_dir / f"{date_part}_{subj_part}"
            n = 1
            while target_dir.exists():
                target_dir = attach_root / sender_dir / f"{date_part}_{subj_part}_{n}"
                n += 1
            target_dir.mkdir(parents=True, exist_ok=True)
            for k in range(1, attachments.Count + 1):
                fname = "?"
                try:
                    att = attachments.Item(k)
                    fname = safe_component(att.FileName or f"attachment_{k}", 100)
                    dest = target_dir / fname
                    m = 1
                    while dest.exists():
                        stem, dot, ext = fname.rpartition(".")
                        dest = target_dir / (f"{stem}_{m}.{ext}" if dot else f"{fname}_{m}")
                        m += 1
                    att.SaveAsFile(str(dest))
                    attach_names.append(str(dest.relative_to(attach_root)).replace("\\", "/"))
                except Exception as exc:
                    attach_names.append(f"[FAILED: {fname} - {exc}]")

        records.append({
            "folder": folder_path,
            "direction": direction,
            "date_cmp": date_cmp,
            "date_str": date_str,
            "from": from_field,
            "from_smtp": s_smtp,
            "to": to_field,
            "cc": cc_field,
            "bcc": bcc_field,
            "subject": subject,
            "is_reply": is_reply,
            "conv_id": conv_id,
            "conv_topic": conv_topic,
            "body": body,
            "attachments": "; ".join(attach_names),
        })

    print(f"Pass 2: computing reply tracking on {len(records)} messages...")
    by_conv = defaultdict(list)
    for r in records:
        if r["conv_id"]:
            by_conv[r["conv_id"]].append(r)
    for conv in by_conv.values():
        conv.sort(key=lambda r: r["date_cmp"] or datetime.min)

    for r in records:
        if not user_email or not r["conv_id"] or r["date_cmp"] is None:
            r["got_reply"] = ""
            r["i_replied"] = ""
            continue
        got_reply = "No"
        i_replied = "No"
        for other in by_conv[r["conv_id"]]:
            if other is r or other["date_cmp"] is None:
                continue
            if other["date_cmp"] <= r["date_cmp"]:
                continue
            if r["direction"] == "Sent" and other["from_smtp"] and other["from_smtp"] != user_email:
                got_reply = "Yes"
            if r["direction"] == "Received" and other["from_smtp"] == user_email:
                i_replied = "Yes"
        r["got_reply"] = got_reply
        r["i_replied"] = i_replied

    print("Pass 3: writing Excel...")
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Emails")
    ws.append(COLUMNS)
    for r in records:
        ws.append([
            r["folder"], r["direction"], r["date_str"],
            truncate(r["from"]), truncate(r["to"]), truncate(r["cc"]), truncate(r["bcc"]),
            truncate(r["subject"]),
            "Yes" if r["is_reply"] else "No",
            r["got_reply"], r["i_replied"],
            truncate(r["conv_topic"]), truncate(r["body"]), truncate(r["attachments"]),
        ])
    wb.save(xlsx_path)
    try:
        outlook.RemoveStore(root)
    except Exception:
        pass
    print(f"\nDone. {len(records)} emails written to {xlsx_path}")
    print(f"Attachments under {attach_root}")


def main():
    if len(sys.argv) < 2:
        sys.exit("Usage: py pst_to_excel.py <path-to-pst> [your-email] [output-dir]")
    pst_path = Path(sys.argv[1]).resolve()
    user_email = sys.argv[2] if len(sys.argv) > 2 else ""
    out_dir = Path(sys.argv[3]).resolve() if len(sys.argv) > 3 else (pst_path.parent / "pst_export")
    if not pst_path.exists():
        sys.exit(f"PST not found: {pst_path}")
    extract(pst_path, user_email, out_dir)


if __name__ == "__main__":
    main()
