# PST → Excel extractor

Extracts every email from a PST file into `emails.xlsx` and dumps attachments
into an `attachments/` folder.

## Requirements

- Windows with **Microsoft Outlook installed** (the script drives Outlook via COM)
- Python 3.8+ (`py --version` to check)

## Install

```powershell
py -m pip install pywin32 openpyxl
```

## Run

```powershell
py pst_to_excel.py "C:\Users\User\OneDrive - Walter Kidde Portable Equipment\emails  backup - inbox.pst"
```

Optional second argument sets the output directory (default: `./pst_export`).

```powershell
py pst_to_excel.py "C:\path\to\file.pst" "C:\path\to\output"
```

## Output

```
pst_export/
├── emails.xlsx                # Folder, Received, From, To, CC, Subject, Body, Attachments
└── attachments/
    └── <12-char-id>/
        └── invoice.pdf
        └── photo.jpg
```

The `Attachments` column in Excel lists each attachment as `<id>/<filename>` so
you can find the corresponding file on disk.

## Notes

- The file you mentioned (`emails  backup - inbox`) has no `.pst` extension in
  the path. If it's actually a PST, just pass the path as-is — Outlook detects
  the format from contents, not the extension. If Outlook refuses to open it,
  rename a copy to end in `.pst` and re-run.
- Body text is plain text (not HTML). Excel cells are capped at 32,767
  characters; anything longer is truncated with a `...[TRUNCATED]` marker — the
  full body is still in the original PST.
- The script attaches the PST to your Outlook profile while it runs and detaches
  it on completion. No emails are modified.
- A large PST (10k+ messages) may take 10–30 minutes; progress prints every 100
  messages.
