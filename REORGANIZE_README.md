# OneDrive Reorganization — Abbott (PARA)

A PowerShell script that reorganizes `C:\Users\FINCHTB\OneDrive - Abbott` into a
**PARA** structure (Projects / Areas / Resources / Archive), built for ~20 years
of Abbott work plus personal files.

## Safety guarantees
- **Moves, never copies** — no duplicates created.
- **Folders move intact** — related context stays together.
- **Dry run by default** — nothing changes until you add `-Execute`.
- **Never deletes** — suspected junk is staged in `_TO_DELETE` for you to review.
- **Reversible** — every run writes a CSV plan and an auto-generated **undo** script
  to `OneDrive\_Reorg_Logs\`.
- **Nothing left behind silently** — anything not covered by the plan is listed as
  "left in place for you to review."

## How to run
```powershell
# 1) Preview — safe, changes nothing:
powershell -ExecutionPolicy Bypass -File .\Reorganize-OneDrive.ps1

# 2) Apply for real once the preview looks right:
powershell -ExecutionPolicy Bypass -File .\Reorganize-OneDrive.ps1 -Execute

# Custom path (if different):
powershell -ExecutionPolicy Bypass -File .\Reorganize-OneDrive.ps1 -Root "D:\OneDrive - Abbott"
```
> Tip: let OneDrive finish syncing first, or pause sync during the move so it
> isn't re-uploading mid-operation.

## Target structure
```
OneDrive - Abbott/
├── 1_Projects/        EV_Charger_Install, US_Citizenship_N-400
├── 2_Areas/           Career, Finances_and_Tax, Health, Home_and_Property,
│                      Vehicle, Identity_and_Travel, Family_and_Legal,
│                      Accounts_and_Licenses, Correspondence
├── 3_Resources/       Templates, Media/{Music,Photos}, Reference
├── 4_Archive/
│   └── Abbott_2005-2024/
│         Strategy_and_Operating_Model, Major_Programs_and_Projects,
│         M_and_A_and_Deals, Compliance_Security_Governance,
│         Brand_Marketing_Comms, Products_and_Divisions, People_and_HR,
│         Presentations_and_Knowledge, IT_Tools_and_Configs,
│         Email_Archive, _Loose_Files
├── _TO_DELETE/        scan_windows_files.bat, empty Copilot folder (review, then delete)
└── _Reorg_Logs/       CSV plan + undo script per run
```

## Renames applied (vague → dated)
Date is read from each file's real `LastWriteTime` at runtime:
- `Invoice 3786653.pdf` → `YYYY_MM_DD_Invoice_3786653.pdf`
- `Swap over.docx` → `YYYY_MM_DD_Swap_Over_REVIEW.docx`

## Left in place on purpose (ambiguous — decide yourself)
- `Plan.xlsx`, `Assets March 21.xlsx` (root) — unclear personal vs. work
- `Personal Files\smithjones.jpg` — unknown image
- `Letter to Tesco …` `.docx` + `.pdf` — kept both (editable + final); not treated as junk

## Editing the plan
All moves live in the `$Map` table at the top of `Reorganize-OneDrive.ps1`.
Each row is `Src` (relative to root) → `Dest` (relative to root), with an optional
`Rename`. Edit, re-run the dry run, then `-Execute`.
