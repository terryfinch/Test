# OneDrive Reorganization — Abbott (PARA) · PowerShell 7

`Reorganize-OneDrive.ps1` reorganizes `C:\Users\FINCHTB\OneDrive - Abbott` into a
**PARA** structure (Projects / Areas / Resources / Archive), built for ~20 years
of Abbott work plus personal files. **Requires PowerShell 7.0+** (`pwsh`).

## Safety guarantees
- **Moves, never copies** — no duplicates created.
- **Folders move intact** — related context stays together.
- **Dry run by default** — nothing changes until you add `-Execute`.
- **Never deletes** — suspected junk/duplicates are staged in `_TO_DELETE\` for review.
- **Reversible** — each run writes a CSV plan and an auto **undo** script to `_Reorg_Logs\`.
- **Nothing left behind silently** — anything not covered is listed as "left in place to review."

## How to run
```powershell
# 1) Preview everything (safe):
pwsh -File .\Reorganize-OneDrive.ps1

# 2) Preview + duplicate report:
pwsh -File .\Reorganize-OneDrive.ps1 -FindDuplicates

# 3) Apply for real (interactive — confirms Extracts mappings + a final go/no-go):
pwsh -File .\Reorganize-OneDrive.ps1 -Execute -FindDuplicates

# Unattended (no prompts; uncertain Extracts items go to 0_REVIEW_Unsorted):
pwsh -File .\Reorganize-OneDrive.ps1 -Execute -NonInteractive
```
> Pause OneDrive sync during the run so it isn't re-uploading mid-move.

## What's new in this version
- **PowerShell 7 optimized** — ternary/`??` operators, `ForEach-Object -Parallel`
  for fast hashing, `Write-Progress` throughout.
- **`Extracts\` multi-pass review** — the script discovers everything under
  `OneDrive - Abbott\Extracts`, then classifies each top-level item in up to 4 passes:
  1. **Pass 1** — folder/file name vs. category keywords
  2. **Pass 2** — immediate child names (for vague folder names)
  3. **Pass 3** — deep file names (capped at 300)
  4. **Pass 4** — file-type hint (e.g. lots of `.jpg` → Photos, `.pst` → Email)
  Each item gets a proposed destination + confidence + the evidence that matched.
- **Asks you to confirm** every uncertain mapping:
  `[A]ccept  [C]hange  [V]iew  [S]kip  [Q]uit`
  (`Change` shows a numbered category menu; `View` lists the folder's contents).
  High-confidence items (score ≥ `-ConfidenceThreshold`, default 3) are shown but
  not blocked on. There's also a final **"Apply N moves now? [y/N]"** gate.
- **Duplicate finder** (`-FindDuplicates`) — pre-filters by **filename + size**, then
  confirms true duplicates by **SHA-256** (parallel). Reports each set (marking the
  oldest copy as the keeper). With `-Execute` it can move the extra copies into
  `_TO_DELETE\_Duplicates\` (kept, never deleted) after you confirm.
- **Progress + summary** — live progress bars for classify/dedupe/move and an
  end-of-run summary (moved / failed / log paths).

## Target structure
```
OneDrive - Abbott/
├── 1_Projects/        EV_Charger_Install, US_Citizenship_N-400
├── 2_Areas/           Career, Finances_and_Tax, Health, Home_and_Property,
│                      Vehicle, Identity_and_Travel, Family_and_Legal,
│                      Accounts_and_Licenses, Correspondence
├── 3_Resources/       Templates, Media/{Music,Photos}, Reference
├── 4_Archive/Abbott_2005-2024/
│                      Strategy_and_Operating_Model, Major_Programs_and_Projects,
│                      M_and_A_and_Deals, Compliance_Security_Governance,
│                      Brand_Marketing_Comms, Products_and_Divisions, People_and_HR,
│                      Presentations_and_Knowledge, IT_Tools_and_Configs,
│                      Email_Archive, _Loose_Files
├── 0_REVIEW_Unsorted/ Extracts items the script couldn't place confidently
├── _TO_DELETE/        junk + _Duplicates/ (review, then delete manually)
└── _Reorg_Logs/       CSV plan, duplicate report, and undo script per run
```

## Tuning the classifier
- **Destinations + keywords** live in the `$Cat` registry at the top of the script.
  Add a keyword to nudge how `Extracts` items are sorted, or change a `Path`.
- **Known Personal/Professional items** live in the `$Map` table — edit `Src`/`Dest`/`Rename`.
- **Confidence gate**: `-ConfidenceThreshold <int>` (default 3). Lower = more auto-accepts.

## Left in place on purpose (ambiguous)
`Plan.xlsx`, `Assets March 21.xlsx` (root), `Personal Files\smithjones.jpg`, and the
`Letter to Tesco` `.docx`/`.pdf` pair (kept both). These are reported, not moved.
