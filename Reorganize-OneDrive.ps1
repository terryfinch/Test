<#
.SYNOPSIS
    Reorganizes a OneDrive folder into a PARA structure
    (Projects / Areas / Resources / Archive) for ~20 years of Abbott + personal files.

.DESCRIPTION
    - MOVES files and folders (never copies/duplicates).
    - Keeps each folder intact so related context stays together.
    - RENAMES a small set of vaguely-named files to "YYYY_MM_DD_Topic"
      using the file's real LastWriteTime.
    - Stages suspected junk into "_TO_DELETE" (NEVER deletes anything).
    - Defaults to a DRY RUN. Nothing is touched until you re-run with -Execute.
    - Writes a timestamped CSV plan/log and an Undo script you can run to reverse moves.

.PARAMETER Root
    Path to the OneDrive root. Defaults to "C:\Users\FINCHTB\OneDrive - Abbott".

.PARAMETER Execute
    Actually perform the moves/renames. Without this switch the script only previews.

.EXAMPLE
    # 1) Preview (safe — changes nothing):
    powershell -ExecutionPolicy Bypass -File .\Reorganize-OneDrive.ps1

.EXAMPLE
    # 2) Do it for real:
    powershell -ExecutionPolicy Bypass -File .\Reorganize-OneDrive.ps1 -Execute
#>

[CmdletBinding()]
param(
    [string] $Root = "C:\Users\FINCHTB\OneDrive - Abbott",
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'
$stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$logDir  = Join-Path $Root "_Reorg_Logs"
$planCsv = Join-Path $logDir "reorg_plan_$stamp.csv"
$undoPs1 = Join-Path $logDir "reorg_undo_$stamp.ps1"

# ---------------------------------------------------------------------------
# MAPPING TABLE  ->  edit freely. Each row: Source (relative to Root), DestDir
# (relative to Root). Folders move intact; files move individually. Add an
# optional Rename to give a file a clean "YYYY_MM_DD_<Rename>" name (date is
# taken from the file at runtime, so it is always accurate).
# ---------------------------------------------------------------------------
$Map = @(
    # ---------------- 1_Projects (active, goal + end date) ----------------
    @{ Src = 'Personal Files\Electrical EV Install';              Dest = '1_Projects\EV_Charger_Install' }
    @{ Src = 'Personal Files\Mr Electric of Highland Park Transaction Receipt.pdf'; Dest = '1_Projects\EV_Charger_Install' }
    @{ Src = 'Personal Files\application N-400.pdf';              Dest = '1_Projects\US_Citizenship_N-400' }

    # ---------------- 2_Areas (ongoing responsibilities) ------------------
    @{ Src = 'Personal Files\CV Resumes';                        Dest = '2_Areas\Career' }
    @{ Src = 'Personal Files\personal deveopment';               Dest = '2_Areas\Career' }
    @{ Src = 'Professional\General Interview';                    Dest = '2_Areas\Career' }
    @{ Src = 'Professional\Goals';                               Dest = '2_Areas\Career' }

    @{ Src = 'Personal Files\Pay and tax';                       Dest = '2_Areas\Finances_and_Tax' }
    @{ Src = 'Personal Files\White Holdings';                    Dest = '2_Areas\Finances_and_Tax' }
    @{ Src = 'Professional\Finances';                            Dest = '2_Areas\Finances_and_Tax' }
    @{ Src = 'US 1099 2025.pdf';                                 Dest = '2_Areas\Finances_and_Tax' }
    @{ Src = 'UK Property Tax 2025.pdf';                         Dest = '2_Areas\Finances_and_Tax' }
    @{ Src = 'US Mortgage 1098.pdf';                             Dest = '2_Areas\Finances_and_Tax' }
    @{ Src = 'Personal Files\Invoice 3786653.pdf';               Dest = '2_Areas\Finances_and_Tax'; Rename = 'Invoice_3786653' }

    @{ Src = 'Personal Files\Health';                            Dest = '2_Areas\Health' }

    @{ Src = 'Personal Files\House';                             Dest = '2_Areas\Home_and_Property' }
    @{ Src = 'Personal Files\Land Registry';                     Dest = '2_Areas\Home_and_Property' }
    @{ Src = 'Personal Files\Control4';                          Dest = '2_Areas\Home_and_Property' }
    @{ Src = 'Personal Files\router';                            Dest = '2_Areas\Home_and_Property' }
    @{ Src = 'Personal Files\Quotation DP2213 Ensuite 12-08-2019 11-00.PDF';  Dest = '2_Areas\Home_and_Property' }
    @{ Src = 'Personal Files\Quotation DP2214 Bathroom 12-08-2019 11-00.PDF'; Dest = '2_Areas\Home_and_Property' }
    @{ Src = 'Personal Files\Swap over.docx';                    Dest = '2_Areas\Home_and_Property'; Rename = 'Swap_Over_REVIEW' }

    @{ Src = 'Personal Files\Car';                               Dest = '2_Areas\Vehicle' }

    @{ Src = 'Personal Files\Passport and Visas';                Dest = '2_Areas\Identity_and_Travel' }

    @{ Src = 'Personal Files\Divorce';                           Dest = '2_Areas\Family_and_Legal' }
    @{ Src = 'Personal Files\150 Finstad';                       Dest = '2_Areas\Family_and_Legal' }
    @{ Src = 'Personal Files\16 Wordsworth';                     Dest = '2_Areas\Family_and_Legal' }
    @{ Src = 'Personal Files\Estate Accounts FINAL.pdf';         Dest = '2_Areas\Family_and_Legal' }
    @{ Src = 'Personal Files\Forrester Hyde let and Portfolio.pdf'; Dest = '2_Areas\Family_and_Legal' }

    @{ Src = 'Personal Files\License Keys';                      Dest = '2_Areas\Accounts_and_Licenses' }
    @{ Src = 'Personal Files\Apps';                              Dest = '2_Areas\Accounts_and_Licenses' }

    @{ Src = 'Letter to Tesco for Terence Finch.docx';           Dest = '2_Areas\Correspondence' }
    @{ Src = 'Letter to Tesco for Terence Finch.pdf';            Dest = '2_Areas\Correspondence' }

    # ---------------- 3_Resources (reusable reference) --------------------
    @{ Src = 'Personal Files\Custom Office Templates';           Dest = '3_Resources\Templates' }
    @{ Src = 'Professional\Project Management Templates';        Dest = '3_Resources\Templates' }
    @{ Src = 'Professional\Office 365';                          Dest = '3_Resources\Templates' }

    @{ Src = 'Personal Files\Music';                             Dest = '3_Resources\Media\Music' }
    @{ Src = 'Personal Files\Pictures';                          Dest = '3_Resources\Media\Photos' }
    @{ Src = 'Personal Files\Terry Finch Photos';               Dest = '3_Resources\Media\Photos' }
    @{ Src = 'Personal Files\Wedding';                          Dest = '3_Resources\Media\Photos' }

    @{ Src = 'Personal Files\Useful';                            Dest = '3_Resources\Reference' }
    @{ Src = 'Personal Files\Keep File';                         Dest = '3_Resources\Reference' }
    @{ Src = 'Personal Files\Stuff';                             Dest = '3_Resources\Reference' }
    @{ Src = 'Personal Files\Attachments';                       Dest = '3_Resources\Reference' }
    @{ Src = 'Personal Files\Public';                            Dest = '3_Resources\Reference' }

    # ---------------- 4_Archive\Abbott_2005-2024 (the 20 years) -----------
    # Strategy & Operating Model
    @{ Src = 'Professional\IT Strategy';                         Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }
    @{ Src = 'Professional\DIT Operating Model';                 Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }
    @{ Src = 'Professional\Transformation';                      Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }
    @{ Src = 'Professional\Leadership Strategy';                 Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }
    @{ Src = 'Professional\Microsoft Strategy';                  Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }
    @{ Src = 'Professional\Procurement Strategy';                Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }
    @{ Src = 'Professional\Reference Architecture';              Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }
    @{ Src = 'Professional\Business Architecture';               Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }
    @{ Src = 'Professional\MCOE Reboot';                         Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }
    @{ Src = 'Professional\Organization Design';                 Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }
    @{ Src = 'Professional\Governance';                          Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }
    @{ Src = 'Professional\100 day plans';                       Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }
    @{ Src = 'Professional\infrastructure technical strategy';   Dest = '4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model' }

    # Major Programs & Projects
    @{ Src = 'Professional\Application Rationalization';         Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\Project Go';                          Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\Project Spring';                      Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\Projects';                            Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\eCommerce';                           Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\Digital Heallth';                     Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\Big Data and Analytics';              Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\Workday';                             Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\Digital Summit & SDMC';               Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\War on Waste';                        Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\TCO Exercise 2016';                   Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\2010 Managed Services Deal';          Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\2014 BSS Savings';                    Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\2014 Task and Mckinsey';              Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\2015 ESP Finance';                    Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }
    @{ Src = 'Professional\Innovation & Hackathorn';             Dest = '4_Archive\Abbott_2005-2024\Major_Programs_and_Projects' }

    # M&A & Deals
    @{ Src = 'Professional\M and A Details';                     Dest = '4_Archive\Abbott_2005-2024\M_and_A_and_Deals' }
    @{ Src = 'Professional\Licensing Deals';                     Dest = '4_Archive\Abbott_2005-2024\M_and_A_and_Deals' }
    @{ Src = 'Professional\Alere';                               Dest = '4_Archive\Abbott_2005-2024\M_and_A_and_Deals' }
    @{ Src = 'Professional\Mylan';                               Dest = '4_Archive\Abbott_2005-2024\M_and_A_and_Deals' }
    @{ Src = 'Professional\Safe Harbor';                         Dest = '4_Archive\Abbott_2005-2024\M_and_A_and_Deals' }
    @{ Src = 'Professional\One Abbott Freight';                  Dest = '4_Archive\Abbott_2005-2024\M_and_A_and_Deals' }

    # Compliance, Security & Risk
    @{ Src = 'Professional\123Compliance';                      Dest = '4_Archive\Abbott_2005-2024\Compliance_Security_Governance' }
    @{ Src = 'Professional\Compliance Audit _ Sox';             Dest = '4_Archive\Abbott_2005-2024\Compliance_Security_Governance' }
    @{ Src = 'Professional\Information Security';                Dest = '4_Archive\Abbott_2005-2024\Compliance_Security_Governance' }
    @{ Src = 'Professional\PCI';                                Dest = '4_Archive\Abbott_2005-2024\Compliance_Security_Governance' }
    @{ Src = 'Professional\Rims';                               Dest = '4_Archive\Abbott_2005-2024\Compliance_Security_Governance' }

    # Brand, Marketing & Comms
    @{ Src = 'Professional\Abbott Corp Branding';               Dest = '4_Archive\Abbott_2005-2024\Brand_Marketing_Comms' }
    @{ Src = 'Professional\Social Media';                       Dest = '4_Archive\Abbott_2005-2024\Brand_Marketing_Comms' }
    @{ Src = 'Professional\Yammer';                             Dest = '4_Archive\Abbott_2005-2024\Brand_Marketing_Comms' }
    @{ Src = 'Professional\eDetailing';                         Dest = '4_Archive\Abbott_2005-2024\Brand_Marketing_Comms' }
    @{ Src = 'Professional\EVP Portal Presentation';            Dest = '4_Archive\Abbott_2005-2024\Brand_Marketing_Comms' }
    @{ Src = 'Professional\Path To Purchase';                   Dest = '4_Archive\Abbott_2005-2024\Brand_Marketing_Comms' }

    # Products & Divisions
    @{ Src = 'Professional\ADD Informatics';                    Dest = '4_Archive\Abbott_2005-2024\Products_and_Divisions' }
    @{ Src = 'Professional\GMO';                                Dest = '4_Archive\Abbott_2005-2024\Products_and_Divisions' }
    @{ Src = 'Professional\Core IT';                            Dest = '4_Archive\Abbott_2005-2024\Products_and_Divisions' }

    # People & HR
    @{ Src = 'Professional\Berce Staff Meetings';               Dest = '4_Archive\Abbott_2005-2024\People_and_HR' }
    @{ Src = 'Professional\staff';                              Dest = '4_Archive\Abbott_2005-2024\People_and_HR' }
    @{ Src = 'Professional\Employment Tribunal';               Dest = '4_Archive\Abbott_2005-2024\People_and_HR' }
    @{ Src = 'Personal Files\Abbott Employment Contracts';     Dest = '4_Archive\Abbott_2005-2024\People_and_HR' }

    # Presentations & Knowledge
    @{ Src = 'Professional\Presentations';                      Dest = '4_Archive\Abbott_2005-2024\Presentations_and_Knowledge' }
    @{ Src = 'Professional\Research';                           Dest = '4_Archive\Abbott_2005-2024\Presentations_and_Knowledge' }
    @{ Src = 'Professional\Training';                           Dest = '4_Archive\Abbott_2005-2024\Presentations_and_Knowledge' }
    @{ Src = 'Professional\pEX';                                Dest = '4_Archive\Abbott_2005-2024\Presentations_and_Knowledge' }
    @{ Src = 'Professional\Operations Readiness';              Dest = '4_Archive\Abbott_2005-2024\Presentations_and_Knowledge' }
    @{ Src = 'Professional\Notability';                        Dest = '4_Archive\Abbott_2005-2024\Presentations_and_Knowledge' }
    @{ Src = 'Professional\Notebooks';                         Dest = '4_Archive\Abbott_2005-2024\Presentations_and_Knowledge' }

    # IT Tools & Configs
    @{ Src = 'Professional\CiscoCfg';                          Dest = '4_Archive\Abbott_2005-2024\IT_Tools_and_Configs' }
    @{ Src = 'Professional\sql';                               Dest = '4_Archive\Abbott_2005-2024\IT_Tools_and_Configs' }
    @{ Src = 'Professional\slc';                               Dest = '4_Archive\Abbott_2005-2024\IT_Tools_and_Configs' }
    @{ Src = 'Professional\Snagit';                            Dest = '4_Archive\Abbott_2005-2024\IT_Tools_and_Configs' }
    @{ Src = 'Professional\IBM Lotus Notes ID';               Dest = '4_Archive\Abbott_2005-2024\IT_Tools_and_Configs' }
    @{ Src = 'Professional\Licence Keys';                      Dest = '4_Archive\Abbott_2005-2024\IT_Tools_and_Configs' }
    @{ Src = 'Professional\Favorites';                         Dest = '4_Archive\Abbott_2005-2024\IT_Tools_and_Configs' }

    # Email archive
    @{ Src = 'Professional\OneEmail';                          Dest = '4_Archive\Abbott_2005-2024\Email_Archive' }

    # Loose Abbott files
    @{ Src = 'Professional\Abbott At Home Testing Experience-Comms and Timeline 022721 v18[93].docx'; Dest = '4_Archive\Abbott_2005-2024\_Loose_Files' }
    @{ Src = 'Professional\APOC_2017_Pepper_Order_Form_21_20170721.docx';      Dest = '4_Archive\Abbott_2005-2024\_Loose_Files' }
    @{ Src = 'Professional\Aug Exec Program Review TBF and VC comments.pptx';  Dest = '4_Archive\Abbott_2005-2024\_Loose_Files' }
    @{ Src = 'Professional\health07_hi-res.wmv';                              Dest = '4_Archive\Abbott_2005-2024\_Loose_Files' }
    @{ Src = 'Professional\NAVICA - Keep it Simple Schools DAFTT for discusion.pptx'; Dest = '4_Archive\Abbott_2005-2024\_Loose_Files' }
    @{ Src = 'Professional\NAVICA - Keep it Simple Schools v1.pptx';          Dest = '4_Archive\Abbott_2005-2024\_Loose_Files' }
    @{ Src = 'Professional\Single Device Only At-Home Testing TF Update (No Visuals).pptx'; Dest = '4_Archive\Abbott_2005-2024\_Loose_Files' }
    @{ Src = 'Professional\Strategic focus and Op Model Value summary Finch Updates.docx';  Dest = '4_Archive\Abbott_2005-2024\_Loose_Files' }
    @{ Src = 'Abbott_GenAi_090825.pptx';                                      Dest = '4_Archive\Abbott_2005-2024\_Loose_Files' }
    @{ Src = 'Landor_AbbottMotion_jb_Pipette_Slide_v002.mp4';                 Dest = '4_Archive\Abbott_2005-2024\_Loose_Files' }
)

# ---------------------------------------------------------------------------
# JUNK  ->  staged into _TO_DELETE. NEVER deleted; you review and empty it.
# ---------------------------------------------------------------------------
$Junk = @(
    'Personal Files\scan_windows_files.bat'   # stray utility script of unknown origin
    'Microsoft Copilot Chat Files'            # empty system folder
)

# ---------------------------------------------------------------------------
# Items deliberately NOT moved (ambiguous — left in place for you to decide).
# Reported at the end so they are visible, not silently ignored.
#   Root\Plan.xlsx            - unclear personal vs work
#   Root\Assets March 21.xlsx - unclear personal vs work
#   Personal Files\smithjones.jpg - unknown image
# ---------------------------------------------------------------------------

# ===========================================================================
#  ENGINE  (no need to edit below)
# ===========================================================================
function New-DirIfNeeded([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) {
        if ($Execute) { New-Item -ItemType Directory -Path $path -Force | Out-Null }
    }
}

function Resolve-Collision([string]$targetPath) {
    # If targetPath already exists, append " (2)", " (3)", ... to the leaf.
    if (-not (Test-Path -LiteralPath $targetPath)) { return $targetPath }
    $dir  = Split-Path $targetPath -Parent
    $name = [IO.Path]::GetFileNameWithoutExtension($targetPath)
    $ext  = [IO.Path]::GetExtension($targetPath)
    $n = 2
    while ($true) {
        $candidate = Join-Path $dir ("{0} ({1}){2}" -f $name, $n, $ext)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
        $n++
    }
}

$actions = New-Object System.Collections.Generic.List[object]
$missing = New-Object System.Collections.Generic.List[string]

function Add-Action($srcRel, $destDirRel, $rename, $kind) {
    $srcFull = Join-Path $Root $srcRel
    if (-not (Test-Path -LiteralPath $srcFull)) { $missing.Add($srcRel); return }

    $item   = Get-Item -LiteralPath $srcFull
    $isDir  = $item.PSIsContainer
    $leaf   = Split-Path $srcRel -Leaf

    if ($rename -and -not $isDir) {
        $datePrefix = $item.LastWriteTime.ToString('yyyy_MM_dd')
        $ext        = [IO.Path]::GetExtension($leaf)
        $leaf       = "{0}_{1}{2}" -f $datePrefix, $rename, $ext
    }

    $destDirFull = Join-Path $Root $destDirRel
    $targetFull  = Join-Path $destDirFull $leaf
    $finalFull   = Resolve-Collision $targetFull

    $actions.Add([pscustomobject]@{
        Kind        = $kind
        Type        = if ($isDir) { 'Folder' } else { 'File' }
        Source      = $srcRel
        Destination = (Join-Path $destDirRel (Split-Path $finalFull -Leaf))
        SourceFull  = $srcFull
        DestDirFull = $destDirFull
        TargetFull  = $finalFull
        Renamed     = [bool]($rename -and -not $isDir)
    })
}

foreach ($e in $Map)  { Add-Action $e.Src $e.Dest $e.Rename 'MOVE' }
foreach ($j in $Junk) { Add-Action $j '_TO_DELETE' $null 'JUNK' }

# ---- Report ---------------------------------------------------------------
$mode = if ($Execute) { 'EXECUTE (changes WILL be made)' } else { 'DRY RUN (no changes)' }
Write-Host ""
Write-Host "OneDrive PARA reorganization — $mode" -ForegroundColor Cyan
Write-Host "Root: $Root"
Write-Host ("Planned actions: {0}   Missing sources: {1}" -f $actions.Count, $missing.Count)
Write-Host ""

$actions | Sort-Object Destination, Source |
    Format-Table Kind, Type, Source, Destination, Renamed -AutoSize | Out-String -Width 4096 | Write-Host

if ($missing.Count) {
    Write-Host "Sources not found (skipped — already moved or renamed?):" -ForegroundColor Yellow
    $missing | ForEach-Object { Write-Host "   $_" -ForegroundColor Yellow }
    Write-Host ""
}

# Surface anything at the top levels that the plan does NOT cover, so new /
# unknown items are never silently left behind.
$plannedTop = @{}
foreach ($a in $actions) {
    $parts = $a.Source -split '\\'
    $top = if ($parts.Count -ge 2) { "$($parts[0])\$($parts[1])" } else { $parts[0] }
    $plannedTop[$top] = $true
}
$unmapped = New-Object System.Collections.Generic.List[string]
foreach ($container in @('.', 'Personal Files', 'Professional')) {
    $base = if ($container -eq '.') { $Root } else { Join-Path $Root $container }
    if (-not (Test-Path -LiteralPath $base)) { continue }
    Get-ChildItem -LiteralPath $base -Force | ForEach-Object {
        $rel = if ($container -eq '.') { $_.Name } else { "$container\$($_.Name)" }
        # ignore the new PARA folders and our log/staging dirs
        if ($_.Name -match '^(1_Projects|2_Areas|3_Resources|4_Archive|_TO_DELETE|_Reorg_Logs)$') { return }
        if (-not $plannedTop.ContainsKey($rel)) { $unmapped.Add($rel) }
    }
}
if ($unmapped.Count) {
    Write-Host "NOT in plan — left in place for you to review:" -ForegroundColor Magenta
    $unmapped | Sort-Object -Unique | ForEach-Object { Write-Host "   $_" -ForegroundColor Magenta }
    Write-Host ""
}

# ---- Execute --------------------------------------------------------------
if (-not $Execute) {
    Write-Host "This was a DRY RUN. Re-run with  -Execute  to apply." -ForegroundColor Green
    Write-Host "Review the plan above first." -ForegroundColor Green
    return
}

New-DirIfNeeded $logDir
$actions | Export-Csv -LiteralPath $planCsv -NoTypeInformation -Encoding UTF8

# Undo script header
$undo = New-Object System.Collections.Generic.List[string]
$undo.Add('# Auto-generated undo script. Run to reverse the moves from this batch.')
$undo.Add('$ErrorActionPreference = ''Stop''')

$done = 0; $failed = 0
foreach ($a in ($actions | Sort-Object Destination)) {
    try {
        New-DirIfNeeded $a.DestDirFull
        Move-Item -LiteralPath $a.SourceFull -Destination $a.TargetFull -Force:$false
        # record reverse move
        $undo.Add(('Move-Item -LiteralPath "{0}" -Destination "{1}"' -f $a.TargetFull, $a.SourceFull))
        $done++
        Write-Host ("[OK]   {0}  ->  {1}" -f $a.Source, $a.Destination) -ForegroundColor Green
    } catch {
        $failed++
        Write-Host ("[FAIL] {0}  ({1})" -f $a.Source, $_.Exception.Message) -ForegroundColor Red
    }
}

Set-Content -LiteralPath $undoPs1 -Value $undo -Encoding UTF8

Write-Host ""
Write-Host ("Done. Moved: {0}   Failed: {1}" -f $done, $failed) -ForegroundColor Cyan
Write-Host "Plan log : $planCsv"
Write-Host "Undo     : $undoPs1   (run this to reverse)"
Write-Host ""
Write-Host "NOTE: _TO_DELETE holds suspected junk. Nothing was deleted." -ForegroundColor Yellow
Write-Host "      Review it, then delete manually when you are satisfied." -ForegroundColor Yellow
