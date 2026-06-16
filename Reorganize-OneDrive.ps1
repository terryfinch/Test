#Requires -Version 7.0
<#
.SYNOPSIS
    Reorganizes "C:\Users\FINCHTB\OneDrive - Abbott" into a PARA structure
    (Projects / Areas / Resources / Archive). PowerShell 7.x.

.DESCRIPTION
    * MOVES files/folders (never copies) — no duplicates created.
    * Folders move intact so related context stays together.
    * Renames vague files to "YYYY_MM_DD_Topic" from their real LastWriteTime.
    * Reviews the unstructured "Extracts\" tree in MULTIPLE PASSES and proposes a
      home for each item, then ASKS YOU to confirm/redirect each one.
    * Optional duplicate finder: groups by filename + size, confirms by SHA-256
      (parallel), and can stage extras into _TO_DELETE\_Duplicates (never deletes).
    * Live PROGRESS reporting + a final summary.
    * DRY RUN by default. Nothing changes until you add -Execute.
    * Reversible: writes a CSV plan and an auto undo script to _Reorg_Logs\.

.PARAMETER Root
    OneDrive root. Default "C:\Users\FINCHTB\OneDrive - Abbott".

.PARAMETER Execute
    Apply changes. Without it, the script previews only.

.PARAMETER FindDuplicates
    Also scan for duplicate files (filename + metadata + hash).

.PARAMETER NonInteractive
    Don't prompt. High-confidence Extracts items are auto-accepted; everything
    uncertain goes to 0_REVIEW_Unsorted instead of asking.

.PARAMETER ConfidenceThreshold
    Score at/above which an Extracts classification is treated as High confidence
    (auto-accepted in -NonInteractive). Default 3.

.EXAMPLE
    pwsh -File .\Reorganize-OneDrive.ps1                       # preview everything
.EXAMPLE
    pwsh -File .\Reorganize-OneDrive.ps1 -FindDuplicates       # preview + dup report
.EXAMPLE
    pwsh -File .\Reorganize-OneDrive.ps1 -Execute             # apply (interactive)
#>

[CmdletBinding()]
param(
    [string] $Root = "C:\Users\FINCHTB\OneDrive - Abbott",
    [string] $ExtractsFolder = "Extracts",
    [switch] $Execute,
    [switch] $FindDuplicates,
    [switch] $HydrateCloudFiles,
    [switch] $NonInteractive,
    [int]    $ConfidenceThreshold = 3
)

$ErrorActionPreference = 'Stop'
$stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$logDir  = Join-Path $Root '_Reorg_Logs'
$planCsv = Join-Path $logDir "reorg_plan_$stamp.csv"
$dupCsv  = Join-Path $logDir "duplicates_$stamp.csv"
$undoPs1 = Join-Path $logDir "reorg_undo_$stamp.ps1"
$canPrompt = -not $NonInteractive -and [Environment]::UserInteractive

# ===========================================================================
#  CATEGORY REGISTRY  — single source of truth for destinations + keywords.
#  Keywords drive the Extracts auto-classifier (name + contents matching).
# ===========================================================================
$Cat = [ordered]@{
  'Projects.EV'        = @{ Path='1_Projects\EV_Charger_Install';        Keywords=@('ev charger','charge point','wallbox','electric vehicle') }
  'Projects.N400'      = @{ Path='1_Projects\US_Citizenship_N-400';      Keywords=@('n-400','n400','citizenship','naturalization') }
  'Areas.Career'       = @{ Path='2_Areas\Career';                       Keywords=@('cv','resume','interview','linkedin','cover letter','appraisal','career','goals','development plan') }
  'Areas.Finance'      = @{ Path='2_Areas\Finances_and_Tax';             Keywords=@('tax','1099','1098','w2','w-2','invoice','receipt','statement','bank','mortgage','payslip','salary','finance','pension','401k','expenses','hmrc','irs','budget','forecast') }
  'Areas.Health'       = @{ Path='2_Areas\Health';                       Keywords=@('health','medical','doctor','prescription','nhs','dental','vaccin','lab result','clinic','therapy') }
  'Areas.Home'         = @{ Path='2_Areas\Home_and_Property';            Keywords=@('house','home','property','land registry','control4','router','electrical','plumbing','quotation','renovation','boiler','ensuite','bathroom','kitchen','builder') }
  'Areas.Vehicle'      = @{ Path='2_Areas\Vehicle';                      Keywords=@('car','vehicle','mot','dvla','tyre','tire','service history','registration') }
  'Areas.Identity'     = @{ Path='2_Areas\Identity_and_Travel';          Keywords=@('passport','visa','green card','immigration','birth certificate','driver license','driving licence','travel') }
  'Areas.Family'       = @{ Path='2_Areas\Family_and_Legal';             Keywords=@('divorce','estate','will','probate','custody','legal','solicitor','finstad','wordsworth','forrester','settlement') }
  'Areas.Accounts'     = @{ Path='2_Areas\Accounts_and_Licenses';        Keywords=@('license key','licence key','password','account','subscription','software key','activation') }
  'Areas.Correspondence' = @{ Path='2_Areas\Correspondence';             Keywords=@('letter','correspondence','complaint','tesco') }
  'Resources.Templates'= @{ Path='3_Resources\Templates';                Keywords=@('template','boilerplate','dotx','potx','xltx') }
  'Resources.Music'    = @{ Path='3_Resources\Media\Music';              Keywords=@('music','album','playlist','mp3','flac','soundtrack') }
  'Resources.Photos'   = @{ Path='3_Resources\Media\Photos';             Keywords=@('photo','picture','img','image','wedding','holiday pics','camera') }
  'Resources.Reference'= @{ Path='3_Resources\Reference';                Keywords=@('useful','misc','reference','keep','manual','guide','how to') }
  'Archive.Strategy'   = @{ Path='4_Archive\Abbott_2005-2024\Strategy_and_Operating_Model'; Keywords=@('strategy','operating model','architecture','roadmap','governance','transformation','org design','100 day','target operating','tom','blueprint','vision') }
  'Archive.Programs'   = @{ Path='4_Archive\Abbott_2005-2024\Major_Programs_and_Projects';  Keywords=@('project','program','programme','rollout','deployment','migration','workday','ecommerce','ariba','sap','sdmc','hackathon','innovation','tco','savings','mckinsey','esp','bss','managed services','path to purchase') }
  'Archive.MandA'      = @{ Path='4_Archive\Abbott_2005-2024\M_and_A_and_Deals';            Keywords=@('acquisition','merger','m&a','due diligence','divestiture','alere','mylan','integration','safe harbor','licensing','deal','freight') }
  'Archive.Compliance' = @{ Path='4_Archive\Abbott_2005-2024\Compliance_Security_Governance';Keywords=@('compliance','sox','audit','pci','gxp','gdpr','hipaa','risk','rims','security','infosec','iso27001','control framework') }
  'Archive.Brand'      = @{ Path='4_Archive\Abbott_2005-2024\Brand_Marketing_Comms';        Keywords=@('brand','branding','marketing','campaign','social media','yammer','edetailing','comms','communication','evp','landor','logo','creative') }
  'Archive.Products'   = @{ Path='4_Archive\Abbott_2005-2024\Products_and_Divisions';        Keywords=@('informatics','add','gmo','diagnostics','device','navica','panbio','freestyle','libre','alinity','core it') }
  'Archive.People'     = @{ Path='4_Archive\Abbott_2005-2024\People_and_HR';                 Keywords=@('staff','hr','people','org chart','headcount','employee','tribunal','recruit','talent','town hall','minutes') }
  'Archive.Knowledge'  = @{ Path='4_Archive\Abbott_2005-2024\Presentations_and_Knowledge';   Keywords=@('training','research','presentation','deck','whitepaper','study','readiness','pex','notebook','notability','course') }
  'Archive.IT'         = @{ Path='4_Archive\Abbott_2005-2024\IT_Tools_and_Configs';          Keywords=@('config','cfg','cisco','sql','server','network','vpn','snagit','lotus notes','backup','install','log file') }
  'Archive.Email'      = @{ Path='4_Archive\Abbott_2005-2024\Email_Archive';                 Keywords=@('email','mailbox','pst','ost','outlook export','inbox','oneemail') }
  'Archive.Loose'      = @{ Path='4_Archive\Abbott_2005-2024\_Loose_Files';                  Keywords=@() }
  'Review'             = @{ Path='0_REVIEW_Unsorted';                                        Keywords=@() }
}
function CatPath([string]$key) { Join-Path $Root $Cat[$key].Path }

# ---- Extension hints used as a late classification pass --------------------
$ExtHint = @{
  '.mp3'='Resources.Music'; '.flac'='Resources.Music'; '.wav'='Resources.Music'; '.m4a'='Resources.Music';
  '.jpg'='Resources.Photos'; '.jpeg'='Resources.Photos'; '.png'='Resources.Photos'; '.heic'='Resources.Photos';
  '.gif'='Resources.Photos'; '.tiff'='Resources.Photos'; '.raw'='Resources.Photos'; '.cr2'='Resources.Photos';
  '.pst'='Archive.Email'; '.ost'='Archive.Email'; '.msg'='Archive.Email'; '.eml'='Archive.Email';
  '.potx'='Resources.Templates'; '.dotx'='Resources.Templates'; '.xltx'='Resources.Templates';
}

# ===========================================================================
#  STATIC MAP  — known Personal Files / Professional / root items.
#  (Same plan as before; edit Src/Dest/Rename freely.)
# ===========================================================================
$Map = @(
  # 1_Projects
  @{ Src='Personal Files\Electrical EV Install';            Dest=$Cat['Projects.EV'].Path }
  @{ Src='Personal Files\Mr Electric of Highland Park Transaction Receipt.pdf'; Dest=$Cat['Projects.EV'].Path }
  @{ Src='Personal Files\application N-400.pdf';            Dest=$Cat['Projects.N400'].Path }
  # 2_Areas\Career
  @{ Src='Personal Files\CV Resumes';                      Dest=$Cat['Areas.Career'].Path }
  @{ Src='Personal Files\personal deveopment';             Dest=$Cat['Areas.Career'].Path }
  @{ Src='Professional\General Interview';                  Dest=$Cat['Areas.Career'].Path }
  @{ Src='Professional\Goals';                             Dest=$Cat['Areas.Career'].Path }
  # 2_Areas\Finances_and_Tax
  @{ Src='Personal Files\Pay and tax';                     Dest=$Cat['Areas.Finance'].Path }
  @{ Src='Personal Files\White Holdings';                  Dest=$Cat['Areas.Finance'].Path }
  @{ Src='Professional\Finances';                          Dest=$Cat['Areas.Finance'].Path }
  @{ Src='US 1099 2025.pdf';                               Dest=$Cat['Areas.Finance'].Path }
  @{ Src='UK Property Tax 2025.pdf';                       Dest=$Cat['Areas.Finance'].Path }
  @{ Src='US Mortgage 1098.pdf';                           Dest=$Cat['Areas.Finance'].Path }
  @{ Src='Personal Files\Invoice 3786653.pdf';             Dest=$Cat['Areas.Finance'].Path; Rename='Invoice_3786653' }
  # 2_Areas\Health
  @{ Src='Personal Files\Health';                          Dest=$Cat['Areas.Health'].Path }
  # 2_Areas\Home_and_Property
  @{ Src='Personal Files\House';                           Dest=$Cat['Areas.Home'].Path }
  @{ Src='Personal Files\Land Registry';                   Dest=$Cat['Areas.Home'].Path }
  @{ Src='Personal Files\Control4';                        Dest=$Cat['Areas.Home'].Path }
  @{ Src='Personal Files\router';                          Dest=$Cat['Areas.Home'].Path }
  @{ Src='Personal Files\Quotation DP2213 Ensuite 12-08-2019 11-00.PDF';  Dest=$Cat['Areas.Home'].Path }
  @{ Src='Personal Files\Quotation DP2214 Bathroom 12-08-2019 11-00.PDF'; Dest=$Cat['Areas.Home'].Path }
  @{ Src='Personal Files\Swap over.docx';                  Dest=$Cat['Areas.Home'].Path; Rename='Swap_Over_REVIEW' }
  # 2_Areas\Vehicle
  @{ Src='Personal Files\Car';                             Dest=$Cat['Areas.Vehicle'].Path }
  # 2_Areas\Identity_and_Travel
  @{ Src='Personal Files\Passport and Visas';              Dest=$Cat['Areas.Identity'].Path }
  # 2_Areas\Family_and_Legal
  @{ Src='Personal Files\Divorce';                         Dest=$Cat['Areas.Family'].Path }
  @{ Src='Personal Files\150 Finstad';                     Dest=$Cat['Areas.Family'].Path }
  @{ Src='Personal Files\16 Wordsworth';                   Dest=$Cat['Areas.Family'].Path }
  @{ Src='Personal Files\Estate Accounts FINAL.pdf';       Dest=$Cat['Areas.Family'].Path }
  @{ Src='Personal Files\Forrester Hyde let and Portfolio.pdf'; Dest=$Cat['Areas.Family'].Path }
  # 2_Areas\Accounts_and_Licenses
  @{ Src='Personal Files\License Keys';                    Dest=$Cat['Areas.Accounts'].Path }
  @{ Src='Personal Files\Apps';                            Dest=$Cat['Areas.Accounts'].Path }
  # 2_Areas\Correspondence
  @{ Src='Letter to Tesco for Terence Finch.docx';         Dest=$Cat['Areas.Correspondence'].Path }
  @{ Src='Letter to Tesco for Terence Finch.pdf';          Dest=$Cat['Areas.Correspondence'].Path }
  # 3_Resources
  @{ Src='Personal Files\Custom Office Templates';         Dest=$Cat['Resources.Templates'].Path }
  @{ Src='Professional\Project Management Templates';      Dest=$Cat['Resources.Templates'].Path }
  @{ Src='Professional\Office 365';                        Dest=$Cat['Resources.Templates'].Path }
  @{ Src='Personal Files\Music';                           Dest=$Cat['Resources.Music'].Path }
  @{ Src='Personal Files\Pictures';                        Dest=$Cat['Resources.Photos'].Path }
  @{ Src='Personal Files\Terry Finch Photos';             Dest=$Cat['Resources.Photos'].Path }
  @{ Src='Personal Files\Wedding';                        Dest=$Cat['Resources.Photos'].Path }
  @{ Src='Personal Files\Useful';                          Dest=$Cat['Resources.Reference'].Path }
  @{ Src='Personal Files\Keep File';                       Dest=$Cat['Resources.Reference'].Path }
  @{ Src='Personal Files\Stuff';                           Dest=$Cat['Resources.Reference'].Path }
  @{ Src='Personal Files\Attachments';                     Dest=$Cat['Resources.Reference'].Path }
  @{ Src='Personal Files\Public';                          Dest=$Cat['Resources.Reference'].Path }
  # 4_Archive : Strategy
  @{ Src='Professional\IT Strategy';                       Dest=$Cat['Archive.Strategy'].Path }
  @{ Src='Professional\DIT Operating Model';               Dest=$Cat['Archive.Strategy'].Path }
  @{ Src='Professional\Transformation';                    Dest=$Cat['Archive.Strategy'].Path }
  @{ Src='Professional\Leadership Strategy';               Dest=$Cat['Archive.Strategy'].Path }
  @{ Src='Professional\Microsoft Strategy';                Dest=$Cat['Archive.Strategy'].Path }
  @{ Src='Professional\Procurement Strategy';              Dest=$Cat['Archive.Strategy'].Path }
  @{ Src='Professional\Reference Architecture';            Dest=$Cat['Archive.Strategy'].Path }
  @{ Src='Professional\Business Architecture';             Dest=$Cat['Archive.Strategy'].Path }
  @{ Src='Professional\MCOE Reboot';                       Dest=$Cat['Archive.Strategy'].Path }
  @{ Src='Professional\Organization Design';               Dest=$Cat['Archive.Strategy'].Path }
  @{ Src='Professional\Governance';                        Dest=$Cat['Archive.Strategy'].Path }
  @{ Src='Professional\100 day plans';                     Dest=$Cat['Archive.Strategy'].Path }
  @{ Src='Professional\infrastructure technical strategy'; Dest=$Cat['Archive.Strategy'].Path }
  # 4_Archive : Programs
  @{ Src='Professional\Application Rationalization';       Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\Project Go';                        Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\Project Spring';                    Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\Projects';                          Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\eCommerce';                         Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\Digital Heallth';                   Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\Big Data and Analytics';            Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\Workday';                           Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\Digital Summit & SDMC';             Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\War on Waste';                      Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\TCO Exercise 2016';                 Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\2010 Managed Services Deal';        Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\2014 BSS Savings';                  Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\2014 Task and Mckinsey';            Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\2015 ESP Finance';                  Dest=$Cat['Archive.Programs'].Path }
  @{ Src='Professional\Innovation & Hackathorn';           Dest=$Cat['Archive.Programs'].Path }
  # 4_Archive : M&A
  @{ Src='Professional\M and A Details';                   Dest=$Cat['Archive.MandA'].Path }
  @{ Src='Professional\Licensing Deals';                   Dest=$Cat['Archive.MandA'].Path }
  @{ Src='Professional\Alere';                             Dest=$Cat['Archive.MandA'].Path }
  @{ Src='Professional\Mylan';                             Dest=$Cat['Archive.MandA'].Path }
  @{ Src='Professional\Safe Harbor';                       Dest=$Cat['Archive.MandA'].Path }
  @{ Src='Professional\One Abbott Freight';                Dest=$Cat['Archive.MandA'].Path }
  # 4_Archive : Compliance
  @{ Src='Professional\123Compliance';                    Dest=$Cat['Archive.Compliance'].Path }
  @{ Src='Professional\Compliance Audit _ Sox';           Dest=$Cat['Archive.Compliance'].Path }
  @{ Src='Professional\Information Security';              Dest=$Cat['Archive.Compliance'].Path }
  @{ Src='Professional\PCI';                              Dest=$Cat['Archive.Compliance'].Path }
  @{ Src='Professional\Rims';                             Dest=$Cat['Archive.Compliance'].Path }
  # 4_Archive : Brand
  @{ Src='Professional\Abbott Corp Branding';             Dest=$Cat['Archive.Brand'].Path }
  @{ Src='Professional\Social Media';                     Dest=$Cat['Archive.Brand'].Path }
  @{ Src='Professional\Yammer';                           Dest=$Cat['Archive.Brand'].Path }
  @{ Src='Professional\eDetailing';                       Dest=$Cat['Archive.Brand'].Path }
  @{ Src='Professional\EVP Portal Presentation';          Dest=$Cat['Archive.Brand'].Path }
  @{ Src='Professional\Path To Purchase';                 Dest=$Cat['Archive.Brand'].Path }
  # 4_Archive : Products
  @{ Src='Professional\ADD Informatics';                  Dest=$Cat['Archive.Products'].Path }
  @{ Src='Professional\GMO';                              Dest=$Cat['Archive.Products'].Path }
  @{ Src='Professional\Core IT';                          Dest=$Cat['Archive.Products'].Path }
  # 4_Archive : People
  @{ Src='Professional\Berce Staff Meetings';             Dest=$Cat['Archive.People'].Path }
  @{ Src='Professional\staff';                            Dest=$Cat['Archive.People'].Path }
  @{ Src='Professional\Employment Tribunal';             Dest=$Cat['Archive.People'].Path }
  @{ Src='Personal Files\Abbott Employment Contracts';   Dest=$Cat['Archive.People'].Path }
  # 4_Archive : Knowledge
  @{ Src='Professional\Presentations';                    Dest=$Cat['Archive.Knowledge'].Path }
  @{ Src='Professional\Research';                         Dest=$Cat['Archive.Knowledge'].Path }
  @{ Src='Professional\Training';                         Dest=$Cat['Archive.Knowledge'].Path }
  @{ Src='Professional\pEX';                              Dest=$Cat['Archive.Knowledge'].Path }
  @{ Src='Professional\Operations Readiness';            Dest=$Cat['Archive.Knowledge'].Path }
  @{ Src='Professional\Notability';                      Dest=$Cat['Archive.Knowledge'].Path }
  @{ Src='Professional\Notebooks';                       Dest=$Cat['Archive.Knowledge'].Path }
  # 4_Archive : IT
  @{ Src='Professional\CiscoCfg';                        Dest=$Cat['Archive.IT'].Path }
  @{ Src='Professional\sql';                             Dest=$Cat['Archive.IT'].Path }
  @{ Src='Professional\slc';                             Dest=$Cat['Archive.IT'].Path }
  @{ Src='Professional\Snagit';                          Dest=$Cat['Archive.IT'].Path }
  @{ Src='Professional\IBM Lotus Notes ID';             Dest=$Cat['Archive.IT'].Path }
  @{ Src='Professional\Licence Keys';                    Dest=$Cat['Archive.IT'].Path }
  @{ Src='Professional\Favorites';                       Dest=$Cat['Archive.IT'].Path }
  # 4_Archive : Email
  @{ Src='Professional\OneEmail';                        Dest=$Cat['Archive.Email'].Path }
  # 4_Archive : Loose Abbott files
  @{ Src='Professional\Abbott At Home Testing Experience-Comms and Timeline 022721 v18[93].docx'; Dest=$Cat['Archive.Loose'].Path }
  @{ Src='Professional\APOC_2017_Pepper_Order_Form_21_20170721.docx';      Dest=$Cat['Archive.Loose'].Path }
  @{ Src='Professional\Aug Exec Program Review TBF and VC comments.pptx';  Dest=$Cat['Archive.Loose'].Path }
  @{ Src='Professional\health07_hi-res.wmv';                              Dest=$Cat['Archive.Loose'].Path }
  @{ Src='Professional\NAVICA - Keep it Simple Schools DAFTT for discusion.pptx'; Dest=$Cat['Archive.Loose'].Path }
  @{ Src='Professional\NAVICA - Keep it Simple Schools v1.pptx';          Dest=$Cat['Archive.Loose'].Path }
  @{ Src='Professional\Single Device Only At-Home Testing TF Update (No Visuals).pptx'; Dest=$Cat['Archive.Loose'].Path }
  @{ Src='Professional\Strategic focus and Op Model Value summary Finch Updates.docx';  Dest=$Cat['Archive.Loose'].Path }
  @{ Src='Abbott_GenAi_090825.pptx';                                      Dest=$Cat['Archive.Loose'].Path }
  @{ Src='Landor_AbbottMotion_jb_Pipette_Slide_v002.mp4';                 Dest=$Cat['Archive.Loose'].Path }
)

$Junk = @(
  'Personal Files\scan_windows_files.bat'
  'Microsoft Copilot Chat Files'
)

# ===========================================================================
#  ENGINE
# ===========================================================================
function New-DirIfNeeded([string]$path) {
  if (-not (Test-Path -LiteralPath $path)) { if ($Execute) { New-Item -ItemType Directory -Path $path -Force | Out-Null } }
}
function Resolve-Collision([string]$targetPath) {
  if (-not (Test-Path -LiteralPath $targetPath)) { return $targetPath }
  $dir  = Split-Path $targetPath -Parent
  $name = [IO.Path]::GetFileNameWithoutExtension($targetPath)
  $ext  = [IO.Path]::GetExtension($targetPath)
  $n = 2
  while (Test-Path -LiteralPath (Join-Path $dir ("{0} ({1}){2}" -f $name,$n,$ext))) { $n++ }
  return (Join-Path $dir ("{0} ({1}){2}" -f $name,$n,$ext))
}

# OneDrive Files On-Demand placeholder? Hashing these forces a cloud download.
#   Offline 0x1000 | RecallOnOpen 0x40000 | RecallOnDataAccess 0x400000
function Test-CloudOnly([System.IO.FileSystemInfo]$f) {
  $a = [int]$f.Attributes
  return (($a -band 0x1000) -ne 0) -or (($a -band 0x40000) -ne 0) -or (($a -band 0x400000) -ne 0)
}

$actions = [System.Collections.Generic.List[object]]::new()
$missing = [System.Collections.Generic.List[string]]::new()

function Add-Action($srcRel, $destDirRel, $rename, $kind) {
  $srcFull = Join-Path $Root $srcRel
  if (-not (Test-Path -LiteralPath $srcFull)) { $missing.Add($srcRel); return }
  $item  = Get-Item -LiteralPath $srcFull
  $isDir = $item.PSIsContainer
  $leaf  = Split-Path $srcRel -Leaf
  if ($rename -and -not $isDir) {
    $leaf = '{0}_{1}{2}' -f $item.LastWriteTime.ToString('yyyy_MM_dd'), $rename, ([IO.Path]::GetExtension($leaf))
  }
  $destDirFull = Join-Path $Root $destDirRel
  $finalFull   = Resolve-Collision (Join-Path $destDirFull $leaf)
  $actions.Add([pscustomobject]@{
    Kind=$kind; Type=($isDir ? 'Folder' : 'File'); Source=$srcRel
    Destination=(Join-Path $destDirRel (Split-Path $finalFull -Leaf))
    SourceFull=$srcFull; DestDirFull=$destDirFull; TargetFull=$finalFull
  })
}

# ---- Multi-pass classifier for the Extracts tree --------------------------
function Score-Text([string]$name, [string]$blob) {
  $best='Review'; $bestScore=0; $matched=@()
  foreach ($k in $Cat.Keys) {
    $kw = $Cat[$k].Keywords; if (-not $kw) { continue }
    $score=0; $hits=@()
    foreach ($w in $kw) {
      $p=[regex]::Escape($w.ToLower())
      if ($name -match $p) { $score+=2; $hits+=$w }     # name match weighted x2
      elseif ($blob -match $p) { $score+=1; $hits+=$w }
    }
    if ($score -gt $bestScore) { $bestScore=$score; $best=$k; $matched=$hits }
  }
  [pscustomobject]@{ Key=$best; Score=$bestScore; Matched=($matched | Select-Object -Unique) }
}
function Invoke-MultiPassClassify([System.IO.FileSystemInfo]$item) {
  $name = $item.Name.ToLower()
  # Pass 1 — name only
  $r = Score-Text $name ''
  $pass = 1
  if ($r.Score -eq 0 -and $item.PSIsContainer) {
    # Pass 2 — immediate children names
    $kids = (Get-ChildItem -LiteralPath $item.FullName -Force -ErrorAction SilentlyContinue | Select-Object -First 100).Name -join ' '
    $r = Score-Text $name $kids.ToLower(); $pass = 2
    if ($r.Score -eq 0) {
      # Pass 3 — deep file names (capped)
      $deep = (Get-ChildItem -LiteralPath $item.FullName -Force -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 300).Name -join ' '
      $r = Score-Text $name $deep.ToLower(); $pass = 3
    }
  }
  # Pass 4 — extension hint (files, or folder dominated by one media type)
  if ($r.Score -eq 0) {
    $exts = if ($item.PSIsContainer) {
      (Get-ChildItem -LiteralPath $item.FullName -Force -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 300).Extension
    } else { @($item.Extension) }
    $hintKey = $exts | Where-Object { $_ } | ForEach-Object { $ExtHint[$_.ToLower()] } |
               Where-Object { $_ } | Group-Object | Sort-Object Count -Descending | Select-Object -First 1
    if ($hintKey) { $r = [pscustomobject]@{ Key=$hintKey.Name; Score=2; Matched=@('[by file type]') }; $pass=4 }
  }
  $conf = $r.Score -ge $ConfidenceThreshold ? 'High' : ($r.Score -ge 1 ? 'Medium' : 'None')
  [pscustomobject]@{ Key=$r.Key; Score=$r.Score; Confidence=$conf; Pass=$pass; Matched=$r.Matched }
}

function Show-Menu {
  $keys = @($Cat.Keys)
  for ($i=0; $i -lt $keys.Count; $i++) { Write-Host ("   {0,2}) {1}" -f ($i+1), $Cat[$keys[$i]].Path) }
  $keys
}

# ---- Build static actions --------------------------------------------------
foreach ($e in $Map)  { Add-Action $e.Src $e.Dest $e.Rename 'MOVE' }
foreach ($j in $Junk) { Add-Action $j '_TO_DELETE' $null 'JUNK' }

# ---- Review & classify Extracts -------------------------------------------
$extractsFull = Join-Path $Root $ExtractsFolder
if (Test-Path -LiteralPath $extractsFull) {
  $items = Get-ChildItem -LiteralPath $extractsFull -Force
  Write-Host "`nReviewing '$ExtractsFolder' — $($items.Count) top-level items (multi-pass)..." -ForegroundColor Cyan
  $idx = 0
  foreach ($it in $items) {
    $idx++
    Write-Progress -Activity "Classifying $ExtractsFolder" -Status $it.Name -PercentComplete (($idx/$items.Count)*100)
    $c = Invoke-MultiPassClassify $it
    $proposedKey = $c.Key
    if ($canPrompt -and $c.Confidence -ne 'High') {
      Write-Host ""
      $itType = $it.PSIsContainer ? 'Folder' : 'File'
      Write-Host ("[{0}/{1}] {2}  ""{3}""" -f $idx, $items.Count, $itType, $it.Name) -ForegroundColor White
      Write-Host ("   Proposed: {0}" -f $Cat[$proposedKey].Path) -ForegroundColor Yellow
      $matchedTxt = ($c.Matched.Count) ? ($c.Matched -join ', ') : 'none'
      Write-Host ("   Confidence: {0}  (pass {1}; matched: {2})" -f $c.Confidence,$c.Pass,$matchedTxt)
      $decided = $false
      while (-not $decided) {
        $ans = (Read-Host "   [A]ccept  [C]hange  [V]iew  [S]kip  [Q]uit").ToUpper()
        if     ($ans -eq 'A') { $decided = $true }
        elseif ($ans -eq 'S') { $proposedKey = $null; $decided = $true }
        elseif ($ans -eq 'Q') { Write-Host 'Aborted by user — nothing changed.' -ForegroundColor Red; return }
        elseif ($ans -eq 'V') {
          Get-ChildItem -LiteralPath $it.FullName -Force -ErrorAction SilentlyContinue |
            Select-Object -First 20 | ForEach-Object { Write-Host "      - $($_.Name)" }
        }
        elseif ($ans -eq 'C') {
          $keys = Show-Menu
          $pick = Read-Host "   Enter number (1-$($keys.Count))"
          if (($pick -as [int]) -and [int]$pick -ge 1 -and [int]$pick -le $keys.Count) { $proposedKey = $keys[[int]$pick-1]; $decided = $true }
          else { Write-Host '   Invalid selection.' -ForegroundColor Red }
        }
        else { Write-Host '   Please choose A, C, V, S, or Q.' -ForegroundColor Red }
      }
    } elseif ($c.Confidence -eq 'None') {
      $proposedKey = 'Review'   # non-interactive / unattended fallback
    }
    if ($proposedKey) {
      $relSrc = Join-Path $ExtractsFolder $it.Name
      Add-Action $relSrc $Cat[$proposedKey].Path $null 'EXTRACT'
    }
  }
  Write-Progress -Activity "Classifying $ExtractsFolder" -Completed
} else {
  Write-Host "`n(No '$ExtractsFolder' folder found at root — skipping that pass.)" -ForegroundColor DarkGray
}

# ===========================================================================
#  REPORT
# ===========================================================================
$mode = $Execute ? 'EXECUTE (changes WILL be made)' : 'DRY RUN (no changes)'
Write-Host "`nOneDrive PARA reorganization — $mode" -ForegroundColor Cyan
Write-Host "Root: $Root"
Write-Host ("Planned actions: {0}   (static: {1}, extracts: {2}, junk: {3})   Missing: {4}" -f `
  $actions.Count, @($actions | Where-Object Kind -eq 'MOVE').Count, @($actions | Where-Object Kind -eq 'EXTRACT').Count, @($actions | Where-Object Kind -eq 'JUNK').Count, $missing.Count)
Write-Host ""
$actions | Sort-Object Destination, Source |
  Format-Table Kind, Type, Source, Destination -AutoSize | Out-String -Width 4096 | Write-Host
if ($missing.Count) {
  Write-Host "Sources not found (skipped):" -ForegroundColor Yellow
  $missing | ForEach-Object { Write-Host "   $_" -ForegroundColor Yellow }
}

# ---- Duplicate finder ------------------------------------------------------
if ($FindDuplicates) {
  Write-Host "`nScanning for duplicates (filename + size, confirmed by SHA-256)..." -ForegroundColor Cyan
  $allFiles = Get-ChildItem -LiteralPath $Root -Force -Recurse -File -ErrorAction SilentlyContinue |
              Where-Object { $_.FullName -notmatch '\\_Reorg_Logs\\' }
  # OneDrive Files On-Demand: online-only placeholders can't be hashed without a
  # cloud download (which may fail). Skip + report them unless -HydrateCloudFiles.
  if (-not $HydrateCloudFiles) {
    $cloudOnly = @($allFiles | Where-Object { Test-CloudOnly $_ })
    if ($cloudOnly.Count) {
      Write-Host ("Skipping {0} online-only (cloud) file(s) — not downloaded. Use -HydrateCloudFiles to include them." -f $cloudOnly.Count) -ForegroundColor DarkYellow
      $allFiles = $allFiles | Where-Object { -not (Test-CloudOnly $_) }
    }
  }
  # cheap pre-filter: same Name AND same Length
  $candidates = $allFiles | Group-Object Name, Length | Where-Object Count -gt 1 |
                ForEach-Object { $_.Group }
  $dupSets = @()
  $hashErrors = [System.Collections.Generic.List[string]]::new()
  if ($candidates) {
    $hashed = $candidates | ForEach-Object -ThrottleLimit 8 -Parallel {
      try {
        [pscustomobject]@{
          Path=$_.FullName; Name=$_.Name; Length=$_.Length; Modified=$_.LastWriteTime
          Hash=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256 -ErrorAction Stop).Hash; Error=$null
        }
      } catch {
        [pscustomobject]@{
          Path=$_.FullName; Name=$_.Name; Length=$_.Length; Modified=$_.LastWriteTime
          Hash=$null; Error=$_.Exception.Message
        }
      }
    }
    foreach ($h in @($hashed | Where-Object Error)) { $hashErrors.Add($h.Path) }
    $dupSets = @($hashed | Where-Object Hash | Group-Object Hash | Where-Object Count -gt 1)
  }
  if ($hashErrors.Count) {
    Write-Host ("{0} file(s) could not be hashed and were skipped:" -f $hashErrors.Count) -ForegroundColor DarkYellow
    $hashErrors | Select-Object -First 15 | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkYellow }
    if ($hashErrors.Count -gt 15) { Write-Host ("   ... and {0} more" -f ($hashErrors.Count - 15)) -ForegroundColor DarkYellow }
  }
  if ($dupSets.Count) {
    Write-Host ("Found {0} duplicate set(s):" -f $dupSets.Count) -ForegroundColor Magenta
    $rows = foreach ($s in $dupSets) {
      $keep = $s.Group | Sort-Object Modified | Select-Object -First 1  # keep oldest original
      foreach ($f in $s.Group) {
        [pscustomobject]@{ Hash=$s.Name.Substring(0,12); Keep=($f.Path -eq $keep.Path); Modified=$f.Modified; Size=$f.Length; Path=$f.Path }
      }
    }
    $rows | Sort-Object Hash, Modified | Format-Table Hash, Keep, Modified, Size, Path -AutoSize | Out-String -Width 4096 | Write-Host
    if ($Execute) { New-DirIfNeeded $logDir; $rows | Export-Csv -LiteralPath $dupCsv -NoTypeInformation -Encoding UTF8; Write-Host "Duplicate report: $dupCsv" }
    $stageDups = $false
    if ($Execute) {
      $stageDups = $canPrompt ? ((Read-Host "`nMove duplicate copies (keeping oldest) into _TO_DELETE\_Duplicates? [y/N]").ToUpper() -eq 'Y') : $false
    }
    if ($stageDups) {
      foreach ($r in ($rows | Where-Object { -not $_.Keep })) {
        $destDir = Join-Path $Root '_TO_DELETE\_Duplicates'
        New-DirIfNeeded $destDir
        $target = Resolve-Collision (Join-Path $destDir (Split-Path $r.Path -Leaf))
        try { Move-Item -LiteralPath $r.Path -Destination $target; Write-Host "[DUP] staged $($r.Path)" -ForegroundColor DarkYellow }
        catch { Write-Host "[DUP-FAIL] $($r.Path) ($($_.Exception.Message))" -ForegroundColor Red }
      }
      Write-Host "Duplicates staged in _TO_DELETE\_Duplicates — nothing deleted." -ForegroundColor Yellow
    }
  } else { Write-Host "No content-confirmed duplicates found." -ForegroundColor Green }
}

# ---- Unmapped top-level items (safety net) --------------------------------
$plannedTop = @{}
foreach ($a in $actions) {
  $parts = $a.Source -split '\\'
  $top = ($parts.Count -ge 2) ? "$($parts[0])\$($parts[1])" : $parts[0]
  $plannedTop[$top] = $true
}
$unmapped = [System.Collections.Generic.List[string]]::new()
foreach ($container in @('.', 'Personal Files', 'Professional', $ExtractsFolder)) {
  $base = ($container -eq '.') ? $Root : (Join-Path $Root $container)
  if (-not (Test-Path -LiteralPath $base)) { continue }
  Get-ChildItem -LiteralPath $base -Force | ForEach-Object {
    if ($_.Name -match '^(1_Projects|2_Areas|3_Resources|4_Archive|0_REVIEW_Unsorted|_TO_DELETE|_Reorg_Logs)$') { return }
    $rel = ($container -eq '.') ? $_.Name : "$container\$($_.Name)"
    if (-not $plannedTop.ContainsKey($rel)) { $unmapped.Add($rel) }
  }
}
if ($unmapped.Count) {
  Write-Host "`nNOT in plan — left in place for you to review:" -ForegroundColor Magenta
  $unmapped | Sort-Object -Unique | ForEach-Object { Write-Host "   $_" -ForegroundColor Magenta }
}

# ===========================================================================
#  EXECUTE
# ===========================================================================
if (-not $Execute) {
  Write-Host "`nDRY RUN complete. Re-run with -Execute to apply (add -FindDuplicates to also dedupe)." -ForegroundColor Green
  return
}
if ($canPrompt) {
  $go = Read-Host "`nApply $($actions.Count) moves now? [y/N]"
  if ($go.ToUpper() -ne 'Y') { Write-Host 'Cancelled — nothing changed.' -ForegroundColor Yellow; return }
}

New-DirIfNeeded $logDir
$actions | Export-Csv -LiteralPath $planCsv -NoTypeInformation -Encoding UTF8
$undo = [System.Collections.Generic.List[string]]::new()
$undo.Add('# Auto-generated undo script — reverses this batch.')
$undo.Add('$ErrorActionPreference = ''Stop''')

$done=0; $failed=0; $n=0
foreach ($a in ($actions | Sort-Object Destination)) {
  $n++
  Write-Progress -Activity 'Moving' -Status $a.Source -PercentComplete (($n/$actions.Count)*100)
  try {
    New-DirIfNeeded $a.DestDirFull
    Move-Item -LiteralPath $a.SourceFull -Destination $a.TargetFull
    $undo.Add(('Move-Item -LiteralPath "{0}" -Destination "{1}"' -f $a.TargetFull, $a.SourceFull))
    $done++; Write-Host ("[OK]   {0} -> {1}" -f $a.Source, $a.Destination) -ForegroundColor Green
  } catch {
    $failed++; Write-Host ("[FAIL] {0} ({1})" -f $a.Source, $_.Exception.Message) -ForegroundColor Red
  }
}
Write-Progress -Activity 'Moving' -Completed
Set-Content -LiteralPath $undoPs1 -Value $undo -Encoding UTF8

Write-Host "`n===== SUMMARY =====" -ForegroundColor Cyan
Write-Host ("Moved: {0}   Failed: {1}   (of {2} planned)" -f $done,$failed,$actions.Count)
Write-Host "Plan log : $planCsv"
Write-Host "Undo     : $undoPs1   (run to reverse)"
Write-Host "_TO_DELETE holds suspected junk/duplicates — nothing deleted. Review then remove." -ForegroundColor Yellow
