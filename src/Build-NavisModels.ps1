<#
.SYNOPSIS
    Navisworks Batch Builder v2.0
    
.DESCRIPTION
    1. Reads configuration from settings.json.
    2. Copies template to local workspace.
    3. Processes folders based on "All" or "Selected" mode.
    4. Assembles Master Model.
    5. Transfers files back.

.NOTES
    Version: 2.0
#>

# ========================================================
# 1. CONFIGURATION
# ========================================================

# Determine Project Root (Go up one level from src/ folder)
$ProjectRoot = (Get-Item $PSScriptRoot).Parent.FullName
$ConfigFile  = Join-Path $ProjectRoot "settings.json"

Clear-Host
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host " NAVISWORKS BATCH AUTOMATION v2.0" -ForegroundColor Cyan
Write-Host "========================================================"

# --- Load JSON ---
if (Test-Path $ConfigFile) {
    try {
        $Config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
        Write-Host " [OK] Configuration loaded." -ForegroundColor Green
    }
    catch {
        Write-Host " [ERROR] JSON Syntax Error!" -ForegroundColor Red; Pause; Exit
    }
} else {
    Write-Host " [ERROR] settings.json not found in: $ProjectRoot" -ForegroundColor Red
    Pause; Exit
}

# Set variables from Config
$MasterName   = $Config.MasterModelName
$TemplateName = $Config.TemplateName
$LocalWorkDir = "C:\Temp\NavisWork_BATCH"
$Recursive    = $Config.RecursiveSearch

Write-Host "Master Model: $MasterName"
Write-Host "Mode: $($Config.Mode) | Recursive: $Recursive"
Write-Host ""

# ========================================================
# 2. INITIALIZATION
# ========================================================

# --- Find Navisworks ---
Write-Host "Searching for Navisworks..." -NoNewline
$NavisPath = Get-ChildItem -Path "C:\Program Files\Autodesk" -Filter "FileToolsTaskRunner.exe" -Recurse -ErrorAction SilentlyContinue | 
             Sort-Object FullName -Descending | 
             Select-Object -First 1

if ($NavisPath) {
    Write-Host " [OK]" -ForegroundColor Green
} else {
    Write-Host " [ERROR] Navisworks FileToolsTaskRunner not found!" -ForegroundColor Red; Pause; Exit
}

# --- Prepare Local Workspace ---
if (Test-Path $LocalWorkDir) { Remove-Item $LocalWorkDir -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $LocalWorkDir -Force | Out-Null

# --- Copy Template ---
$TemplatePath = Join-Path $ProjectRoot $TemplateName
if (Test-Path $TemplatePath) {
    Copy-Item -Path $TemplatePath -Destination $LocalWorkDir -Force
    Write-Host "Template copied to workspace." -ForegroundColor Green
} else {
    Write-Host "Template '$TemplateName' not found. Proceeding without it." -ForegroundColor Yellow
}

# ========================================================
# 3. SELECT FOLDERS
# ========================================================

Write-Host ""
Write-Host "--- Selecting Folders ---" -ForegroundColor Cyan

$FoldersToProcess = @()

if ($Config.Mode -eq "All") {
    # Scan project root for folders, excluding system folders
    $FoldersToProcess = Get-ChildItem -Path $ProjectRoot -Directory | 
                        Where-Object { $_.Name -ne "Temp" -and $_.Name -ne "src" -and $_.Name -ne ".git" }
}
elseif ($Config.Mode -eq "Selected") {
    # Use the list defined in settings.json
    foreach ($FolderName in $Config.SelectedFolders) {
        $FullPath = Join-Path $ProjectRoot $FolderName
        if (Test-Path $FullPath) { $FoldersToProcess += (Get-Item $FullPath) }
        else { Write-Host " [WARNING] Configured folder not found: $FolderName" -ForegroundColor Yellow }
    }
}

# ========================================================
# 4. PROCESSING LOOP
# ========================================================

foreach ($Folder in $FoldersToProcess) {
    $FolderName = $Folder.Name
    $FileListPath = Join-Path $LocalWorkDir "temp_filelist.txt"
    $OutputNwd    = Join-Path $LocalWorkDir "$FolderName.nwd"
    
    # Check recursion setting from JSON
    if ($Recursive -eq $true) {
        $Files = Get-ChildItem -Path $Folder.FullName -Include $Config.FileExtensions -Recurse -File | 
                 Where-Object { $_.Name -ne "$FolderName.nwd" }
    }
    else {
        $SearchPath = Join-Path $Folder.FullName "*"
        $Files = Get-ChildItem -Path $SearchPath -Include $Config.FileExtensions -File | 
                 Where-Object { $_.Name -ne "$FolderName.nwd" }
    }

    if ($Files) {
        Write-Host "Processing: $FolderName ($($Files.Count) files)" -NoNewline
        
        $Files.FullName | Set-Content -Path $FileListPath -Encoding UTF8
        
        $ArgsList = "/i `"$FileListPath`" /of `"$OutputNwd`""
        $Process = Start-Process -FilePath $NavisPath.FullName -ArgumentList $ArgsList -Wait -NoNewWindow -PassThru

        if ($Process.ExitCode -eq 0) { Write-Host " -> Done" -ForegroundColor Green }
        else { Write-Host " -> FAILED" -ForegroundColor Red }
    }
}

# ========================================================
# 5. ASSEMBLE MASTER MODEL
# ========================================================

Write-Host ""
Write-Host "--- Assembling Master Model ---" -ForegroundColor Cyan

$MasterListFile = Join-Path $LocalWorkDir "master_list.txt"
$MasterOutput   = Join-Path $LocalWorkDir $MasterName
$MasterFiles    = @()

# 1. Add Template
if (Test-Path (Join-Path $LocalWorkDir $TemplateName)) {
    $MasterFiles += (Join-Path $LocalWorkDir $TemplateName)
}

# 2. Add Sub-models
$SubModels = Get-ChildItem -Path $LocalWorkDir -Filter "*.nwd" | 
             Where-Object { $_.Name -ne $TemplateName -and $_.Name -ne $MasterName }

foreach ($m in $SubModels) { $MasterFiles += $m.FullName }

if ($MasterFiles.Count -gt 0) {
    $MasterFiles | Set-Content -Path $MasterListFile -Encoding UTF8
    
    Write-Host "Creating Master Model..." -NoNewline
    $ArgsList = "/i `"$MasterListFile`" /of `"$MasterOutput`""
    Start-Process -FilePath $NavisPath.FullName -ArgumentList $ArgsList -Wait -NoNewWindow
    
    if (Test-Path $MasterOutput) { Write-Host " [OK]" -ForegroundColor Green }
    else { Write-Host " [ERROR]" -ForegroundColor Red }
} else {
    Write-Host "No files to assemble." -ForegroundColor Yellow
}

# ========================================================
# 6. TRANSFER FILES
# ========================================================

Write-Host ""
Write-Host "--- Transferring Files ---" -ForegroundColor Cyan

# 1. Move Sub-models back
foreach ($Folder in $FoldersToProcess) {
    $NwdName  = "$($Folder.Name).nwd"
    $LocalNwd = Join-Path $LocalWorkDir $NwdName
    $DestNwd  = Join-Path $Folder.FullName $NwdName

    if (Test-Path $LocalNwd) {
        try {
            Move-Item -Path $LocalNwd -Destination $DestNwd -Force -ErrorAction Stop
            Write-Host "Updated: $NwdName" -ForegroundColor Gray
        } catch { Write-Host "LOCKED: $NwdName (File in use)" -ForegroundColor Red }
    }
}

# 2. Move Master Model
if (Test-Path $MasterOutput) {
    try {
        Move-Item -Path $MasterOutput -Destination (Join-Path $ProjectRoot $MasterName) -Force -ErrorAction Stop
        Write-Host "Updated Master: $MasterName" -ForegroundColor Green
    } catch { Write-Host "LOCKED: Master Model (File in use)" -ForegroundColor Red }
}

# 3. Cleanup
if (Test-Path $LocalWorkDir) { Remove-Item $LocalWorkDir -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ""
Write-Host "Job Completed." -ForegroundColor Cyan
Pause
