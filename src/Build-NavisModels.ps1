<#
.SYNOPSIS
    Navisworks Batch Builder - Automates NWD creation and Master Model assembly.
    
.DESCRIPTION
    1. Copies template and resources to a local temporary workspace (SSD speed).
    2. Iterates through sub-folders, converts CAD files to NWD.
    3. Assembles a Master Model from the created sub-models.
    4. Transfers updated files back to the server/project directory.

.NOTES
    Version: 1.0
#>

# ========================================================
# 1. CONFIGURATION
# ========================================================

$MasterName   = "MASTER_MODEL.nwd"
$TemplateName = "_TEMPLATE.nwd"
$LocalWorkDir = "C:\Temp\NavisWork_BATCH"

# Get the script execution directory (Project Root)
# If running from a separate src folder, we might need to go up one level. 
# Here assuming script is run with project root context or files are relative.
$ProjectRoot = Get-Location

# ========================================================
# 2. INITIALIZATION & CHECKS
# ========================================================

Clear-Host
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host " NAVISWORKS BATCH AUTOMATION" -ForegroundColor Cyan
Write-Host "========================================================"
Write-Host "Master Model: $MasterName"
Write-Host "Project Root: $ProjectRoot"
Write-Host ""

# --- Find Navisworks (FileToolsTaskRunner.exe) ---
Write-Host "Searching for Navisworks..." -NoNewline
$NavisPath = Get-ChildItem -Path "C:\Program Files\Autodesk" -Filter "FileToolsTaskRunner.exe" -Recurse -ErrorAction SilentlyContinue | 
             Sort-Object FullName -Descending | 
             Select-Object -First 1

if ($NavisPath) {
    Write-Host " [OK]" -ForegroundColor Green
    Write-Host "Found: $($NavisPath.FullName)" -ForegroundColor DarkGray
} else {
    Write-Host " [ERROR]" -ForegroundColor Red
    Write-Host "Navisworks FileToolsTaskRunner not found! Please install Navisworks Manage or Simulate."
    Pause
    Exit
}

# --- Prepare Local Workspace ---
# We use a local temp folder to avoid network lag and file locking issues during processing
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
# 3. PROCESSING SUB-FOLDERS
# ========================================================

Write-Host ""
Write-Host "--- Processing Sub-folders ---" -ForegroundColor Cyan

# Get all sub-directories in the project root
$SubFolders = Get-ChildItem -Path $ProjectRoot -Directory

foreach ($Folder in $SubFolders) {
    $FolderName = $Folder.Name
    
    # Skip the Temp folder and hidden folders (like .git)
    if ($FolderName -eq "Temp" -or $FolderName.StartsWith(".")) { continue }

    $FileListPath = Join-Path $LocalWorkDir "temp_filelist.txt"
    $OutputNwd    = Join-Path $LocalWorkDir "$FolderName.nwd"
    $SkipNwdName  = "$FolderName.nwd"

    # Find source files (DWG, IFC, DGN, NWD)
    # Exclude the NWD we are about to create to avoid recursive loops
    $Files = Get-ChildItem -Path $Folder.FullName -Include *.dwg, *.ifc, *.dgn, *.nwd -Recurse | 
             Where-Object { $_.Name -ne $SkipNwdName }

    if ($Files) {
        Write-Host "Processing: $FolderName" -NoNewline
        
        # Create input list for Navisworks
        $Files.FullName | Set-Content -Path $FileListPath -Encoding ASCII

        # Run Navisworks Task Runner
        # Syntax: /i "input_list.txt" /of "output.nwd"
        $ArgsList = "/i `"$FileListPath`" /of `"$OutputNwd`""
        
        $Process = Start-Process -FilePath $NavisPath.FullName -ArgumentList $ArgsList -Wait -NoNewWindow -PassThru

        if ($Process.ExitCode -eq 0 -and (Test-Path $OutputNwd)) {
            Write-Host " -> Done" -ForegroundColor Green
        } else {
            Write-Host " -> FAILED (Code: $($Process.ExitCode))" -ForegroundColor Red
        }
    }
}

# ========================================================
# 4. ASSEMBLING MASTER MODEL
# ========================================================

Write-Host ""
Write-Host "--- Assembling Master Model ---" -ForegroundColor Cyan

$MasterListFile = Join-Path $LocalWorkDir "master_list.txt"
$MasterOutput   = Join-Path $LocalWorkDir $MasterName

# Create assembly list
$MasterFiles = @()

# 1. Add Template FIRST (controls settings/coordinates)
if (Test-Path (Join-Path $LocalWorkDir $TemplateName)) {
    $MasterFiles += (Join-Path $LocalWorkDir $TemplateName)
}

# 2. Add Sub-models (excluding template and master itself)
$SubModels = Get-ChildItem -Path $LocalWorkDir -Filter "*.nwd" | 
             Where-Object { $_.Name -ne $TemplateName -and $_.Name -ne $MasterName }

foreach ($Model in $SubModels) {
    $MasterFiles += $Model.FullName
}

if ($MasterFiles.Count -gt 0) {
    $MasterFiles | Set-Content -Path $MasterListFile -Encoding ASCII
    
    Write-Host "Creating Master Model..." -NoNewline
    $ArgsList = "/i `"$MasterListFile`" /of `"$MasterOutput`""
    Start-Process -FilePath $NavisPath.FullName -ArgumentList $ArgsList -Wait -NoNewWindow
    
    if (Test-Path $MasterOutput) {
        Write-Host " [OK]" -ForegroundColor Green
    } else {
        Write-Host " [ERROR] Failed to create Master." -ForegroundColor Red
    }
} else {
    Write-Host "No files to assemble." -ForegroundColor Yellow
}

# ========================================================
# 5. TRANSFERRING FILES
# ========================================================

Write-Host ""
Write-Host "--- Transferring Files ---" -ForegroundColor Cyan

# 1. Move Sub-models back to their respective folders
foreach ($Folder in $SubFolders) {
    $NwdName = "$($Folder.Name).nwd"
    $LocalNwd = Join-Path $LocalWorkDir $NwdName
    $DestNwd  = Join-Path $Folder.FullName $NwdName

    if (Test-Path $LocalNwd) {
        try {
            Move-Item -Path $LocalNwd -Destination $DestNwd -Force -ErrorAction Stop
            Write-Host "Updated: $NwdName" -ForegroundColor Gray
        }
        catch {
            Write-Host "ERROR updating $NwdName. File might be in use." -ForegroundColor Red
        }
    }
}

# 2. Move Master Model to Project Root
if (Test-Path $MasterOutput) {
    try {
        Move-Item -Path $MasterOutput -Destination (Join-Path $ProjectRoot $MasterName) -Force -ErrorAction Stop
        Write-Host "Updated Master: $MasterName" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR updating Master Model. File might be in use." -ForegroundColor Red
    }
}

# 3. Cleanup
Write-Host ""
Write-Host "Cleaning up temp files..." -NoNewline
if (Test-Path $LocalWorkDir) { Remove-Item $LocalWorkDir -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host " Done."

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host " JOB COMPLETED" -ForegroundColor Cyan
Write-Host "========================================================"

# Keep window open
Pause