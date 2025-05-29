<#
.SYNOPSIS
    Scans for orphaned Hyper-V VM files, attempts to identify their previous owner,
    and offers to delete them.
.DESCRIPTION
    This script identifies Hyper-V virtual machine configuration files (.vmcx, .vmrs)
    and virtual hard disk/ISO files (.vhd, .vhdx, .avhd, .avhdx, .iso) that are not
    currently associated with any registered virtual machine on the host.

    If scan paths are not provided as parameters, the script will interactively prompt for them.

    HOW IT WORKS:
    1. GATHERS ACTIVE FILES: It first queries Hyper-V for all registered VMs and their
       snapshots. To identify active VMCX/VMRS files, it constructs paths by combining
       your provided VMConfigScanPath with the VM's ID, and your SnapshotConfigScanPath
       with the Snapshot's ID. VHDs and ISOs are taken directly from Get-VMHardDiskDrive
       and Get-VMDvdDrive.
    2. SCANS DISK LOCATIONS: You must provide the scan paths (either via parameters or
       interactive prompts). The script will scan these locations recursively for
       VM-related file extensions.
    3. COMPARES & IDENTIFIES ORPHANS: It compares the list of files found on disk with
       the list of actively used files. Any file on disk not in the active list is
       considered a potential orphan.
    4. ATTEMPTS TO IDENTIFY PREVIOUS OWNER: For orphaned files, it tries to guess the
       original VM name or ID based on filename or parent directory names and checks
       if a VM with that identifier is currently registered.
    5. PRESENTS LIST & PROMPTS (if not -DryRun): It shows a list of potential orphans
       and, if -DryRun is $false, prompts for deletion with Yes/No/All/Quit options.

    This script supports -WhatIf and -Confirm common PowerShell parameters.

    WARNING: Deleting files can lead to data loss. ALWAYS use -DryRun first,
             carefully review the output, and ensure you have backups before
             running with -DryRun:$false.
.PARAMETER VMConfigScanPath
    [Optional] The full, direct path to the directory where main virtual machine
    configuration files (.vmcx, .vmrs) are stored and should be scanned.
    If not provided, the script will prompt for this path.
    Example: "H:\Virtual Machine Configurations\Virtual Machines"
.PARAMETER SnapshotConfigScanPath
    [Optional] The full, direct path to the directory where snapshot configuration
    files (.vmcx, .vmrs for snapshots) are stored and should be scanned.
    If not provided, the script will prompt for this path.
    Example: "H:\Virtual Machine Configurations\Snapshots"
.PARAMETER VHDScanPaths
    [Optional] An array of full, direct paths to directories that should be scanned
    for virtual hard disk files (.vhd, .vhdx, .avhd, .avhdx).
    If not provided, the script will prompt interactively to add VHD scan paths.
    Example: @("H:\Virtual Hard Disks", "D:\ExtraVHDStorage")
.PARAMETER DryRun
    If $true (the default), the script will only list orphaned files and will NOT prompt for deletion.
    Set to $false to enable deletion prompts. USE WITH EXTREME CAUTION.
.PARAMETER IncludeISOs
    If $true, the script will also scan for .iso files in the specified scan paths.
    ISOs are often shared or kept for reinstallation, so be cautious. Default is false.
.PARAMETER Help
    Displays this help message and exits.
.INPUTS
    None. This script does not accept pipeline input for its main parameters.
.OUTPUTS
    System.Management.Automation.PSCustomObject
    The script outputs custom objects to the console, formatted as a table,
    listing potential orphaned files with details.
    It also outputs informational messages, warnings, and errors to the console streams.
.EXAMPLE
    # Run the script; it will prompt for necessary scan paths if not provided.
    .\Find-OrphanedVMFilesWithHistory.ps1 -Verbose

.EXAMPLE
    # Specify all scan paths as parameters to bypass interactive prompts.
    .\Find-OrphanedVMFilesWithHistory.ps1 -VMConfigScanPath "H:\Virtual Machine Configurations\Virtual Machines" `
                                       -SnapshotConfigScanPath "H:\Virtual Machine Configurations\Snapshots" `
                                       -VHDScanPaths "H:\Virtual Hard Disks" `
                                       -Verbose

.EXAMPLE
    # Enable deletion (EXTREME CAUTION! ENSURE BACKUPS!)
    .\Find-OrphanedVMFilesWithHistory.ps1 -VMConfigScanPath "H:\Configs" -SnapshotConfigScanPath "H:\Snapshots" -VHDScanPaths "H:\VHDs" -DryRun:$false -Verbose
.NOTES
    Author: rescrack
    Version: 1.9.6 (Final polish)
    Requires Administrator privileges.
#>
[CmdletBinding(SupportsShouldProcess = $true, HelpUri = "https://example.com/Find-OrphanedVMFilesWithHistory-Help")]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Full path to scan for main VM config files. Will prompt if omitted.")]
    [string]$VMConfigScanPath,

    [Parameter(Mandatory = $false, HelpMessage = "Full path to scan for snapshot config files. Will prompt if omitted.")]
    [string]$SnapshotConfigScanPath,

    [Parameter(Mandatory = $false, HelpMessage = "Array of full paths to scan for VHD files. Will prompt if omitted.")]
    [string[]]$VHDScanPaths,

    [Parameter(Mandatory = $false, HelpMessage = "Perform a dry run without deleting. Default: true.")]
    [bool]$DryRun = $true,

    [Parameter(Mandatory = $false, HelpMessage = "Include ISO files in the scan. Default: false.")]
    [bool]$IncludeISOs = $false,

    [Parameter(Mandatory = $false, HelpMessage = "Displays help information and exits.")]
    [switch]$Help
)

# Handle -Help switch
if ($Help.IsPresent) {
    Get-Help $MyInvocation.MyCommand.Definition -Full
    exit 0
}

# --- Script Body ---
Write-Verbose "Starting orphaned VM file check..."
Write-Verbose "DryRun mode: $DryRun"
Write-Verbose "Include ISOs: $IncludeISOs"

$FileExtensionsToScan = @("*.vmcx", "*.vmrs", "*.vhd", "*.vhdx", "*.avhd", "*.avhdx")
if ($IncludeISOs) { $FileExtensionsToScan += "*.iso"; Write-Warning "ISO file scanning enabled." }
$CommonSubFoldersForNamingInference = @("Virtual Hard Disks", "Snapshots", "Virtual Machines")

# --- Interactive Prompting for Scan Paths if not provided ---
if (-not $PSBoundParameters.ContainsKey('VMConfigScanPath')) {
    Write-Host "`n--- VM Configuration Scan Path ---" -ForegroundColor Cyan
    Write-Host "Please provide the full path to the directory where your main VM configuration files"
    Write-Host "(.vmcx, .vmrs) are stored. This script will scan this directory recursively."
    Write-Host "Example: C:\Hyper-V\Virtual Machines  or  H:\VMConfigs\VMs"
    $VMConfigScanPath = Read-Host -Prompt "Enter Main VM Config Scan Path"
}

if (-not $PSBoundParameters.ContainsKey('SnapshotConfigScanPath')) {
    Write-Host "`n--- Snapshot Configuration Scan Path ---" -ForegroundColor Cyan
    Write-Host "Please provide the full path to the directory where your VM snapshot configuration files"
    Write-Host "(.vmcx, .vmrs for snapshots) are stored. This script will scan this directory recursively."
    Write-Host "Example: C:\Hyper-V\Snapshots  or  H:\VMConfigs\Snapshots"
    $SnapshotConfigScanPath = Read-Host -Prompt "Enter Snapshot Config Scan Path"
}

if (-not $PSBoundParameters.ContainsKey('VHDScanPaths')) {
    Write-Host "`n--- Virtual Hard Disk (VHD/VHDX) Scan Paths ---" -ForegroundColor Cyan
    $tempVHDPathsList = [System.Collections.Generic.List[string]]::new() 
    $addMoreVHDPaths = $true
    $firstVHDPathPrompt = $true 
    
    while ($addMoreVHDPaths) {
        if ($firstVHDPathPrompt) {
            Write-Host "Please provide the full path to a directory where your virtual hard disk files"
            Write-Host "(.vhd, .vhdx, .avhd, .avhdx) are stored. This script will scan this directory recursively."
            Write-Host "Example: C:\Hyper-V\Virtual Hard Disks  or  H:\VHDStore"
            $currentVHDPath = Read-Host -Prompt "Enter VHD Scan Path (or press Enter if no VHD paths to add)"
            $firstVHDPathPrompt = $false
        } else {
            $currentVHDPath = Read-Host -Prompt "Enter another VHD Scan Path (or press Enter to finish)"
        }

        if ([string]::IsNullOrWhiteSpace($currentVHDPath)) {
            $addMoreVHDPaths = $false
        } else {
            $tempVHDPathsList.Add($currentVHDPath) 
        }
    }
    $VHDScanPaths = $tempVHDPathsList.ToArray() 
}


# --- Validate and Collect Final Scan Locations ---
$ScanLocations = [System.Collections.Generic.List[string]]::new()
$validationFailed = $false

if ([string]::IsNullOrWhiteSpace($VMConfigScanPath) -or -not (Test-Path $VMConfigScanPath -PathType Container)) {
    Write-Error "VMConfigScanPath '$VMConfigScanPath' is not provided or not a valid directory."
    $validationFailed = $true
} else {
    $ScanLocations.Add($VMConfigScanPath)
    Write-Verbose "Using VMConfigScanPath: $VMConfigScanPath"
}

if ([string]::IsNullOrWhiteSpace($SnapshotConfigScanPath) -or -not (Test-Path $SnapshotConfigScanPath -PathType Container)) {
    Write-Error "SnapshotConfigScanPath '$SnapshotConfigScanPath' is not provided or not a valid directory."
    $validationFailed = $true
} else {
    $ScanLocations.Add($SnapshotConfigScanPath)
    Write-Verbose "Using SnapshotConfigScanPath: $SnapshotConfigScanPath"
}

if ($VHDScanPaths -and $VHDScanPaths.Count -gt 0) {
    foreach ($path in $VHDScanPaths) {
        if (-not ([string]::IsNullOrWhiteSpace($path)) -and (Test-Path $path -PathType Container)) {
            $ScanLocations.Add($path)
            Write-Verbose "Using VHDScanPath: $path"
        } else {
            Write-Error "A provided VHDScanPath '$path' is not a valid directory."
            $validationFailed = $true
        }
    }
} else {
    Write-Warning "No VHD scan paths were specified. VHDs/AVHDs will not be scanned for orphans unless they happen to be in config/snapshot scan paths."
}

if ($validationFailed) {
    Write-Error "One or more scan paths are invalid. Please check the paths and try again."
    exit 1
}

$FinalScanLocations = $ScanLocations | Get-Unique
if ($FinalScanLocations.Count -eq 0) {
    Write-Error "No valid scan locations were ultimately determined. Exiting."
    exit 1
}
Write-Verbose "Final locations to scan for potential orphans: $($FinalScanLocations -join ', ')"


# 1. Get all files actively used by current VMs and cache VM info
Write-Host "`nGathering information about active VM files and current VMs..." -ForegroundColor Green
$ActiveVMFiles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$RegisteredVMsCache = @{}

try {
    $AllVMs = Get-VM -ErrorAction Stop
} catch {
    Write-Error "Failed to get list of VMs. Ensure Hyper-V module is available and you have permissions. Error: $($_.Exception.Message)"
    exit 1
}

if (-not $AllVMs) { Write-Warning "No virtual machines found on this host." }
else {
    foreach ($VM in $AllVMs) {
        $RegisteredVMsCache[$VM.VMId.ToString()] = $VM
        $RegisteredVMsCache[$VM.VMName.ToLowerInvariant()] = $VM
        Write-Verbose "Processing VM: $($VM.VMName) (ID: $($VM.VMId))"

        $vmcxFile = Join-Path -Path $VMConfigScanPath -ChildPath ($VM.VMId.ToString() + ".vmcx")
        Write-Verbose "  Attempting to locate Main VMCX at: '$vmcxFile'"
        if (Test-Path $vmcxFile -PathType Leaf) {
            if ($ActiveVMFiles.Add($vmcxFile)) { Write-Verbose "    Added active Main VMCX: $vmcxFile" }
            $vmrsFile = Join-Path -Path $VMConfigScanPath -ChildPath ($VM.VMId.ToString() + ".vmrs") 
            if (Test-Path $vmrsFile -PathType Leaf) {
                if ($ActiveVMFiles.Add($vmrsFile)) { Write-Verbose "      Added active Main VMRS: $vmrsFile" }
            } else { Write-Verbose "      Main VMRS for '$($VM.VMName)' not found at '$vmrsFile' (normal if VM is 'Off')." }
        } else {
            Write-Warning "  Main VMCX for VM '$($VM.VMName)' not found at expected path '$vmcxFile'."
        }

        ($VM | Get-VMHardDiskDrive) | ForEach-Object {
            if ($_.Path) {
                if ($ActiveVMFiles.Add($_.Path)) { Write-Verbose "    Added active VHD: $($_.Path)" }
                $CurrentDiskPath = $_.Path
                while ($CurrentDiskPath) {
                    try {
                        $DiskInfo = Get-VHD -Path $CurrentDiskPath -ErrorAction Stop
                        if ($DiskInfo.ParentPath) {
                            if ($ActiveVMFiles.Add($DiskInfo.ParentPath)) { Write-Verbose "      Added active parent VHD: $($DiskInfo.ParentPath)" }
                            $CurrentDiskPath = $DiskInfo.ParentPath
                        } else { $CurrentDiskPath = $null }
                    } catch { Write-Warning "      Could not get VHD info for '$CurrentDiskPath' (VM: $($VM.VMName)). Error: $($_.Exception.Message)"; $CurrentDiskPath = $null }
                }
            }
        }
        ($VM | Get-VMDvdDrive) | ForEach-Object { if ($_.Path) { if ($ActiveVMFiles.Add($_.Path)) { Write-Verbose "    Added active ISO: $($_.Path)" } } }
        
        $VMSnapshots = $VM | Get-VMSnapshot
        foreach ($Snapshot in $VMSnapshots) {
            Write-Verbose "    Processing Snapshot: $($Snapshot.Name) (ID: $($Snapshot.Id))"
            $snapshotVmcxFile = Join-Path -Path $SnapshotConfigScanPath -ChildPath ($Snapshot.Id.ToString() + ".vmcx")
            Write-Verbose "      Attempting to locate Snapshot VMCX at: '$snapshotVmcxFile'"
            if (Test-Path $snapshotVmcxFile -PathType Leaf) {
                if ($ActiveVMFiles.Add($snapshotVmcxFile)) { Write-Verbose "        Added active snapshot VMCX: $snapshotVmcxFile" }
                $snapshotVmrsFile = Join-Path -Path $SnapshotConfigScanPath -ChildPath ($Snapshot.Id.ToString() + ".vmrs") 
                if (Test-Path $snapshotVmrsFile -PathType Leaf) {
                    if ($ActiveVMFiles.Add($snapshotVmrsFile)) { Write-Verbose "          Added active snapshot VMRS: $snapshotVmrsFile" }
                } else { Write-Verbose "          Snapshot VMRS for '$($Snapshot.Name)' not found at '$snapshotVmrsFile'." }
            } else {
                Write-Warning "        Snapshot VMCX for '$($Snapshot.Name)' (VM '$($VM.VMName)') not found at '$snapshotVmcxFile'."
            }
        }
    }
}
Write-Host "Found $($ActiveVMFiles.Count) actively used files." -ForegroundColor Green
Write-Host "Cached $($RegisteredVMsCache.Count) registered VM details." -ForegroundColor Green

Write-Host "`nScanning disk locations for potential VM files..." -ForegroundColor Green
$PotentialDiskFiles = @()
foreach ($location in $FinalScanLocations) {
    Write-Verbose "Scanning location: $location"
    Get-ChildItem -Path $location -Include $FileExtensionsToScan -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        $PotentialDiskFiles += $_
    }
}
$PotentialDiskFiles = $PotentialDiskFiles | Get-Unique -AsString 
Write-Host "Found $($PotentialDiskFiles.Count) potential VM-related files in scanned locations." -ForegroundColor Green

Write-Host "`nIdentifying orphaned files and their potential previous owners..." -ForegroundColor Green
$OrphanedFilesDetailed = @()
foreach ($FileItem in $PotentialDiskFiles) {
    if (-not $ActiveVMFiles.Contains($FileItem.FullName)) {
        $fileNameNoExt = $FileItem.BaseName
        $parentDir = $FileItem.Directory
        $parentDirName = $parentDir.Name
        $orphanDetails = [PSCustomObject]@{
            Path                         = $FileItem.FullName; Name = $FileItem.Name
            SizeMB                       = [math]::Round($FileItem.Length / 1MB, 2); LastWriteTime = $FileItem.LastWriteTime
            Directory                    = $FileItem.DirectoryName; PotentialOriginalVMIdentifier = "Unknown"
            OriginalVMRegistered          = "N/A"; OriginalVMCurrentState = "N/A"; SourceFileItem = $FileItem
        }
        $foundAssociatedVM = $null
        if ($FileItem.Extension -in ".vmcx", ".vmrs") {
            if ($fileNameNoExt -match "^([0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12})$") {
                $vmGuid = $Matches[1]; $orphanDetails.PotentialOriginalVMIdentifier = "ID: $vmGuid"
                if ($RegisteredVMsCache.ContainsKey($vmGuid)) {
                    $foundAssociatedVM = $RegisteredVMsCache[$vmGuid]
                    $orphanDetails.PotentialOriginalVMIdentifier = "ID: $($foundAssociatedVM.VMId) (Name: $($foundAssociatedVM.VMName))"
                }
            }
        } else { 
            if ($RegisteredVMsCache.ContainsKey($fileNameNoExt.ToLowerInvariant())) {
                $foundAssociatedVM = $RegisteredVMsCache[$fileNameNoExt.ToLowerInvariant()]
                $orphanDetails.PotentialOriginalVMIdentifier = "Name (from file): $($foundAssociatedVM.VMName) (ID: $($foundAssociatedVM.VMId))"
            }
            if (-not $foundAssociatedVM -and $RegisteredVMsCache.ContainsKey($parentDirName.ToLowerInvariant())) {
                $foundAssociatedVM = $RegisteredVMsCache[$parentDirName.ToLowerInvariant()]
                $orphanDetails.PotentialOriginalVMIdentifier = "Name (from dir): $($foundAssociatedVM.VMName) (ID: $($foundAssociatedVM.VMId))"
            }
            if (-not $foundAssociatedVM -and $parentDir.Parent -and ($CommonSubFoldersForNamingInference -contains $parentDirName)) {
                $grandParentDirName = $parentDir.Parent.Name
                if ($RegisteredVMsCache.ContainsKey($grandParentDirName.ToLowerInvariant())) {
                    $foundAssociatedVM = $RegisteredVMsCache[$grandParentDirName.ToLowerInvariant()]
                    $orphanDetails.PotentialOriginalVMIdentifier = "Name (from grandparent dir): $($foundAssociatedVM.VMName) (ID: $($foundAssociatedVM.VMId))"
                }
            }
            if (-not $foundAssociatedVM -and $fileNameNoExt) { $orphanDetails.PotentialOriginalVMIdentifier = "Name (inferred from file): $fileNameNoExt" }
        }
        if ($foundAssociatedVM) {
            $orphanDetails.OriginalVMRegistered = "Yes (but this file is orphaned)"; $orphanDetails.OriginalVMCurrentState = $foundAssociatedVM.State.ToString()
        } else {
            if ($orphanDetails.PotentialOriginalVMIdentifier -ne "Unknown" -and $orphanDetails.PotentialOriginalVMIdentifier -notmatch "^ID: ([0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12})$") {
                 $orphanDetails.OriginalVMRegistered = "No (VM not found by identifier: $($orphanDetails.PotentialOriginalVMIdentifier))"
            } elseif ($orphanDetails.PotentialOriginalVMIdentifier -match "^ID: ([0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12})$") {
                 $orphanDetails.OriginalVMRegistered = "No (VM with ID '$($Matches[1])' not found)"
            } else { $orphanDetails.OriginalVMRegistered = "No (Could not infer identifier or VM not found)" }
        }
        $OrphanedFilesDetailed += $orphanDetails
    }
}

if ($OrphanedFilesDetailed.Count -eq 0) { Write-Host "`nNo orphaned VM files found in the scanned locations." -ForegroundColor Green; exit 0 }
Write-Host "`nFound $($OrphanedFilesDetailed.Count) potential orphaned file(s):" -ForegroundColor Yellow
$OrphanedFilesDetailed | Format-Table Name, SizeMB, LastWriteTime, PotentialOriginalVMIdentifier, OriginalVMRegistered, OriginalVMCurrentState, Directory -AutoSize -Wrap

if ($DryRun) { Write-Host "`nDryRun mode. No files will be deleted. Re-run with -DryRun:`$false to enable." -ForegroundColor Cyan; exit 0 }
Write-Warning "`nWARNING: You are about to be prompted to delete files. Ensure backups and understanding."
$ConfirmGlobal = $null; $FilesToDelete = @(); $FilesToKeep = @()
for ($i = 0; $i -lt $OrphanedFilesDetailed.Count; $i++) {
    $Orphan = $OrphanedFilesDetailed[$i]; $FilePath = $Orphan.Path; $FileSizeMB = $Orphan.SizeMB
    $PotentiallyFrom = if ($Orphan.PotentialOriginalVMIdentifier -ne "Unknown") { " (Potentially from: $($Orphan.PotentialOriginalVMIdentifier))" } else { "" }
    if ($ConfirmGlobal -eq "NoToAll") { Write-Host "Skipping '$FilePath'$PotentiallyFrom (NoToAll)" -ForegroundColor Gray; $FilesToKeep += $Orphan; continue }
    $ShouldProcessMessage = "Delete orphaned file '$FilePath' ($($FileSizeMB)MB)$PotentiallyFrom?"
    if ($ConfirmGlobal -eq "YesToAll") {
        Write-Host "Deleting '$FilePath'$PotentiallyFrom (YesToAll)" -ForegroundColor Magenta
        if ($PSCmdlet.ShouldProcess($FilePath, "Delete Orphaned File")) {
            try { Remove-Item -Path $FilePath -Force -ErrorAction Stop; Write-Host "  OK: $FilePath" -ForegroundColor Green; $FilesToDelete += $Orphan }
            catch { Write-Error "  FAIL: $FilePath. Error: $($_.Exception.Message)" }
        } else { Write-Warning "  SKIP (WhatIf/Confirm): $FilePath"; $FilesToKeep += $Orphan }
        continue
    }
    $Choice = Read-Host -Prompt "Delete '$FilePath' ($($FileSizeMB)MB)$PotentiallyFrom? (Y)es, (N)o, (A)ll Yes, (L)all No, (Q)uit"
    switch ($Choice.ToUpper()) {
        "Y" { if ($PSCmdlet.ShouldProcess($FilePath, "Delete Orphaned File")) { try { Remove-Item -Path $FilePath -Force -ErrorAction Stop; Write-Host "  OK: $FilePath" -ForegroundColor Green; $FilesToDelete += $Orphan } catch { Write-Error "  FAIL: $FilePath. Error: $($_.Exception.Message)" } } else { Write-Warning "  SKIP (WhatIf/Confirm): $FilePath"; $FilesToKeep += $Orphan } }
        "N" { Write-Host "  Skipped: $FilePath" -ForegroundColor Gray; $FilesToKeep += $Orphan }
        "A" { $ConfirmGlobal = "YesToAll"; Write-Host "Deleting '$FilePath'$PotentiallyFrom and subsequent..." -ForegroundColor Magenta; if ($PSCmdlet.ShouldProcess($FilePath, "Delete Orphaned File")) { try { Remove-Item -Path $FilePath -Force -ErrorAction Stop; Write-Host "  OK: $FilePath" -ForegroundColor Green; $FilesToDelete += $Orphan } catch { Write-Error "  FAIL: $FilePath. Error: $($_.Exception.Message)" } } else { Write-Warning "  SKIP (WhatIf/Confirm): $FilePath"; $FilesToKeep += $Orphan } }
        "L" { $ConfirmGlobal = "NoToAll"; Write-Host "Skipping '$FilePath'$PotentiallyFrom and subsequent..." -ForegroundColor Gray; $FilesToKeep += $Orphan }
        "Q" { Write-Host "Quitting." -ForegroundColor Yellow; for ($j = $i; $j -lt $OrphanedFilesDetailed.Count; $j++) { $FilesToKeep += $OrphanedFilesDetailed[$j] }; $i = $OrphanedFilesDetailed.Count }
        default { Write-Warning "Invalid choice. Assuming 'No'."; Write-Host "  Skipped: $FilePath" -ForegroundColor Gray; $FilesToKeep += $Orphan }
    }
}
Write-Host "`n--- Summary ---" -ForegroundColor Cyan
Write-Host "Processed: $($OrphanedFilesDetailed.Count), Deleted: $($FilesToDelete.Count), Kept/Skipped: $($FilesToKeep.Count)"
if ($FilesToDelete.Count -gt 0) { Write-Host "Deleted files:"; $FilesToDelete.Path | ForEach-Object { Write-Host "  $_" -ForegroundColor Red } }
if ($FilesToKeep.Count -gt 0) { Write-Host "Kept/Skipped files:"; $FilesToKeep.Path | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray } }
Write-Host "`nOrphaned file check complete." -ForegroundColor Green