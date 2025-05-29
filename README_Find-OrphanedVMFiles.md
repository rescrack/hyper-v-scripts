# [Find-OrphanedVMFilesWithHistory.ps1](https://github.com/rescrack/hyper-v-scripts/blob/main/Find-OrphanedVMFiles.ps1)

This PowerShell script identifies **orphaned Hyper-V virtual machine files** that are no longer associated with any registered VM on the host. These include:

* Configuration files: `.vmcx`, `.vmrs`
* Virtual hard disks: `.vhd`, `.vhdx`, `.avhd`, `.avhdx`
* ISO files: `.iso` (optional)

The script can **scan directories recursively**, **compare file usage**, and **prompt for deletion** of unassociated files.
Supports interactive mode or parameter-based usage, along with PowerShell's `-WhatIf` and `-Confirm`.

> ⚠️ **WARNING:** Deleting files can result in data loss. Always run with `-DryRun` first, verify the output, and ensure you have backups before using `-DryRun:$false`.

---

## 🛠️ How It Works

1. **Gathers Active Files**

   * Queries all registered VMs and snapshots on the host.
   * Constructs paths for `.vmcx`/`.vmrs` files based on provided config paths and VM/snapshot IDs.
   * Collects active `.vhd(x)` and `.iso` files from VM drives.

2. **Scans Disk Locations**

   * Recursively scans provided paths (or prompts you) for VM-related file types.

3. **Compares & Identifies Orphans**

   * Compares files found on disk against actively used files.
   * Flags anything not in use as a **potential orphan**.

4. **Attempts to Identify Previous Owner**

   * Tries to match orphaned files to known VMs using filenames or directory names.

5. **Presents Orphans & Prompts for Deletion**

   * Lists orphans in the console.
   * If `-DryRun:$false`, it will prompt to delete with `Yes/No/All/Quit` options.

---

## 🔧 Parameters

| Parameter                 | Description                                                                                                                                                                         |
| ------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `-VMConfigScanPath`       | *(Optional)* Path to scan for main VM configuration files (`.vmcx`, `.vmrs`). Prompts if omitted.<br>Example: `"H:\Virtual Machine Configurations\Virtual Machines"`                |
| `-SnapshotConfigScanPath` | *(Optional)* Path to scan for snapshot configuration files. Prompts if omitted.<br>Example: `"H:\Virtual Machine Configurations\Snapshots"`                                         |
| `-VHDScanPaths`           | *(Optional)* Array of paths to scan for virtual disk files (`.vhd`, `.vhdx`, `.avhd`, `.avhdx`). Prompts if omitted.<br>Example: `@("H:\Virtual Hard Disks", "D:\ExtraVHDStorage")` |
| `-DryRun`                 | *(Default: `$true`)* When `$true`, only lists orphaned files. Set to `$false` to enable deletion prompts. **Use with caution!**                                                     |
| `-IncludeISOs`            | *(Default: `$false`)* If `$true`, includes `.iso` files in scan. Use carefully—ISOs are often reused.                                                                               |
| `-Help`                   | Displays the help text and exits.                                                                                                                                                   |

---

## 📥 Inputs

This script **does not accept pipeline input** for its main parameters.

## 📤 Outputs

* Outputs `PSCustomObject` entries with details about potential orphaned files.
* Console output includes:

  * Informational messages
  * Warnings
  * Errors

---

## 💡 Examples

```powershell
# Run the script interactively (will prompt for paths)
.\Find-OrphanedVMFilesWithHistory.ps1 -Verbose
```

```powershell
# Use parameterized paths (no prompts)
.\Find-OrphanedVMFilesWithHistory.ps1 `
    -VMConfigScanPath "H:\Virtual Machine Configurations\Virtual Machines" `
    -SnapshotConfigScanPath "H:\Virtual Machine Configurations\Snapshots" `
    -VHDScanPaths "H:\Virtual Hard Disks" `
    -Verbose
```

```powershell
# Enable deletion (EXTREME CAUTION)
.\Find-OrphanedVMFilesWithHistory.ps1 `
    -VMConfigScanPath "H:\Configs" `
    -SnapshotConfigScanPath "H:\Snapshots" `
    -VHDScanPaths "H:\VHDs" `
    -DryRun:$false `
    -Verbose
```
