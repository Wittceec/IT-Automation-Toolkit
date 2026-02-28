# File Management & Automation

A collection of utility scripts for workstation maintenance and bulk document processing.

## 📦 Included Scripts

### 1. `CleanDownloads.ps1`
An intelligent cleanup utility that automatically sorts an overcrowded Downloads folder into categorized archive directories (Installers, Documents, Code, etc.) based on file extensions. It safely routes temporary/trash files (like `.rdp`) directly to the Windows Recycle Bin using the VisualBasic FileIO namespace, rather than permanently deleting them.

### 2. `Print-Schedules_v2.ps1`
A robust bulk-printing script that automatically locates and prints operational schedules across different enterprise sync environments.
* **Dynamic Path Discovery:** Scans common `OneDrive` and `SharePoint` environmental variables to dynamically locate the correct synced folder, regardless of the user's local machine setup.
* **Office COM Automation:** Interacts directly with Word and Excel COM objects to silently spool print jobs.
* **Intervention Prompts:** Contains logic to pause the automated loop and prompt the user for manual intervention on specific documents requiring duplex (2-sided) printing.
