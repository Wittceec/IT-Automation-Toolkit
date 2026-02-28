# --- CONFIGURATION ---
Add-Type -AssemblyName Microsoft.VisualBasic

$sourceDir = "$env:USERPROFILE\Downloads"
$archiveDir = "$env:USERPROFILE\Downloads\Sorted"

# Define your file types to be SORTED
$mapping = @{
    "Installers"   = @(".exe", ".msi", ".bat", ".sh", ".plp")
    "Documents"    = @(".pdf", ".docx", ".xlsx", ".txt", ".pptx", ".csv", ".md", ".rtf")
    "Images"       = @(".jpg", ".jpeg", ".png", ".gif", ".svg", ".bmp", ".webp")
    "Archives"     = @(".zip", ".rar", ".7z", ".tar", ".gz")
    "Code"         = @(".py", ".js", ".html", ".css", ".cpp", ".ps1", ".json", ".xml", ".yaml", ".yml", ".sql", ".vbs")
    "EmailDrafts"  = @(".msg", ".oft")
    "IT_Configs"   = @(".log", ".ini", ".conf", ".pem", ".cer", ".crt", ".pcap", ".pcapng")
    "Virtual_OS"   = @(".iso", ".ova", ".vmdk")
    "Media"        = @(".mp4", ".mov", ".wmv", ".mp3", ".wav")
}

# Define your file types to be TRASHED (Recycle Bin)
$trashExtensions = @(".rdp", ".ica")

# --- THE LOGIC ---
# Create the base Archive folder if it doesn't exist
if (-not (Test-Path -Path $archiveDir)) {
    New-Item -ItemType Directory -Path $archiveDir | Out-Null
}

# Get all files in the Downloads folder (exclude folders)
$files = Get-ChildItem -Path $sourceDir -File

foreach ($file in $files) {
    # Skip the script itself
    if ($file.Name -eq "CleanDownloads.ps1") { continue }
    
    $extension = $file.Extension.ToLower()

    # --- CHECK 1: TRASH ---
    if ($trashExtensions -contains $extension) {
        try {
            # This sends to Recycle Bin instead of permanent delete
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($file.FullName, 'OnlyErrorDialogs', 'SendToRecycleBin')
            Write-Host "Recycled $($file.Name)" -ForegroundColor Yellow
            continue # Skip to the next file
        }
        catch {
            Write-Host "Failed to recycle $($file.Name): $_" -ForegroundColor Red
        }
    }

    # --- CHECK 2: SORT ---
    foreach ($category in $mapping.Keys) {
        if ($mapping[$category] -contains $extension) {
            $targetDir = Join-Path -Path $archiveDir -ChildPath $category
            
            if (-not (Test-Path -Path $targetDir)) {
                New-Item -ItemType Directory -Path $targetDir | Out-Null
            }

            $destination = Join-Path -Path $targetDir -ChildPath $file.Name
            
            try {
                Move-Item -Path $file.FullName -Destination $destination -ErrorAction Stop
                Write-Host "Moved $($file.Name) to $category" -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to move $($file.Name): $_" -ForegroundColor Red
            }
            break
        }
    }
}
