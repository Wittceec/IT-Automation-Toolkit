<# Print all schedules from the correct synced folder (auto-detect path) #>

param(
  [string]$TargetFolderName = "Operations Schedules",
  [string[]]$Extensions = @(".docx",".xlsx")
)

# --- CRASH CATCHER (Keeps window open on error) ---
Trap {
    Write-Host "`n[CRITICAL ERROR] The script crashed:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit 1
}

# ==============================================================================
# CONFIGURATION
# SANITIZED: Replaced specific internal document name
$DuplexFileName = "[Special_Print_Handling_Doc]" 
# ==============================================================================

Write-Host "============================================" -ForegroundColor Green
Write-Host "  PRINTING ALL SCHEDULES IN THIS FOLDER" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""

# 1) Build candidate roots where OneDrive/SharePoint usually syncs
# SANITIZED: Replaced specific organization names with placeholders
$roots = @(
  $env:OneDriveCommercial,                      
  $env:OneDrive,                                
  (Join-Path $env:USERPROFILE "OneDrive - [Organization]"),
  (Join-Path $env:USERPROFILE "[Organization]"),  
  (Join-Path $env:USERPROFILE "SharePoint"),          
  $env:USERPROFILE                              
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

# 2) Try known SharePoint site/library patterns first (fast)
# SANITIZED: Replaced specific department routing names
$preferredSubpaths = @(
  "[Department] - Documents\General\$TargetFolderName",
  "Documents\General\$TargetFolderName",                        
  "[Department] - Operations Schedules"                
)

$found = $null
foreach ($r in $roots) {
  foreach ($sp in $preferredSubpaths) {
    $candidate = Join-Path $r $sp
    if (Test-Path $candidate) { $found = (Get-Item $candidate).FullName; break }
  }
  if ($found) { break }
}

# 3) If still not found, do a constrained search
if (-not $found) {
  foreach ($r in $roots) {
    try {
      $hit = Get-ChildItem -Path $r -Directory -Recurse -Depth 4 -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -eq $TargetFolderName } |
             Select-Object -First 1
      if ($hit) { $found = $hit.FullName; break }
    } catch { }
  }
}

if (-not $found) {
  Write-Host "I couldn't find the synced '$TargetFolderName' folder." -ForegroundColor Yellow
  Read-Host "Press Enter to exit..."
  exit 1
}

$folderPath = $found
Write-Host "Found folder: $folderPath" -ForegroundColor Cyan

# 4) Pin the folder so files are local
try {
  Write-Host "Pinning files locally..."
  attrib +P "$folderPath"
  attrib +P "$folderPath\*" /S /D
} catch {
  Write-Host "Pin attempt skipped." -ForegroundColor DarkYellow
}

# 5) List and print files
Write-Host "Listing all files in the folder:"
Get-ChildItem -Path $folderPath

Write-Host "Filtering for $($Extensions -join ', ') ..."
$filesToPrint = Get-ChildItem -Path $folderPath -File -Recurse:$false | Where-Object {
  $Extensions -contains $_.Extension.ToLower()
}

if (-not $filesToPrint) {
  Write-Host "No schedule files found." -ForegroundColor Yellow
  Read-Host "Press Enter to exit..."
  exit
}

# Use Office COM
$word = $null; $excel = $null
try { $word = New-Object -ComObject Word.Application;  $word.Visible = $false } catch {}
try { $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false } catch {}

foreach ($file in $filesToPrint) {
  Write-Host " -> Processing $($file.Name)..."
  try {
    switch ($file.Extension.ToLower()) {
      ".docx" {
        if ($word) {
          $doc = $word.Documents.Open($file.FullName, $false, $true)
          $doc.PrintOut()
          $doc.Close($false)
        } else {
          Start-Process -FilePath $file.FullName -Verb Print
        }
      }
      ".xlsx" {
        if ($excel) {
          # --- DUPLEX FIX LOGIC START ---
          if ($file.Name -match $DuplexFileName) {
              # If this is the "Problem File", we make Excel visible and show the Print Dialog
              Write-Host "    [!] MANUAL INTERVENTION: Please select '2-Sided' and Print." -ForegroundColor Yellow
              
              $excel.Visible = $true  # Show Excel so they can see the dialog
              $wb = $excel.Workbooks.Open($file.FullName, $false, $true)
              
              # Show the classic Print Dialog (Dialog ID 8)
              $excel.Dialogs.Item(8).Show() 
              
              $wb.Close($false)
              $excel.Visible = $false # Hide Excel again for the next files
          } 
          else {
              # Standard Automatic Printing for everything else
              $wb = $excel.Workbooks.Open($file.FullName, $false, $true)
              $wb.PrintOut()
              
              # === FIX: WAIT FOR EXCEL TO SPOOL ===
              Write-Host "    ...spooling..." -NoNewline
              Start-Sleep -Seconds 3 
              # ====================================
              
              $wb.Close($false)
          }
          # --- DUPLEX FIX LOGIC END ---
        } else {
          Start-Process -FilePath $file.FullName -Verb Print
        }
      }
    }
    # Increased global delay to prevent printer queue jams
    Start-Sleep -Seconds 5
  } catch {
    Write-Host "   Printing failed for $($file.Name): $($_.Exception.Message)" -ForegroundColor Yellow
  }
}

if ($excel) { $excel.Quit() | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
if ($word)  { $word.Quit()  | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)  }

Write-Host ""
Write-Host "All print jobs have been sent!" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to exit..."
}