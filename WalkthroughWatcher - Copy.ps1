<#
.SYNOPSIS
    Watcher script for Data Center Walkthroughs.
    - v3: Fixes double-posting, cleans up formatting.
#>

# --- WRAPPER TO CATCH CRASHES ---
try {
    # --- CONFIGURATION ---
    # SANITIZED: Replaced internal network path with generic placeholder
    $FolderToWatch = "\\[SERVER_NAME]\[Share]\[Path]\Operations\Walkthrough"
    # SANITIZED: Replaced live Teams channel routing email with generic placeholder
    $TeamsEmail    = "Walkthrough Photos Channel <your-teams-channel-email@amer.teams.ms>" 

    Write-Host "Initializing Watcher..." -ForegroundColor Cyan

    # --- TRACKING VARIABLES (To prevent double posts) ---
    # We use 'Global' scope so these persist between events
    $Global:LastProcessedFile = ""
    $Global:LastProcessedTime = (Get-Date).AddMinutes(-1)

    # --- THE LOGIC ---
    $Action = {
        try {
            $Path = $Event.SourceEventArgs.FullPath
            $Name = $Event.SourceEventArgs.Name
            $Now  = Get-Date

            # --- DEBOUNCE CHECK (The Duplicate Fix) ---
            # If we processed this exact file less than 60 seconds ago, STOP.
            $SecondsSinceLast = ($Now - $Global:LastProcessedTime).TotalSeconds
            if ($Path -eq $Global:LastProcessedFile -and $SecondsSinceLast -lt 60) {
                Write-Warning "Duplicate trigger detected for '$Name'. Ignoring."
                return
            }

            # Update our trackers immediately
            $Global:LastProcessedFile = $Path
            $Global:LastProcessedTime = $Now

            # 1. Wait a moment for the file save to finish
            Start-Sleep -Seconds 5
            
            # 2. Calculate "Shift Date" for the Title
            # Subtract 7 hours so 2 AM Friday becomes "Thursday Night"
            $ShiftDate = $Now.AddHours(-7) 
            $DateString = $ShiftDate.ToString("MM/dd/yy")
            
            # 3. Format the Subject Line
            # Example: "Walkthrough Night 01/16/26"
            $SubjectLine = "Walkthrough Night $DateString"
            
            Write-Host "New file detected: $Name" -ForegroundColor Cyan
            Write-Host "Posting as: $SubjectLine" -ForegroundColor Yellow

            # 4. Send to Teams
            $Outlook = New-Object -ComObject Outlook.Application
            $Mail = $Outlook.CreateItem(0)
            $Mail.To = $TeamsEmail
            $Mail.Subject = $SubjectLine
            
            # Empty body so only the image shows
            $Mail.Body = " " 
            
            $Mail.Attachments.Add($Path)
            $Mail.Send()
            
            Write-Host "Success! Posted to Teams." -ForegroundColor Green
        }
        catch {
            Write-Error "Error during processing: $_"
        }
    }

    # --- THE WATCHER ---
    $Watcher = New-Object System.IO.FileSystemWatcher
    $Watcher.Path = $FolderToWatch
    # Watch for JPGs
    $Watcher.Filter = "*Walkthrough*Night*.jpg" 
    $Watcher.IncludeSubdirectories = $false
    $Watcher.EnableRaisingEvents = $true

    # Clear old events to prevent ghost triggers
    Unregister-Event -SourceIdentifier "FileCreated" -ErrorAction SilentlyContinue

    # Register the new event
    Register-ObjectEvent $Watcher "Created" -SourceIdentifier "FileCreated" -Action $Action

    Write-Host "---------------------------------------------------" -ForegroundColor Green
    Write-Host "  WATCHER RUNNING (v3 - No Duplicates)"
    Write-Host "  Monitoring: $FolderToWatch"
    Write-Host "  Format: Walkthrough Night MM/DD/YY"
    Write-Host "---------------------------------------------------"
    Write-Host "You can minimize this window, but DO NOT close it."

    # Keep script running indefinitely
    while ($true) { Start-Sleep -Seconds 5 }

}
catch {
    Write-Host "CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    $null = Read-Host "Press ENTER to exit..."
}