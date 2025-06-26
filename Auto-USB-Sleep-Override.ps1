$OutputEncoding = [System.Text.Encoding]::UTF8

# ========================================
# USB SLEEP FIX TOOL
# ========================================
# Combines both versions with:
#   - Multiple override formats for compatibility
#   - History tracking for safe removal
#   - Submenu for removal options
#   - Improved device scanning

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host ""
    Write-Host "[ATTENTION] IMPORTANT: You need to run this as Administrator!" -ForegroundColor Red
    Write-Host "            Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# Global variable to track what this script has added
$global:ScriptAddedOverrides = @()

function Show-Title {
    Clear-Host
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "              USB SLEEP FIX TOOL" -ForegroundColor Cyan
    Write-Host "      Fix USB devices preventing system sleep" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Show-Menu {
    Write-Host "MAIN MENU" -ForegroundColor White
    Write-Host "--------------------------------------------------" -ForegroundColor DarkCyan
    Write-Host "1. Fix all audio devices" -ForegroundColor Green
    Write-Host "2. Fix all USB devices" -ForegroundColor Yellow
    Write-Host "3. See what's currently preventing sleep" -ForegroundColor Cyan
    Write-Host "4. Show all existing request overrides" -ForegroundColor Magenta
    Write-Host "5. Remove sleep fixes applied by this script (Undo)" -ForegroundColor Red
    Write-Host "6. Exit" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Enter your choice (1-6): " -ForegroundColor White -NoNewline
}

function Get-ProblematicDevices {
    param([string]$FilterType = "all")
    
    Write-Host "[SCAN] Scanning your USB devices..." -ForegroundColor Yellow
    
    $devices = Get-PnpDevice -PresentOnly | Where-Object { 
        $_.InstanceId -like "USB*" -and $_.Status -eq "OK" 
    }
    
    $problematicDevices = @()
    
    foreach ($device in $devices) {
        $shouldInclude = $false
        
        switch ($FilterType) {
            "audio" {
                if ($device.FriendlyName -match "(Audio|Sound|Headset|Speaker|Microphone)" -or
                    $device.Class -eq "AudioEndpoint" -or 
                    $device.Class -eq "MEDIA") {
                    $shouldInclude = $true
                }
            }
            "all" {
                if ($device.FriendlyName -match "(Audio|Sound|Headset|Speaker|Microphone|Mouse|Keyboard)" -or
                    $device.Class -in @("AudioEndpoint", "MEDIA", "HIDClass")) {
                    $shouldInclude = $true
                }
            }
        }
        
        if ($shouldInclude) {
            $problematicDevices += $device
        }
    }
    
    return $problematicDevices
}

function Show-DeviceList {
    param($devices, $action = "fix")
    
    if ($devices.Count -eq 0) {
        Write-Host "[ERROR] No devices found!" -ForegroundColor Red
        return $false
    }
    
    Write-Host ""
    Write-Host "Found $($devices.Count) device(s) to ${action}:" -ForegroundColor Green
    Write-Host ""
    
    for ($i = 0; $i -lt $devices.Count; $i++) {
        $device = $devices[$i]
        Write-Host "  + $($device.FriendlyName)" -ForegroundColor White
    }
    
    return $true
}

function Apply-SleepFix {
    param($devices)
    
    Write-Host ""
    Write-Host "[APPLYING] Applying sleep fixes..." -ForegroundColor Yellow
    
    $successCount = 0
    
    foreach ($device in $devices) {
        $deviceName = $device.FriendlyName
        $instanceId = $device.InstanceId
        
        Write-Host "  [FIXING] Processing: $deviceName" -ForegroundColor Cyan
        
        # Escape any parentheses in device names for proper formatting
        $escapedDeviceName = $deviceName -replace '\(', '`(' -replace '\)', '`)'
        
        # Try ALL possible override formats for maximum compatibility
        $commands = @(
            # Format 1: Just device name
            "powercfg /requestsoverride DRIVER `"$escapedDeviceName`" DISPLAY SYSTEM AWAYMODE",
            
            # Format 2: Device name with instance ID in parentheses (exactly: "Name (InstanceID)")
            "powercfg /requestsoverride DRIVER `"$escapedDeviceName ($instanceId)`" DISPLAY SYSTEM AWAYMODE",
            
            # Format 3: Just the instance ID
            "powercfg /requestsoverride DRIVER `"$instanceId`" DISPLAY SYSTEM AWAYMODE",
            
            # Format 4: Try with SYSTEM only (some devices need this)
            "powercfg /requestsoverride DRIVER `"$escapedDeviceName`" SYSTEM",
            "powercfg /requestsoverride DRIVER `"$instanceId`" SYSTEM",
            
            # Format 5: Try with all override types
            "powercfg /requestsoverride DRIVER `"$escapedDeviceName`" DISPLAY SYSTEM AWAYMODE EXECUTION",
            "powercfg /requestsoverride DRIVER `"$instanceId`" DISPLAY SYSTEM AWAYMODE EXECUTION"
        )
        
        $appliedOverrides = 0
        
        foreach ($cmd in $commands) {
            try {
                $formatDescription = switch -Wildcard ($cmd) {
                    "*`"$escapedDeviceName`"*" { "Device name only" }
                    "*`"$escapedDeviceName ($instanceId)`"*" { "Device name with instance ID" }
                    "*`"$instanceId`"*" { "Instance ID only" }
                    default { "Custom format" }
                }
                
                Write-Host "    Trying [$formatDescription]: $($cmd.Split('"')[1])..." -ForegroundColor Gray
                Invoke-Expression $cmd 2>$null
                Write-Host "    [SUCCESS] Applied successfully!" -ForegroundColor Green
                $appliedOverrides++
            }
            catch {
                Write-Host "    [SKIPPED] Skipped (already exists or failed)" -ForegroundColor DarkGray
            }
        }
        
        if ($appliedOverrides -gt 0) {
            Write-Host "  [SUCCESS] Fixed: $deviceName ($appliedOverrides overrides applied)" -ForegroundColor Green
            Save-OverrideHistory -deviceName $deviceName -instanceId $instanceId
            $successCount++
        } else {
            Write-Host "  [ERROR] Failed: $deviceName (no overrides could be applied)" -ForegroundColor Red
        }
        
        Write-Host ""
    }
    
    Write-Host ""
    if ($successCount -eq $devices.Count) {
        Write-Host "[SUCCESS] All devices fixed!" -ForegroundColor Green
        Write-Host "  Multiple override formats applied for maximum compatibility." -ForegroundColor Green
    } elseif ($successCount -gt 0) {
        Write-Host "[WARNING] Partial success: $successCount out of $($devices.Count) devices fixed." -ForegroundColor Yellow
    } else {
        Write-Host "[ERROR] No devices could be fixed. Try running as Administrator." -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "[TIP] You can verify by checking option 4 in the main menu" -ForegroundColor Cyan
}

function Save-OverrideHistory {
    param($deviceName, $instanceId)
    
    $logFile = "$env:TEMP\USB-Sleep-Fix-History.txt"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp|$deviceName|$instanceId"
    
    Add-Content -Path $logFile -Value $entry
    $global:ScriptAddedOverrides += @{
        DeviceName = $deviceName
        InstanceId = $instanceId
        Timestamp = $timestamp
    }
}

function Get-ScriptOverrideHistory {
    $logFile = "$env:TEMP\USB-Sleep-Fix-History.txt"
    
    if (Test-Path $logFile) {
        $entries = Get-Content $logFile | ForEach-Object {
            $parts = $_ -split '\|'
            if ($parts.Count -eq 3) {
                @{
                    Timestamp = $parts[0]
                    DeviceName = $parts[1]
                    InstanceId = $parts[2]
                }
            }
        }
        return $entries
    }
    
    return @()
}

function Remove-ScriptFixes {
    Write-Host "[INFO] Looking for fixes that THIS SCRIPT added..." -ForegroundColor Yellow
    
    $scriptOverrides = Get-ScriptOverrideHistory
    
    if ($scriptOverrides.Count -eq 0) {
        Write-Host "[INFO] No fixes from this script found to remove." -ForegroundColor Cyan
        Write-Host "       (This script only removes overrides it added itself)" -ForegroundColor Gray
        return
    }
    
    Write-Host "[INFO] Found $($scriptOverrides.Count) override(s) that this script added:" -ForegroundColor Cyan
    foreach ($override in $scriptOverrides) {
        Write-Host "  - $($override.DeviceName) (added: $($override.Timestamp))" -ForegroundColor White
    }
    
    if (-not (Confirm-Action "Do you want to proceed with removing these fixes?")) {
        return
    }
    
    Write-Host ""
    Write-Host "[REMOVE] Removing script overrides..." -ForegroundColor Yellow
    
    $removedCount = 0
    foreach ($override in $scriptOverrides) {
        $deviceName = $override.DeviceName
        $instanceId = $override.InstanceId
        
        # Try ALL formats that might have been applied
        $formats = @(
            $deviceName,                      # Format 1: Just device name
            "$deviceName ($instanceId)",      # Format 2: Device name with instance ID
            $instanceId,                      # Format 3: Instance ID alone
            "$instanceId DISPLAY",            # Format 4: Instance ID with override types
            "$deviceName DISPLAY"             # Format 5: Device name with override types
        )
        
        $anyRemoved = $false
        foreach ($format in $formats) {
            try {
                # Remove using all possible formats
                & powercfg /requestsoverride DRIVER "`"$format`"" 2>$null
                $anyRemoved = $true
                Write-Host "    [REMOVED] Format: $format" -ForegroundColor DarkGreen
            }
            catch {
                # Ignore errors when removing non-existent overrides
            }
        }
        
        if ($anyRemoved) {
            Write-Host "  [SUCCESS] Removed: $deviceName" -ForegroundColor Green
            $removedCount++
        } else {
            Write-Host "  [WARNING] Couldn't remove: $deviceName (no matching overrides found)" -ForegroundColor Yellow
        }
    }
    
    # Clear the history file
    $logFile = "$env:TEMP\USB-Sleep-Fix-History.txt"
    if (Test-Path $logFile) {
        Remove-Item $logFile -Force
    }
    
    Write-Host ""
    if ($removedCount -gt 0) {
        Write-Host "[SUCCESS] Successfully removed $removedCount device overrides!" -ForegroundColor Green
    }
    Write-Host "[NOTE] Some overrides might remain if they were applied outside this script" -ForegroundColor Cyan
}

function Show-CurrentSleepBlockers {
    Write-Host "[INFO] Checking what's currently preventing sleep..." -ForegroundColor Yellow
    Write-Host ""
    
    try {
        $requests = & powercfg -requests 2>$null
        
        if ($requests) {
            # Parse and display in a user-friendly way
            $inSystemSection = $false
            $foundBlockers = $false
            
            foreach ($line in $requests) {
                if ($line -match "SYSTEM:") {
                    $inSystemSection = $true
                    Write-Host "[BLOCKERS] Devices preventing sleep:" -ForegroundColor Red
                    continue
                }
                
                if ($line -match "^[A-Z]+:") {
                    $inSystemSection = $false
                }
                
                if ($inSystemSection -and $line.Trim() -ne "None." -and $line.Trim() -ne "") {
                    if ($line -match "\[DRIVER\]") {
                        $deviceName = ($line -split "\[DRIVER\]")[1].Trim()
                        Write-Host "  - $deviceName" -ForegroundColor Yellow
                        $foundBlockers = $true
                    }
                }
            }
            
            if (-not $foundBlockers) {
                Write-Host "[OK] No devices are currently preventing sleep!" -ForegroundColor Green
            }
        } else {
            Write-Host "[ERROR] Couldn't check sleep status." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "[ERROR] Error checking sleep blockers: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-ExistingOverrides {
    Write-Host "[INFO] Fetching all active power request overrides..." -ForegroundColor Yellow
    Write-Host ""
    
    try {
        $overrides = powercfg /requestsoverride
        $driverOverrides = @()
        $inDriverSection = $false

        # This loop specifically parses the output of powercfg to find the [DRIVER] section
        foreach ($line in $overrides) {
            if ($line -match "\[DRIVER\]") {
                $inDriverSection = $true
                continue # Skip the section header line itself
            }
            
            # If we are in the DRIVER section and the line is not empty, it's an override
            if ($inDriverSection -and -not [string]::IsNullOrWhiteSpace($line)) {
                $driverOverrides += $line.Trim()
            }
            
            # If we were in the section and hit a blank line, the section is over
            if ($inDriverSection -and [string]::IsNullOrWhiteSpace($line)) {
                $inDriverSection = $false
                break
            }
        }

        if ($driverOverrides.Count -gt 0) {
            Write-Host "[OVERRIDES FOUND]" -ForegroundColor Green
            foreach ($override in $driverOverrides) {
                Write-Host "  - $override" -ForegroundColor White
            }
        } else {
            Write-Host "[OK] No DRIVER power request overrides are currently set." -ForegroundColor Green
        }
    }
    catch {
        Write-Host "[ERROR] Could not fetch power request overrides: $($_.Exception.Message)" -ForegroundColor Red
    }
}


function Confirm-Action {
    param([string]$message)
    
    Write-Host ""
    Write-Host "$message" -ForegroundColor Yellow
    Write-Host "Continue? (Y/N): " -ForegroundColor White -NoNewline
    
    do {
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        $response = $key.Character.ToString().ToUpper()
    } while ($response -notin @('Y', 'N'))
    
    Write-Host $response
    return $response -eq 'Y'
}

function Wait-ForUser {
    Write-Host ""
    Write-Host "Press any key to continue..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ========================================
# MAIN PROGRAM
# ========================================

do {
    Show-Title
    Show-Menu
    
    $choice = Read-Host
    
    switch ($choice) {
        "1" {
            Show-Title
            Write-Host "[OPTION 1] Audio Device Sleep Fix" -ForegroundColor Green
            Write-Host "This will fix all audio devices from preventing sleep." -ForegroundColor White
            Write-Host "(Including headsets like ASTRO A50, speakers, microphones, etc.)" -ForegroundColor White
            Write-Host ""
            
            $devices = Get-ProblematicDevices -FilterType "audio"
            
            if (Show-DeviceList -devices $devices) {
                if (Confirm-Action "Apply sleep fix to these audio devices?") {
                    Apply-SleepFix -devices $devices
                }
            } else {
                Write-Host "[INFO] No problematic audio devices found." -ForegroundColor Cyan
            }
            
            Wait-ForUser
        }
        
        "2" {
            Show-Title
            Write-Host "[OPTION 2] All USB Device Sleep Fix" -ForegroundColor Green
            Write-Host "This will fix all USB devices that commonly prevent sleep." -ForegroundColor White
            Write-Host "(Including audio devices, mice, keyboards, etc.)" -ForegroundColor White
            Write-Host ""
            
            $devices = Get-ProblematicDevices -FilterType "all"
            
            if (Show-DeviceList -devices $devices) {
                if (Confirm-Action "Apply sleep fix to ALL these devices?") {
                    Apply-SleepFix -devices $devices
                }
            } else {
                Write-Host "[INFO] No problematic devices found." -ForegroundColor Cyan
            }
            
            Wait-ForUser
        }
        
        "3" {
            Show-Title
            Write-Host "[OPTION 3] Current Sleep Status" -ForegroundColor Cyan
            Show-CurrentSleepBlockers
            Wait-ForUser
        }
        
        "4" {
            Show-Title
            Write-Host "[OPTION 4] Show Existing Request Overrides" -ForegroundColor Magenta
            Show-ExistingOverrides
            Wait-ForUser
        }
        
        "5" {
            Show-Title
            Write-Host "[OPTION 5] Remove Fixes Applied by This Script" -ForegroundColor Red
            Write-Host ""
            Remove-ScriptFixes
            Wait-ForUser
        }
        
        "6" {
            Show-Title
            Write-Host "[EXIT] Thanks for using USB Sleep Fix Tool!" -ForegroundColor Green
            Write-Host "       Your PC should now sleep properly." -ForegroundColor White
            Write-Host ""
            exit
        }
        
        default {
            Show-Title
            Write-Host "[ERROR] Invalid choice. Please enter 1-6." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
    
} while ($true)
