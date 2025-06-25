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
    Write-Host "⚠️  IMPORTANT: You need to run this as Administrator!" -ForegroundColor Red
    Write-Host "   Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# Global variable to track what this script has added
$global:ScriptAddedOverrides = @()

function Show-Title {
    Clear-Host
    Write-Host "╔════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║               USB SLEEP FIX TOOL               ║" -ForegroundColor Cyan
    Write-Host "║     Fix USB devices preventing system sleep    ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Show-Menu {
    Write-Host "MAIN MENU" -ForegroundColor White
    Write-Host "════════════════════════════════════════════════" -ForegroundColor DarkCyan
    Write-Host "1️ Fix all audio devices" -ForegroundColor Green
    Write-Host "2️ Fix all USB devices" -ForegroundColor Yellow
    Write-Host "3️ See what's currently preventing sleep" -ForegroundColor Cyan
    Write-Host "4️ Remove sleep fixes (Undo changes)" -ForegroundColor Red
    Write-Host "5️ Exit" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Enter your choice (1-5): " -ForegroundColor White -NoNewline
}

function Get-ProblematicDevices {
    param([string]$FilterType = "all")
    
    Write-Host "🔍 Scanning your USB devices..." -ForegroundColor Yellow
    
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
        Write-Host "❌ No devices found!" -ForegroundColor Red
        return $false
    }
    
    Write-Host ""
    Write-Host "Found $($devices.Count) device(s) to ${action}:" -ForegroundColor Green
    Write-Host ""
    
    for ($i = 0; $i -lt $devices.Count; $i++) {
        $device = $devices[$i]
        Write-Host "   ✓ $($device.FriendlyName)" -ForegroundColor White
    }
    
    return $true
}

function Apply-SleepFix {
    param($devices)
    
    Write-Host ""
    Write-Host "🔧 Applying sleep fixes..." -ForegroundColor Yellow
    
    $successCount = 0
    
    foreach ($device in $devices) {
        $deviceName = $device.FriendlyName
        $instanceId = $device.InstanceId
        
        Write-Host "   🔧 Processing: $deviceName" -ForegroundColor Cyan
        
        # Try ALL possible override formats for maximum compatibility
        $commands = @(
            # Format 1: Just device name
            "powercfg /requestsoverride DRIVER `"$deviceName`" DISPLAY SYSTEM AWAYMODE",
            
            # Format 2: Device name with instance ID in parentheses
            "powercfg /requestsoverride DRIVER `"$deviceName ($instanceId)`" DISPLAY SYSTEM AWAYMODE",
            
            # Format 3: Just the instance ID
            "powercfg /requestsoverride DRIVER `"$instanceId`" DISPLAY SYSTEM AWAYMODE",
            
            # Format 4: Try with SYSTEM only (some devices need this)
            "powercfg /requestsoverride DRIVER `"$deviceName`" SYSTEM",
            "powercfg /requestsoverride DRIVER `"$instanceId`" SYSTEM",
            
            # Format 5: Try with all override types
            "powercfg /requestsoverride DRIVER `"$deviceName`" DISPLAY SYSTEM AWAYMODE EXECUTION",
            "powercfg /requestsoverride DRIVER `"$instanceId`" DISPLAY SYSTEM AWAYMODE EXECUTION"
        )
        
        $appliedOverrides = 0
        
        foreach ($cmd in $commands) {
            try {
                Write-Host "      Trying: $($cmd.Split('"')[1])..." -ForegroundColor Gray
                Invoke-Expression $cmd 2>$null
                Write-Host "      ✅ Applied successfully!" -ForegroundColor Green
                $appliedOverrides++
            }
            catch {
                Write-Host "      ⚠️  Skipped (already exists or failed)" -ForegroundColor DarkGray
            }
        }
        
        if ($appliedOverrides -gt 0) {
            Write-Host "   ✅ Fixed: $deviceName ($appliedOverrides overrides applied)" -ForegroundColor Green
            Save-OverrideHistory -deviceName $deviceName -instanceId $instanceId
            $successCount++
        } else {
            Write-Host "   ❌ Failed: $deviceName (no overrides could be applied)" -ForegroundColor Red
        }
        
        Write-Host ""
    }
    
    Write-Host ""
    if ($successCount -eq $devices.Count) {
        Write-Host "🎉 SUCCESS! All devices fixed!" -ForegroundColor Green
        Write-Host "   Multiple override formats applied for maximum compatibility." -ForegroundColor Green
    } elseif ($successCount -gt 0) {
        Write-Host "⚠️  Partial success: $successCount out of $($devices.Count) devices fixed." -ForegroundColor Yellow
    } else {
        Write-Host "❌ No devices could be fixed. Try running as Administrator." -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "💡 Tip: You can verify with 'powercfg /requestsoverride' command" -ForegroundColor Cyan
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
    Write-Host "🔍 Looking for fixes that THIS SCRIPT added..." -ForegroundColor Yellow
    
    $scriptOverrides = Get-ScriptOverrideHistory
    
    if ($scriptOverrides.Count -eq 0) {
        Write-Host "ℹ️  No fixes from this script found to remove." -ForegroundColor Cyan
        Write-Host "   (This script only removes overrides it added itself)" -ForegroundColor Gray
        return
    }
    
    Write-Host "📋 Found $($scriptOverrides.Count) override(s) that this script added:" -ForegroundColor Cyan
    foreach ($override in $scriptOverrides) {
        Write-Host "   • $($override.DeviceName) (added: $($override.Timestamp))" -ForegroundColor White
    }
    
    if (-not (Confirm-Action "Do you want to proceed?")) {
        return
    }
    
    Write-Host ""
    Write-Host "🗑️  Removing script overrides..." -ForegroundColor Yellow
    
    $removedCount = 0
    foreach ($override in $scriptOverrides) {
        $deviceName = $override.DeviceName
        $instanceId = $override.InstanceId
        
        # Try ALL formats that might have been applied
        $formats = @(
            $deviceName,                        # Format 1: Just device name
            "$deviceName ($instanceId)",         # Format 2: Device name with instance ID
            $instanceId,                         # Format 3: Instance ID alone
            "$instanceId DISPLAY",               # Format 4: Instance ID with override types
            "$deviceName DISPLAY"                # Format 5: Device name with override types
        )
        
        $anyRemoved = $false
        foreach ($format in $formats) {
            try {
                # Remove using all possible formats
                & powercfg /requestsoverride DRIVER "`"$format`"" 2>$null
                $anyRemoved = $true
                Write-Host "      ✅ Removed format: $format" -ForegroundColor DarkGreen
            }
            catch {
                # Ignore errors when removing non-existent overrides
            }
        }
        
        if ($anyRemoved) {
            Write-Host "   ✅ Removed: $deviceName" -ForegroundColor Green
            $removedCount++
        } else {
            Write-Host "   ⚠️  Couldn't remove: $deviceName (no matching overrides found)" -ForegroundColor Yellow
        }
    }
    
    # Clear the history file
    $logFile = "$env:TEMP\USB-Sleep-Fix-History.txt"
    if (Test-Path $logFile) {
        Remove-Item $logFile -Force
    }
    
    Write-Host ""
    if ($removedCount -gt 0) {
        Write-Host "✅ Successfully removed $removedCount device overrides!" -ForegroundColor Green
    }
    Write-Host "💡 Note: Some overrides might remain if they were applied outside this script" -ForegroundColor Cyan
}

function Remove-AllUSBFixes {
    Write-Host "⚠️  NUCLEAR OPTION: Remove ALL USB overrides" -ForegroundColor Red
    Write-Host "   This will remove EVERYTHING, including your manual ones!" -ForegroundColor Yellow
    
    if (-not (Confirm-Action "Are you SURE you want to remove ALL USB overrides?")) {
        return
    }
    
    Write-Host ""
    Write-Host "💥 Removing ALL USB device overrides..." -ForegroundColor Red
    
    # Get all USB devices and try to remove their overrides
    $devices = Get-PnpDevice -PresentOnly | Where-Object { $_.InstanceId -like "USB*" }
    
    foreach ($device in $devices) {
        try {
            & powercfg /requestsoverride DRIVER "$($device.FriendlyName)" 2>$null
            & powercfg /requestsoverride DRIVER "$($device.FriendlyName) ($($device.InstanceId))" 2>$null
        }
        catch {
            # Ignore errors when removing non-existent overrides
        }
    }
    
    # Also clear our history
    $logFile = "$env:TEMP\USB-Sleep-Fix-History.txt"
    if (Test-Path $logFile) {
        Remove-Item $logFile -Force
    }
    
    Write-Host "💥 Attempted to remove all USB overrides!" -ForegroundColor Red
}

function Show-CurrentSleepBlockers {
    Write-Host "🔍 Checking what's currently preventing sleep..." -ForegroundColor Yellow
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
                    Write-Host "🚫 Devices preventing sleep:" -ForegroundColor Red
                    continue
                }
                
                if ($line -match "^[A-Z]+:") {
                    $inSystemSection = $false
                }
                
                if ($inSystemSection -and $line.Trim() -ne "None." -and $line.Trim() -ne "") {
                    if ($line -match "\[DRIVER\]") {
                        $deviceName = ($line -split "\[DRIVER\]")[1].Trim()
                        Write-Host "   • $deviceName" -ForegroundColor Yellow
                        $foundBlockers = $true
                    }
                }
            }
            
            if (-not $foundBlockers) {
                Write-Host "✅ No devices are currently preventing sleep!" -ForegroundColor Green
            }
        } else {
            Write-Host "❌ Couldn't check sleep status." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "❌ Error checking sleep blockers: $($_.Exception.Message)" -ForegroundColor Red
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
            Write-Host "🔊 Audio Device Sleep Fix" -ForegroundColor Green
            Write-Host "This will fix all audio devices from preventing sleep." -ForegroundColor White
            Write-Host "(Including headsets like ASTRO A50, speakers, microphones, etc.)" -ForegroundColor White
            Write-Host ""
            
            $devices = Get-ProblematicDevices -FilterType "audio"
            
            if (Show-DeviceList -devices $devices) {
                if (Confirm-Action "Apply sleep fix to these audio devices?") {
                    Apply-SleepFix -devices $devices
                }
            } else {
                Write-Host "❌ No audio devices found." -ForegroundColor Red
            }
            
            Wait-ForUser
        }
        
        "2" {
            Show-Title
            Write-Host "🔧 All USB Device Sleep Fix" -ForegroundColor Green
            Write-Host "This will fix all USB devices that commonly prevent sleep." -ForegroundColor White
            Write-Host "(Including audio devices, mice, keyboards, etc.)" -ForegroundColor White
            Write-Host ""
            
            $devices = Get-ProblematicDevices -FilterType "all"
            
            if (Show-DeviceList -devices $devices) {
                if (Confirm-Action "Apply sleep fix to ALL these devices?") {
                    Apply-SleepFix -devices $devices
                }
            } else {
                Write-Host "❌ No problematic devices found." -ForegroundColor Red
            }
            
            Wait-ForUser
        }
        
        "3" {
            Show-Title
            Write-Host "🔍 Current Sleep Status" -ForegroundColor Cyan
            Show-CurrentSleepBlockers
            Wait-ForUser
        }
        
        "4" {
            Show-Title
            Write-Host "🗑️  Remove Sleep Fixes" -ForegroundColor Red
            Write-Host ""
            Write-Host "Choose removal option:" -ForegroundColor White
            Write-Host ""
            Write-Host "A) Remove only fixes that THIS SCRIPT added (SAFE)" -ForegroundColor Green
            Write-Host "B) Remove ALL USB overrides (NUCLEAR - removes manual ones too!)" -ForegroundColor Red
            Write-Host "C) Cancel" -ForegroundColor Gray
            Write-Host ""
            Write-Host "Enter choice (A/B/C): " -ForegroundColor White -NoNewline
            
            $removeChoice = Read-Host
            
            switch ($removeChoice.ToUpper()) {
                "A" {
                    Remove-ScriptFixes
                }
                "B" {
                    Remove-AllUSBFixes
                }
                "C" {
                    Write-Host "Cancelled." -ForegroundColor Gray
                }
                default {
                    Write-Host "Invalid choice." -ForegroundColor Red
                }
            }
            
            Wait-ForUser
        }
        
        "5" {
            Show-Title
            Write-Host "👋 Thanks for using USB Sleep Fix Tool!" -ForegroundColor Green
            Write-Host "   Your PC should now sleep properly." -ForegroundColor White
            Write-Host ""
            exit
        }
        
        default {
            Show-Title
            Write-Host "❌ Invalid choice. Please enter 1-5." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
    
} while ($true)