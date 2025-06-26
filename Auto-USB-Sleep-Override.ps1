param(
    [switch]$silentFixAll
)

$OutputEncoding = [System.Text.Encoding]::UTF8

# Robust script directory detection
$scriptPath = $MyInvocation.MyCommand.Path
if (-not $scriptPath) { $scriptPath = $PSCommandPath }
if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Definition }
$scriptDir = Split-Path -Parent $scriptPath

if (-not (Test-Path $scriptDir)) {
    New-Item -ItemType Directory -Path $scriptDir -Force | Out-Null
}

# --- SILENT MODE (for Scheduled Task) ---
if ($silentFixAll) {
    # Create detailed log file in script directory
    $logPath = "$scriptDir\USB-SleepFix-Persistence.log"
    "===== Silent Mode Started at $(Get-Date) =====" | Out-File $logPath
    "System Uptime: $((Get-CimInstance Win32_OperatingSystem).LastBootUpTime)" | Out-File $logPath -Append
    
    try {
        # Wait for device initialization
        "Step 0: Waiting for system readiness" | Out-File $logPath -Append
        Start-Sleep -Seconds 45
        
        # Enhanced device verification
        "Step 0.5: Verifying device initialization" | Out-File $logPath -Append
        $audioDevices = Get-PnpDevice -PresentOnly | Where-Object { 
            $_.InstanceId -like "USB*" -and $_.Status -eq "OK" -and 
            ($_.FriendlyName -match "(Audio|Sound|Headset|Speaker|Microphone)" -or $_.Class -in @("AudioEndpoint", "MEDIA"))
        }
        $hidDevices = Get-PnpDevice -PresentOnly | Where-Object { 
            $_.InstanceId -like "USB*" -and $_.Status -eq "OK" -and $_.Class -eq "HIDClass"
        }
        "Device check: Audio devices: $($audioDevices.Count), HID devices: $($hidDevices.Count)" | Out-File $logPath -Append
        
        "Step 1: Removing existing script fixes" | Out-File $logPath -Append
        $historyFile = "$scriptDir\USB-Sleep-Fix-History.txt"
        $removedCount = 0
        
        if (Test-Path $historyFile) {
            $scriptOverrides = Get-Content $historyFile
            "Found $($scriptOverrides.Count) existing overrides to remove" | Out-File $logPath -Append
            
            foreach ($entry in $scriptOverrides) {
                $parts = $entry -split '\|'
                if ($parts.Count -eq 3) { 
                    $deviceName = $parts[1]
                    $instanceId = $parts[2]
                    "  Removing: $deviceName" | Out-File $logPath -Append
                    
                    # Remove using different approaches with proper error capture
                    try {
                        $cmd1 = "powercfg /requestsoverride DRIVER `"$deviceName`""
                        $result1 = Invoke-Expression $cmd1 2>&1
                        $exitCode1 = $LASTEXITCODE
                        "    Removal output 1: $result1 (Exit: $exitCode1)" | Out-File $logPath -Append
                        
                        $cmd2 = "powercfg /requestsoverride DRIVER `"$deviceName ($instanceId)`""
                        $result2 = Invoke-Expression $cmd2 2>&1
                        $exitCode2 = $LASTEXITCODE
                        "    Removal output 2: $result2 (Exit: $exitCode2)" | Out-File $logPath -Append
                        
                        $cmd3 = "powercfg /requestsoverride DRIVER `"$instanceId`""
                        $result3 = Invoke-Expression $cmd3 2>&1
                        $exitCode3 = $LASTEXITCODE
                        "    Removal output 3: $result3 (Exit: $exitCode3)" | Out-File $logPath -Append
                        
                        $removedCount++
                    }
                    catch {
                        "    Removal error: $($_.Exception.Message)" | Out-File $logPath -Append
                    }
                }
            }
            "Removed $removedCount existing overrides" | Out-File $logPath -Append
        } else {
            "No existing fixes found" | Out-File $logPath -Append
        }
        
        # Reset power configuration
        "Resetting power configuration..." | Out-File $logPath -Append
        try {
            $resetResult = Invoke-Expression "powercfg /restoredefaultschemes" 2>&1
            "Reset result: $resetResult" | Out-File $logPath -Append
        }
        catch {
            "Reset error: $($_.Exception.Message)" | Out-File $logPath -Append
        }
        Start-Sleep -Seconds 5
        
        # Step 2: Fix ALL USB devices
        "Step 2: Fixing all USB devices" | Out-File $logPath -Append
        $devicesToFix = Get-PnpDevice -PresentOnly | Where-Object { 
            $_.InstanceId -like "USB*" -and $_.Status -eq "OK" -and 
            ($_.FriendlyName -match "(Audio|Sound|Headset|Speaker|Microphone|Mouse|Keyboard)" -or 
            $_.Class -in @("AudioEndpoint", "MEDIA", "HIDClass"))
        }
        
        "Found $($devicesToFix.Count) USB devices to fix" | Out-File $logPath -Append
        
        # Track processed devices and results
        $processedDevices = @()
        $successCount = 0
        $newHistoryFile = "$scriptDir\USB-Sleep-Fix-History.txt"
        
        # Clear the history file before adding new entries
        if (Test-Path $newHistoryFile) {
            Remove-Item $newHistoryFile -Force
        }
        
        foreach ($device in $devicesToFix) {
            $deviceName = $device.FriendlyName
            $instanceId = $device.InstanceId
            
            if ($processedDevices -contains $instanceId) {
                "  Skipping duplicate: $deviceName" | Out-File $logPath -Append
                continue
            }
            
            $processedDevices += $instanceId
            
            "  Fixing: $deviceName" | Out-File $logPath -Append
            
            # Apply overrides using multiple methods and track success
            $anySuccess = $false
            
            try {
                # Method 1: Use device name
                $cmd1 = "powercfg /requestsoverride DRIVER `"$deviceName`" DISPLAY SYSTEM AWAYMODE EXECUTION"
                $result1 = Invoke-Expression $cmd1 2>&1
                $exitCode1 = $LASTEXITCODE
                "    Command output 1: $result1 (Exit: $exitCode1)" | Out-File $logPath -Append
                if ($exitCode1 -eq 0) { $anySuccess = $true }
                
                # Method 2: Use device name with instance ID
                $cmd2 = "powercfg /requestsoverride DRIVER `"$deviceName ($instanceId)`" DISPLAY SYSTEM AWAYMODE EXECUTION"
                $result2 = Invoke-Expression $cmd2 2>&1
                $exitCode2 = $LASTEXITCODE
                "    Command output 2: $result2 (Exit: $exitCode2)" | Out-File $logPath -Append
                if ($exitCode2 -eq 0) { $anySuccess = $true }
                
                # Method 3: Use instance ID only
                $cmd3 = "powercfg /requestsoverride DRIVER `"$instanceId`" DISPLAY SYSTEM AWAYMODE EXECUTION"
                $result3 = Invoke-Expression $cmd3 2>&1
                $exitCode3 = $LASTEXITCODE
                "    Command output 3: $result3 (Exit: $exitCode3)" | Out-File $logPath -Append
                if ($exitCode3 -eq 0) { $anySuccess = $true }
                
                if ($anySuccess) {
                    "    [SUCCESS] At least one override method succeeded" | Out-File $logPath -Append
                    $successCount++
                    # Log for history
                    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')|$deviceName|$instanceId" | Add-Content -Path $newHistoryFile
                } else {
                    "    [ERROR] All override methods failed" | Out-File $logPath -Append
                }
            }
            catch {
                "    [EXCEPTION] Error processing device: $($_.Exception.Message)" | Out-File $logPath -Append
            }
        }
        
        # Verify applied overrides
        "Step 3: Verifying applied overrides" | Out-File $logPath -Append
        try {
            $overrideResult = Invoke-Expression "powercfg /requestsoverride" 2>&1
            $currentOverrides = $overrideResult | Where-Object { $_ -match "\[DRIVER\]" }
            "Current overrides count: $($currentOverrides.Count)" | Out-File $logPath -Append
            if ($currentOverrides.Count -gt 0) {
                $currentOverrides | Out-File $logPath -Append
            } else {
                "No DRIVER overrides found" | Out-File $logPath -Append
            }
        }
        catch {
            "Error verifying overrides: $($_.Exception.Message)" | Out-File $logPath -Append
        }
        
        "Successfully processed $successCount of $($processedDevices.Count) devices" | Out-File $logPath -Append
        "===== Silent Mode Completed Successfully =====" | Out-File $logPath -Append
    }
    catch {
        "ERROR: $($_.Exception.Message)" | Out-File $logPath -Append
        "ERROR DETAILS: $($_.ScriptStackTrace)" | Out-File $logPath -Append
        "===== Silent Mode Failed =====" | Out-File $logPath -Append
    }
    
    exit
}

# --- INTERACTIVE MODE ---

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host ""
    Write-Host "[ATTENTION] IMPORTANT: You need to run this as Administrator!" -ForegroundColor Red
    Write-Host "            Right-click the script and select 'Run with PowerShell'" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# --- SCRIPT FUNCTIONS ---

function Show-Title {
    Clear-Host
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "                USB SLEEP FIX TOOL                " -ForegroundColor Cyan
    Write-Host "      Fix USB devices preventing system sleep" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Show-Menu {
    Write-Host "MAIN MENU" -ForegroundColor White
    Write-Host "--------------------------------------------------" -ForegroundColor DarkCyan
    Write-Host "--- APPLY FIXES ---" -ForegroundColor DarkCyan
    Write-Host "1. Fix all audio devices" -ForegroundColor Green
    Write-Host "2. Fix all USB devices (Recommended)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "--- DIAGNOSTICS & REMOVAL ---" -ForegroundColor DarkCyan
    Write-Host "3. See what's currently preventing sleep" -ForegroundColor Cyan
    Write-Host "4. Show all existing fixes (request overrides)" -ForegroundColor Magenta
    Write-Host "5. Remove fixes applied by this script" -ForegroundColor Red
    Write-Host ""
    Write-Host "--- PERSISTENCE (RUNS FIX ON STARTUP) ---" -ForegroundColor DarkCyan
    $taskExists = Get-ScheduledTask -TaskName "UsbSleepFix" -ErrorAction SilentlyContinue
    if ($taskExists) {
        Write-Host "6. [ACTIVE] Remove Persistent Fix (Disable running on startup)" -ForegroundColor DarkRed
    } else {
        Write-Host "6. Make Fixes Persistent (Run on startup)" -ForegroundColor DarkGreen
    }
    Write-Host ""
    Write-Host "7. Exit" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Enter your choice (1-7): " -ForegroundColor White -NoNewline
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
        Write-Host "[ERROR] No devices found matching the criteria!" -ForegroundColor Red
        return $false
    }
    
    Write-Host ""
    Write-Host "Found $($devices.Count) device(s) to ${action}:" -ForegroundColor Green
    Write-Host "--------------------------------------------------"
    
    foreach ($device in $devices) {
        $driverDate = (Get-PnpDeviceProperty -InstanceId $device.InstanceId -KeyName DEVPKEY_Device_DriverDate).Data
        $driverProvider = (Get-PnpDeviceProperty -InstanceId $device.InstanceId -KeyName DEVPKEY_Device_DriverProvider).Data
        Write-Host "  + $($device.FriendlyName)" -ForegroundColor White
        Write-Host "    (Provider: $driverProvider, Date: $driverDate)" -ForegroundColor Gray
    }
    Write-Host "--------------------------------------------------"
    
    return $true
}

function Apply-SleepFix {
    param($devices)
    
    Write-Host ""
    Write-Host "[APPLYING] Applying sleep fixes..." -ForegroundColor Yellow
    
    # Reset power configuration first
    Write-Host "  [PREP] Resetting power configuration..." -ForegroundColor Cyan
    try {
        $resetResult = Invoke-Expression "powercfg /restoredefaultschemes" 2>&1
        Write-Host "    Reset result: $resetResult" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "    Reset error: $($_.Exception.Message)" -ForegroundColor Red
    }
    Start-Sleep -Seconds 2
    
    $successCount = 0
    $historyFile = "$scriptDir\USB-Sleep-Fix-History.txt"
    
    # Clear existing history
    if (Test-Path $historyFile) {
        Remove-Item $historyFile -Force
    }
    
    foreach ($device in $devices) {
        $deviceName = $device.FriendlyName
        $instanceId = $device.InstanceId
        
        Write-Host "  [FIXING] Processing: $deviceName" -ForegroundColor Cyan
        
        # Try multiple override methods
        $anySuccess = $false
        
        try {
            # Method 1: Use device name
            $cmd1 = "powercfg /requestsoverride DRIVER `"$deviceName`" DISPLAY SYSTEM AWAYMODE EXECUTION"
            Write-Host "    Trying method 1: Device name" -ForegroundColor DarkGray
            $result1 = Invoke-Expression $cmd1 2>&1
            $exitCode1 = $LASTEXITCODE
            if ($exitCode1 -eq 0) { 
                Write-Host "      [SUCCESS] Method 1 worked" -ForegroundColor Green
                $anySuccess = $true 
            } else {
                Write-Host "      [FAILED] Method 1 failed (Exit: $exitCode1)" -ForegroundColor DarkYellow
            }
            
            # Method 2: Use device name with instance ID
            $cmd2 = "powercfg /requestsoverride DRIVER `"$deviceName ($instanceId)`" DISPLAY SYSTEM AWAYMODE EXECUTION"
            Write-Host "    Trying method 2: Device name + Instance ID" -ForegroundColor DarkGray
            $result2 = Invoke-Expression $cmd2 2>&1
            $exitCode2 = $LASTEXITCODE
            if ($exitCode2 -eq 0) { 
                Write-Host "      [SUCCESS] Method 2 worked" -ForegroundColor Green
                $anySuccess = $true 
            } else {
                Write-Host "      [FAILED] Method 2 failed (Exit: $exitCode2)" -ForegroundColor DarkYellow
            }
            
            # Method 3: Use instance ID only
            $cmd3 = "powercfg /requestsoverride DRIVER `"$instanceId`" DISPLAY SYSTEM AWAYMODE EXECUTION"
            Write-Host "    Trying method 3: Instance ID only" -ForegroundColor DarkGray
            $result3 = Invoke-Expression $cmd3 2>&1
            $exitCode3 = $LASTEXITCODE
            if ($exitCode3 -eq 0) { 
                Write-Host "      [SUCCESS] Method 3 worked" -ForegroundColor Green
                $anySuccess = $true 
            } else {
                Write-Host "      [FAILED] Method 3 failed (Exit: $exitCode3)" -ForegroundColor DarkYellow
            }
            
            if ($anySuccess) {
                Write-Host "      [OVERALL SUCCESS] Device override applied" -ForegroundColor Green
                # Log for history
                "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')|$deviceName|$instanceId" | Add-Content -Path $historyFile
                $successCount++
            } else {
                Write-Host "      [OVERALL FAILED] All methods failed for this device" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "      [EXCEPTION] Error: $($_.Exception.Message)" -ForegroundColor Red
        }
        Write-Host ""
    }
    
    # Final verification
    Write-Host "[VERIFY] Checking applied overrides..." -ForegroundColor Cyan
    try {
        $overrideResult = Invoke-Expression "powercfg /requestsoverride" 2>&1
        $currentOverrides = $overrideResult | Where-Object { $_ -match "\[DRIVER\]" }
        if ($currentOverrides.Count -gt 0) {
            Write-Host "  Current overrides:" -ForegroundColor Gray
            $currentOverrides | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
        } else {
            Write-Host "  [WARNING] No overrides detected!" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "  [ERROR] Could not verify overrides: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host ""
    if ($successCount -eq $devices.Count) {
        Write-Host "[SUCCESS] All devices fixed!" -ForegroundColor Green
    } elseif ($successCount -gt 0) {
        Write-Host "[WARNING] Partial success: $successCount out of $($devices.Count) devices fixed." -ForegroundColor Yellow
    } else {
        Write-Host "[ERROR] No devices could be fixed." -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "[TIP] You can now make this fix persistent using Option 6 from the main menu." -ForegroundColor Cyan
}

function Get-ScriptOverrideHistory {
    $logFile = "$scriptDir\USB-Sleep-Fix-History.txt"
    if (Test-Path $logFile) {
        return Get-Content $logFile | ForEach-Object {
            $parts = $_ -split '\|'
            if ($parts.Count -eq 3) { @{ Timestamp = $parts[0]; DeviceName = $parts[1]; InstanceId = $parts[2] } }
        }
    }
    return @()
}

function Remove-ScriptFixes {
    Write-Host "[INFO] Looking for fixes that THIS SCRIPT added..." -ForegroundColor Yellow
    $scriptOverrides = Get-ScriptOverrideHistory
    
    if ($scriptOverrides.Count -eq 0) {
        Write-Host "[INFO] No fixes from this script found to remove." -ForegroundColor Cyan
        return
    }
    
    Write-Host "[INFO] Found $($scriptOverrides.Count) override(s) that this script added:" -ForegroundColor Cyan
    foreach ($override in $scriptOverrides) {
        Write-Host "  - $($override.DeviceName) (added: $($override.Timestamp))"
    }
    
    if (-not (Confirm-Action "Do you want to proceed with removing these fixes?")) { return }
    
    Write-Host ""
    Write-Host "[REMOVE] Removing script overrides..." -ForegroundColor Yellow
    
    # Reset power configuration first
    Write-Host "  [PREP] Resetting power configuration..." -ForegroundColor Cyan
    try {
        $resetResult = Invoke-Expression "powercfg /restoredefaultschemes" 2>&1
        Write-Host "    Reset result: $resetResult" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "    Reset error: $($_.Exception.Message)" -ForegroundColor Red
    }
    Start-Sleep -Seconds 2
    
    foreach ($override in $scriptOverrides) {
        $deviceName = $override.DeviceName
        $instanceId = $override.InstanceId
        
        Write-Host "  Removing: $deviceName" -ForegroundColor DarkYellow
        
        # Try multiple removal methods
        try {
            $cmd1 = "powercfg /requestsoverride DRIVER `"$deviceName`""
            $result1 = Invoke-Expression $cmd1 2>&1
            $exitCode1 = $LASTEXITCODE
            Write-Host "    Method 1 result: $result1 (Exit: $exitCode1)" -ForegroundColor DarkGray
            
            $cmd2 = "powercfg /requestsoverride DRIVER `"$deviceName ($instanceId)`""
            $result2 = Invoke-Expression $cmd2 2>&1
            $exitCode2 = $LASTEXITCODE
            Write-Host "    Method 2 result: $result2 (Exit: $exitCode2)" -ForegroundColor DarkGray
            
            $cmd3 = "powercfg /requestsoverride DRIVER `"$instanceId`""
            $result3 = Invoke-Expression $cmd3 2>&1
            $exitCode3 = $LASTEXITCODE
            Write-Host "    Method 3 result: $result3 (Exit: $exitCode3)" -ForegroundColor DarkGray
            
            if ($exitCode1 -eq 0 -or $exitCode2 -eq 0 -or $exitCode3 -eq 0) {
                Write-Host "      [SUCCESS] Override removed" -ForegroundColor Green
            } else {
                Write-Host "      [WARNING] All removal methods failed" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "      [ERROR] Exception during removal: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # Remove history file after removal attempt
    $historyFile = "$scriptDir\USB-Sleep-Fix-History.txt"
    if (Test-Path $historyFile) {
        Remove-Item $historyFile -Force -ErrorAction SilentlyContinue
        Write-Host ""
        Write-Host "[SUCCESS] History cleared." -ForegroundColor Green
    }
    
    # Final verification
    Write-Host "[VERIFY] Checking remaining overrides..." -ForegroundColor Cyan
    try {
        $overrideResult = Invoke-Expression "powercfg /requestsoverride" 2>&1
        $currentOverrides = $overrideResult | Where-Object { $_ -match "\[DRIVER\]" }
        if ($currentOverrides.Count -gt 0) {
            Write-Host "  Remaining overrides:" -ForegroundColor Gray
            $currentOverrides | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
        } else {
            Write-Host "  [SUCCESS] No overrides remain" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  [ERROR] Could not verify remaining overrides: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-CurrentSleepBlockers {
    Write-Host "[INFO] Checking what's currently preventing sleep..." -ForegroundColor Yellow
    Write-Host ""
    
    try {
        $requests = powercfg -requests 2>$null
        
        if ($requests) {
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
        foreach ($line in $overrides) {
            if ($line -match "\[DRIVER\]") { $inDriverSection = $true; continue }
            if ($inDriverSection -and -not [string]::IsNullOrWhiteSpace($line)) { $driverOverrides += $line.Trim() }
            if ($inDriverSection -and [string]::IsNullOrWhiteSpace($line)) { break }
        }

        if ($driverOverrides.Count -gt 0) {
            Write-Host "[OVERRIDES FOUND]" -ForegroundColor Green
            foreach ($override in $driverOverrides) { Write-Host "  - $override" -ForegroundColor White }
        } else {
            Write-Host "[OK] No DRIVER power request overrides are currently set." -ForegroundColor Green
        }
    } catch {
        Write-Host "[ERROR] Could not fetch power request overrides: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Manage-Persistence {
    $taskName = "UsbSleepFix"
    $taskExists = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    if ($taskExists) {
        # --- REMOVE PERSISTENCE ---
        Write-Host "[INFO] A scheduled task for persistence is currently ACTIVE." -ForegroundColor Yellow
        if (Confirm-Action "Do you want to remove it to disable the automatic startup fix?") {
            try {
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                Write-Host "[SUCCESS] The persistent fix task has been removed." -ForegroundColor Green
            }
            catch {
                Write-Host "[ERROR] Failed to remove task: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        # --- ADD PERSISTENCE ---
        Write-Host "[INFO] This will create a SYSTEM task that runs on startup" -ForegroundColor Yellow
        Write-Host "       with proper timing to ensure devices are ready" -ForegroundColor Gray
        Write-Host ""
        
        # Get script path - use precomputed path
        if (-not $scriptPath) {
            Write-Host "[ERROR] Could not determine script path!" -ForegroundColor Red
            Write-Host "        Please save this script to a permanent location." -ForegroundColor Yellow
            return
        }
        
        Write-Host "Script path: $scriptPath" -ForegroundColor Cyan
        
        if (Confirm-Action "Do you want to create the persistent fix task?") {
            try {
                # Create the action
                $action = New-ScheduledTaskAction -Execute 'powershell.exe' `
                    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"& { & '$scriptPath' -silentFixAll }`""
                
                # Create the trigger (At startup with sufficient delay)
                $trigger = New-ScheduledTaskTrigger -AtStartup
                $trigger.Delay = "PT3M"  # 3 minute delay after startup
                
                # Create settings
                $settings = New-ScheduledTaskSettingsSet `
                    -AllowStartIfOnBatteries `
                    -DontStopIfGoingOnBatteries `
                    -StartWhenAvailable `
                    -RunOnlyIfNetworkAvailable
                
                # Create the principal (run as SYSTEM)
                $principal = New-ScheduledTaskPrincipal `
                    -UserId "NT AUTHORITY\SYSTEM" `
                    -LogonType ServiceAccount `
                    -RunLevel Highest
                
                # Register the task
                $task = Register-ScheduledTask `
                    -TaskName $taskName `
                    -Action $action `
                    -Trigger $trigger `
                    -Principal $principal `
                    -Settings $settings `
                    -Description "Applies USB sleep fixes at startup with proper timing" `
                    -Force
                
                Write-Host "[SUCCESS] Persistent fix task created" -ForegroundColor Green
                Write-Host "           Task Name: UsbSleepFix" -ForegroundColor Cyan
                Write-Host "           Run As: SYSTEM" -ForegroundColor Cyan
                Write-Host "           Trigger: At startup with 3 min delay" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "Log file location: $scriptDir\USB-SleepFix-Persistence.log" -ForegroundColor Cyan
            }
            catch {
                Write-Host "[ERROR] Failed to create task: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "        Make sure you're running as Administrator" -ForegroundColor Yellow
            }
        }
    }
}

function Confirm-Action {
    param([string]$message)
    Write-Host ""
    Write-Host "$message (Y/N): " -ForegroundColor Yellow -NoNewline
    do {
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        $response = $key.Character.ToString().ToUpper()
    } while ($response -notin @('Y', 'N'))
    Write-Host $response
    return $response -eq 'Y'
}

function Wait-ForUser {
    Write-Host ""
    Write-Host "Press any key to return to the menu..." -ForegroundColor Gray
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
            $devices = Get-ProblematicDevices -FilterType "audio"
            if (Show-DeviceList -devices $devices) {
                if (Confirm-Action "Apply sleep fix to these audio devices?") { Apply-SleepFix -devices $devices }
            }
            Wait-ForUser
        }
        "2" {
            Show-Title
            Write-Host "[OPTION 2] All USB Device Sleep Fix" -ForegroundColor Green
            $devices = Get-ProblematicDevices -FilterType "all"
            if (Show-DeviceList -devices $devices) {
                if (Confirm-Action "Apply sleep fix to ALL these devices?") { Apply-SleepFix -devices $devices }
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
            Remove-ScriptFixes
            Wait-ForUser
        }
        "6" {
            Show-Title
            Write-Host "[OPTION 6] Manage Fix Persistence" -ForegroundColor DarkCyan
            Manage-Persistence
            Wait-ForUser
        }
        "7" {
            Show-Title
            Write-Host "[EXIT] Goodbye!" -ForegroundColor Green
            exit
        }
        default {
            Write-Host " [ERROR] Invalid choice. Please enter 1-7." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
    
} while ($true)