# USB Sleep Fix Tool

A PowerShell utility that fixes USB devices preventing your Windows PC from entering sleep mode.

![Main Menu](https://i.imgur.com/7kVRKNd.png)

---

## ✨ Features

- 🔧 **Fix audio devices** (headsets, speakers, microphones)
- 🔌 **Fix all USB devices** (mice, keyboards, etc.)
- 🔍 **Check current sleep blockers**
- ⚠️ **Safe removal options** (undo changes)
- 🚀 **Persistence mode** (auto-run on startup)

---

## 🚀 Usage

1. **Open run.bat and accept admin privileges**
2. **Select from simple menu options**
3. **Let the tool handle the rest**

> _Ideal for troubleshooting headsets (like ASTRO A50), audio interfaces, and other USB devices that interfere with sleep mode._

## 🔄 Persistent Fix Mode (Run on Startup)
To automatically apply fixes on system startup:
1. Select option 6: "Make Fixes Persistent"
2. The tool will create a scheduled task that:
   - Runs at startup with 3-minute delay
   - Executes in SYSTEM context
   - Logs results to `USB-SleepFix-Persistence.log`

> **Note:** Persistent mode waits 45 seconds for device initialization before applying fixes

## 📂 File Overview
- `USB-SleepFix.ps1` - Main PowerShell script
- `run.bat` - Launcher (Always Run as Administrator)
- `USB-Sleep-Fix-History.txt` - Tracks fixed devices (for safe removal)
- `USB-SleepFix-Persistence.log` - Silent mode operation log

