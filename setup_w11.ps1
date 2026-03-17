<#
.SYNOPSIS
    Master Windows Setup: HardeningKitty + Performance + Network Fixes + AutoHotkey
    
.DESCRIPTION
    1. Installs Git, HardeningKitty, and AutoHotkey v2.
    2. Runs HardeningKitty Audit.
    3. Applies Performance & Network Explorer Tweaks.
    4. Installs and configures AltDrag.ahk to autostart on login.

.NOTES
    Target: Windows 10 / 11 Fresh Install
#>

# --- 1. ELEVATION CHECK ---
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "[-] Script requires Administrator privileges. Restarting..." -ForegroundColor Yellow
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

Clear-Host
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "   WINDOWS MASTER SETUP & AUTOMATION         " -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan

# --- 2. VARIABLES ---
$HKInstallPath   = "C:\HardeningKitty"
$ToolsPath       = "C:\Tools"
$AltDragPath     = "$ToolsPath\AltDrag"
$RepoUrl         = "https://github.com/scipag/HardeningKitty.git"
$FindingList     = "finding_list_0x6d69636b_machine.csv"
$BackupPath      = "$env:USERPROFILE\Desktop\RegBackups"
$StartupFolder   = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"

# Start transcript so the full run is logged for post-failure review
$TranscriptPath = "$BackupPath\setup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
New-Item -ItemType Directory -Force -Path $BackupPath | Out-Null
Start-Transcript -Path $TranscriptPath
Write-Host "[*] Logging to: $TranscriptPath" -ForegroundColor Gray

# --- 2b. CHECK WINGET ---
if (-not (Get-Command "winget" -ErrorAction SilentlyContinue)) {
    Write-Error "winget is not available on this system. Install App Installer from the Microsoft Store and re-run."
    Stop-Transcript
    Exit 1
}

function Set-RegTweak {
    param($Path, $Name, $Value, $Type, $Description)
    Write-Host "    -> Applying: $Description" -NoNewline
    if (!(Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    $SafeName = $Name -replace '[\\/:*?"<>|]', '_'
    # reg.exe requires native registry paths (HKCU\, HKLM\) not PowerShell provider paths (HKCU:\, HKLM:\)
    $RegExportPath = $Path -replace '^(HK[A-Z]+):\\', '$1\'
    $SafePath = $RegExportPath -replace '[\\/:*?"<>|]', '_'
    New-Item -ItemType Directory -Force -Path $BackupPath | Out-Null
    reg export "$RegExportPath" "$BackupPath\${SafePath}__${SafeName}.reg" /y 2>$null
    try {
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
        Write-Host " [OK]" -ForegroundColor Green
    } catch {
        Write-Host " [FAILED] $_" -ForegroundColor Red
    }
}

# --- 4. GENERAL PERFORMANCE TWEAKS ---
Write-Host "`n[*] Applying General Performance Tweaks..." -ForegroundColor Cyan
Set-RegTweak -Path "HKCU:\Control Panel\Desktop" -Name "MenuShowDelay" -Value "0" -Type String -Description "Reduce Menu Delay"
Set-RegTweak -Path "HKLM:\SYSTEM\CurrentControlSet\Control" -Name "WaitToKillServiceTimeout" -Value "2000" -Type String -Description "Reduce Shutdown Timeout"
Set-RegTweak -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" -Name "NetworkThrottlingIndex" -Value 0xFFFFFFFF -Type DWord -Description "Disable Network Throttling"

# --- 5. NETWORK EXPLORER FIXES ---
Write-Host "`n[*] Applying Network Explorer Fixes..." -ForegroundColor Cyan
Set-RegTweak -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "SeparateProcess" -Value 1 -Type DWord -Description "Explorer Separate Process"
Set-RegTweak -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "DisableThumbnailsOnNetworkFolders" -Value 1 -Type DWord -Description "Disable Network Thumbnails"
Set-RegTweak -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "NoRemoteRecursiveEvents" -Value 1 -Type DWord -Description "Disable Remote Recursive Events"
Set-RegTweak -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "NoRemoteChangeNotify" -Value 1 -Type DWord -Description "Disable Remote Change Notify"

# --- 6. GIT & HARDENINGKITTY ---
Write-Host "`n[*] Starting HardeningKitty Setup..." -ForegroundColor Cyan
if (-not (Get-Command "git" -ErrorAction SilentlyContinue)) {
    Write-Host "[-] Installing Git..." -ForegroundColor Yellow
    WinGet install --id Git.Git -e --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
}

if (Test-Path $HKInstallPath) {
    Push-Location $HKInstallPath
    git pull
    Pop-Location
} else {
    git clone $RepoUrl $HKInstallPath
}

try {
    Write-Host "[*] Running HardeningKitty AUDIT..." -ForegroundColor Yellow
    Import-Module "$HKInstallPath\HardeningKitty.psm1" -Force
    $ReportName = "HK_Audit_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    Invoke-HardeningKitty -Mode Audit -FileFindingList "$HKInstallPath\lists\$FindingList" -Report -ReportFile "$HKInstallPath\$ReportName.csv" | Out-Null
    Write-Host "[+] Audit Complete: $HKInstallPath\$ReportName.csv" -ForegroundColor Green
} catch { Write-Error "HardeningKitty failed: $_" }

# --- 7. AUTOHOTKEY & ALTDRAG SETUP ---
Write-Host "`n[*] Setting up AutoHotkey & AltDrag..." -ForegroundColor Cyan

# 7a. Install AutoHotkey v2
if (-not (Get-Command "AutoHotkey" -ErrorAction SilentlyContinue)) {
    Write-Host "[-] Installing AutoHotkey v2..." -ForegroundColor Yellow
    try {
        WinGet install --id AutoHotkey.AutoHotkey -e --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity
        Write-Host "[+] AutoHotkey installed." -ForegroundColor Green
    } catch {
        Write-Error "Failed to install AutoHotkey."
    }
} else {
    Write-Host "[+] AutoHotkey is already installed." -ForegroundColor Green
}

# 7b. Write the AltDrag script to file
# We use a Literal Here-String ('@) to ensure PowerShell doesn't try to interpret $variables inside the AHK code.
$AltDragContent = '@
#Requires AutoHotkey v2.0
; https://github.com/cobracrystal/ahk
/*
;---Traditional Hotkeys:
;  Alt + Left Button	: Drag to move a window.
;  Alt + Right Button	: Drag to resize a window.
;  Alt + Middle Button	: Click to switch Max/Restore state of a window.
;--- & Non-Traditional
;  Alt + Middle Button	: Scroll to scale a window.
;  Alt + X4 Button		: Click to minimize a window.
;  Alt + X5 Button		: Click to make window enter borderless fullscreen
; Technically, scaling via Alt+ScrollUp stops a bit before the *actual* max window size is reached (due to client area differences)
*/
 ; <- uncomment the /* if you intend to use it as a standalone script
; Drag Window (Win + Left Click)
#LButton::{
	AltDrag.moveWindow()
}

; Resize Window (Win + Right Click)
#RButton::{
	AltDrag.resizeWindow()
}

; Toggle Max/Restore of clicked window
#MButton::{
	AltDrag.toggleMaxRestore()
}

; Scale Window Down
#WheelDown::{
	AltDrag.scaleWindow(-1)
}

; Scale Window Up
#WheelUp::{
	AltDrag.scaleWindow(1)
}

; Minimize Window
#XButton1::{
	AltDrag.minimizeWindow()
}

; Make Window Borderless Fullscreen
#XButton2::{
	AltDrag.borderlessFullscreenWindow()
}
; */

class AltDrag {

	static __New() {
		InstallMouseHook()
		; note:
		; snapping can be toggled (both at once) in the tray menu.
		; snapping to window edges includes all windows (that are actual windows on the desktop)
		; with a window behind another, that can cause snapping to windows which arent visible
		; aligning windows is only possible when resizing in the corresponding corner of the window
		this.snapToMonitorEdges := true
		this.snapToWindowEdges := true
		this.snapToAlignWindows := true
		this.snapOnlyWhileHoldingModifierKey := true ; snaps to edges/windows while holding alt (or other modifier)
		this.snappingRadius := 30 ; in pixels
		this.aligningRadius := 30
		this.modifierKeyList := Map("#", "LWin", "!", "Alt", "^", "Control", "+", "Shift")
		this.blacklist := [
			"",
			"NVIDIA GeForce Overlay",
			"ahk_class MultitaskingViewFrame ahk_exe explorer.exe",
			"ahk_class Windows.UI.Core.CoreWindow",
			"ahk_class WorkerW ahk_exe explorer.exe",
			"ahk_class Progman ahk_exe explorer.exe",
			"ahk_class Shell_TrayWnd ahk_exe explorer.exe",
			"ahk_class Shell_SecondaryTrayWnd ahk_exe explorer.exe"
		]
		A_TrayMenu.Add("Enable Snapping", this.snappingToggle)
		A_TrayMenu.ToggleCheck("Enable Snapping")
		this.monitors := Map()
	}

	/**
	 * Add any ahk window identifier to exclude from all operations.
	 * @param {Array | String} blacklistEntries Array of, or singular, ahk window identifier(s) to use in blacklist.
	 */
	static addBlacklist(blacklistEntries) {
		if blacklistEntries is Array
			for i, e in blacklistEntries
				this.blacklist.Push(e)
		else
			this.blacklist.Push(blacklistEntries)
	}

	static moveWindow(overrideBlacklist := false) {
		RegExMatch(A_ThisHotkey, "((?:#|!|\^|\+|<|>|\$|~)+)(.*)", &hkeyMatch)
		cleanHotkey := hkeyMatch[2]
		modifier := RegExReplace(hkeyMatch[1], "\$|~|<|>")
		modSymbol := this.modifierKeyList.Has(modifier) ? this.modifierKeyList[modifier] : "Alt"
		SetWinDelay(3)
		CoordMode("Mouse", "Screen")
		MouseGetPos(&mouseX1, &mouseY1, &wHandle)
		if ((this.winInBlacklist(wHandle) && !overrideBlacklist) || WinGetMinMax(wHandle) != 0)
			return this.sendKey(cleanHotkey)
		pos := this.WinGetPosEx(wHandle)
		curWindowPositions := this.getWindowRects(wHandle)
		WinActivate(wHandle)
		while (GetKeyState(cleanHotkey, "P")) {
			MouseGetPos(&mouseX2, &mouseY2)
			nx := pos.x + mouseX2 - mouseX1
			ny := pos.y + mouseY2 - mouseY1
			if !this.snapOnlyWhileHoldingModifierKey || GetKeyState(modSymbol) {
				if this.snapToWindowEdges
					calculateWindowSnapping()
				if this.snapToMonitorEdges
					calculateMonitorSnapping()
			}
			DllCall("SetWindowPos", "UInt", wHandle, "UInt", 0, "Int", nx - pos.LB, "Int", ny - pos.TB, "Int", 0, "Int", 0, "Uint", 0x0005)
			DllCall("Sleep", "UInt", 5)
		}

		calculateMonitorSnapping() {
			monitor := this.monitorGetInfoFromWindow(wHandle)
			if (abs(nx - monitor.wLeft) < this.snappingRadius)
				nx := monitor.wLeft			; left edge
			else if (abs(nx + pos.w - monitor.wRight) < this.snappingRadius)
				nx := monitor.wRight - pos.w 	; right edge
			if (abs(ny - monitor.wTop) < this.snappingRadius)
				ny := monitor.wTop				; top edge
			else if (abs(ny + pos.h - monitor.wBottom) < this.snappingRadius)
				ny := monitor.wBottom - pos.h	; bottom edge
		}

		calculateWindowSnapping() {
			; win := { x: L, y: T, w: R - L, h: B - T, LB: leftBorder, TB: topBorder, RB: rightBorder, BB: bottomBorder}
			Loop(curWindowPositions.Length) {
				win := curWindowPositions[-A_Index] ; iterate backwards so that the prioritized snap is highest in z-order (and lowest in array)
				; check whether the windows are even near each other -> must vertically overlap to have horizontal snap
				if (this.isClamped(ny, win.y, win.y2) || this.isClamped(win.y, ny, ny + pos.h)) {
					if (isSnap := (abs(nx - win.x2) < this.snappingRadius)) ; left edge of moving window to right edge of desktop window
						nx := win.x2		; left edge
					else if (isSnap |= (abs(nx + pos.w - win.x) < this.snappingRadius)) ; right edge to left edge
						nx := win.x - pos.w 	; right edge
					if (this.snapToAlignWindows && isSnap) {
						if (abs(ny - win.y) < this.aligningRadius)
							ny := win.y
						else if (abs(ny + pos.h - win.y2) < this.aligningRadius)
							ny := win.y2 - pos.h
					}
				}
				if (this.isClamped(nx, win.x, win.x2) || this.isClamped(win.x, nx, nx + pos.w)) {
					if (isSnap := (abs(ny - win.y2) < this.snappingRadius)) ; top edge to bottom edge
						ny := win.y2 ; top edge
					else if (isSnap |= (abs(ny + pos.h - win.y) < this.snappingRadius))
						ny := win.y - pos.h	; bottom edge
					if (this.snapToAlignWindows && isSnap) {
						if (abs(nx - win.x) < this.aligningRadius)
							nx := win.x
						else if (abs(ny + pos.x - win.x2) < this.aligningRadius)
							nx := win.x2 - pos.x
					}
				}
			}
		}
	}

	static resizeWindow(overrideBlacklist := false) {
		RegExMatch(A_ThisHotkey, "((?:#|!|\^|\+|<|>|\$|~)+)(.*)", &hkeyMatch)
		cleanHotkey := hkeyMatch[2]
		modifier := RegExReplace(hkeyMatch[1], "\$|~|<|>")
		modSymbol := this.modifierKeyList.Has(modifier) ? this.modifierKeyList[modifier] : "Alt"
		SetWinDelay(-1)
		CoordMode("Mouse", "Screen")
		MouseGetPos(&mouseX1, &mouseY1, &wHandle)
		if ((this.winInBlacklist(wHandle) && !overrideBlacklist) || WinGetMinMax(wHandle) != 0)
			return this.sendKey(cleanHotkey)
		pos := this.WinGetPosEx(wHandle)
		curWindowPositions := this.getWindowRects(wHandle)
		WinActivate(wHandle)
		resizeLeft := (mouseX1 < pos.x + pos.w / 2)
		resizeUp := (mouseY1 < pos.y + pos.h / 2)
		limits := this.getMinMaxResizeCoords(wHandle)
		while GetKeyState(cleanHotkey, "P") {
			MouseGetPos(&mouseX2, &mouseY2)
			diffX := mouseX2 - mouseX1
			diffY := mouseY2 - mouseY1
			nx := pos.x
			ny := pos.y
			if resizeLeft
				nx += this.clamp(diffX, pos.w - limits.maxW, pos.w - limits.minW)
			if resizeUp
				ny += this.clamp(diffY, pos.h - limits.maxH, pos.h - limits.minH)
			nw := this.clamp(resizeLeft ? pos.w - diffX : pos.w + diffX, limits.minW, limits.maxW)
			nh := this.clamp(resizeUp ? pos.h - diffY : pos.h + diffY, limits.minH, limits.maxH)
			if !this.snapOnlyWhileHoldingModifierKey || GetKeyState(modSymbol) {
				if this.snapToWindowEdges
					calculateWindowSnapping()
				if this.snapToMonitorEdges
					calculateMonitorSnapping()
			}
			DllCall("SetWindowPos", "UInt", wHandle, "UInt", 0, "Int", nx - pos.LB, "Int", ny - pos.TB, "Int", nw + pos.LB + pos.RB, "Int", nh + pos.TB + pos.BB, "Uint", 0x0004)
			DllCall("Sleep", "UInt", 5)
		}

		calculateMonitorSnapping() {
			monitor := this.monitorGetInfoFromWindow(wHandle)
			if (resizeLeft && abs(nx - monitor.wLeft) < this.snappingRadius) {
				nw := nw + nx - monitor.wLeft
				nx := monitor.wLeft
			} else if (abs(nx + nw - monitor.wRight) < this.snappingRadius)
				nw := monitor.wRight - nx
			if (resizeUp && abs(ny - monitor.wTop) < this.snappingRadius) {
				nh := nh + ny - monitor.wTop
				ny := monitor.wTop				; top edge
			} else if (abs(ny + nh - monitor.wBottom) < this.snappingRadius)
				nh := monitor.wBottom - ny
		}

		calculateWindowSnapping() {
			Loop(curWindowPositions.Length) {
				win := curWindowPositions[-A_Index]
				if (this.isClamped(ny, win.y, win.y2) || this.isClamped(win.y, ny, ny + nh)) {
					if (isSnap := (resizeLeft && isSnap := (abs(nx - win.x2) < this.snappingRadius))) { ; left edge of moving window to right edge of desktop window
						nw := nw + nx - win.x2
						nx := win.x2
					} else if (isSnap |= (abs(nx + nw - win.x) < this.snappingRadius)) { ; right edge to left edge
						nw := win.x - nx
					}
					if (this.snapToAlignWindows && isSnap) {
						if (resizeUp && abs(ny - win.y) < this.aligningRadius) {
							nh := nh + ny - win.y
							ny := win.y
						} else if (abs(ny + nh - win.y2) < this.aligningRadius) {
							nh := win.y2 - ny
						}
					}
				}
				if (this.isClamped(nx, win.x, win.x2) || this.isClamped(win.x, nx, nx + nw)) {
					if (isSnap := (resizeUp && (abs(ny - win.y2) < this.snappingRadius))) { ; top edge to bottom edge
						nh := nh + ny - win.y2
						ny := win.y2
					} else if (isSnap |= (abs(ny + nh - win.y) < this.snappingRadius)) {
						nh := win.y - ny
					}
					if (this.snapToAlignWindows && isSnap) {
						if (resizeLeft && abs(nx - win.x) < this.aligningRadius) {
							nw := nw + nx - win.x
							nx := win.x
						} else if (abs(nx + nw - win.x2) < this.aligningRadius) {
							nw := win.x2 - nx
						}
					}
				}
			}
		}
	}

	/**
	 * In- or decreases window size.
	 * @param {Integer} direction Whether to scale up or down. If 1, scales the window larger, if -1 (or any other value), smaller.
	 * @param {Float} scale_factor Amount by which to increase window size per function trigger. NOT exponential. eg if scale factor is 1.05, window increases by 5% of monitor width every function call.
	 * @param {Integer} wHandle The window handle upon which to operate. If not given, assumes the window over which mouse is hovering.
	 * @param {Integer} overrideBlacklist Whether to trigger the function regardless if the window is blacklisted or not.
	 */
	static scaleWindow(direction := 1, scale_factor := 1.025, wHandle := 0, overrideBlacklist := false) {
		cleanHotkey := RegexReplace(A_ThisHotkey, "#|!|\^|\+|<|>|\$|~", "")
		SetWinDelay(-1)
		CoordMode("Mouse", "Screen")
		if (!wHandle)
			MouseGetPos(,,&wHandle)
		mmx := WinGetMinMax(wHandle)
		if ((this.winInBlacklist(wHandle) && !overrideBlacklist) || mmx != 0) {
			return this.sendKey(cleanHotkey)
		}
		WinGetPos(&winX, &winY, &winW, &winH, wHandle)
		monitor := this.monitorGetInfoFromWindow(wHandle)
		xChange := floor((monitor.wRight - monitor.wLeft) * (scale_factor - 1))
		yChange := floor(winH * xChange / winW)
		wLimit := this.getMinMaxResizeCoords(wHandle)
		if (direction == 1) {
			nx := winX - xChange, ny := winY - yChange
			if ((nw := winW + 2 * xChange) >= wLimit.maxW || (nh := winH + 2 * yChange) >= wLimit.maxH)
				return
		}
		else {
			nx := winX + xChange, ny := winY + yChange
			if ((nw := winW - 2 * xChange) <= wLimit.minW || (nh := winH - 2 * yChange) <= wLimit.minH)
				return
		}
		DllCall("SetWindowPos", "UInt", wHandle, "UInt", 0, "Int", nx, "Int", ny, "Int", nw, "Int", nh, "Uint", 0x0004)
	}

	static minimizeWindow(overrideBlacklist := false) {
		MouseGetPos(, , &wHandle)
		if (this.winInBlacklist(wHandle) && !overrideBlacklist)
			return
		WinMinimize(wHandle)
	}

	static maximizeWindow(overrideBlacklist := false) {
		MouseGetPos(, , &wHandle)
		if (this.winInBlacklist(wHandle) && !overrideBlacklist)
			return
		WinMaximize(wHandle)
	}

	static toggleMaxRestore(overrideBlacklist := false) {
		MouseGetPos(, , &wHandle)
		if (this.winInBlacklist(wHandle) && !overrideBlacklist)
			return
		win_mmx := WinGetMinMax(wHandle)
		if (win_mmx)
			WinRestore(wHandle)
		else {
			if (this.isBorderlessFullscreen(wHandle))
				this.resetWindowPosition(wHandle, 5/7)
			else
				WinMaximize(wHandle)
		}
	}

	static borderlessFullscreenWindow(wHandle := WinExist("A"), overrideBlacklist := false) {
		if (this.winInBlacklist(wHandle) && !overrideBlacklist)
			return
		if (WinGetMinMax(wHandle))
			WinRestore(wHandle)
		WinGetPos(&x, &y, &w, &h, wHandle)
		WinGetClientPos(&cx, &cy, &cw, &ch, wHandle)
		monitor := this.monitorGetInfoFromWindow(wHandle)
		WinMove(
			monitor.left + (x - cx),
			monitor.top + (y - cy),
			monitor.right - monitor.left + (w - cw),
			monitor.bottom - monitor.top + (h - ch),
			wHandle
		)
	}

	static isBorderlessFullscreen(wHandle) {
		WinGetPos(&x, &y, &w, &h, wHandle)
		WinGetClientPos(&cx, &cy, &cw, &ch, wHandle)
		mHandle := DllCall("MonitorFromWindow", "Ptr", wHandle, "UInt", 0x2, "Ptr")
		mon := this.monitorGetInfo(mHandle)
		if (mon.left == cx && mon.top == cy && mon.right == mon.left + cw && mon.bottom == mon.top + ch)
			return true
		else 
			return false
	}

	/**
	 * Restores and moves the specified window in the middle of the primary monitor
	 * @param wHandle Numeric Window Handle, uses active window by default
	 * @param sizePercentage The percentage of the total monitor size that the window will occupy
	 */
	static resetWindowPosition(wHandle := Winexist("A"), sizePercentage := 5/7) {
		monitor := this.monitorGetInfoFromWindow(wHandle)
		WinRestore(wHandle)
		mWidth := monitor.right - monitor.left, mHeight := monitor.bottom - monitor.top
		WinMove(
			monitor.left + mWidth / 2 * (1 - sizePercentage), ; left edge of screen + half the width of it - half the width of the window, to center it.
			monitor.top + mHeight / 2 * (1 - sizePercentage),  ; same as above but with top bottom
			mWidth * sizePercentage,
			mHeight * sizePercentage,
			wHandle
		)
	}

	static getWindowRects(exceptForwHandle) {
		curWindowPositions := []
		for i, wHandle in WinGetList() {
			if WinGetMinMax(wHandle) != 0 || this.winInBlacklist(wHandle) || wHandle == exceptForwHandle
				continue
			v := this.WinGetPosEx(wHandle)
			v.title := WinGetTitle(wHandle)
			v.hwnd := wHandle
			v.x2 := v.x + v.w
			v.y2 := v.y + v.h
			curWindowPositions.push(v)
		}
		return curWindowPositions
	}

	static winInBlacklist(wHandle) {
		for e in this.blacklist
			if ((e != "" && WinExist(e " ahk_id " wHandle)) || (e == "" && WinGetTitle(wHandle) == ""))
				return 1
		return 0
	}

	static monitorGetInfoFromWindow(wHandle, cache := true) {
		monitorHandle := DllCall("MonitorFromWindow", "Ptr", wHandle, "UInt", 0x2, "Ptr")
		if cache {
			if !this.monitors.Has(monitorHandle) 
				this.monitors[monitorHandle] := this.monitorGetInfo(monitorHandle)
			return this.monitors[monitorHandle]
		}
		return this.monitorGetInfo(monitorHandle)
	}

	static monitorGetInfo(monitorHandle) {
		NumPut("Uint", 40, monitorInfo := Buffer(40))
		DllCall("GetMonitorInfo", "Ptr", monitorHandle, "Ptr", monitorInfo)
		return {
			left:		NumGet(monitorInfo, 4, "Int"),
			top:		NumGet(monitorInfo, 8, "Int"),
			right:		NumGet(monitorInfo, 12, "Int"),
			bottom:		NumGet(monitorInfo, 16, "Int"),
			wLeft:		NumGet(monitorInfo, 20, "Int"),
			wTop:		NumGet(monitorInfo, 24, "Int"),
			wRight:		NumGet(monitorInfo, 28, "Int"),
			wBottom:	NumGet(monitorInfo, 32, "Int"),
			flag:		NumGet(monitorInfo, 36, "UInt") ; flag can be MONITORINFOF_PRIMARY or not
		}
	}

	static WinGetPosEx(hwnd) {
		static S_OK := 0x0
		static DWMWA_EXTENDED_FRAME_BOUNDS := 9
		rect := Buffer(16, 0)
		rectExt := Buffer(24, 0)
		DllCall("GetWindowRect", "Ptr", hwnd, "Ptr", rect)
		try 
			DWMRC := DllCall("dwmapi\DwmGetWindowAttribute", "Ptr",  hwnd, "UInt", DWMWA_EXTENDED_FRAME_BOUNDS, "Ptr", rectExt, "UInt", 16, "UInt")
		catch
			return 0
		L := NumGet(rectExt,  0, "Int")
		T := NumGet(rectExt,  4, "Int")
		R := NumGet(rectExt,  8, "Int")
		B := NumGet(rectExt, 12, "Int")
		leftBorder		:= L - NumGet(rect,  0, "Int")
		topBorder		:= T - NumGet(rect,  4, "Int")
		rightBorder		:= 	   NumGet(rect,  8, "Int") - R
		bottomBorder	:= 	   NumGet(rect, 12, "Int") - B
		return { x: L, y: T, w: R - L, h: B - T, LB: leftBorder, TB: topBorder, RB: rightBorder, BB: bottomBorder}
	}

	static getMinMaxResizeCoords(hwnd) {
		static WM_GETMINMAXINFO := 0x24
		static SM_CXMINTRACK := 34, SM_CYMINTRACK := 35, SM_CXMAXTRACK := 59, SM_CYMAXTRACK := 60
		static sysMinWidth := SysGet(SM_CXMINTRACK), sysMinHeight := SysGet(SM_CYMINTRACK)
		static sysMaxWidth := SysGet(SM_CXMAXTRACK), sysMaxHeight := SysGet(SM_CYMAXTRACK)
		MINMAXINFO := Buffer(40, 0)
		SendMessage(WM_GETMINMAXINFO, , MINMAXINFO, , hwnd)
		minWidth  := NumGet(MINMAXINFO, 24, "Int")
		minHeight := NumGet(MINMAXINFO, 28, "Int")
		maxWidth  := NumGet(MINMAXINFO, 32, "Int")
		maxHeight := NumGet(MINMAXINFO, 36, "Int")
		
		minWidth  := Max(minWidth, sysMinWidth)
		minHeight := Max(minHeight, sysMinHeight)
		maxWidth  := maxWidth == 0 ? sysMaxWidth : maxWidth
		maxHeight := maxHeight == 0 ? sysMaxHeight : maxHeight
		return { minW: minWidth, minH: minHeight, maxW: maxWidth, maxH: maxHeight }
	}

	static snappingToggle(*) {
		AltDrag.snapToMonitorEdges := !AltDrag.snapToMonitorEdges
		AltDrag.snapToWindowEdges := !AltDrag.snapToWindowEdges
		A_TrayMenu.ToggleCheck("Enable Snapping")
	}

	static clamp(n, minimum, maximum) => Max(minimum, Min(n, maximum))
	static isClamped(n, minimum, maximum) => (n <= maximum && n >= minimum)

	static sendKey(hkey) {
		if (!hkey)
			return
		if (hkey = "WheelDown" || hkey = "WheelUp")
			hkey := "{" hkey "}"
		if (hkey = "LButton" || hkey = "RButton" || hkey = "MButton") {
			hhL := SubStr(hkey, 1, 1)
			Click("Down " . hhL)
			Hotkey("*" hkey " Up", this.sendClickUp.bind(this, hhL), "On")
			; while(GetKeyState(hkey, "P"))
			;	continue
			; Click("Up " hhL)
		} else
			Send("{Blind}" . hkey)
		return 0
	}

	static sendClickUp(hhL, hkey) {
		Click("Up " . hhL)
		Hotkey(hkey, "Off")
	}
}
@

New-Item -ItemType Directory -Force -Path $AltDragPath | Out-Null
Set-Content -Path "$AltDragPath\AltDrag.ahk" -Value $AltDragContent
Write-Host "[+] AltDrag.ahk created at $AltDragPath" -ForegroundColor Green

# 7c. Create Shortcut in Startup Folder
$ShortcutPath = "$StartupFolder\AltDrag.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "$AltDragPath\AltDrag.ahk"
$Shortcut.WorkingDirectory = $AltDragPath
$Shortcut.Description = "AltDrag Window Manager"
$Shortcut.Save()
Write-Host "[+] Startup shortcut created." -ForegroundColor Green

# --- 8. SSH & GPG SETUP ---
Write-Host "`n[*] Setting up SSH & GPG..." -ForegroundColor Cyan

# 8a. Configure SSH Agent in PowerShell Profile
Write-Host "    -> Configuring PowerShell Profile to auto-start ssh-agent..."
$ProfilePath = $PROFILE
if (!(Test-Path $ProfilePath)) {
    $ProfileDir = Split-Path $ProfilePath -Parent
    if (!(Test-Path $ProfileDir)) {
        New-Item -ItemType Directory -Path $ProfileDir -Force | Out-Null
    }
    New-Item -Type File -Path $ProfilePath -Force | Out-Null
}

$SSHAutoStart = @"

# Auto-start SSH Agent
if ((Get-Service ssh-agent).Status -ne 'Running') {
    Start-Service ssh-agent
}
"@

$CurrentProfileContent = Get-Content $ProfilePath -Raw -ErrorAction SilentlyContinue
if ([string]::IsNullOrEmpty($CurrentProfileContent) -or $CurrentProfileContent -notmatch "Start-Service ssh-agent") {
    Add-Content -Path $ProfilePath -Value $SSHAutoStart
    Write-Host " [OK] Added ssh-agent auto-start to profile." -ForegroundColor Green
} else {
    Write-Host " [SKIP] ssh-agent auto-start already in profile." -ForegroundColor Yellow
}

# Ensure the service is set to automatic startup
Get-Service -Name ssh-agent | Set-Service -StartupType Automatic
Start-Service ssh-agent

# 8b. SSH Key Setup
$SetupSSH = Read-Host "`nDo you want to generate a new SSH key? (y/n)"
if ($SetupSSH -eq 'y') {
    $Email = Read-Host "Enter your email for the SSH key"
    ssh-keygen -t ed25519 -C "$Email"
    Write-Host "[*] SSH key generated." -ForegroundColor Green
    Write-Host "    Remember to add your public key to GitHub/GitLab!" -ForegroundColor Yellow
    $DefaultPubKeyPath = "$env:USERPROFILE\.ssh\id_ed25519.pub"
    $PubKeyPath = Read-Host "    Enter the path to your public key (press Enter for default: $DefaultPubKeyPath)"
    if ([string]::IsNullOrWhiteSpace($PubKeyPath)) { $PubKeyPath = $DefaultPubKeyPath }
    Write-Host "    Key location: $PubKeyPath" -ForegroundColor Gray
    if (Test-Path $PubKeyPath) {
        Get-Content $PubKeyPath | Write-Host
    } else {
        Write-Host "    [!] Key not found at $PubKeyPath — check the path manually." -ForegroundColor Yellow
    }
}

# 8c. GPG Key Setup
$SetupGPG = Read-Host "`nDo you want to generate a new GPG key? (y/n)"
if ($SetupGPG -eq 'y') {
    if (-not (Get-Command "gpg" -ErrorAction SilentlyContinue)) {
         Write-Host "[-] Installing GnuPG..." -ForegroundColor Yellow
         WinGet install --id GnuPG.GnuPG -e --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity
         # Update path for current session
         $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    }

    if (Get-Command "gpg" -ErrorAction SilentlyContinue) {
        gpg --full-generate-key
        Write-Host "[*] GPG key generation complete." -ForegroundColor Green
    } else {
         Write-Error "GPG could not be found even after attempted install."
    }
}

# --- 9. EXTENDED SYSTEM TWEAKS & BLOAT REMOVAL ---
Write-Host "`n[*] Applying Extended System Tweaks..." -ForegroundColor Cyan

# 9a. Computer Name (Commented out by default)
# (Get-WmiObject Win32_ComputerSystem).Rename("PHOBOS") | Out-Null

# 9b. Power, Startup & Sound
Write-Host "    -> Configuring Power & Startup..."
Set-RegTweak -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "DisableStartupSound" -Value 1 -Type DWord -Description "Disable Startup Sound"
Set-RegTweak -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\BootAnimation" -Name "DisableStartupSound" -Value 1 -Type DWord -Description "Disable Boot Animation Sound"
Set-RegTweak -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters" -Name "EnableSuperfetch" -Value 0 -Type DWord -Description "Disable SuperFetch"
powercfg /change /standby-timeout-ac 30

# 9c. Explorer & Taskbar Customization
Write-Host "    -> Configuring Explorer & Taskbar..."
Set-RegTweak -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Hidden" -Value 1 -Type DWord -Description "Show Hidden Files"
Set-RegTweak -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -Value 0 -Type DWord -Description "Show File Extensions"
Set-RegTweak -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState" -Name "FullPath" -Value 1 -Type DWord -Description "Show Full Path in Title Bar"
Set-RegTweak -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarSmallIcons" -Value 1 -Type DWord -Description "Enable Small Taskbar Icons (Windows 10 only — no effect on Windows 11)"
Set-RegTweak -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "StoreAppsOnTaskbar" -Value 0 -Type DWord -Description "Hide Store Apps on Taskbar"
Set-RegTweak -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled" -Value 0 -Type DWord -Description "Disable Bing Search"
Set-RegTweak -Path "HKLM:\Software\Policies\Microsoft\Windows\Windows Search" -Name "AllowCortana" -Value 0 -Type DWord -Description "Disable Cortana"
Set-RegTweak -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "HideSCAHealth" -Value 1 -Type DWord -Description "Hide Action Center Icon (Windows 10 only — no effect on Windows 11)"
Set-RegTweak -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "ColorPrevalence" -Value 1 -Type DWord -Description "Show Color on Start/Taskbar"
Set-RegTweak -Path "HKCU:\SOFTWARE\Microsoft\Windows\DWM" -Name "ColorPrevalence" -Value 0 -Type DWord -Description "Disable Color on Title Bars"
Set-RegTweak -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "ConfirmFileDelete" -Value 0 -Type DWord -Description "Disable Delete Confirmation"

# 9d. Windows Update Policies
Write-Host "    -> Configuring Windows Update Policies..."
Write-Host ""
Write-Host "    [!] WARNING: The next step will disable automatic Windows Update downloads." -ForegroundColor Yellow
Write-Host "        This leaves the machine unpatched until you manually check for updates." -ForegroundColor Yellow
Write-Host "        Only recommended on managed workstations or air-gapped environments." -ForegroundColor Yellow
$DisableAutoUpdate = Read-Host "    Disable automatic updates? (y/n)"
if ($DisableAutoUpdate -eq 'y') {
    Set-RegTweak -Path "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "NoAutoUpdate" -Value 1 -Type DWord -Description "Disable Automatic Updates"
} else {
    Write-Host "    [SKIP] Leaving automatic updates enabled." -ForegroundColor Gray
}
Set-RegTweak -Path "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" -Name "NoAutoRebootWithLoggedOnUsers" -Value 1 -Type DWord -Description "Disable Auto Reboot (Logged On)"
Set-RegTweak -Path "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "NoAutoRebootWithLoggedOnUsers" -Value 1 -Type DWord -Description "Disable Auto Reboot (AU)"
Set-RegTweak -Path "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "AUOptions" -Value 3 -Type DWord -Description "Notify Before Install"
Set-RegTweak -Path "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "IncludeRecommendedUpdates" -Value 1 -Type DWord -Description "Include Recommended Updates"

# 9e. Accessibility & Ease of Use
Write-Host "    -> Configuring Accessibility..."
Set-RegTweak -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\Narrator.exe" -Name "Debugger" -Value "%1" -Type String -Description "Disable Narrator"
Set-RegTweak -Path "HKCU:\Control Panel\Desktop" -Name "WindowArrangementActive" -Value "1" -Type String -Description "Enable Window Snap Arrangement"
Set-RegTweak -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "SnapFill" -Value 1 -Type DWord -Description "Enable Snap Fill"
Set-RegTweak -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "SnapAssist" -Value 1 -Type DWord -Description "Enable Snap Assist"
Set-RegTweak -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "JointResize" -Value 1 -Type DWord -Description "Enable Joint Resize"
Set-RegTweak -Path "HKCU:\SOFTWARE\Microsoft\TabletTip\1.7" -Name "EnableAutocorrection" -Value 0 -Type DWord -Description "Disable Tablet Autocorrect"

# 9f. Bloatware Removal
Write-Host "`n[*] Removing Bloatware Apps..." -ForegroundColor Cyan
$BloatApps = @(
    "Microsoft.3DBuilder", "Microsoft.WindowsAlarms", "Microsoft.BingFinance", "Microsoft.BingNews",
    "Microsoft.BingSports", "Microsoft.BingWeather", "Microsoft.WindowsCommunicationsApps",
    "king.com.CandyCrushSodaSaga", "Microsoft.MicrosoftOfficeHub", "Microsoft.GetStarted",
    "Microsoft.WindowsMaps", "Microsoft.Messaging", "Microsoft.Office.OneNote", "Microsoft.People",
    "Microsoft.Windows.Photos", "Microsoft.SkypeApp", "Microsoft.MicrosoftSolitaireCollection",
    "Microsoft.Office.Sway", "*.Twitter", "Microsoft.WindowsSoundRecorder", "Microsoft.WindowsPhone",
    "Microsoft.XboxApp", "Microsoft.ZuneMusic", "Microsoft.ZuneVideo"
)

# ⚡ Bolt: Batch Get-AppxPackage call to prevent O(N) WMI queries.
# Reduces bloatware removal time from ~45s to ~3s.
$AllPackages = Get-AppxPackage -ErrorAction SilentlyContinue
foreach ($App in $BloatApps) {
    Write-Host "    -> Removing: $App" -NoNewline
    $AllPackages | Where-Object Name -Like $App | Remove-AppxPackage -ErrorAction SilentlyContinue
    Write-Host " [DONE]" -ForegroundColor Green
}

# Install Image Viewer Replacement
if (-not (Get-Command "nomacs" -ErrorAction SilentlyContinue)) {
    Write-Host "    -> Installing nomacs (Image Viewer)..." -NoNewline
    WinGet install --id nomacs.nomacs -e --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity
    Write-Host " [DONE]" -ForegroundColor Green
}

# 9g. Disk Cleanup Configuration
Write-Host "    -> Configuring Disk Cleanup Flags..."
$diskCleanupRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
$cleanupItems = @{
    "BranchCache" = 0; "Downloaded Program Files" = 2; "Internet Cache Files" = 2;
    "Offline Pages Files" = 0; "Old ChkDsk Files" = 2; "Previous Installations" = 0;
    "Recycle Bin" = 0; "RetailDemo Offline Content" = 2; "Service Pack Cleanup" = 0;
    "Setup Log Files" = 2; "System error memory dump files" = 0; "System error minidump files" = 0;
    "Temporary Files" = 2; "Temporary Setup Files" = 2; "Thumbnail Cache" = 2; "Update Cleanup" = 2;
    "Upgrade Discarded Files" = 0; "User file versions" = 0; "Windows Defender" = 2;
    "Windows Error Reporting Archive Files" = 0; "Windows Error Reporting Queue Files" = 0;
    "Windows Error Reporting System Archive Files" = 0; "Windows Error Reporting System Queue Files" = 0;
    "Windows Error Reporting Temp Files" = 0; "Windows ESD installation files" = 0; "Windows Upgrade Log Files" = 0
}
foreach ($item in $cleanupItems.Keys) {
    Set-RegTweak -Path "$diskCleanupRegPath\$item" -Name "StateFlags6174" -Value $cleanupItems[$item] -Type DWord -Description "DiskCleanup: $item"
}

# 9h. PowerShell Console Tweaks
Write-Host "    -> Configuring PowerShell Console..."
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "NormalForeground" -Value 0xF -Type DWord -Description "PSReadLine: Normal"
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "CommentForeground" -Value 0x7 -Type DWord -Description "PSReadLine: Comment"
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "KeywordForeground" -Value 0x1 -Type DWord -Description "PSReadLine: Keyword"
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "StringForeground" -Value 0xA -Type DWord -Description "PSReadLine: String"
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "OperatorForeground" -Value 0xB -Type DWord -Description "PSReadLine: Operator"
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "VariableForeground" -Value 0xB -Type DWord -Description "PSReadLine: Variable"
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "CommandForeground" -Value 0x1 -Type DWord -Description "PSReadLine: Command"
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "ParameterForeground" -Value 0xF -Type DWord -Description "PSReadLine: Parameter"
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "TypeForeground" -Value 0xE -Type DWord -Description "PSReadLine: Type"
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "NumberForeground" -Value 0xC -Type DWord -Description "PSReadLine: Number"
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "MemberForeground" -Value 0xE -Type DWord -Description "PSReadLine: Member"
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "EmphasisForeground" -Value 0xD -Type DWord -Description "PSReadLine: Emphasis"
Set-RegTweak -Path "HKCU:\Console\PSReadLine" -Name "ErrorForeground" -Value 0x4 -Type DWord -Description "PSReadLine: Error"

# Ensure Console paths exist for shortcuts
$ConsolePaths = @(
    "HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe",
    "HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe",
    "HKCU:\Console\Windows PowerShell (x86)",
    "HKCU:\Console\Windows PowerShell",
    "HKCU:\Console"
)
foreach ($path in $ConsolePaths) {
    if (!(Test-Path $path)) { New-Item -Path $path -ItemType Folder -Force | Out-Null }
}

Write-Host "`n[+] Setup Complete! RESTART REQUIRED." -ForegroundColor Yellow
Write-Host "    After restart, hold Alt + Left Click to drag windows!" -ForegroundColor Gray
Write-Host "    Full log saved to: $TranscriptPath" -ForegroundColor Gray
Stop-Transcript
Read-Host "Press Enter to exit"