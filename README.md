<p align="center">
  <img src="logo.png" width="128" alt="SnapCherry Logo">
</p>

<h1 align="center">SnapCherry</h1>

<p align="center">
  <strong>High-Performance, Ultra-Lightweight Screenshot tool for Windows.</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Language-Pure%20C-red?style=flat-square" alt="C">
  <img src="https://img.shields.io/badge/Platform-Windows-blue?style=flat-square" alt="Windows">
  <img src="https://img.shields.io/badge/RAM-%3C5MB-orange?style=flat-square" alt="RAM">
  <img src="https://img.shields.io/badge/Binary-~200KB-green?style=flat-square" alt="Binary">
</p>

---

SnapCherry is a minimalist, single-binary screenshot utility built with pure C and the Win32 API. It delivers an instant, Android-inspired editing workflow with zero background bloat and near-zero memory footprint.

## Features

- **Performance**: Native Win32 execution with less than 5MB RAM usage.
- **Instant Editor**: A floating toolbar appears immediately after selection for fast annotations.
- **Drawing Tools**:
  - Precision Pen (5 Dynamic Colors)
  - Intelligent Eraser (restores original pixels)
- **Native PNG Support**: Built-in Windows Imaging Component (WIC) encoding for high-quality PNGs without external libraries.
- **Global Hotkey**: System-wide CTRL + SHIFT + S (with automatic fallback to ALT + SHIFT + S if occupied).
- **Auto-Start**: Integrated registry installer for Windows boot persistence.
- **Fail-Safe**: Right-click to cancel instantly and Mutex protection for single-instance stability.

## Usage

1. **Launch**: Open SnapCherry.exe (it runs silently in the background).
2. **Capture**: Press CTRL + SHIFT + S from anywhere in Windows.
3. **Select**: Drag to select a region, or click once for a full-screen capture.
4. **Annotate**: Use the floating toolbar at the bottom center to draw or erase.
5. **Finalize**: Click **Done** or **Save** to export your capture.

## Build Instructions

Ensure you have **GCC (MinGW)** installed and configured.

```powershell
gcc SnapCherry.c -o SnapCherry.exe -mwindows -luser32 -lgdi32 -lole32 -lwindowscodecs -lmsimg32 -lshlwapi -lshell32 -loleaut32 -luuid
```

## Shortcuts

| Key | Action |
| :-- | :--- |
| `CTRL + SHIFT + S` | Trigger Capture Mode |
| `ALT + SHIFT + S` | Trigger Capture Mode (Fallback) |
| `ESC` | Cancel current capture |
| `Right Click` | Instant Close |

## Save Location

SnapCherry resolves your Windows **Screenshots** folder dynamically (supporting OneDrive integration). Typically found at:
`C:\Users\[User]\Pictures\Screenshots`

---

**SnapCherry** - Minimal. Fast. Pure Red.
