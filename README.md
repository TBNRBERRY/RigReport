# RigReport

**RigReport** is a standalone, single-file PowerShell utility designed for system administrators and power users who need hardware details fast. No installation, no dependencies—just run and retrieve.

![License](https://img.shields.io/github/license/TBNRBERRY/RigReport?color=blue)
![Powershell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)

## ✨ Features

- **Dynamic UI:** Automatic column resizing ensures long GPU or Motherboard names are never cut off.
- **Smart RAM Detection:** Reports total capacity, manufacturer, speed, and active channel configuration (Single/Dual/Quad).
- **Dark Mode Default:** Sleek, modern interface designed for Windows 10/11.
- **One-Click Export:** Copy a clean report to your clipboard or save it as a formatted `.txt` file.
- **Portable:** Zero installation required. Does not leave files behind.

## 🚀 How to Run

1. Download the `RigReport.ps1` file.
2. Right-click the file and select **Run with PowerShell**.
3. *Note: If prompted about Execution Policy, run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` in your terminal first.*

## 📸 Screenshots
<img src="https://github.com/TBNRBERRY/RigReport/blob/main/Screenshots/Dark%20Mode.png" width="546" height="548" /> <img src="https://github.com/TBNRBERRY/RigReport/blob/main/Screenshots/Light%20Mode.png" width="546" height="548" />

## 🛠️ Data Points Captured
- **Operating System** & Computer Name
- **Processor** (Full Name & Model)
- **Memory** (GB, Speed, Slots used, and Channel Mode)
- **Graphics Card(s)** (Supports multi-GPU setups)
- **Storage** (Available space on C: drive)
- **Motherboard** (Manufacturer, Model, and Revision)
- **BIOS** (Version, Release Date, and Boot Mode)

## 📄 License
Distributed under the MIT License. See `LICENSE` for more information.
