# Navisworks Batch Automation Tool

A robust PowerShell automation tool designed to streamline the creation of Navisworks (`.nwd`) models from various CAD sources. It processes sub-folders, converts files, and assembles a federated Master Model.

## 🚀 Features

* **Local Processing:** Copies files to a local SSD workspace (`C:\Temp`) during processing to avoid network latency and file locking issues.
* **Auto-Detection:** Automatically finds the installed version of Navisworks Manage or Simulate.
* **Master Assembly:** Automatically compiles a `MASTER_MODEL.nwd` from all processed sub-models.
* **Template Support:** Uses a `_TEMPLATE.nwd` to enforce settings, colors, and coordinate systems.
* **Error Handling:** Provides clear color-coded feedback on success/failure and prevents script crashes if files are locked.

## 📂 Project Structure

Recommended folder structure for usage:

```text
Project-Folder/
│
├── Run-NavisBuilder.bat       <-- Run this file
├── _TEMPLATE.nwd              <-- Your project template (optional but recommended)
│
├── src/
│   └── Build-NavisModels.ps1  <-- The main script
│
├── Architecture/              <-- Sub-folder
│   ├── Level1.ifc
│   └── Roof.dwg
│
├── HVAC/                      <-- Sub-folder
│   └── Ventilation.ifc
│
└── ...
```

## 🛠️ Usage

1.  **Clone or Download** this repository.
2.  Place the scripts in your project root (or keep them central and copy them when needed).
3.  Ensure you have a `_TEMPLATE.nwd` in the root if you want to use specific Navisworks settings.
4.  Double-click **`Run-NavisBuilder.bat`**.

The script will:
1.  Scan all sub-folders for `.dwg`, `.ifc`, `.dgn`, and `.nwd` files.
2.  Convert each folder's contents into a single NWD (e.g., `Architecture.nwd`).
3.  Combine all NWDs into `MASTER_MODEL.nwd`.
4.  Place the new NWD files back into their respective folders on the server.

## 📋 Requirements

* Windows 10/11
* Autodesk Navisworks Manage OR Simulate (2020 or newer recommended)
* PowerShell 5.1 (Standard on Windows)