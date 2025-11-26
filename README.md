# Vikabh Dashboard

Vikabh Dashboard is a desktop network-monitoring tool built in Python using Tkinter.  
It is designed for IT support and system administrators to monitor devices, check availability (ping/TCP), and quickly launch remote sessions via VNC from a single dashboard.

---

## Features

- **Device management**
  - Add, edit, enable/disable devices
  - Store device name, host/IP, monitoring method (Ping/TCP), port and team/department
  - Teams management with a simple team selector and filter

- **Network monitoring**
  - Periodic ping/TCP checks with configurable interval
  - Per-device online/offline status and latency display
  - Automatic detection of VNC service (port 5900)
  - Manual **“Ping Now”** button for on-demand refresh

- **Monitoring engine**
  - Background monitoring thread using SQLite for persistence
  - Uses a configurable polling interval (1–10 seconds via toolbar, or extended range via Settings)
  - Graceful logging and error handling to avoid UI freezes

- **VNC integration**
  - **Launch TightVNC/RealVNC** directly from the dashboard
  - Launch viewer for a single device (context menu)
  - Launch viewer for **multiple selected devices** from the toolbar

- **Data storage & export**
  - Uses SQLite database (`devices.db`) for devices, settings and teams
  - Stores data in `C:\Users\Administrator\Documents\Vikabh Dashboard` by default on Windows  
    (falls back to current user’s `Documents` or app folder)
  - On-demand export of:
    - `devices.db` (device configuration)
    - `monitor.txt` (summary log derived from `monitor.log`)

- **UI & personalization**
  - Tkinter + ttk-based responsive UI
  - Theme support with multiple palettes (Light, Dark, Teal)
  - Customizable colors via **Personalize** dialog
  - Per-device checkbox selection and **Select All / Unselect All**
  - Status bar updates and logging of key operations

---

## Technology stack

- **Language:** Python 3
- **GUI:** Tkinter / ttk
- **Database:** SQLite
- **OS target:** Windows (primary), should run on Linux/macOS with minor changes
- **Optional:** Pillow (PIL) for generating logo and icon

---

## Project structure

```text
vikabh-dashboard/
├─ src/
│  └─ vikabh_dashboard.py      # main application (single-file app)
├─ assets/
│  ├─ logo.png                 # generated/used by the app (optional)
│  └─ screenshots/             # add a few screenshots here (optional)
├─ README.md
├─ requirements.txt
├─ .gitignore
└─ LICENSE


===============================================================
                Vikabh Dashboard - Installation Guide
===============================================================

This document explains how to run Vikabh Dashboard from source
and how to build a standalone EXE for Windows using PyInstaller.

---------------------------------------------------------------
1. REQUIREMENTS
---------------------------------------------------------------
Operating System : Windows 10 / 11
Python Version   : Python 3.9 or higher recommended

Main Dependencies:
------------------
- Tkinter (comes pre-installed with Python on Windows)
- SQLite (comes built-in with Python)
- Pillow (optional, required only for logo/icon generation)

Install Pillow (Optional):
--------------------------
pip install Pillow

---------------------------------------------------------------
2. HOW TO RUN FROM SOURCE
---------------------------------------------------------------

Step 1: Clone the project

    git clone https://github.com/<your-username>/vikabh-dashboard.git
    cd vikabh-dashboard/src

Step 2: Install optional dependency

    pip install Pillow

Step 3: Run application

    python vikabh_dashboard.py

On first launch, the app will automatically:
- Create required folders and database in Documents
- Generate logo/icon (if Pillow installed)
- Start UI dashboard

---------------------------------------------------------------
3. BUILD WINDOWS EXE (SINGLE FILE)
---------------------------------------------------------------

Requirement:
    pip install pyinstaller

Run this command inside the 'src' folder:

---------------------------------------------------------------
BASH/TERMINAL COMMAND TO GENERATE EXE
---------------------------------------------------------------

    pyinstaller vikabh_dashboard.py --noconsole --onefile --icon=../assets/vd.ico --name "Vikabh Dashboard"

Notes:
- Output executable will be created inside:  dist/
- Copy assets/logo.png or icon manually if needed
- App will still generate folders + database automatically in Documents

To include logo/icon into bundle permanently, use:

    pyinstaller vikabh_dashboard.py --noconsole --onefile --add-data "logo.png;." --icon=../assets/vd.ico --name "Vikabh Dashboard"

---------------------------------------------------------------
4. AFTER BUILDING
---------------------------------------------------------------

Your EXE location:

    dist/Vikabh Dashboard.exe

This file can be used independently on any Windows system with no Python installed.

---------------------------------------------------------------
5. OPTIONAL: CLEAN BUILD FILES
---------------------------------------------------------------

    rmdir /s /q build
    rmdir /s /q __pycache__
    del *.spec


===============================================================
       END OF INSTALLATION + EXE BUILD DOCUMENT
===============================================================


