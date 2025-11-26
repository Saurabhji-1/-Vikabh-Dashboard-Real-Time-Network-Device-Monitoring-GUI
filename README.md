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
