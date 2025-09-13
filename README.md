# 🏷️ MVA Label Printing Software

A Windows desktop application for generating and printing
Built with **Python**, **Tkinter**, and **ReportLab**, with **silent PDF printing via the portable version of SumatraPDF**.

---

 Prerequisites

- **Windows 10/11**
- **Python 3.10+**
- **Portable SumatraPDF (REQUIRED)**  
  This app prints silently by launching SumatraPDF in the background.  
  You **must** download the **portable** executable (no installer) and place it where the app can find it.

  Put one of these files in your repo’s `assets/` folder:
  - `assets/SumatraPDF.exe` (64-bit portable)
  - `assets/SumatraPDF-32.exe` (32-bit portable)

  The app searches in this order (by default it prefers 32-bit first):
  1. `assets/SumatraPDF-32.exe`
  2. `assets/SumatraPDF.exe`
  3. `C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe`
  4. `C:\Program Files\SumatraPDF\SumatraPDF.exe`

  > Tip: If you want to prefer 64-bit first, set `USE_SUMATRA_32BIT_FIRST = False` in the script.

 Features

- **Preview Panel**
  - Real-time preview using Pillow (falls back gracefully if not installed).
- **Silent Printing (Windows)**
  - Uses **portable SumatraPDF** for zero-UI, background printing.
  - Multiple copies via `"loop"` (most reliable) or `"nx"` mode.
  - Temporary PDFs auto-clean after ~25s.
- **SQLite Database**
  - All labels stored in `labels.db` (auto-created).
  - Unique codes with timestamps.
- **Admin Panel** (password protected)
  - Add single label.
  - Bulk import from pasted lines.
  - View/search/delete labels.
- **Friendly UI**
  - Read-only dropdown, copy selector, `Ctrl+P` shortcut.

## 📦 Installation

1. Clone the repository:
 ```bash
 git clone https://github.com/EvoTheSecond/MVALabelPrinting.git
 cd MVA-Label-Printing
 ```

2. Install Python dependencies:
  ```bash
pip install reportlab pillow
```

3.Run The Program:
```bash
python Material.py
```


