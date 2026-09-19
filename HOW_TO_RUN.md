# How to Run NMDC Document Index in Microsoft Excel

Welcome to the **NMDC Document Index** tool. This workbook consolidates, normalizes, and filters project deliverable registers across all methods and technical registers.

---

## 🚀 Quick Start in 3 Easy Steps

### Step 1: Open the Workbook
- Double-click **`NMDC_Document_Index.xlsm`** in this folder.
- If Microsoft Excel shows a yellow security banner asking to **Enable Content** or **Enable Macros**, click **"Enable Content"**.

### Step 2: Verify the Data Folder
- Go to the **`Configuration`** worksheet.
- Confirm that the **Data Folder** path points to your `DATA` folder (by default, it automatically resolves to the local `DATA` folder inside this project).

### Step 3: Run Full Rescan
- Go to the **`Home`** sheet.
- Click the **"Full Rescan / Rebuild All"** button.
- Excel will call the backend engine (`engine/nmdc_index_engine.exe`) and index all project workbooks. Once finished, a confirmation message will appear, and all 15 sheets will refresh automatically.

---

## 🔍 Key Features & How to Use Them

| Feature | Where to Find It | What It Does |
|:---|:---|:---|
| **Live Filter** | Press `Ctrl+Shift+F` or click **Live Filter** on Home | Instant multi-column search across all master documents and revision registers |
| **Update Changed Files** | Button on **Home** sheet | Quickly checks if any workbooks in `DATA/` were edited and indexes only changed files |
| **Pending Update** | **`Pending Update`** sheet | Review newly extracted deliverables before approving them |
| **Review Flags** | **`Review Flags`** sheet | Inspect any documents with ambiguous classifications or project number warnings |
| **Admin Reset** | Button on **Home** sheet | Clears indexed tables if you ever want a completely fresh start |

---

## 📁 Essential Folder Layout

- **`NMDC_Document_Index.xlsm`**: Your primary Excel application.
- **`DATA/`**: Source deliverable Excel workbooks (Methods and Tech). **Read-only.**
- **`engine/`**: Contains `nmdc_index_engine.exe` (the fast backend processing engine).
- **`config/`**: Contains `classification_rules.csv` and other editable taxonomy mappings.
- **`_ARCHIVE/`**: Safe storage containing past test runs, logs, and temporary build files.
