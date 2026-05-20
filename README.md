# chm2docx-power-converter 🚀

A robust PowerShell tool designed to bypass standard Microsoft Word limitations and seamlessly convert massive `.chm` (Compiled HTML Help) documentation into a single, cohesive `.docx` file. Perfectly optimized for ingestion into AI assistants like **Microsoft Copilot Studio**.

## 🌟 Why This Tool Exists
Standard COM-interop scripts often cause Microsoft Word to crash or run out of memory when trying to stitch together hundreds of separate HTML pages sequentially (throwing generic errors like *"Word has encountered a problem"*).

This script solves the issue by bypassing Word during the heavy lifting: it decompiles the CHM, combines all 900+ pages instantly on a code level via PowerShell into a single Master-HTML file, and instructs Word to perform just **one single open-and-save action**.

### Features:
- 🔓 **Auto-Unblocks:** Automatically unblocks downloaded internet files to prevent silent extraction failures.
- 🏷️ **Filename Normalization:** Handles brackets, spaces, and special characters that cause the legacy `hh.exe` utility to crash.
- 🖼️ **Preserves Assets:** Keeps tables, formatting, and embedded images perfectly intact.
- 🧠 **AI-Optimized:** Inserts clean page-break markers for perfect indexing by LLMs/Copilot.

---

## 🛠️ Prerequisites
- **Windows OS** (with the built-in `hh.exe` utility)
- **Microsoft Word** installed locally
- **PowerShell 5.1+** or **PowerShell Core**

---

## 🚀 How to Use

1. **Clone or download** this repository.
2. Open **PowerShell** (ideally as Administrator).
3. If script execution is blocked on your machine, temporarily bypass it using:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

4.Run the script:
open PowerShell
.\chm2docx-power-converter.ps1

5.Enter the requested parameters:

ChmFilePath: The full path to your .chm file (e.g., C:\Users\Name\Downloads\Manual.chm).

OutputDirectory: The destination folder where you want your files saved.

📂 Output Structure
After a successful run, your destination folder will look like this:

Plaintext
📁 Your-Output-Folder/
├── 📁 CHM_Entpackt/            # Raw decompiled HTML pages and images
└── 📄 Copilot_Wissensbasis.docx # The final file ready for Copilot Studio!

⚙️ How it Works under the Hood
Clean & Prepare: Uses Unblock-File and creates a temporary processing file without spaces or parentheses to respect the limitations of the ancient hh.exe decompiler.

Decompile: Extracts all internal HTML files and image folders.

PowerShell Stitching: Reads the raw <body> of all files and merges them sequentially into a single Copilot_Master_Wissen.html using UTF-8 encoding to preserve special characters.

Word Conversion: Launches a headless Word instance (Word.Application), opens the massive HTML file, and saves it as a standard, high-compatibility .docx document before cleaning up the workspace.

📄 License
This project is open-source and available under the MIT License.
