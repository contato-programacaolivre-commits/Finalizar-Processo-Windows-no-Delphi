🌐 **Language / Idioma:** [Português (Brasil)](README.md) | [English](README.en.md)

# Kill Windows Process in Delphi (3 Safe Methods via WinAPI)

Practical and compilable VCL project demonstrating how to safely terminate Windows processes in **Delphi** without leaking system resources (*Handle Leaks*) or closing files the user is currently editing.

> 📖 **Full tutorial with line-by-line explanation and prevention tips (in Portuguese):**  
> [Finalizar Processo Windows no Delphi: 3 Métodos (PID e .EXE)](https://programacaolivre.com.br)

---

## 🎯 What This Project Solves

1. **Surgical Termination by PID (`FinalizarProcessoPorPID`):**
   * Requests termination permission (`PROCESS_TERMINATE`), ends a specific process by its numeric identifier (`PID`), and immediately releases the *Handle* using `CloseHandle` inside a `finally` block.
   * Includes a safety guard against accidentally terminating critical system processes (`PID <= 4`) or the application itself (`GetCurrentProcessId`).

2. **Safe OLE Automation for Excel and Word (`ExportarExcelComSeguranca`):**
   * Captures the exclusive `HWND` of the window created via `CreateOleObject` and extracts its `PID` using `GetWindowThreadProcessId`.
   * Clears the COM reference (`Unassigned`) **before** terminating the process in the `finally` block (preventing *RPC Server is unavailable* exceptions) and closes **only** the hidden instance created by your system, preserving spreadsheets opened by the user.

3. **Batch Termination by Name (`FinalizarProcessoPorNome`):**
   * Uses `CreateToolhelp32Snapshot` (`Winapi.TlHelp32`) to scan active Windows processes and terminate all instances of a specific executable (`.exe`) by reusing the safe function from Method 1.

---

## 🛠️ Compatibility & Code Standards

* **Developed and tested on:** Embarcadero Delphi Community Edition (RAD Studio 11+).
* **Backward compatible with:** Delphi 10 Seattle, Delphi XE, and earlier versions (written in classic syntax, without *inline variables* `var x :=` and zero compiler *warnings*).
* **Platform:** Windows (Win32 / Win64).

---

## 🚀 How to Test It

1. Clone or download this repository:
   ```bash
   git clone https://github.com
   ```
2. Open the `.dproj` file in your Delphi IDE (`F11` / `Ctrl + F9`) and run (`F9`).
3. **To test Method 1 (PID):**
   * Open two separate Notepad windows (`notepad.exe`).
   * In the Command Prompt (CMD), run:
     ```bat
     tasklist /FI "IMAGENAME eq notepad.exe"
     ```
   * Enter just one of the returned PIDs into the application's input field and click **Finalizar por PID Direto**. Only the matching window will be closed.
4. **To test Method 2 (Excel/OLE):**
   * Leave a Microsoft Excel spreadsheet open on your desktop.
   * Click **Testar Excel/OLE** — the project will create a hidden instance, simulate an exception mid-routine, and terminate only the hidden process while keeping your open spreadsheet intact.

---

## 📂 File Structure

* `uPrincipal.pas` / `uPrincipal.dfm` — VCL Form and full commented implementation of all 3 methods.

---

## 📄 License

Code provided under the MIT License for educational purposes and free use in personal or commercial projects by the **[Programação Livre](https://programacaolivre.com.br)** hub.
