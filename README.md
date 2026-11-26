🌐 **Language / Idioma:** [English](README.md) | [Português (Brasil)](README.pt-BR.md)

# Terminate Windows Process in Delphi (3 Safe WinAPI Methods)

Practical, compilable Delphi VCL demonstration project showing how to safely terminate Windows processes without leaking system resources (*Handle Leaks*) or closing files the user is actively editing.

> 📖 **Full Technical Guide with Line-by-Line Explanation (Portuguese):**  
> [Finalizar Processo Windows no Delphi: 3 Métodos (PID e .EXE)](https://programacaolivre.com.br/finalizar-processo-windows-no-delphi/)

---

## 🎯 What This Project Solves

1. **Surgical Termination by PID (`FinalizarProcessoPorPID`):**
   * Requests termination rights (`PROCESS_TERMINATE`), terminates a specific process by its numeric Process ID (`PID`), and immediately releases the Windows handle inside a `finally` block using `CloseHandle`.
   * Includes built-in safety guards against terminating critical OS processes (`PID <= 4`) or the application itself (`GetCurrentProcessId`).

2. **Safe OLE Automation for Excel & Word (`ExportarExcelComSeguranca`):**
   * Captures the exclusive window handle (`HWND`) created via `CreateOleObject` and extracts its exact `PID` using `GetWindowThreadProcessId`.
   * Clears the COM reference (`Unassigned`) **before** terminating the process in the `finally` block (preventing *RPC server is unavailable* exceptions) and kills **only** the hidden background instance created by your app, leaving user-opened spreadsheets untouched.

3. **Batch Termination by Executable Name (`FinalizarProcessoPorNome`):**
   * Uses `CreateToolhelp32Snapshot` (`Winapi.TlHelp32`) to enumerate active Windows processes and safely terminate all running instances of a target executable (`.exe`) by reusing the safe Method 1 routine.

---

## 🛠️ Compatibility & Code Standards

* **Developed and tested on:** Embarcadero Delphi Community Edition (RAD Studio 11+).
* **Backward compatible with:** Delphi 10 Seattle, Delphi XE, and earlier versions (written in classic Object Pascal syntax without inline variables `var x :=` and zero compiler warnings).
* **Platform:** Windows (Win32 / Win64).

---

## 🚀 How to Test in Practice

1. Clone or download this repository:
   ```bash
   git clone https://github.com
   ```
2. Open the `.dproj` file in Delphi (`F11` / `Ctrl + F9`) and run (`F9`).
3. **Testing Method 1 (Direct PID):**
   * Open two separate Notepad windows (`notepad.exe`).
   * In Command Prompt (CMD), run:
     ```bat
     tasklist /FI "IMAGENAME eq notepad.exe"
     ```
   * Enter only one of the returned PIDs into the application field and click **Finalizar por PID Direto**. Only that specific window will close.
4. **Testing Method 2 (Excel / OLE Isolation):**
   * Keep a Microsoft Excel workbook open on your desktop.
   * Click **Testar Excel/OLE** — the demo will spawn a hidden Excel instance, simulate an exception mid-routine, and terminate only the background process while keeping your active workbook intact.

---

## 📂 Project Structure

* `uPrincipal.pas` / `uPrincipal.dfm` — Main VCL form and complete commented implementation of all 3 methods.

---

## 📄 License

Source code provided under the MIT License for educational and commercial use by **[Programação Livre](https://programacaolivre.com.br)**.
