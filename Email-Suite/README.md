# Email Suite GUI

A full-featured graphical user interface (GUI) built with PowerShell and WPF/XAML, designed to automate the creation of Data Center Operations notifications.

## 🎯 Problem Solved
Manual generation of maintenance and patching notifications is error-prone and time-consuming. This tool provides a single pane of glass to parse raw system alerts, calculate dynamic patching dates, and generate standardized formatting.

## ✨ Key Features
* **WPF/XAML Interface:** A custom, dark-themed UI with tabbed navigation for Patching, Service Alerts, and Admin Tools.
* **Smart Excel Parsing:** Reads and updates `.xlsx` maintenance schedules, dynamically calculating upcoming patch windows based on complex business logic (e.g., "3rd Sunday of the month").
* **Regex Data Extraction:** Automatically extracts key fields (Subject, Dates, Affected Systems) from pasted text or raw `.msg`/`.docx` files.
* **Async Outlook Dispatch:** Uses background worker jobs to generate and display Outlook drafts without freezing the primary GUI.

## ⚙️ Usage
Run `EmailSuiteGUI.ps1`. Ensure `templates/`, `logs/`, and `dlists/` directories exist in the same root path, along with the required `Patches.xlsx` and `ChangeRequests.xlsx` files.
