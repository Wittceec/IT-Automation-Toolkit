# IT Operations Automation Toolkit

A collection of PowerShell scripts and tools designed to streamline Data Center Operations, automate system administration tasks, and integrate local workflows with enterprise environments (Microsoft 365, SharePoint, and Teams).

## 📁 Repository Structure

* **`/Email-Suite`**: Contains a robust WPF/XAML graphical user interface for parsing system alerts, parsing Excel maintenance schedules, and generating standardized Outlook notification templates.
* **`/Monitoring`**: Includes filesystem watcher scripts designed to monitor network shares and automate internal Microsoft Teams reporting.
* **`/File-Management`**: Scripts dedicated to local workstation hygiene, automated file sorting, and bulk document printing with intelligent path discovery.

## 🛠️ Skills & Technologies Demonstrated
* **Languages & Frameworks:** PowerShell 5.1+, WPF/XAML (UI Design), Regex
* **Enterprise Integration:** Microsoft Office COM Objects (Word, Excel, Outlook), SharePoint/OneDrive path resolution
* **Systems Administration:** File system monitoring, background job dispatching, patch schedule management
* **Core Competencies:** Business workflow automation, system configuration, Azure/Linux environments

## 🚀 Getting Started
Most scripts in this repository utilize standard Windows environments and Microsoft Office COM objects. Ensure you have the appropriate execution policies set for local scripts:
`Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`

---
*Author: Chris (@Wittceec)*
