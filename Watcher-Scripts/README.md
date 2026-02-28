# Walkthrough Watcher

An automated filesystem monitoring script that bridges local/network file storage with Microsoft Teams communications.

## 🎯 Problem Solved
Data Center walkthrough logs and images needed to be manually uploaded and reported to the team. This script automates the reporting pipeline while preventing duplicate alerts.

## ✨ Key Features
* **Filesystem Watcher:** Utilizes `System.IO.FileSystemWatcher` to monitor specific network paths for new image files matching a defined naming convention.
* **Debounce Logic:** Implements global state tracking to prevent duplicate triggers (a common issue with FileSystemWatcher events) by enforcing a 60-second cooldown per file.
* **Date-Shift Calculations:** Automatically calculates shift dates (e.g., treating a 2:00 AM Friday file creation as a "Thursday Night" shift) for accurate subject line generation.
* **Outlook-to-Teams Integration:** Programmatically dispatches an email with the image attachment to a specific Microsoft Teams channel routing address.

## ⚙️ Usage
Run the script in a persistent PowerShell window. The script uses a global crash guard (`Try/Catch`) and an infinite `while` loop to ensure continuous background monitoring.
