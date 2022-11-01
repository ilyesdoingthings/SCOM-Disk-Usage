## Microsoft SCOM Server - Disk Usage and Report 

This article provides a PowerShell script that gathers disk usage and then sends the report. I also describe the steps to configure a diagnostic task on a "disk full" alert in SCOM.

## introduction : 

It may happen because we sized a disk wrongly, did not expect growth of the WinSxS folder, neglected to stop IIS logging during troubleshooting, or because the developer did not implement a log-cleanup routine in her application. Full disks can have many reasons.

The first step when receiving an alert about a filling disk is to RDP into a server and use tools to analyze the disk consumption. This takes time and reoccurs—an ideal candidate for scripting and automation.

Diagnostic tasks in SCOM can run scripts or commands directly on the affected machine when an alert occurs.

![image](https://user-images.githubusercontent.com/26825056/199241994-07480194-39bb-4974-b22e-67772656cf0a.png)

## Preparing SCOM : 
