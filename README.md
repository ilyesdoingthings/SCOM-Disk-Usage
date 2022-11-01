## Microsoft SCOM Server - Disk Usage and Report 

This article provides a PowerShell script that gathers disk usage and then sends the report. I also describe the steps to configure a diagnostic task on a "disk full" alert in SCOM.

## introduction

It may happen because we sized a disk wrongly, did not expect growth of the WinSxS folder, neglected to stop IIS logging during troubleshooting, or because the developer did not implement a log-cleanup routine in her application. Full disks can have many reasons.

The first step when receiving an alert about a filling disk is to RDP into a server and use tools to analyze the disk consumption. This takes time and reoccurs—an ideal candidate for scripting and automation.

Diagnostic tasks in SCOM can run scripts or commands directly on the affected machine when an alert occurs.

![image](https://user-images.githubusercontent.com/26825056/199241994-07480194-39bb-4974-b22e-67772656cf0a.png)

## Preparing SCOM 

By default, only VBScript can create diagnostic and recovery tasks. Download and install a free, open-source PowerShell community management pack on GitHub to use PowerShell. You likely have already imported management packs for Windows Server operating systems.

## PowerShell script 

![image](https://user-images.githubusercontent.com/26825056/199244227-e85c3588-68dd-4eff-bbca-b225612cee00.png)

## Create the diagnostic task 

In the SCOM console, switch to the Authoring pane and expand Management Pack Objects and Monitors. Limit the search to logical disk. Right click Windows [20XX] Logical Disk Free Space Monitor.

![image](https://user-images.githubusercontent.com/26825056/199244423-cccb938e-b875-4c69-b00b-674f62e54bf1.png)

Switch to the Diagnostic and Recovery tab, click Add…, and choose Diagnostic for warning health state.

![image](https://user-images.githubusercontent.com/26825056/199244510-b9e9d3f1-a0ee-4782-8fab-1f52b0f3a2cf.png)

Chose Run a PowerShell script (Community) and click New… to create a management pack to store the task in.

![image](https://user-images.githubusercontent.com/26825056/199244647-877f2d38-dc5a-49f9-958d-91382f713ed2.png)

Choose a fitting name like Windows.Server.Custom.Tasks and proceed by clicking Next. As an optional step, you can specify additional information in the Knowledge section. Proceed by clicking Create.
Back in the Task Wizard, ensure that Run a PowerShell script (Community) is still selected and click Next.

In the General section, specify a task name, such as Disk Full - Troubleshooting Assistance and optionally a description, such as Scans the partition and sends a utilization report to the admin team. Proceed by clicking Next.

In the Script section, enter a File Name, such as Get-LargeDirectoriesAndFiles.ps1, set the timeout to 5 minutes, and paste the Script into it.

Set the variable $emailTo to specify the recipient of the disk usage report, such as adminteam@contoso.msft. The variable $smtpSrv specifies the name that will send the mail, and set $emailFrom specifies the sender address.

Optionally you can adjust the values for the number of directories ($numberOfTopDirectories) and files ($numberOfTopFiles) to show on the report.

Finaly, Initiate the task creation by clicking Create. Depending on your environment, this may take a while.

![image](https://user-images.githubusercontent.com/26825056/199244885-ae994c12-046e-45ba-a13f-1e91b259aede.png)

![image](https://user-images.githubusercontent.com/26825056/199244932-33821fee-cc00-4bcb-8b6d-820b13c19fc8.png)

![image](https://user-images.githubusercontent.com/26825056/199245208-0ff98cc5-bf18-45ae-94c9-d67cffec4e96.png)



