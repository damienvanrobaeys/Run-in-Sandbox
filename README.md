# Run in Sandbox: a quick way to test EXE, PS1, VBS, MSI, Intunewin in Windows Sandbox
This allows you to do the below things in Windows Sandbox **just from a right-click** by adding context menus:
- Run PS1, VBS, EXE, MSI in the Sandbox
- Extract ZIP directly in the Sandbox
- Share a specific folder in the Sandbox

> *View the full blog post here*
http://www.systanddeploy.com/2019/06/run-file-in-windows-sandbox-from-right.html


**Note that this project has been build on personal time, it's not a professional project. Use it at your own risk, and please read How to install it before running it.**

**How to install it ?**
- Download the ZIP Run-in-Sandbox project (this is the main prerequiste)
- Extract the ZIP
- The Run-in-Sandbox-master **should contain** at least Add_Structure.ps1  and a Sources folder
- Please **do not download only** Add_Structure.ps1
- The Sources folder **should contain** folder Run_in_Sandbox containing 16 files
- Once you have downloaded the folder structure, **check if files have not be blocked after download**
- Do a right-click on Add_Structure.ps1 and check if needed check Unblocked
- Run Add_Structure.ps1 **with admin rights**

**Update (08/03/21): Run intunewin file in sandbox**
- Add a context menu for intunewin file, to run them in Sandbox
- Add ability to choose which content menu to add

**Update (07/27/21): Change default WSB location**
- Change the default path where WSB are saved after running Sandbox: now in %temp%

**Update (07/21/21): Change GUI for MSI, EXE, PS1**
- Updated the GUI when running EXE or MSI for more understanding
- Updated the GUI when running PS1 for more understanding

**Update (07/16/21): Add more controls to avoid association EXE issue**
- The Add_Structure.ps1 will now create a restore point
- It will then check if Sources folder exists

**Update (06/02/20): Add new WSB config options for Windows 10 2004**
- Those settings can be managed in the **Sources\Run_in_Sandbox\Sandbox_Config.xml**
- New options: AudioInput, VideoInput, ProtectedClient, PrinterRedirection, ClipboardRedirection, MemoryInMB



![alt text](https://github.com/damienvanrobaeys/Run-in-Sandbox/blob/master/run_ps1_preview.gif.gif)
