# Run in Sandbox
This allows you to do the below things in Windows Sandbox **just from a right-click** by adding context menus:
- Run PS1, VBS, EXE, MSI in the Sandbox
- Extract ZIP directly in the Sandbox
- Share a specific folder in the Sandbox

> *View the full blog post here*
http://www.systanddeploy.com/2019/06/run-file-in-windows-sandbox-from-right.html

**How to install it ?**
- Download the ZIP Run-in-Sandbox project (this is the main prerequiste)
- Extract the ZIP
- The Run-in-Sandbox-master **should contain** Add_Structure.ps1, Remove_Structure.ps1 and a Sources folder
- The Sources folder **should contain** a folder Run_in_Sandbox containing 13 files and 2 folders
- Once you have downloaded the folder structure, check if files have not be blocked after download
- Do a right-click on Add_Structure.ps1 and check if needed check Unblocked
- Run Add_Structure.ps1 **with admin rights**


**Update (06/02/20): Add new WSB config options for Windows 10 2004**
- Those settings can be managed in the **Sources\Run_in_Sandbox\Sandbox_Config.xml**
- New options: AudioInput, VideoInput, ProtectedClient, PrinterRedirection, ClipboardRedirection, MemoryInMB


**Update (05/19/20): Add other languages for context menus**
- Those languages are available: French, Italian, Spanish, English, German
- To configure language, go to **Sources\Run_in_Sandbox\Sandbox_Config.xml**
- Add language in **Main_Language**
- This language should be the language code name


![alt text](https://github.com/damienvanrobaeys/Run-in-Sandbox/blob/master/run_ps1_preview.gif.gif)
