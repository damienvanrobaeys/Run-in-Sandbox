# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).


## Unreleased


## 2023-05-03
### Added
- Added option to run .intunewin via sdbapp
### Changed
- Changed Intunewin_Content_File and Intunewin_Command_File to be parameters for IntuneWin_Install.ps1


## 2023-03-29
### Added
- Added context menu entry for opening PDF in Sandbox
### Changed
- Completely rewrote alot of code in Add_Structure.ps1


## 2023-03-22
### Added
- Added context menu entry for running CMD/BAT in Sandbox


## 2023-03-21
### Changed
- Readded 7z part and adjusted 7z reg key path


## 2023-03-20
### Changed
- Completly refactored RunInSandbox.ps1 to use switch instead of ifelse and rearranged alot of code
### Fixed
- Fixed some issues with loading iso´s, exe´s and zip´s
### Removed
- Removed 7z part of RunInSandbox.ps1 because non-functional


## 2023-03-07
### Added
- Added ServiceUI
### Changed
- Replaced PSexec with ServiceUI for intunewin sandbox
### Removed
- Removed PSexec in favor of ServiceUI


## 2023-03-06
### Added
- Added option to Sandbox_Config.xml to cleanup leftover .wsb file afterwards (default is true)
### Changed
- .wsb is not executed by the "Start-Process"-cmdlet with -wait parameter


## 2023-03-03
### Added
- Added -noprofile to powershell commands to improve performance
### Changed
- Applied formatting of scripts and applied best practices
### Fixed
- Fixed .ps1 conext menu


## 2021-11-16
### Added
- Add a context menu for running PS1 as system in Sandbox
- Add a context menu for running MSIX in Sandbox
- Add a context menu for running PPKG in Sandbox
- Add a context menu for opening URL in Sandbox
- Add a context menu for extracting ISO in Sandbox
- Add a context menu for extracting 7z file in Sandbox
### Fixed
- Fix a bug where context menu for PS1 does not appear on Windows 11 


## 2021-09-21
### Added
- Add a context menu for reg file, to run them in Sandbox
- Add ability to run multiple apps in the same Sandbox session


## 2021-08-03
### Added
- Add a context menu for intunewin file, to run them in Sandbox
- Add ability to choose which content menu to add


## 2021-07-27
### Changed
- Change the default path where WSB are saved after running Sandbox: now in %temp%


## 2021-07-21
### Changed
- Updated the GUI when running EXE or MSI for more understanding
- Updated the GUI when running PS1 for more understanding


## 2021-07-16
### Added
- The Add_Structure.ps1 will now create a restore point
- It will then check if Sources folder exists


## 2020-06-24
### Removed
- Temporarily removed the main file [#9](https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/9)
### Changed
- Fixed detail language setting being French


## 2020-06-02 
### Added
 - Add new WSB config options for Windows 10 2004. These new settings can be managed in the **Sources\Run_in_Sandbox\Sandbox_Config.xml**
 - New options: AudioInput, VideoInput, ProtectedClient, PrinterRedirection, ClipboardRedirection, MemoryInMB


## 2020-05-19
### Added
- Added French, Italian, Spanish, English, and German languages for context menus. To configure language, edit **Main_Language** in **Sources\Run_in_Sandbox\Sandbox_Config.xml**