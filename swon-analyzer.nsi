; NSIS Installer for SWON-Analyzer
; The name of the installer
Name "Swim Ontario"

OutFile "swon-install.exe"

;--------------------------------

; Build Unicode installer
Unicode True

; The default installation directory
InstallDir "$APPDATA\Swim Ontario"

; Registry key to check for directory
InstallDirRegKey HKLM "Software\SwimOntario" "Install_Dir"

;--------------------------------

; Pages

Page components
Page directory
Page instfiles

UninstPage uninstConfirm
UninstPage instfiles

;--------------------------------

; The stuff to install
Section "Officials Utilities (required)"

  SectionIn RO

  ; Set output path to the installation directory.
  SetOutPath $INSTDIR

  ; Put file there
  File "swon-analyzer.exe"

  ; Write the installation path into the registry
  WriteRegStr HKLM SOFTWARE\SwimOntario "Install_Dir" "$INSTDIR"

  ; Write the uninstall keys for Windows
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\SwimOntario" "DisplayName" "Swim Ontario - Officials Utilities"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\SwimOntario" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\SwimOntario" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\SwimOntario" "NoRepair" 1
  WriteUninstaller "$INSTDIR\uninstall.exe"

SectionEnd

; Optional section (can be disabled by the user)
Section "Start Menu Shortcuts"

  CreateDirectory "$SMPROGRAMS\Swim Ontario"
  CreateShortcut "$SMPROGRAMS\Swim Ontario\Officials Utilities.lnk" "$INSTDIR\swon-analyzer.exe"

SectionEnd

;--------------------------------

; Uninstaller

Section "Uninstall"

  ; Remove registry keys
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\SwimOntario"
  DeleteRegKey HKLM SOFTWARE\SwimOntario

  ; Remove files and uninstaller
  Delete $INSTDIR\swon-analyzer.exe
  Delete $INSTDIR\uninstall.exe

  ; Remove shortcuts, if any
  Delete "$SMPROGRAMS\Swim Ontario\*.lnk"

  ; Remove .ini files, if any
  Delete "$INSTDIR\*.ini"

  ; Remove directories
  RMDir "$SMPROGRAMS\Swim Ontario"
  RMDir "$INSTDIR"

SectionEnd
