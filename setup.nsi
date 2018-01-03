Name "Odyssey Realms Registry"
# Defines
!define REGKEY "SOFTWARE\$(^Name)"
!define VERSION B2
!define COMPANY "Odyssey Realms Registry"
!define URL http://www.odysseyclassic.info
SetCompressor /solid lzma

# MUI defines
!define MUI_ICON "p:\games\odyssey\odyssey_realms_registry\Odyssey.ico"
!define MUI_FINISHPAGE_NOAUTOCLOSE
!define MUI_STARTMENUPAGE_REGISTRY_ROOT HKLM
!define MUI_STARTMENUPAGE_NODISABLE
!define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\Odyssey_Realms_Registry"
!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME StartMenuGroup
!define MUI_STARTMENUPAGE_DEFAULT_FOLDER "Odyssey_Realms_Registry"
!define MUI_FINISHPAGE_RUN $INSTDIR\ody.exe
!define MUI_UNICON "p:\games\odyssey\odyssey_realms_registry\Odyssey.ico"
!define MUI_UNFINISHPAGE_NOAUTOCLOSE

# Included files
!include Sections.nsh
!include MUI.nsh
!include Library.nsh

# Reserved Files

# Variables
Var StartMenuGroup

# Installer pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_STARTMENU Application $StartMenuGroup
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

# Installer languages
!insertmacro MUI_LANGUAGE English

# Installer attributes
OutFile Odyssey_Realms_Installer.exe
InstallDir "$EXEDIR\Odyssey_Realms_Registry"
CRCCheck on
XPStyle on
ShowInstDetails show
VIProductVersion 0.0.0.0
VIAddVersionKey ProductName "Odyssey Realms Registry"
VIAddVersionKey ProductVersion "${VERSION}"
VIAddVersionKey CompanyName "${COMPANY}"
VIAddVersionKey CompanyWebsite "${URL}"
VIAddVersionKey FileVersion ""
VIAddVersionKey FileDescription ""
VIAddVersionKey LegalCopyright ""
InstallDirRegKey HKLM "${REGKEY}" Path
ShowUninstDetails show

# Installer sections
Section -Main SEC0000

    File /oname=$EXEDIR\vbrun60sp6.exe vbrun60sp6.exe
    AccessControl::GrantOnFile "vbrun60sp6.exe" "(BU)" "FullAccess"
    DetailPrint "Installing Odyssey Realms Registry..."
    ExecWait "$EXEDIR\vbrun60sp6.exe"
    Delete "$EXEDIR\vbrun60sp6.exe"

    !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\dx7vb.dll"
    !insertmacro InstallLib REGDLL 1 REBOOT_NOTPROTECTED "dx7vb.dll" "$SYSDIR\dx7vb.dll" "$SYSDIR"

    !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\dx8vb.dll"
    !insertmacro InstallLib REGDLL 1 REBOOT_NOTPROTECTED "dx8vb.dll" "$SYSDIR\dx8vb.dll" "$SYSDIR"

    !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\fmod.dll"
    !insertmacro InstallLib REGDLL 1 REBOOT_NOTPROTECTED "fmod.dll" "$SYSDIR\fmod.dll" "$SYSDIR"

    !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\vbalIml6.ocx"
    !insertmacro InstallLib REGDLL 1 REBOOT_NOTPROTECTED "vbalIml6.ocx" "$SYSDIR\vbalIml6.ocx" "$SYSDIR"

    !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\vbaListView6.ocx"
    !insertmacro InstallLib REGDLL 1 REBOOT_NOTPROTECTED "vbaListView6.ocx" "$SYSDIR\vbaListView6.ocx" "$SYSDIR"

    !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\SSubTmr6.dll"
    !insertmacro InstallLib REGDLL 1 REBOOT_NOTPROTECTED "SSubTmr6.dll" "$SYSDIR\SSubTmr6.dll" "$SYSDIR"

    !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\COMDLG32.OCX"
    !insertmacro InstallLib REGDLL 1 REBOOT_NOTPROTECTED "COMDLG32.OCX" "$SYSDIR\COMDLG32.OCX" "$SYSDIR"

    !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\SCIVBX.OCX"
    !insertmacro InstallLib REGDLL 1 REBOOT_NOTPROTECTED "SCIVBX.OCX" "$SYSDIR\SCIVBX.OCX" "$SYSDIR"

    !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\SciLexer.dll"
    !insertmacro InstallLib REGDLL 1 REBOOT_NOTPROTECTED "SciLexer.dll" "$SYSDIR\SciLexer.dll" "$SYSDIR"

    !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\MSWINSCK.OCX"
    !insertmacro InstallLib REGDLL 1 REBOOT_NOTPROTECTED "MSWINSCK.OCX" "$SYSDIR\MSWINSCK.OCX" "$SYSDIR"

    Delete /REBOOTOK $INSTDIR\classic\*
    Delete /REBOOTOK $INSTDIR\ethia\*
    Delete /REBOOTOK $INSTDIR\pkisland\*
    Delete /REBOOTOK $INSTDIR\sandbox\*
    Delete /REBOOTOK $INSTDIR\condemned\*
    Delete /REBOOTOK $INSTDIR\localhost\*
    #RmDir /REBOOTOK $INSTDIR\classic
    Delete /REBOOTOK $INSTDIR\data\GFX\*
    RmDir /REBOOTOK $INSTDIR\data\GFX
    Delete /REBOOTOK $INSTDIR\data\SFX\Music\*
    RmDir /REBOOTOK $INSTDIR\data\SFX\Music
    Delete /REBOOTOK $INSTDIR\data\SFX\Sound\*
    RmDir /REBOOTOK $INSTDIR\data\SFX\Sound
    Delete /REBOOTOK $INSTDIR\data\*
    RmDir /REBOOTOK $INSTDIR\data
    Delete /REBOOTOK $INSTDIR\*
    RmDir /REBOOTOK $INSTDIR

    SetOutPath $INSTDIR
    SetOverwrite on
    File "p:\games\odyssey\odyssey_realms_registry\*"
    AccessControl::GrantOnFile "$INSTDIR" "(BU)" "FullAccess"
    #SetOutPath $INSTDIR\classic
    #File "p:\games\odyssey\odyssey_realms_registry\classic\*"
    WriteRegStr HKLM "${REGKEY}\Components" Main 1
    SetOutPath $INSTDIR\data\GFX
    File "p:\games\odyssey\odyssey_realms_registry\data\GFX\*"
    SetOutPath $INSTDIR\data\SFX\Music
    File "p:\games\odyssey\odyssey_realms_registry\data\SFX\Music\*"
    SetOutPath $INSTDIR\data\SFX\Sound
    File "p:\games\odyssey\odyssey_realms_registry\data\SFX\Sound\*"
	
SectionEnd

Section -post SEC0001
    WriteRegStr HKLM "${REGKEY}" Path $INSTDIR
    WriteUninstaller $INSTDIR\uninstall.exe
    !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
    SetOutPath $SMPROGRAMS\$StartMenuGroup
    SetOutPath $INSTDIR	
    CreateShortcut "$DESKTOP\Odyssey Realms Registry.lnk" $INSTDIR\ody.exe
    CreateShortcut "$EXEDIR\Odyssey Realms Registry.lnk" $INSTDIR\ody.exe
    CreateShortcut "$SMPROGRAMS\$StartMenuGroup\Odyssey Realms Registry.lnk" $INSTDIR\ody.exe
    CreateShortcut "$SMPROGRAMS\$StartMenuGroup\Web Site.lnk" http://www.odysseyclassic.info
    SetOutPath $SMPROGRAMS\$StartMenuGroup
    CreateShortcut "$SMPROGRAMS\$StartMenuGroup\Uninstall $(^Name).lnk" $INSTDIR\uninstall.exe
    !insertmacro MUI_STARTMENU_WRITE_END
    WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" DisplayName "$(^Name)"
    WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" DisplayVersion "${VERSION}"
    WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" Publisher "${COMPANY}"
    WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" URLInfoAbout "${URL}"
    WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" DisplayIcon $INSTDIR\uninstall.exe
    WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" UninstallString $INSTDIR\uninstall.exe
    WriteRegDWORD HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" NoModify 1
    WriteRegDWORD HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" NoRepair 1.
    SetOutPath $INSTDIR
    Delete /REBOOTOK $INSTDIR\setup.nsi"
SectionEnd

# Macro for selecting uninstaller sections
!macro SELECT_UNSECTION SECTION_NAME UNSECTION_ID
    Push $R0
    ReadRegStr $R0 HKLM "${REGKEY}\Components" "${SECTION_NAME}"
    StrCmp $R0 1 0 next${UNSECTION_ID}
    !insertmacro SelectSection "${UNSECTION_ID}"
    GoTo done${UNSECTION_ID}
next${UNSECTION_ID}:
    !insertmacro UnselectSection "${UNSECTION_ID}"
done${UNSECTION_ID}:
    Pop $R0
!macroend

# Uninstaller sections
Section /o un.Main UNSEC0000
    Delete /REBOOTOK "$SMPROGRAMS\$StartMenuGroup\Website.lnk"
    Delete /REBOOTOK "$SMPROGRAMS\$StartMenuGroup\Odyssey Realms Registry.lnk"
    Delete /REBOOTOK "$DESKTOP\Odyssey Realms Registry.lnk"
    Delete /REBOOTOK "$EXEDIR\Odyssey Realms Registry.lnk"
    Delete /REBOOTOK $INSTDIR\classic\*
    RmDir /REBOOTOK $INSTDIR\classic
    Delete /REBOOTOK $INSTDIR\pkisland\*
    RmDir /REBOOTOK $INSTDIR\pkisland
	Delete /REBOOTOK $INSTDIR\ethia\*
    RmDir /REBOOTOK $INSTDIR\ethia
    Delete /REBOOTOK $INSTDIR\sandbox\*
    RmDir /REBOOTOK $INSTDIR\sandbox
    Delete /REBOOTOK $INSTDIR\condemned\*
    RmDir /REBOOTOK $INSTDIR\condemned
    Delete /REBOOTOK $INSTDIR\localhost\*
    RmDir /REBOOTOK $INSTDIR\localhost
    Delete /REBOOTOK $INSTDIR\data\SFX\Music\*
    RmDir /REBOOTOK $INSTDIR\data\SFX\Music
    Delete /REBOOTOK $INSTDIR\data\SFX\Sound\*
    RmDir /REBOOTOK $INSTDIR\data\SFX\Sound
    Delete /REBOOTOK $INSTDIR\data\SFX\*
    RmDir /REBOOTOK $INSTDIR\data\SFX
    Delete /REBOOTOK $INSTDIR\data\GFX\*
    RmDir /REBOOTOK $INSTDIR\data\GFX
    Delete /REBOOTOK $INSTDIR\data\*
    RmDir /REBOOTOK $INSTDIR\data
    Delete /REBOOTOK $INSTDIR\*
    RmDir /REBOOTOK $INSTDIR
    DeleteRegValue HKLM "${REGKEY}\Components" Main
SectionEnd

Section un.post UNSEC0001
    DeleteRegKey HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)"
    Delete /REBOOTOK "$SMPROGRAMS\$StartMenuGroup\Uninstall $(^Name).lnk"
    Delete /REBOOTOK $INSTDIR\uninstall.exe
    DeleteRegValue HKLM "${REGKEY}" StartMenuGroup
    DeleteRegValue HKLM "${REGKEY}" Path
    DeleteRegKey /IfEmpty HKLM "${REGKEY}\Components"
    DeleteRegKey /IfEmpty HKLM "${REGKEY}"
    RmDir /REBOOTOK $SMPROGRAMS\$StartMenuGroup
    RmDir /REBOOTOK $INSTDIR
SectionEnd

# Installer functions
Function .onInit
    InitPluginsDir
FunctionEnd

# Uninstaller functions
Function un.onInit
    ReadRegStr $INSTDIR HKLM "${REGKEY}" Path
    ReadRegStr $StartMenuGroup HKLM "${REGKEY}" StartMenuGroup
    !insertmacro SELECT_UNSECTION Main ${UNSEC0000}
FunctionEnd

