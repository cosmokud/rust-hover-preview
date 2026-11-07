; Template for the NSIS installer cargo-packager builds, pointed at by
; `[package.metadata.packager.nsis] template` in Cargo.toml. The placeholders
; in double braces are filled in by cargo-packager.
;
; This is cargo-packager 0.11.8's own template with three differences:
;   - there is no "Already Installed" page: a previous installation is removed
;     automatically, before anything is written, instead of being offered as a
;     choice;
;   - a running copy of the app is terminated rather than prompted for;
;   - a page of its own is shown after the shortcuts page, listing the optional
;     engines this
;     app can drive but does not ship, marked with whether this machine already
;     has each of them and carrying an `Install` link where `winget` is here to run
;     and a `Manual Download` link where it is not (see "Optional engines page"
;     below). It writes one batch script into `%TEMP%` and nothing else: an install
;     is winget's, shown in a console window, and a manual download is the page
;     opened in the user's own browser. Office and f3d are the rows with no
;     `Install`, having no package id, so they carry the manual download alone.
; Upgrading cargo-packager means diffing this file against the upstream
; template at crates/packager/src/package/nsis/installer.nsi.

; Set the compression algorithm.
!if "{{compression}}" == ""
  SetCompressor /SOLID lzma
!else
  SetCompressor /SOLID "{{compression}}"
!endif

Unicode true

!include MUI2.nsh
!include FileFunc.nsh
!include x64.nsh
!include WordFunc.nsh
!include "FileAssociation.nsh"
!include "StrFunc.nsh"
!include "StrFunc.nsh"
${StrCase}
${StrLoc}

!define MANUFACTURER "{{manufacturer}}"
!define PRODUCTNAME "{{product_name}}"
!define VERSION "{{version}}"
!define VERSIONWITHBUILD "{{version_with_build}}"
!define SHORTDESCRIPTION "{{short_description}}"
!define INSTALLMODE "{{install_mode}}"
!define LICENSE "{{license}}"
!define INSTALLERICON "{{installer_icon}}"
!define SIDEBARIMAGE "{{sidebar_image}}"
!define HEADERIMAGE "{{header_image}}"
!define MAINBINARYNAME "{{main_binary_name}}"
!define MAINBINARYSRCPATH "{{main_binary_path}}"
!define IDENTIFIER "{{identifier}}"
!define COPYRIGHT "{{copyright}}"
!define OUTFILE "{{out_file}}"
!define ARCH "{{arch}}"
!define PLUGINSPATH "{{additional_plugins_path}}"
!define ALLOWDOWNGRADES "{{allow_downgrades}}"
!define DISPLAYLANGUAGESELECTOR "{{display_language_selector}}"
!define UNINSTKEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCTNAME}"
!define MANUPRODUCTKEY "Software\${MANUFACTURER}\${PRODUCTNAME}"
!define UNINSTALLERSIGNCOMMAND "{{uninstaller_sign_cmd}}"
!define ESTIMATEDSIZE "{{estimated_size}}"

Name "${PRODUCTNAME}"
BrandingText "${COPYRIGHT}"
OutFile "${OUTFILE}"

VIProductVersion "${VERSIONWITHBUILD}"
VIAddVersionKey "ProductName" "${PRODUCTNAME}"
VIAddVersionKey "FileDescription" "${SHORTDESCRIPTION}"
VIAddVersionKey "LegalCopyright" "${COPYRIGHT}"
VIAddVersionKey "FileVersion" "${VERSION}"
VIAddVersionKey "ProductVersion" "${VERSION}"

; Plugins path, currently exists for linux only
!if "${PLUGINSPATH}" != ""
    !addplugindir "${PLUGINSPATH}"
!endif

!if "${UNINSTALLERSIGNCOMMAND}" != ""
  !uninstfinalize '${UNINSTALLERSIGNCOMMAND}'
!endif

; Handle install mode, `perUser`, `perMachine` or `both`
!if "${INSTALLMODE}" == "perMachine"
  RequestExecutionLevel highest
!endif

!if "${INSTALLMODE}" == "currentUser"
  RequestExecutionLevel user
!endif

!if "${INSTALLMODE}" == "both"
  !define MULTIUSER_MUI
  !define MULTIUSER_INSTALLMODE_INSTDIR "${PRODUCTNAME}"
  !define MULTIUSER_INSTALLMODE_COMMANDLINE
  !if "${ARCH}" == "x64"
    !define MULTIUSER_USE_PROGRAMFILES64
  !else if "${ARCH}" == "arm64"
    !define MULTIUSER_USE_PROGRAMFILES64
  !endif
  !define MULTIUSER_INSTALLMODE_DEFAULT_REGISTRY_KEY "${UNINSTKEY}"
  !define MULTIUSER_INSTALLMODE_DEFAULT_REGISTRY_VALUENAME "CurrentUser"
  !define MULTIUSER_INSTALLMODEPAGE_SHOWUSERNAME
  !define MULTIUSER_INSTALLMODE_FUNCTION RestorePreviousInstallLocation
  !define MULTIUSER_EXECUTIONLEVEL Highest
  !include MultiUser.nsh
!endif

; installer icon
; The uninstaller is compiled into the installer as its own EXE, so without
; MUI_UNICON it keeps NSIS's stock uninstall glyph no matter what MUI_ICON says.
!if "${INSTALLERICON}" != ""
  !define MUI_ICON "${INSTALLERICON}"
  !define MUI_UNICON "${INSTALLERICON}"
!endif

; installer sidebar image
!if "${SIDEBARIMAGE}" != ""
  !define MUI_WELCOMEFINISHPAGE_BITMAP "${SIDEBARIMAGE}"
!endif

; installer header image
!if "${HEADERIMAGE}" != ""
  !define MUI_HEADERIMAGE
  !define MUI_HEADERIMAGE_BITMAP  "${HEADERIMAGE}"
!endif

; Define registry key to store installer language
!define MUI_LANGDLL_REGISTRY_ROOT "HKCU"
!define MUI_LANGDLL_REGISTRY_KEY "${MANUPRODUCTKEY}"
!define MUI_LANGDLL_REGISTRY_VALUENAME "Installer Language"

; Installer pages, must be ordered as they appear
; 1. Welcome Page
!define MUI_PAGE_CUSTOMFUNCTION_PRE SkipIfPassive
!insertmacro MUI_PAGE_WELCOME

; 2. License Page (if defined)
!if "${LICENSE}" != ""
  !define MUI_PAGE_CUSTOMFUNCTION_PRE SkipIfPassive
  !insertmacro MUI_PAGE_LICENSE "${LICENSE}"
!endif

; 3. Install mode (if it is set to `both`)
!if "${INSTALLMODE}" == "both"
  !define MUI_PAGE_CUSTOMFUNCTION_PRE SkipIfPassive
  !insertmacro MULTIUSER_PAGE_INSTALLMODE
!endif

; 4. Choose install directory page
!define MUI_PAGE_CUSTOMFUNCTION_PRE SkipIfPassive
!insertmacro MUI_PAGE_DIRECTORY

; 5. Start menu shortcut page
!define MUI_PAGE_CUSTOMFUNCTION_PRE SkipIfPassive
Var AppStartMenuFolder
!insertmacro MUI_PAGE_STARTMENU Application $AppStartMenuFolder

; 6. Optional engines page
;
; The engines this app can drive but does not ship, and whether this machine already has each
; of them: one row per engine, saying what the engine is called and a few of the formats it
; would give previews to, marked as detected or left greyed with an `Install` link and a `Manual
; Download` one.
;
; `Install` runs `winget` in a console window of its own, from a batch script written here into
; `%TEMP%` — the same script, character for character, that the app's own `Codecs → Install` rows
; run (see `src/app/winget.rs`), so the two doors install the same thing the same way and both
; leave the retry, and the reason a failure is readable, to the window. Nothing is elevated:
; `installMode = "currentUser"` above gives `RequestExecutionLevel user`, so `winget` installs
; into this user's own profile, which is the profile the app then looks in. A machine-wide package
; — LibreOffice, calibre, ImageMagick, PeaZip — raises its own UAC prompt, and that is the
; package's, not this page's.
;
; `Manual Download` opens the page the engine is installed from in the browser the user already
; has: `ExecShell "open"`, since this installer is not elevated (see the note above the link
; functions). A row an engine is present for offers nothing, since there is nothing to offer.
;
; What is asked about each engine is what the app itself asks before it drives one, so a page
; that says a name is missing and the app's own `Codecs` row agree: the folders each installer
; writes to, beside the app or the `PATH` for FFmpeg's player, the two folders mpv's own
; installers write to and the three places winget puts one (its `Links` shim, `…\Local\Programs`
; and the portable package under `…\WinGet\Packages`), the console archiver inside a PeaZip
; installation, the `f3d\bin` folder f3d's own installer writes, and the ProgID of each Office
; application — as a row of Office's own, which has no `Install` because Office is not a package
; and no free download, and as the one thing decided here about LibreOffice: the Office formats are
; worth naming only on a machine with no Office to draw them, and what LibreOffice adds to a
; machine that has one is the formats beside them.
;
; Like the pages above it, this page is not shown to an unattended installer at all: an update
; put on by the app's own `Auto` answer runs this installer silently, and a page nobody is
; there to read is not shown to anyone.
Var EnginesLibreText
Var EnginesWinget
Var EnginesIntro
Var EnginesBusy
Var EnginesId
Var EnginesCheck
Var EnginesScript
Var EnginesRc
Var EnginesCode
Var EnginesWaited
Var EnginesLabelText
Var EnginesRowLabel
Var EnginesRowPos
Var EnginesRowInstall
Var EnginesRowDownload

; One row's own controls: the label, where it sits, and its two links. A row is inserted with the
; same `NAME` here, and the same name again in `EnginesInstall<NAME>` — see `EnginesRow`.
!macro EnginesRowVars NAME
  Var Engines${NAME}Label
  Var Engines${NAME}Pos
  Var Engines${NAME}Install
  Var Engines${NAME}Download
!macroend
!insertmacro EnginesRowVars Mpv
!insertmacro EnginesRowVars FFmpeg
!insertmacro EnginesRowVars LibreOffice
!insertmacro EnginesRowVars ImageMagick
!insertmacro EnginesRowVars PeaZip
!insertmacro EnginesRowVars Calibre
!insertmacro EnginesRowVars F3D
!insertmacro EnginesRowVars Office

Page custom EnginesShow EnginesLeave

; 7. Installation page
!insertmacro MUI_PAGE_INSTFILES

; 8. Finish page
;
; Don't auto jump to finish page after installation page,
; because the installation page has useful info that can be used debug any issues with the installer.
!define MUI_FINISHPAGE_NOAUTOCLOSE
; Use show readme button in the finish page as a button create a desktop shortcut
!define MUI_FINISHPAGE_SHOWREADME
!define MUI_FINISHPAGE_SHOWREADME_TEXT "$(createDesktop)"
!define MUI_FINISHPAGE_SHOWREADME_FUNCTION CreateDesktopShortcut
; Show run app after installation.
!define MUI_FINISHPAGE_RUN "$INSTDIR\${MAINBINARYNAME}.exe"
!define MUI_PAGE_CUSTOMFUNCTION_PRE SkipIfPassive
!insertmacro MUI_PAGE_FINISH

; Uninstaller Pages
; 1. Confirm uninstall page
{{#if appdata_paths}}
Var DeleteAppDataCheckbox
Var DeleteAppDataCheckboxState
!define /ifndef WS_EX_LAYOUTRTL         0x00400000
!define MUI_PAGE_CUSTOMFUNCTION_SHOW un.ConfirmShow
Function un.ConfirmShow
    FindWindow $1 "#32770" "" $HWNDPARENT ; Find inner dialog
    ${If} $(^RTL) == 1
      System::Call 'USER32::CreateWindowEx(i${__NSD_CheckBox_EXSTYLE}|${WS_EX_LAYOUTRTL},t"${__NSD_CheckBox_CLASS}",t "$(deleteAppData)",i${__NSD_CheckBox_STYLE},i 50,i 100,i 400, i 25,i$1,i0,i0,i0)i.s'
    ${Else}
      System::Call 'USER32::CreateWindowEx(i${__NSD_CheckBox_EXSTYLE},t"${__NSD_CheckBox_CLASS}",t "$(deleteAppData)",i${__NSD_CheckBox_STYLE},i 0,i 100,i 400, i 25,i$1,i0,i0,i0)i.s'
    ${EndIf}
    Pop $DeleteAppDataCheckbox
    SendMessage $HWNDPARENT ${WM_GETFONT} 0 0 $1
    SendMessage $DeleteAppDataCheckbox ${WM_SETFONT} $1 1
FunctionEnd
!define MUI_PAGE_CUSTOMFUNCTION_LEAVE un.ConfirmLeave
Function un.ConfirmLeave
    SendMessage $DeleteAppDataCheckbox ${BM_GETCHECK} 0 0 $DeleteAppDataCheckboxState
FunctionEnd
{{/if}}
!insertmacro MUI_UNPAGE_CONFIRM

; 2. Uninstalling Page
!insertmacro MUI_UNPAGE_INSTFILES

;Languages
{{#each languages}}
!insertmacro MUI_LANGUAGE "{{this}}"
{{/each}}
!insertmacro MUI_RESERVEFILE_LANGDLL
{{#each language_files}}
  !include "{{this}}"
{{/each}}

; The words this installer's own page carries. Every other string it shows is an id one of
; the language files above defines — `$(createDesktop)`, `$(deleteAppData)` — and these are
; the page's own, so they are written here. A build that adds a language needs a line for it
; here as much as it needs the language file above.
; The words the optional-engines page carries: what each row is called, with a few of the
; formats that engine would give previews to, and the two words a row can carry to its right.
; `enginesLibreOffice` is the name a machine with no Office sees and `enginesLibreNiche` the one
; a machine with it does: the Office formats are covered there already, and what LibreOffice
; adds is the formats beside them. `enginesMPV` names the formats `enginesFFmpeg` does: mpv plays
; what FFmpeg's player plays, and the app's own `Codecs` row for it carries no list of its own.
;
; The checkmark is a square root sign, which is a fact about the font rather than a choice: the
; dialog font these pages are drawn with carries `√` and has no glyph at all for `✔` or `✓`,
; which would be drawn as a box.
LangString enginesTitle ${LANG_ENGLISH} "Optional engines"
LangString enginesSubtitle ${LANG_ENGLISH} "Previews that need something this app does not ship"
LangString enginesIntro ${LANG_ENGLISH} "These engines are optional: without one, the files it draws simply show no preview. Install runs winget here, in a console window and without administrator rights; Manual Download opens the page instead."
LangString enginesDetected ${LANG_ENGLISH} "Detected"
LangString enginesInstall ${LANG_ENGLISH} "Install"
LangString enginesManualDownload ${LANG_ENGLISH} "Manual Download"
LangString enginesInstallRunning ${LANG_ENGLISH} "Installing in the console window that just opened. This page waits until it finishes: closing that window aborts the install."
LangString enginesInstallFailed ${LANG_ENGLISH} "The install did not succeed. The console window it opened says why, and the row is unchanged."
LangString enginesInstallCode ${LANG_ENGLISH} "Exit code:"
LangString enginesInstallTimedOut ${LANG_ENGLISH} "The install did not finish within half an hour. Its console window may still be running, and the row is unchanged."
LangString enginesInstallNotFound ${LANG_ENGLISH} "The install reported success, but the engine is not where this page looks. Restarting the installer may see it."
LangString enginesInstallDone ${LANG_ENGLISH} "Installed. The engine is used by the app without a restart."
LangString enginesMPV ${LANG_ENGLISH} "MPV (flv, rmvb, mxf, ogv, swf, ...)"
LangString enginesFFmpeg ${LANG_ENGLISH} "FFmpeg (flv, rmvb, mxf, ogv, swf, ...)"
LangString enginesLibreOffice ${LANG_ENGLISH} "LibreOffice (doc, docx, xls, xlsx, ppt, pptx, ...)"
LangString enginesLibreNiche ${LANG_ENGLISH} "LibreOffice (cdr, odt, ods, odp, odg, ...)"
LangString enginesImageMagick ${LANG_ENGLISH} "ImageMagick (nef, cr2, cr3, arw, dng, raf, ...)"
LangString enginesPeaZip ${LANG_ENGLISH} "PeaZip (cab, iso, rpm, deb, arj, lzh, ...)"
LangString enginesCalibre ${LANG_ENGLISH} "Calibre (mobi, azw3, epub, djvu, fb2, ...)"
LangString enginesF3D ${LANG_ENGLISH} "f3d (stl, step, glb, gltf, fbx, obj, ...)"
LangString enginesOffice ${LANG_ENGLISH} "Microsoft Office (doc, docx, xls, xlsx, ppt, pptx, ...)"

!macro SetContext
  !if "${INSTALLMODE}" == "currentUser"
    SetShellVarContext current
  !else if "${INSTALLMODE}" == "perMachine"
    SetShellVarContext all
  !endif

  ${If} ${RunningX64}
    !if "${ARCH}" == "x64"
      SetRegView 64
    !else if "${ARCH}" == "arm64"
      SetRegView 64
    !else
      SetRegView 32
    !endif
  ${EndIf}
!macroend

Var PassiveMode
Function .onInit
  ${GetOptions} $CMDLINE "/P" $PassiveMode
  IfErrors +2 0
    StrCpy $PassiveMode 1

  !if "${DISPLAYLANGUAGESELECTOR}" == "true"
    !insertmacro MUI_LANGDLL_DISPLAY
  !endif

  !insertmacro SetContext

  ${If} $INSTDIR == ""
    ; Set default install location
    !if "${INSTALLMODE}" == "perMachine"
      ${If} ${RunningX64}
        !if "${ARCH}" == "x64"
          StrCpy $INSTDIR "$PROGRAMFILES64\${PRODUCTNAME}"
        !else if "${ARCH}" == "arm64"
          StrCpy $INSTDIR "$PROGRAMFILES64\${PRODUCTNAME}"
        !else
          StrCpy $INSTDIR "$PROGRAMFILES\${PRODUCTNAME}"
        !endif
      ${Else}
        StrCpy $INSTDIR "$PROGRAMFILES\${PRODUCTNAME}"
      ${EndIf}
    !else if "${INSTALLMODE}" == "currentUser"
      StrCpy $INSTDIR "$LOCALAPPDATA\${PRODUCTNAME}"
    !endif

    Call RestorePreviousInstallLocation
  ${EndIf}


  !if "${INSTALLMODE}" == "both"
    !insertmacro MULTIUSER_INIT
  !endif
FunctionEnd


Section EarlyChecks
  ; Abort silent installer if downgrades is disabled
  !if "${ALLOWDOWNGRADES}" == "false"
  IfSilent 0 silent_downgrades_done
    ; If downgrading
    ${If} $R0 == -1
      System::Call 'kernel32::AttachConsole(i -1)i.r0'
      ${If} $0 != 0
        System::Call 'kernel32::GetStdHandle(i -11)i.r0'
        System::call 'kernel32::SetConsoleTextAttribute(i r0, i 0x0004)' ; set red color
        FileWrite $0 "$(silentDowngrades)"
      ${EndIf}
      Abort
    ${EndIf}
  silent_downgrades_done:
  !endif

SectionEnd

{{#if preinstall_section}}
{{unescape_newlines preinstall_section}}
{{/if}}

!macro CheckIfAppIsRunning
  nsis_tauri_utils::FindProcess "${MAINBINARYNAME}.exe"
  Pop $R0
  ${If} $R0 = 0
      ; A running copy is terminated without asking: this installer is a
      ; replacement for it, and a prompt only stands between the user and the
      ; update.
      StrCpy $R1 0
      kill:
        nsis_tauri_utils::KillProcess "${MAINBINARYNAME}.exe"
        Pop $R0
        Sleep 500
        nsis_tauri_utils::FindProcess "${MAINBINARYNAME}.exe"
        Pop $R0
        ${If} $R0 = 0
          IntOp $R1 $R1 + 1
          ${If} $R1 < 3
            Goto kill
          ${EndIf}
          IfSilent silent ui
          silent:
            System::Call 'kernel32::AttachConsole(i -1)i.r0'
            ${If} $0 != 0
              System::Call 'kernel32::GetStdHandle(i -11)i.r0'
              System::call 'kernel32::SetConsoleTextAttribute(i r0, i 0x0004)' ; set red color
              FileWrite $0 "$(appRunning)$\n"
            ${EndIf}
            Abort
          ui:
            Abort "$(failedToKillApp)"
        ${EndIf}
  ${EndIf}
!macroend

Section Install
  ; A running copy is closed, and a previous installation of the product is
  ; removed, before anything is written.
  !insertmacro CheckIfAppIsRunning
  Call UninstallPreviousInstallation

  SetOutPath $INSTDIR

  ; Copy main executable
  File "${MAINBINARYSRCPATH}"

  ; Create resources directory structure
  {{#each resources_dirs}}
    CreateDirectory "$INSTDIR\\{{this}}"
  {{/each}}

  ; Copy resources
  {{#each resources}}
    File /a "/oname={{this}}" "{{@key}}"
  {{/each}}

  ; Copy external binaries
  {{#each binaries}}
    File /a "/oname={{this}}" "{{@key}}"
  {{/each}}

  ; Installed apps draws DisplayIcon from a file that stays put: the exe is
  ; replaced by an update mid-draw, which is how the entry lost its icon.
  !if "${INSTALLERICON}" != ""
    File /oname=app.ico "${INSTALLERICON}"
  !endif

  ; Create file associations
  {{#each file_associations as |association| ~}}
    {{#each association.extensions as |ext| ~}}
       !insertmacro APP_ASSOCIATE "{{ext}}" "{{or association.name ext}}" "{{association-description association.description ext}}" "$INSTDIR\${MAINBINARYNAME}.exe,0" "Open with ${PRODUCTNAME}" "$INSTDIR\${MAINBINARYNAME}.exe $\"%1$\""
    {{/each}}
  {{/each}}

  ; Register deep links
  {{#each deep_link_protocols as |protocol| ~}}
    WriteRegStr SHCTX "Software\Classes\\{{protocol}}" "URL Protocol" ""
    WriteRegStr SHCTX "Software\Classes\\{{protocol}}" "" "URL:${BUNDLEID} protocol"
    WriteRegStr SHCTX "Software\Classes\\{{protocol}}\DefaultIcon" "" "$\"$INSTDIR\${MAINBINARYNAME}.exe$\",0"
    WriteRegStr SHCTX "Software\Classes\\{{protocol}}\shell\open\command" "" "$\"$INSTDIR\${MAINBINARYNAME}.exe$\" $\"%1$\""
  {{/each}}

  ; Create uninstaller
  WriteUninstaller "$INSTDIR\uninstall.exe"

  ; Save $INSTDIR in registry for future installations
  WriteRegStr SHCTX "${MANUPRODUCTKEY}" "" $INSTDIR

  !if "${INSTALLMODE}" == "both"
    ; Save install mode to be selected by default for the next installation such as updating
    ; or when uninstalling
    WriteRegStr SHCTX "${UNINSTKEY}" $MultiUser.InstallMode 1
  !endif

  ; Registry information for add/remove programs
  WriteRegStr SHCTX "${UNINSTKEY}" "DisplayName" "${PRODUCTNAME}"
  !if "${INSTALLERICON}" != ""
    WriteRegStr SHCTX "${UNINSTKEY}" "DisplayIcon" "$\"$INSTDIR\app.ico$\",0"
  !else
    WriteRegStr SHCTX "${UNINSTKEY}" "DisplayIcon" "$\"$INSTDIR\${MAINBINARYNAME}.exe$\",0"
  !endif
  WriteRegStr SHCTX "${UNINSTKEY}" "DisplayVersion" "${VERSION}"
  WriteRegStr SHCTX "${UNINSTKEY}" "Publisher" "${MANUFACTURER}"
  WriteRegStr SHCTX "${UNINSTKEY}" "InstallLocation" "$\"$INSTDIR$\""
  WriteRegStr SHCTX "${UNINSTKEY}" "UninstallString" "$\"$INSTDIR\uninstall.exe$\""
  WriteRegDWORD SHCTX "${UNINSTKEY}" "NoModify" "1"
  WriteRegDWORD SHCTX "${UNINSTKEY}" "NoRepair" "1"
  WriteRegDWORD SHCTX "${UNINSTKEY}" "EstimatedSize" "${ESTIMATEDSIZE}"

  ; Create start menu shortcut (GUI)
  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
    Call CreateStartMenuShortcut
  !insertmacro MUI_STARTMENU_WRITE_END

  ; Create shortcuts for silent and passive installers, which
  ; can be disabled by passing `/NS` flag
  ; GUI installer has buttons for users to control creating them
  IfSilent check_ns_flag 0
  ${IfThen} $PassiveMode == 1 ${|} Goto check_ns_flag ${|}
  Goto shortcuts_done
  check_ns_flag:
    ${GetOptions} $CMDLINE "/NS" $R0
    IfErrors 0 shortcuts_done
      Call CreateDesktopShortcut
      Call CreateStartMenuShortcut
  shortcuts_done:

  ; Auto close this page for passive mode
  ${IfThen} $PassiveMode == 1 ${|} SetAutoClose true ${|}
SectionEnd

Function .onInstSuccess
  ; Check for `/R` flag only in silent and passive installers because
  ; GUI installer has a toggle for the user to (re)start the app
  IfSilent check_r_flag 0
  ${IfThen} $PassiveMode == 1 ${|} Goto check_r_flag ${|}
  Goto run_done
  check_r_flag:
    ${GetOptions} $CMDLINE "/R" $R0
    IfErrors run_done 0
      Exec '"$INSTDIR\${MAINBINARYNAME}.exe"'
  run_done:
FunctionEnd

Function un.onInit
  !insertmacro SetContext

  !if "${INSTALLMODE}" == "both"
    !insertmacro MULTIUSER_UNINIT
  !endif

  !insertmacro MUI_UNGETLANGUAGE
FunctionEnd

Section Uninstall
  !insertmacro CheckIfAppIsRunning

  ; Delete the app directory and its content from disk
  ; Copy main executable
  Delete "$INSTDIR\${MAINBINARYNAME}.exe"
  !if "${INSTALLERICON}" != ""
    Delete "$INSTDIR\app.ico"
  !endif

  ; Delete resources
  {{#each resources}}
    Delete "$INSTDIR\\{{this}}"
  {{/each}}

  ; Delete external binaries
  {{#each binaries}}
    Delete "$INSTDIR\\{{this}}"
  {{/each}}

  ; Delete app associations
  {{#each file_associations as |association| ~}}
    {{#each association.ext as |ext| ~}}
      !insertmacro APP_UNASSOCIATE "{{ext}}" "{{or association.name ext}}"
    {{/each}}
  {{/each}}

  ; Delete deep links
  {{#each deep_link_protocols as |protocol| ~}}
    ReadRegStr $R7 SHCTX "Software\Classes\\{{protocol}}\shell\open\command" ""
    !if $R7 == "$\"$INSTDIR\${MAINBINARYNAME}.exe$\" $\"%1$\""
      DeleteRegKey SHCTX "Software\Classes\\{{protocol}}"
    !endif
  {{/each}}

  ; Delete uninstaller
  Delete "$INSTDIR\uninstall.exe"

  {{#each resources_dirs}}
  RMDir /REBOOTOK "$INSTDIR\\{{this}}"
  {{/each}}
  RMDir "$INSTDIR"

  ; Remove start menu shortcut
  !insertmacro MUI_STARTMENU_GETFOLDER Application $AppStartMenuFolder
  Delete "$SMPROGRAMS\$AppStartMenuFolder\${PRODUCTNAME}.lnk"
  RMDir "$SMPROGRAMS\$AppStartMenuFolder"

  ; Remove desktop shortcuts
  Delete "$DESKTOP\${PRODUCTNAME}.lnk"

  ; Remove registry information for add/remove programs
  !if "${INSTALLMODE}" == "both"
    DeleteRegKey SHCTX "${UNINSTKEY}"
  !else if "${INSTALLMODE}" == "perMachine"
    DeleteRegKey HKLM "${UNINSTKEY}"
  !else
    DeleteRegKey HKCU "${UNINSTKEY}"
  !endif

  DeleteRegValue HKCU "${MANUPRODUCTKEY}" "Installer Language"

  ; Delete app data
  {{#if appdata_paths}}
  ${If} $DeleteAppDataCheckboxState == 1
      SetShellVarContext current
      {{#each appdata_paths}}
      RmDir /r "{{unescape_dollar_sign this}}"
      {{/each}}
  ${EndIf}
  {{/if}}

  ${GetOptions} $CMDLINE "/P" $R0
  IfErrors +2 0
    SetAutoClose true
SectionEnd

Function RestorePreviousInstallLocation
  ReadRegStr $4 SHCTX "${MANUPRODUCTKEY}" ""
  StrCmp $4 "" +2 0
    StrCpy $INSTDIR $4
FunctionEnd

; Removes a previous installation without asking: this installer replaces it,
; so the maintenance page the upstream template shows has nothing to ask.
Function UninstallPreviousInstallation
  ; An installation made by the WiX/MSI installer keeps its entry under a UUID
  ; and is removed by Windows Installer.
  StrCpy $0 0
  wix_loop:
    EnumRegKey $1 HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" $0
    StrCmp $1 "" wix_done ; Exit loop if there is no more keys to loop on
    IntOp $0 $0 + 1
    ReadRegStr $R0 HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$1" "DisplayName"
    ReadRegStr $R1 HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$1" "Publisher"
    StrCmp "$R0$R1" "${PRODUCTNAME}${MANUFACTURER}" 0 wix_loop
    ReadRegStr $R2 HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$1" "UninstallString"
    ${StrCase} $R1 $R2 "L"
    ${StrLoc} $R0 $R1 "msiexec" ">"
    StrCmp $R0 0 0 wix_done
    ; The product code sits between the braces of the uninstall string, which
    ; is all Windows Installer needs to remove the product silently.
    ${StrLoc} $R1 $R2 "{" ">"
    StrCmp $R1 "" wix_done
    StrCpy $R2 $R2 -1 $R1
    ${StrLoc} $R3 $R2 "}" ">"
    StrCmp $R3 "" wix_done
    IntOp $R3 $R3 + 1
    StrCpy $R2 $R2 $R3
    ExecWait 'msiexec /x "$R2" /qn /norestart' $0
  wix_done:

  ; An installation made by this installer is removed by its own uninstaller,
  ; run in the directory it was installed into. It is run silently, which is
  ; also what keeps the user's configuration: the uninstaller asks about the
  ; application data, and a silent one is answered with no.
  ReadRegStr $R0 SHCTX "${UNINSTKEY}" ""
  ReadRegStr $R1 SHCTX "${UNINSTKEY}" "UninstallString"
  ${IfThen} "$R0$R1" == "" ${|} Return ${|}

  ReadRegStr $4 SHCTX "${MANUPRODUCTKEY}" ""
  ${IfThen} $4 == "" ${|} StrCpy $4 $INSTDIR ${|}

  StrCpy $0 $R1 1
  ${IfThen} $0 == '"' ${|} StrCpy $R1 $R1 -1 1 ${|} ; Strip quotes from UninstallString
  ; `_?=` is read to the end of the command line, so the path is left
  ; unquoted and kept last.
  ExecWait '"$R1" /S _?=$4' $0

  ${If} $0 <> 0
  ${OrIf} ${FileExists} "$4\${MAINBINARYNAME}.exe"
    MessageBox MB_ICONEXCLAMATION "$(unableToUninstall)"
    Abort
  ${EndIf}

  Delete "$R1"
  RMDir "$4"
FunctionEnd

Function SkipIfPassive
  ${IfThen} $PassiveMode == 1  ${|} Abort ${|}
FunctionEnd

; The optional-engines page: one row per engine this app can drive, marked with whether this
; machine already has it and carrying a link to the page it is installed from where it does not.
;
; A row that is present is the same label with `√` in front of it, and a row that is missing is
; that label greyed — which is what disabling a static does, since the dialog manager draws a
; disabled control's text in the grey the platform keeps for it.
;
; Each engine is asked about one at a time, and each detector answers in `$0`: 1 where the engine
; is installed and 0 where it is not.
;
; Whether `$3` is inside a winget package under `$1` whose folder name starts with `$2`: a package
; folder is named for the package and the source it came from
; (`shinchiro.mpv_Microsoft.Winget.Source_8wekyb3d8bbwe`), so the name is what is matched. The
; payload is asked for both in that folder and in a `bin` beneath it, which is where a portable
; package that ships its own tree puts it.
;
; Answers in `$0`, like every detector here, and uses `$4`/`$5` for the walk.
Function EnginesWinGetScan
  StrCpy $0 0
  FindFirst $4 $5 "$1\$2*"
  engines_winget_scan:
    StrCmp $5 "" engines_winget_scan_done
    ${If} ${FileExists} "$1\$5\$3"
      StrCpy $0 1
      Goto engines_winget_scan_done
    ${EndIf}
    ${If} ${FileExists} "$1\$5\bin\$3"
      StrCpy $0 1
      Goto engines_winget_scan_done
    ${EndIf}
    FindNext $4 $5
    Goto engines_winget_scan
  engines_winget_scan_done:
  FindClose $4
FunctionEnd

; mpv is asked for where the app asks for it (see `codecs::mpv_installs`): beside the app, in a
; folder on the `PATH`, in the two folders its own installers write to — `MPV Player` under
; Program Files and SVP 4's `mpv64` — and in what winget leaves, which is where an mpv installed
; from this page lands: its shim in `…\WinGet\Links`, its per-user Inno destination under
; `…\Local\Programs`, and the portable package's own folder. mpv's per-user destination is not
; verified, so every plausible folder under `Programs` named for it is asked for rather than one
; guessed at. What the app also honors is a copy the user picked in its own settings, which is a
; path in the app's configuration; this installer reads none of that, so a machine whose only mpv
; is a chosen copy is one this page does not know about.
Function MPVInstalled
  StrCpy $0 0
  ${If} ${FileExists} "$INSTDIR\mpv.exe"
    StrCpy $0 1
    Return
  ${EndIf}
  SearchPath $1 "mpv.exe"
  ${IfNot} ${Errors}
    StrCpy $0 1
    Return
  ${EndIf}
  ${If} ${FileExists} "$LOCALAPPDATA\Microsoft\WinGet\Links\mpv.exe"
    StrCpy $0 1
    Return
  ${EndIf}
  StrCpy $1 "$LOCALAPPDATA\Programs"
  StrCpy $2 "mpv"
  StrCpy $3 "mpv.exe"
  Call EnginesWinGetScan
  ${If} $0 == 1
    Return
  ${EndIf}
  StrCpy $1 "$LOCALAPPDATA\Microsoft\WinGet\Packages"
  StrCpy $2 "shinchiro.mpv"
  StrCpy $3 "mpv.exe"
  Call EnginesWinGetScan
  ${If} $0 == 1
    Return
  ${EndIf}
  ${If} ${FileExists} "$PROGRAMFILES64\MPV Player\mpv.exe"
    StrCpy $0 1
    Return
  ${EndIf}
  ${If} ${FileExists} "$PROGRAMFILES\SVP 4\mpv64\mpv.exe"
    StrCpy $0 1
  ${EndIf}
FunctionEnd

; FFmpeg is asked for the player beside the app first, which is where an unpacked copy sits and
; where the app looks before it looks anywhere else (see `codecs::find_ffplay`), then in a folder
; on the `PATH`, and then in what winget leaves: the `Links` shim is what a portable package
; always gets, while the package folder itself holds a build folder one level deeper still
; (`ffmpeg-7.1-full_build\bin`), which the scan below does not walk.
Function FFmpegInstalled
  StrCpy $0 0
  ${If} ${FileExists} "$INSTDIR\ffplay.exe"
    StrCpy $0 1
    Return
  ${EndIf}
  SearchPath $1 "ffplay.exe"
  ${IfNot} ${Errors}
    StrCpy $0 1
    Return
  ${EndIf}
  ${If} ${FileExists} "$LOCALAPPDATA\Microsoft\WinGet\Links\ffplay.exe"
    StrCpy $0 1
    Return
  ${EndIf}
  StrCpy $1 "$LOCALAPPDATA\Microsoft\WinGet\Packages"
  StrCpy $2 "Gyan.FFmpeg"
  StrCpy $3 "ffplay.exe"
  Call EnginesWinGetScan
FunctionEnd

Function LibreOfficeInstalled
  StrCpy $0 0
  ${If} ${FileExists} "$PROGRAMFILES64\LibreOffice\program\soffice.exe"
    StrCpy $0 1
    Return
  ${EndIf}
  ${If} ${FileExists} "$PROGRAMFILES\LibreOffice\program\soffice.exe"
    StrCpy $0 1
  ${EndIf}
FunctionEnd

; ImageMagick installs into a folder named for its version — `ImageMagick-7.1.2-Q16-HDRI` and
; the like — so the folders under a program directory are read rather than a name guessed at.
; A folder counts where one of the two tools inside it is there, which is the same two names, in
; the same order, that the app itself looks for.
Function ImageMagickScan
  ${If} $1 == ""
    Return
  ${EndIf}
  FindFirst $2 $3 "$1\ImageMagick*"
  imagemagick_loop:
    StrCmp $3 "" imagemagick_done
    ${If} ${FileExists} "$1\$3\magick.exe"
      StrCpy $0 1
      Goto imagemagick_done
    ${EndIf}
    ${If} ${FileExists} "$1\$3\convert.exe"
      StrCpy $0 1
      Goto imagemagick_done
    ${EndIf}
    FindNext $2 $3
    Goto imagemagick_loop
  imagemagick_done:
  FindClose $2
FunctionEnd

Function ImageMagickInstalled
  StrCpy $0 0
  StrCpy $1 "$PROGRAMFILES64"
  Call ImageMagickScan
  ${If} $0 == 1
    Return
  ${EndIf}
  StrCpy $1 "$PROGRAMFILES"
  Call ImageMagickScan
FunctionEnd

; PeaZip is asked for the console archiver the app runs rather than for the application's own
; windowed frontend: a window opening is what a PeaZip preview is not, and what the app looks
; for is `res\bin\7z\7z.exe` inside the installation.
Function PeaZipInstalled
  StrCpy $0 0
  ${If} ${FileExists} "$PROGRAMFILES64\PeaZip\res\bin\7z\7z.exe"
    StrCpy $0 1
    Return
  ${EndIf}
  ${If} ${FileExists} "$PROGRAMFILES\PeaZip\res\bin\7z\7z.exe"
    StrCpy $0 1
  ${EndIf}
FunctionEnd

Function CalibreInstalled
  StrCpy $0 0
  ${If} ${FileExists} "$PROGRAMFILES64\Calibre2\ebook-convert.exe"
    StrCpy $0 1
    Return
  ${EndIf}
  ${If} ${FileExists} "$PROGRAMFILES\Calibre2\ebook-convert.exe"
    StrCpy $0 1
  ${EndIf}
FunctionEnd

; f3d is asked for where the app asks for it (see `f3d_render::program`): beside the app, in a
; folder on the `PATH`, and in the folder its own Windows installer writes, `f3d\bin` under
; Program Files. What the app also honors is a copy the user named in its own settings
; (`f3d_path`), which this installer reads none of — the same silence a chosen mpv gets.
Function F3DInstalled
  StrCpy $0 0
  ${If} ${FileExists} "$INSTDIR\f3d.exe"
    StrCpy $0 1
    Return
  ${EndIf}
  SearchPath $1 "f3d.exe"
  ${IfNot} ${Errors}
    StrCpy $0 1
    Return
  ${EndIf}
  ${If} ${FileExists} "$PROGRAMFILES64\f3d\bin\f3d.exe"
    StrCpy $0 1
    Return
  ${EndIf}
  ${If} ${FileExists} "$PROGRAMFILES\f3d\bin\f3d.exe"
    StrCpy $0 1
  ${EndIf}
FunctionEnd

; Whether every Office application this app draws a page with is installed. It marks the Office
; row, and what the LibreOffice row is called follows it: with all three here the Office formats
; are covered and what LibreOffice adds is the formats beside them, while a machine with none —
; or with only some — is shown the Office formats LibreOffice draws.
;
; What is read is the ProgID each application registers, which is the same `CLSIDFromProgID`
; question the app asks before it drives one.
Function OfficeInstalled
  StrCpy $0 1
  ReadRegStr $1 HKCR "Word.Application\CLSID" ""
  ${If} $1 == ""
    StrCpy $0 0
  ${EndIf}
  ReadRegStr $1 HKCR "Excel.Application\CLSID" ""
  ${If} $1 == ""
    StrCpy $0 0
  ${EndIf}
  ReadRegStr $1 HKCR "PowerPoint.Application\CLSID" ""
  ${If} $1 == ""
    StrCpy $0 0
  ${EndIf}
FunctionEnd

; One row of the page, at `POS` units down: `LABEL`, marked with `√` and answered `Detected`
; where the detector `CHECK` says the engine is here, and — where it is not — the same label greyed
; beside an `Install` link where `PKG` is a winget package and winget was found, and a
; `Manual Download` link that opens the page the engine is installed from. `-` for `PKG` means no
; package, which is f3d's row and Office's.
;
; Every link's text carries anchor markup, and that is not decoration: `${NSD_CreateLink}` makes a
; SysLink, and a SysLink whose text is plain draws itself as a link but holds no link at all — it
; never sends `NM_CLICK`, so a function hung off `NSD_OnNotify` is one that never runs. The `id`
; names the anchor within `NMLINK`; a click on a row's two links is two controls and so two
; callbacks, because a WM_NOTIFY does not say which anchor of one control was hit.
;
; `NAME` names the globals this row keeps its own controls in, so that the row can be redrawn
; detected in place once an install has put the engine there (see `EnginesInstallStart`). They are
; declared at the top, beside the other `Var`s, because a `Var` inside a function — which is where
; this macro is inserted — is not one NSIS makes.
!macro EnginesRow NAME POS CHECK LABEL PKG LINK
  Call ${CHECK}
  ${If} $0 == 1
    ${NSD_CreateLabel} 0 ${POS}u 74% 12u "√ ${LABEL}"
    Pop $0
    ${NSD_CreateLabel} 75% ${POS}u 25% 12u "$(enginesDetected)"
    Pop $0
  ${Else}
    ${NSD_CreateLabel} 0 ${POS}u 52% 12u "${LABEL}"
    Pop $Engines${NAME}Label
    EnableWindow $Engines${NAME}Label 0
    StrCpy $Engines${NAME}Pos "${POS}u"
    StrCpy $Engines${NAME}Install ""
    ; `!if` and not `${If}`: the package being `-` has to keep the handler's name out of the
    ; compiled page, because there is no `EnginesInstallF3D` or `EnginesInstallOffice` function to
    ; resolve and a runtime branch would still name one. Whether winget was found is the runtime
    ; half, and it needs no name of its own.
    !if "${PKG}" != "-"
      ${If} $EnginesWinget != ""
        ${NSD_CreateLink} 54% ${POS}u 16% 12u '<a id="i">$(enginesInstall)</a>'
        Pop $Engines${NAME}Install
        ${NSD_OnNotify} $Engines${NAME}Install EnginesInstall${NAME}
      ${EndIf}
    !endif
    ${NSD_CreateLink} 71% ${POS}u 29% 12u '<a id="d">$(enginesManualDownload)</a>'
    Pop $Engines${NAME}Download
    ${NSD_OnNotify} $Engines${NAME}Download ${LINK}
  ${EndIf}
!macroend

Function EnginesShow
  ${If} ${Silent}
    Abort
  ${EndIf}
  ${IfThen} $PassiveMode == 1 ${|} Abort ${|}

  !insertmacro MUI_HEADER_TEXT "$(enginesTitle)" "$(enginesSubtitle)"

  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}

  ${NSD_CreateLabel} 0 0 100% 30u "$(enginesIntro)"
  Pop $EnginesIntro

  StrCpy $EnginesBusy 0

  ; winget, found the way a shell finds it and without starting it: the alias Windows keeps in
  ; this user's `WindowsApps`, the Store package's own folder — named for its version, so the folder
  ; is what is walked — and the `PATH`. Empty means absent, and then no row is offered an `Install`
  ; link at all rather than one that cannot work.
  StrCpy $EnginesWinget ""
  ${If} ${FileExists} "$LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
    StrCpy $EnginesWinget "$LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
  ${Else}
    FindFirst $0 $1 "$PROGRAMFILES\WindowsApps\Microsoft.DesktopAppInstaller_*"
    engines_winget_folder:
      StrCmp $1 "" engines_winget_folder_done
      ${If} ${FileExists} "$PROGRAMFILES\WindowsApps\$1\winget.exe"
        StrCpy $EnginesWinget "$PROGRAMFILES\WindowsApps\$1\winget.exe"
        Goto engines_winget_folder_done
      ${EndIf}
      FindNext $0 $1
      Goto engines_winget_folder
    engines_winget_folder_done:
    FindClose $0
    ${If} $EnginesWinget == ""
      SearchPath $1 "winget.exe"
      ${IfNot} ${Errors}
        StrCpy $EnginesWinget $1
      ${EndIf}
    ${EndIf}
  ${EndIf}

  ; mpv's row is the first, as mpv is the first of the engines the app lists for a video: where
  ; more than one player is installed, it is the one `Best` prefers.
  !insertmacro EnginesRow Mpv 32 MPVInstalled $(enginesMPV) shinchiro.mpv EnginesLinkMPV
  !insertmacro EnginesRow FFmpeg 46 FFmpegInstalled $(enginesFFmpeg) Gyan.FFmpeg EnginesLinkFFmpeg

  ; What LibreOffice is called here is decided before it is asked about: the Office formats are
  ; worth naming where there is no Office to draw them, and the formats beside them are what the
  ; engine is worth where there is one.
  StrCpy $EnginesLibreText "$(enginesLibreOffice)"
  Call OfficeInstalled
  ${If} $0 == 1
    StrCpy $EnginesLibreText "$(enginesLibreNiche)"
  ${EndIf}

  !insertmacro EnginesRow LibreOffice 60 LibreOfficeInstalled $EnginesLibreText TheDocumentFoundation.LibreOffice EnginesLinkLibreOffice
  !insertmacro EnginesRow ImageMagick 74 ImageMagickInstalled $(enginesImageMagick) ImageMagick.ImageMagick EnginesLinkImageMagick
  !insertmacro EnginesRow PeaZip 88 PeaZipInstalled $(enginesPeaZip) Giorgiotani.Peazip EnginesLinkPeaZip
  !insertmacro EnginesRow Calibre 102 CalibreInstalled $(enginesCalibre) calibre.calibre EnginesLinkCalibre

  ; f3d has no winget package, so its row is the manual download alone — the same page the app's
  ; own `Codecs` row opens, and the one its releases are taken from by hand.
  !insertmacro EnginesRow F3D 116 F3DInstalled $(enginesF3D) - EnginesLinkF3D

  ; Office's row offers no Install either: it is not a free package, so what it carries is the page
  ; Microsoft sells it from.
  !insertmacro EnginesRow Office 130 OfficeInstalled $(enginesOffice) - EnginesLinkOffice

  nsDialogs::Show
FunctionEnd

; Nothing is read from this page as it is left: it says what the machine has, and its links are
; what it does, so there is no state to collect. The page above has two boxes and this one has
; none — which is why this function is empty rather than absent, since a page is declared with
; both of its functions.
Function EnginesLeave
FunctionEnd

; The page's links, a function each, so that the address a row opens is written by the row rather
; than looked up from a table. The control handle the notification carries is the first thing on
; the stack and is dropped: what a click does is one address, and which control it was is not part
; of the question.
;
; `ExecShell "open"` and not `explorer.exe`, because this installer is unelevated:
; `installMode = "currentUser"` in Cargo.toml takes the branch above that asks for
; `RequestExecutionLevel user`, so the page opens in the browser of the user who is looking at it.
; (The per-machine branch that would ask for `highest`, and with it the reason the address used to
; be handed to Explorer, is the one this build never takes.)
Function EnginesLinkMPV
  Pop $0
  ExecShell "open" "https://mpv.io/installation/"
FunctionEnd

Function EnginesLinkFFmpeg
  Pop $0
  ExecShell "open" "https://ffmpeg.org/download.html"
FunctionEnd

Function EnginesLinkLibreOffice
  Pop $0
  ExecShell "open" "https://www.libreoffice.org/download/"
FunctionEnd

Function EnginesLinkImageMagick
  Pop $0
  ExecShell "open" "https://imagemagick.org/download/"
FunctionEnd

Function EnginesLinkPeaZip
  Pop $0
  ExecShell "open" "https://peazip.github.io/peazip-64bit.html"
FunctionEnd

Function EnginesLinkCalibre
  Pop $0
  ExecShell "open" "https://calibre-ebook.com/download_windows"
FunctionEnd

Function EnginesLinkF3D
  Pop $0
  ExecShell "open" "https://github.com/f3d-app/f3d/releases"
FunctionEnd

; Office has nowhere to install from but Microsoft's own page, and nothing this page can run.
Function EnginesLinkOffice
  Pop $0
  ExecShell "open" "https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products"
FunctionEnd

; What one engine's `Install` link does: name its own row — its label, where it sits, its two links
; — and its package and the detector that asks whether it arrived, which is everything
; `EnginesInstallStart` knows about any engine, and hand over. Only the engines with a winget
; package have one of these, since the row does not create the link otherwise.
Function EnginesInstallMpv
  Pop $0
  StrCpy $EnginesId "shinchiro.mpv"
  StrCpy $EnginesCheck "MPVInstalled"
  StrCpy $EnginesRowLabel $EnginesMpvLabel
  StrCpy $EnginesRowPos $EnginesMpvPos
  StrCpy $EnginesRowInstall $EnginesMpvInstall
  StrCpy $EnginesRowDownload $EnginesMpvDownload
  Call EnginesInstallStart
FunctionEnd

Function EnginesInstallFFmpeg
  Pop $0
  StrCpy $EnginesId "Gyan.FFmpeg"
  StrCpy $EnginesCheck "FFmpegInstalled"
  StrCpy $EnginesRowLabel $EnginesFFmpegLabel
  StrCpy $EnginesRowPos $EnginesFFmpegPos
  StrCpy $EnginesRowInstall $EnginesFFmpegInstall
  StrCpy $EnginesRowDownload $EnginesFFmpegDownload
  Call EnginesInstallStart
FunctionEnd

Function EnginesInstallLibreOffice
  Pop $0
  StrCpy $EnginesId "TheDocumentFoundation.LibreOffice"
  StrCpy $EnginesCheck "LibreOfficeInstalled"
  StrCpy $EnginesRowLabel $EnginesLibreOfficeLabel
  StrCpy $EnginesRowPos $EnginesLibreOfficePos
  StrCpy $EnginesRowInstall $EnginesLibreOfficeInstall
  StrCpy $EnginesRowDownload $EnginesLibreOfficeDownload
  Call EnginesInstallStart
FunctionEnd

Function EnginesInstallImageMagick
  Pop $0
  StrCpy $EnginesId "ImageMagick.ImageMagick"
  StrCpy $EnginesCheck "ImageMagickInstalled"
  StrCpy $EnginesRowLabel $EnginesImageMagickLabel
  StrCpy $EnginesRowPos $EnginesImageMagickPos
  StrCpy $EnginesRowInstall $EnginesImageMagickInstall
  StrCpy $EnginesRowDownload $EnginesImageMagickDownload
  Call EnginesInstallStart
FunctionEnd

Function EnginesInstallPeaZip
  Pop $0
  StrCpy $EnginesId "Giorgiotani.Peazip"
  StrCpy $EnginesCheck "PeaZipInstalled"
  StrCpy $EnginesRowLabel $EnginesPeaZipLabel
  StrCpy $EnginesRowPos $EnginesPeaZipPos
  StrCpy $EnginesRowInstall $EnginesPeaZipInstall
  StrCpy $EnginesRowDownload $EnginesPeaZipDownload
  Call EnginesInstallStart
FunctionEnd

Function EnginesInstallCalibre
  Pop $0
  StrCpy $EnginesId "calibre.calibre"
  StrCpy $EnginesCheck "CalibreInstalled"
  StrCpy $EnginesRowLabel $EnginesCalibreLabel
  StrCpy $EnginesRowPos $EnginesCalibrePos
  StrCpy $EnginesRowInstall $EnginesCalibreInstall
  StrCpy $EnginesRowDownload $EnginesCalibreDownload
  Call EnginesInstallStart
FunctionEnd

; Half an hour of 250 ms polls. winget's own LibreOffice download is minutes of that; what
; outlasts it is answered as not finished, and the console window is left to finish on its own.
!define ENGINES_INSTALL_POLLS 7200

; The install itself, shared by every engine's `Install` link: the same batch script the app's own
; `Codecs → Install` rows write — one text and two doors, so the flags and the retry cannot drift
; apart — started in a console window the user watches, with the exit code read back out of the
; `.rc` the script leaves beside itself before it waits for a key.
;
; This function sleeps while the window works, so the page answers nothing until the install is
; done and the console window is the whole of the progress display; that is deliberate, and the
; status line is written before the wait and forced out with `UpdateWindow` because nothing the
; dialog manager would deliver is processed here while the loop runs.
;
; Nothing is timed out of its work: the cap below only ends this page's wait, and a row that was
; not re-probed detected is left exactly as it was rather than claimed.
Function EnginesInstallStart
  ${If} $EnginesBusy == 1
    Return
  ${EndIf}
  ${If} $EnginesWinget == ""
    Return
  ${EndIf}

  StrCpy $EnginesBusy 1
  ; The same two names the app's own `Codecs → Install` writes (see `src/app/winget.rs`), which is
  ; what makes the two doors one install rather than two that resemble each other.
  StrCpy $EnginesScript "$TEMP\rust-hover-preview-winget-$EnginesId.cmd"
  StrCpy $EnginesRc "$TEMP\rust-hover-preview-winget-$EnginesId.rc"

  ; An exit code an earlier run left would be read as this install's the moment it appeared.
  Delete "$EnginesRc"

  ClearErrors
  FileOpen $0 "$EnginesScript" w
  ${If} ${Errors}
    ClearErrors
    StrCpy $EnginesBusy 0
    ${NSD_SetText} $EnginesIntro "$(enginesInstallFailed)"
    Return
  ${EndIf}
  FileWrite $0 "@echo off$\r$\n"
  FileWrite $0 "title Rust Hover Preview - installing $EnginesId$\r$\n"
  FileWrite $0 '"$EnginesWinget" install -e --id $EnginesId --source winget --silent --accept-package-agreements --accept-source-agreements --disable-interactivity$\r$\n'
  FileWrite $0 "if errorlevel 1 ($\r$\n"
  FileWrite $0 "    echo.$\r$\n"
  FileWrite $0 "    echo First attempt failed. Updating the winget sources and retrying once.$\r$\n"
  FileWrite $0 "    echo.$\r$\n"
  FileWrite $0 '    "$EnginesWinget" source update$\r$\n'
  FileWrite $0 '    "$EnginesWinget" install -e --id $EnginesId --source winget --silent --accept-package-agreements --accept-source-agreements --disable-interactivity$\r$\n'
  FileWrite $0 ")$\r$\n"
  FileWrite $0 '> "$EnginesRc" echo %ERRORLEVEL%$\r$\n'
  FileWrite $0 "echo.$\r$\n"
  FileWrite $0 "echo Finished. This window can be closed.$\r$\n"
  FileWrite $0 "pause >nul$\r$\n"
  FileClose $0

  ${NSD_SetText} $EnginesIntro "$(enginesInstallRunning)"
  System::Call 'USER32::UpdateWindow(p $EnginesIntro)'

  ClearErrors
  Exec '"$SYSDIR\cmd.exe" /c ""$EnginesScript""'
  ${If} ${Errors}
    ClearErrors
    StrCpy $EnginesBusy 0
    ${NSD_SetText} $EnginesIntro "$(enginesInstallFailed)"
    Return
  ${EndIf}

  EnableWindow $HWNDPARENT 0
  StrCpy $EnginesCode ""
  StrCpy $EnginesWaited 0

  engines_install_wait:
    Sleep 250
    IntOp $EnginesWaited $EnginesWaited + 1
    ${If} ${FileExists} "$EnginesRc"
      ClearErrors
      FileOpen $1 "$EnginesRc" r
      ${IfNot} ${Errors}
        FileRead $1 $2
        FileClose $1
        StrCpy $EnginesCode $2 -2
      ${EndIf}
    ${EndIf}
    ${If} $EnginesCode == ""
    ${AndIf} $EnginesWaited < ${ENGINES_INSTALL_POLLS}
      Goto engines_install_wait
    ${EndIf}

  EnableWindow $HWNDPARENT 1
  StrCpy $EnginesBusy 0

  ${If} $EnginesCode == ""
    ${NSD_SetText} $EnginesIntro "$(enginesInstallTimedOut)"
    Return
  ${EndIf}

  ${If} $EnginesCode != "0"
    ${NSD_SetText} $EnginesIntro "$(enginesInstallFailed) $(enginesInstallCode) $EnginesCode"
    Return
  ${EndIf}

  ; The row is asked again rather than trusted: a package's installer reporting success and the
  ; engine standing where this page looks for it are two different claims.
  StrCmp $EnginesCheck "MPVInstalled" 0 +3
    Call MPVInstalled
    Goto engines_install_probed
  StrCmp $EnginesCheck "FFmpegInstalled" 0 +3
    Call FFmpegInstalled
    Goto engines_install_probed
  StrCmp $EnginesCheck "LibreOfficeInstalled" 0 +3
    Call LibreOfficeInstalled
    Goto engines_install_probed
  StrCmp $EnginesCheck "ImageMagickInstalled" 0 +3
    Call ImageMagickInstalled
    Goto engines_install_probed
  StrCmp $EnginesCheck "PeaZipInstalled" 0 +3
    Call PeaZipInstalled
    Goto engines_install_probed
  StrCmp $EnginesCheck "CalibreInstalled" 0 +3
    Call CalibreInstalled
    Goto engines_install_probed
  engines_install_probed:

  ${If} $0 != 1
    ${NSD_SetText} $EnginesIntro "$(enginesInstallNotFound)"
    Return
  ${EndIf}

  ; Redrawn detected in place: the row's own controls are hidden and a detected pair is put where
  ; they were, which is the layout the row would have been built with had the engine been here.
  ${NSD_GetText} $EnginesRowLabel $EnginesLabelText
  System::Call 'USER32::ShowWindow(p $EnginesRowLabel, i 0)'
  ${If} $EnginesRowInstall != ""
    System::Call 'USER32::ShowWindow(p $EnginesRowInstall, i 0)'
  ${EndIf}
  ${If} $EnginesRowDownload != ""
    System::Call 'USER32::ShowWindow(p $EnginesRowDownload, i 0)'
  ${EndIf}
  ${NSD_CreateLabel} 0 $EnginesRowPos 74% 12u "√ $EnginesLabelText"
  Pop $0
  ${NSD_CreateLabel} 75% $EnginesRowPos 25% 12u "$(enginesDetected)"
  Pop $0
  ${NSD_SetText} $EnginesIntro "$(enginesInstallDone)"
FunctionEnd

Function CreateDesktopShortcut
  CreateShortcut "$DESKTOP\${PRODUCTNAME}.lnk" "$INSTDIR\${MAINBINARYNAME}.exe"
  ApplicationID::Set "$DESKTOP\${PRODUCTNAME}.lnk" "${IDENTIFIER}"
FunctionEnd

Function CreateStartMenuShortcut
  CreateDirectory "$SMPROGRAMS\$AppStartMenuFolder"
  CreateShortcut "$SMPROGRAMS\$AppStartMenuFolder\${PRODUCTNAME}.lnk" "$INSTDIR\${MAINBINARYNAME}.exe"
  ApplicationID::Set "$SMPROGRAMS\$AppStartMenuFolder\${PRODUCTNAME}.lnk" "${IDENTIFIER}"
FunctionEnd
