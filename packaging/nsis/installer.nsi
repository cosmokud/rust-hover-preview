; Template for the NSIS installer cargo-packager builds, pointed at by
; `[package.metadata.packager.nsis] template` in Cargo.toml. The placeholders
; in double braces are filled in by cargo-packager.
;
; This is cargo-packager 0.11.8's own template with four differences:
;   - there is no "Already Installed" page: a previous installation is removed
;     automatically, before anything is written, instead of being offered as a
;     choice;
;   - a running copy of the app is terminated rather than prompted for;
;   - a page of its own is shown after the shortcuts page, asking whether the
;     app's settings are to be put back to the defaults this version recommends
;     (see "Recommended settings page" below). It writes no configuration of its
;     own: it leaves two files for the app to find, and the app is what puts the
;     settings back.
;   - a page of its own is shown after that one, listing the optional engines this
;     app can drive but does not ship, marked with whether this machine already
;     has each of them and carrying a link to the page it is installed from where
;     it does not (see "Optional engines page" below). It installs nothing and
;     writes nothing: it is the README's install steps, offered where a user is
;     already looking.
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
!if "${INSTALLERICON}" != ""
  !define MUI_ICON "${INSTALLERICON}"
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

; 6. Recommended settings page
;
; Two boxes, both clear, and nothing else asked: what is offered is a reset of the app's
; own settings, and a reinstall is not a reason to want one. This installer writes no
; configuration — a second writer of `config.ini` would carry the defaults of whenever the
; installer was built, and an installation made that way would keep them for good — so what
; a box does is leave a file beside that configuration, and the app reads it as it starts.
;
; The page is not shown to an unattended installer at all: an update put on by the app's
; own `Auto` answer runs this installer silently, and must never change a setting.
;
; nsDialogs is not included here because MUI2 already brings it in — the uninstaller's own
; checkbox below is built out of its constants.
Var RecommendedSettingsCheckbox
Var RecommendedListsCheckbox
Var RecommendedSettingsState
Var RecommendedListsState
Page custom RecommendedShow RecommendedLeave

; 7. Optional engines page
;
; The engines this app can drive but does not ship, and whether this machine already has each
; of them: one row per engine, saying what the engine is called and a few of the formats it
; would give previews to, marked as detected or left greyed with a link to the page it is
; installed from.
;
; Nothing is installed from here and nothing is downloaded. What a link does is open a page in
; the browser the user already has — the same pages the README names, and the same ones the
; tray's `Codecs` rows open. A row an engine is present for offers nothing, since there is
; nothing to offer.
;
; What is asked about each engine is what the app itself asks before it drives one, so a page
; that says a name is missing and the app's own `Codecs` row agree: the folders each installer
; writes to, the `PATH` for FFmpeg's player, the console archiver inside a PeaZip installation,
; and the ProgID of each Office application for the one thing decided here about LibreOffice —
; the Office formats are worth naming only on a machine with no Office to draw them, and what
; LibreOffice adds to a machine that has one is the formats beside them.
;
; Like the page above it, this page is not shown to an unattended installer at all: an update
; put on by the app's own `Auto` answer runs this installer silently, and a page nobody is
; there to read is not shown to anyone.
Var EnginesLibreText
Page custom EnginesShow EnginesLeave

; 8. Installation page
!insertmacro MUI_PAGE_INSTFILES

; 9. Finish page
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
LangString recommendedTitle ${LANG_ENGLISH} "Recommended settings"
LangString recommendedSubtitle ${LANG_ENGLISH} "Put the app's settings back to this version's defaults"
LangString recommendedIntro ${LANG_ENGLISH} "Put the app's settings back to the defaults this version recommends. A box left clear changes nothing."
LangString recommendedSettings ${LANG_ENGLISH} "Reset to Recommended Settings"
LangString recommendedLists ${LANG_ENGLISH} "Reset Extension Lists"
LangString recommendedNote ${LANG_ENGLISH} "The first puts the settings back and leaves your extension lists alone. The second puts the lists back and leaves every other setting alone."

; The words the optional-engines page carries: what each row is called, with a few of the
; formats that engine would give previews to, and the two words a row can carry to its right.
; `enginesLibreOffice` is the name a machine with no Office sees and `enginesLibreNiche` the one
; a machine with it does: the Office formats are covered there already, and what LibreOffice
; adds is the formats beside them.
;
; The checkmark is a square root sign, which is a fact about the font rather than a choice: the
; dialog font these pages are drawn with carries `√` and has no glyph at all for `✔` or `✓`,
; which would be drawn as a box.
LangString enginesTitle ${LANG_ENGLISH} "Optional engines"
LangString enginesSubtitle ${LANG_ENGLISH} "Previews that need something this app does not ship"
LangString enginesIntro ${LANG_ENGLISH} "These engines are optional: without one, the files it draws simply show no preview. Nothing is installed here — a link opens the page it is installed from, and an engine installed later is used without a restart."
LangString enginesDetected ${LANG_ENGLISH} "Detected"
LangString enginesDownload ${LANG_ENGLISH} "Download"
LangString enginesFFmpeg ${LANG_ENGLISH} "FFmpeg (flv, rmvb, mxf, ogv, swf, ...)"
LangString enginesLibreOffice ${LANG_ENGLISH} "LibreOffice (doc, docx, xls, xlsx, ppt, pptx, ...)"
LangString enginesLibreNiche ${LANG_ENGLISH} "LibreOffice (cdr, odt, ods, odp, odg, ...)"
LangString enginesImageMagick ${LANG_ENGLISH} "ImageMagick (nef, cr2, cr3, arw, dng, raf, ...)"
LangString enginesPeaZip ${LANG_ENGLISH} "PeaZip (cab, iso, rpm, deb, arj, lzh, ...)"
LangString enginesCalibre ${LANG_ENGLISH} "Calibre (mobi, azw3, epub, djvu, fb2, ...)"

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
  WriteRegStr SHCTX "${UNINSTKEY}" "DisplayIcon" "$\"$INSTDIR\${MAINBINARYNAME}.exe$\""
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

  ; What the recommended-settings page's boxes asked for, left where the app will find it:
  ; the folder the app keeps its own configuration in. It is the path `appdata-paths` names
  ; in the packaging metadata in Cargo.toml, written out here because it is not a file the
  ; app ships — and it is the only place this installer writes anything about the app's
  ; settings, which is why the boxes are files rather than a configuration of its own.
  ;
  ; A box left clear writes nothing, so a reinstall leaves the app's settings alone.
  SetShellVarContext current
  CreateDirectory "$APPDATA\rust-hover-preview"

  ${If} $RecommendedSettingsState == 1
    FileOpen $0 "$APPDATA\rust-hover-preview\reset-settings.marker" w
    FileClose $0
  ${EndIf}

  ${If} $RecommendedListsState == 1
    FileOpen $0 "$APPDATA\rust-hover-preview\reset-extensions.marker" w
    FileClose $0
  ${EndIf}

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

; The recommended-settings page: two boxes and nothing else, both clear. What they were set
; to is read as the page is left, and the files they stand for are written once the
; installation has gone through (see the end of the Install section).
;
; A custom page is skipped by aborting where it is built, which is where an unattended
; installer leaves it: no page is ever shown to one, so an update put on silently cannot
; change a setting.
Function RecommendedShow
  ${If} ${Silent}
    Abort
  ${EndIf}
  ${IfThen} $PassiveMode == 1 ${|} Abort ${|}

  !insertmacro MUI_HEADER_TEXT "$(recommendedTitle)" "$(recommendedSubtitle)"

  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}

  ${NSD_CreateLabel} 0 0 100% 20u "$(recommendedIntro)"
  Pop $0

  ${NSD_CreateCheckbox} 0 26u 100% 12u "$(recommendedSettings)"
  Pop $RecommendedSettingsCheckbox
  ${NSD_CreateCheckbox} 0 40u 100% 12u "$(recommendedLists)"
  Pop $RecommendedListsCheckbox

  ${NSD_CreateLabel} 0 60u 100% 24u "$(recommendedNote)"
  Pop $0

  nsDialogs::Show
FunctionEnd

Function RecommendedLeave
  SendMessage $RecommendedSettingsCheckbox ${BM_GETCHECK} 0 0 $RecommendedSettingsState
  SendMessage $RecommendedListsCheckbox ${BM_GETCHECK} 0 0 $RecommendedListsState
FunctionEnd

; The optional-engines page: one row per engine this app can drive, marked with whether this
; machine already has it and carrying a link to the page it is installed from where it does not.
;
; A row that is present is the same label with `√` in front of it, and a row that is missing is
; that label greyed — which is what disabling a static does, since the dialog manager draws a
; disabled control's text in the grey the platform keeps for it.
;
; The five engines are asked about one at a time, and each detector answers in `$0`: 1 where the
; engine is installed and 0 where it is not.
Function FFmpegInstalled
  StrCpy $0 0
  SearchPath $1 "ffplay.exe"
  IfErrors +2 0
    StrCpy $0 1
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

; Whether every Office application this app draws a page with is installed. It is asked for the
; LibreOffice row and nothing else, and what that row is called follows it: with all three here
; the Office formats are covered and what LibreOffice adds is the formats beside them, while a
; machine with none — or with only some — is shown the Office formats LibreOffice draws.
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

  ${NSD_CreateLabel} 0 0 100% 26u "$(enginesIntro)"
  Pop $0

  Call FFmpegInstalled
  ${If} $0 == 1
    StrCpy $1 "√ $(enginesFFmpeg)"
    ${NSD_CreateLabel} 0 32u 74% 12u "$1"
    Pop $0
    ${NSD_CreateLabel} 75% 32u 25% 12u "$(enginesDetected)"
    Pop $0
  ${Else}
    ${NSD_CreateLabel} 0 32u 74% 12u "$(enginesFFmpeg)"
    Pop $0
    EnableWindow $0 0
    ${NSD_CreateLink} 75% 32u 25% 12u "$(enginesDownload)"
    Pop $0
    ${NSD_OnClick} $0 EnginesLinkFFmpeg
  ${EndIf}

  ; What LibreOffice is called here is decided before it is asked about: the Office formats are
  ; worth naming where there is no Office to draw them, and the formats beside them are what the
  ; engine is worth where there is one.
  StrCpy $EnginesLibreText "$(enginesLibreOffice)"
  Call OfficeInstalled
  ${If} $0 == 1
    StrCpy $EnginesLibreText "$(enginesLibreNiche)"
  ${EndIf}

  Call LibreOfficeInstalled
  ${If} $0 == 1
    StrCpy $1 "√ $EnginesLibreText"
    ${NSD_CreateLabel} 0 46u 74% 12u "$1"
    Pop $0
    ${NSD_CreateLabel} 75% 46u 25% 12u "$(enginesDetected)"
    Pop $0
  ${Else}
    ${NSD_CreateLabel} 0 46u 74% 12u "$EnginesLibreText"
    Pop $0
    EnableWindow $0 0
    ${NSD_CreateLink} 75% 46u 25% 12u "$(enginesDownload)"
    Pop $0
    ${NSD_OnClick} $0 EnginesLinkLibreOffice
  ${EndIf}

  Call ImageMagickInstalled
  ${If} $0 == 1
    StrCpy $1 "√ $(enginesImageMagick)"
    ${NSD_CreateLabel} 0 60u 74% 12u "$1"
    Pop $0
    ${NSD_CreateLabel} 75% 60u 25% 12u "$(enginesDetected)"
    Pop $0
  ${Else}
    ${NSD_CreateLabel} 0 60u 74% 12u "$(enginesImageMagick)"
    Pop $0
    EnableWindow $0 0
    ${NSD_CreateLink} 75% 60u 25% 12u "$(enginesDownload)"
    Pop $0
    ${NSD_OnClick} $0 EnginesLinkImageMagick
  ${EndIf}

  Call PeaZipInstalled
  ${If} $0 == 1
    StrCpy $1 "√ $(enginesPeaZip)"
    ${NSD_CreateLabel} 0 74u 74% 12u "$1"
    Pop $0
    ${NSD_CreateLabel} 75% 74u 25% 12u "$(enginesDetected)"
    Pop $0
  ${Else}
    ${NSD_CreateLabel} 0 74u 74% 12u "$(enginesPeaZip)"
    Pop $0
    EnableWindow $0 0
    ${NSD_CreateLink} 75% 74u 25% 12u "$(enginesDownload)"
    Pop $0
    ${NSD_OnClick} $0 EnginesLinkPeaZip
  ${EndIf}

  Call CalibreInstalled
  ${If} $0 == 1
    StrCpy $1 "√ $(enginesCalibre)"
    ${NSD_CreateLabel} 0 88u 74% 12u "$1"
    Pop $0
    ${NSD_CreateLabel} 75% 88u 25% 12u "$(enginesDetected)"
    Pop $0
  ${Else}
    ${NSD_CreateLabel} 0 88u 74% 12u "$(enginesCalibre)"
    Pop $0
    EnableWindow $0 0
    ${NSD_CreateLink} 75% 88u 25% 12u "$(enginesDownload)"
    Pop $0
    ${NSD_OnClick} $0 EnginesLinkCalibre
  ${EndIf}

  nsDialogs::Show
FunctionEnd

; Nothing is read from this page as it is left: it says what the machine has and the links are
; all it does, so there is no state to collect. The page above has two boxes and this one has
; none — which is why this function is empty rather than absent, since a page is declared with
; both of its functions.
Function EnginesLeave
FunctionEnd

; The page's five links, a function each, so that the address a row opens is written by the row
; rather than looked up from a table. The control handle the click carries is the first thing on
; the stack and is dropped: what a click does is one address, and which control it was is not
; part of the question.
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

Function CreateDesktopShortcut
  CreateShortcut "$DESKTOP\${PRODUCTNAME}.lnk" "$INSTDIR\${MAINBINARYNAME}.exe"
  ApplicationID::Set "$DESKTOP\${PRODUCTNAME}.lnk" "${IDENTIFIER}"
FunctionEnd

Function CreateStartMenuShortcut
  CreateDirectory "$SMPROGRAMS\$AppStartMenuFolder"
  CreateShortcut "$SMPROGRAMS\$AppStartMenuFolder\${PRODUCTNAME}.lnk" "$INSTDIR\${MAINBINARYNAME}.exe"
  ApplicationID::Set "$SMPROGRAMS\$AppStartMenuFolder\${PRODUCTNAME}.lnk" "${IDENTIFIER}"
FunctionEnd
