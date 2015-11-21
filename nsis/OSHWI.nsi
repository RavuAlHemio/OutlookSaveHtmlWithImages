!include "MUI2.nsh"
!include "x64.nsh"

; installer name and EXE
Name "Outlook Save HTML With Images"
OutFile "OutlookSaveHtmlWithImagesInstaller.exe"

; request admin
RequestExecutionLevel admin

; differentiate between 32 and 64
Function .onInit
${If} "$InstDir" == ""  ; don't override /D=C:\bleh
  ${If} ${RunningX64}
    StrCpy $InstDir "$ProgramFiles64\OutlookSaveHtmlWithImages"
  ${Else}
    StrCpy $InstDir "$ProgramFiles32\OutlookSaveHtmlWithImages"
  ${EndIf}
${EndIf}
FunctionEnd

; UI settings
!define MUI_ABORTWARNING

; UI pages
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; language(s)
!insertmacro MUI_LANGUAGE "English"

; constants
!define OSHWIADDINREGPATH "SOFTWARE\Microsoft\Office\Outlook\Addins\RavuAlHemio.OutlookSaveHtmlWithImages"
!define OSHWIADDINREG32PATH "SOFTWARE\Wow6432Node\Microsoft\Office\Outlook\Addins\RavuAlHemio.OutlookSaveHtmlWithImages"
!define OSHWIUNINSTALLREGPATH "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\RavuAlHemio.OutlookSaveHtmlWithImages"
!define OSHWISOURCEPATH "..\OutlookSaveHtmlWithImages\bin\Release\"

; prerequisite: .NET v4
Section "-.NET Framework v4" SecDotNet4
  IfFileExists "$WINDIR\Microsoft.NET\Framework\v4.0.30319" NETFrameworkInstalled 0

  File "/oname=$TEMP\dotNetFx40_Full_x86_x64.exe" "Redist\dotNetFx40_Full_x86_x64.exe"
  DetailPrint "Installing .NET Framework v4.0..."
  ExecWait '"$TEMP\dotNetFx40_Full_x86_x64.exe" /q /norestart'
  Return

NETFrameworkInstalled:
  DetailPrint ".NET Framework v4.0 already installed."
SectionEnd

; prerequisite: VSTOR
Section "-Visual Studio 2010 Tools for Office" SecVSTOR
  ; enable registry redirection
  SetRegView 32

  ClearErrors
  EnumRegKey $0 HKLM "SOFTWARE\Microsoft\VSTO Runtime Setup\v4R" 0
  IfErrors 0 VstorInstalled

  File "/oname=$TEMP\vstor_redist_10_0_60724.exe" "Redist\vstor_redist_10_0_60724.exe"
  DetailPrint "Installing Visual Studio 2010 Tools for Office Runtime..."
  ExecWait '"$TEMP\vstor_redist_10_0_60724.exe" /q /norestart'
  Return

VstorInstalled:
  DetailPrint "Visual Studio 2010 Tools for Office Runtime already installed."
SectionEnd

; the addin itself
Section "!Outlook Save HTML With Images" SecOSHWI
  SetOutPath "$INSTDIR"
  ; disable registry redirection
  SetRegView 64

  ; addin DLL + manifest
  File "${OSHWISOURCEPATH}\OutlookSaveHtmlWithImages.dll"
  File "${OSHWISOURCEPATH}\OutlookSaveHtmlWithImages.dll.manifest"

  ; addin VSTO info
  File "${OSHWISOURCEPATH}\OutlookSaveHtmlWithImages.vsto"

  ; dependent VSTO DLLs
  File "${OSHWISOURCEPATH}\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
  File "${OSHWISOURCEPATH}\Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll"

  ; other dependent DLLs
  File "${OSHWISOURCEPATH}\HtmlAgilityPack.dll"

  ; create uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  ; register with Outlook
  WriteRegStr HKLM "${OSHWIADDINREGPATH}" "FriendlyName" "Save HTML With Images"
  WriteRegStr HKLM "${OSHWIADDINREGPATH}" "Description" "Add-in that allows saving HTML e-mails, integrating external images into the HTML file."
  WriteRegDWORD HKLM "${OSHWIADDINREGPATH}" "LoadBehavior" 3
  WriteRegStr HKLM "${OSHWIADDINREGPATH}" "Manifest" "$INSTDIR\OutlookSaveHtmlWithImages.vsto|vstolocal"
  
  ${If} ${RunningX64}
    ; register with 32-bit Outlook too
    WriteRegStr HKLM "${OSHWIADDINREG32PATH}" "FriendlyName" "Save HTML With Images"
    WriteRegStr HKLM "${OSHWIADDINREG32PATH}" "Description" "Add-in that allows saving HTML e-mails, integrating external images into the HTML file."
    WriteRegDWORD HKLM "${OSHWIADDINREG32PATH}" "LoadBehavior" 3
    WriteRegStr HKLM "${OSHWIADDINREG32PATH}" "Manifest" "$INSTDIR\OutlookSaveHtmlWithImages.vsto|vstolocal"
  ${EndIf}
  
  ; register for uninstall
  WriteRegStr HKLM "${OSHWIUNINSTALLREGPATH}" "DisplayName" "Outlook Save HTML With Images"
  WriteRegStr HKLM "${OSHWIUNINSTALLREGPATH}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKLM "${OSHWIUNINSTALLREGPATH}" "UninstallString" "$\"$INSTDIR\Uninstall.exe$\""
  WriteRegStr HKLM "${OSHWIUNINSTALLREGPATH}" "QuietUninstallString" "$\"$INSTDIR\Uninstall.exe$\" /S"
  WriteRegDWORD HKLM "${OSHWIUNINSTALLREGPATH}" "NoModify" 1
  WriteRegDWORD HKLM "${OSHWIUNINSTALLREGPATH}" "NoRepair" 1

  ; make Office 2007 load machine-local add-ins
  ClearErrors
  EnumRegKey $0 HKLM "SOFTWARE\Microsoft\Office\12.0\Common\General" 0
  IfErrors NoOffice2012
  WriteRegDWORD HKLM "SOFTWARE\Microsoft\Office\12.0\Common\General" "EnableLocalMachineVSTO" 1
NoOffice2012:
  ClearErrors
  EnumRegKey $0 HKLM "SOFTWARE\Wow6432Node\Microsoft\Office\12.0\Common\General" 0
  IfErrors NoOffice2012x6432
  WriteRegDWORD HKLM "SOFTWARE\Wow6432Node\Microsoft\Office\12.0\Common\General" "EnableLocalMachineVSTO" 1
NoOffice2012x6432:
SectionEnd

; descriptions
LangString DESC_SecDotNet4 ${LANG_ENGLISH} ".NET Framework v4"
LangString DESC_SecVSTOR ${LANG_ENGLISH} "Visual Studio 2010 Tools for Office"
LangString DESC_SecOSHWI ${LANG_ENGLISH} "Outlook Save HTML With Images"

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecDotNet4} $(DESC_SecDotNet4)
  !insertmacro MUI_DESCRIPTION_TEXT ${SecVSTOR} $(DESC_SecVSTOR)
  !insertmacro MUI_DESCRIPTION_TEXT ${SecOSHWI} $(DESC_SecOSHWI)
!insertmacro MUI_FUNCTION_DESCRIPTION_END

; uninstaller
Section "Uninstall"
  ; disable registry redirection
  SetRegView 64

  ; unregister from Outlook
  DeleteRegKey HKLM "${OSHWIADDINREGPATH}"
  DeleteRegKey HKLM "${OSHWIADDINREG32PATH}"

  ; addin DLL + manifest
  Delete "$INSTDIR\OutlookSaveHtmlWithImages.dll"
  Delete "$INSTDIR\OutlookSaveHtmlWithImages.dll.manifest"

  ; addin VSTO info
  Delete "$INSTDIR\OutlookSaveHtmlWithImages.vsto"

  ; dependent VSTO DLLs
  Delete "$INSTDIR\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
  Delete "$INSTDIR\Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll"

  ; other dependent DLLs
  Delete "$INSTDIR\HtmlAgilityPack.dll"

  ; uninstaller
  Delete "$INSTDIR\Uninstall.exe"
  
  ; unregister uninstaller
  DeleteRegKey HKLM "${OSHWIUNINSTALLREGPATH}"

  RMDir "$INSTDIR"
SectionEnd
