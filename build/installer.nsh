; Kertas Office: extra installer steps (electron-builder includes this file).
;
; Windows 10 and 11 do not let an installer quietly make itself the default app for a
; file type. What an installer can do is register the app properly, so it shows up in
; Settings > Default apps with all its file types and in every "Open with" list. Kertas
; then offers to open that Settings page for the user (see electron/main.cjs).

!macro customInstall
  ; Capabilities: the file types Kertas can open, pointing at the ProgIDs electron-builder made
  WriteRegStr SHELL_CONTEXT "Software\Kertas Office\Capabilities" "ApplicationName" "Kertas Office"
  WriteRegStr SHELL_CONTEXT "Software\Kertas Office\Capabilities" "ApplicationDescription" "Free documents, spreadsheets and PDFs that stay on your computer"
  WriteRegStr SHELL_CONTEXT "Software\Kertas Office\Capabilities" "ApplicationIcon" "$appExe,0"
  !insertmacro kertasCapability "docx"
  !insertmacro kertasCapability "doc"
  !insertmacro kertasCapability "rtf"
  !insertmacro kertasCapability "odt"
  !insertmacro kertasCapability "xlsx"
  !insertmacro kertasCapability "xlsm"
  !insertmacro kertasCapability "xls"
  !insertmacro kertasCapability "ods"
  !insertmacro kertasCapability "csv"
  !insertmacro kertasCapability "pdf"
  WriteRegStr SHELL_CONTEXT "Software\RegisteredApplications" "Kertas Office" "Software\Kertas Office\Capabilities"
  ; Kertas is also listed under "Open with" for these types even where another app is the default
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\Kertas Office.exe" "FriendlyAppName" "Kertas Office"
  ; tell Explorer that file associations changed, so icons refresh straight away
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro kertasCapability EXT
  WriteRegStr SHELL_CONTEXT "Software\Kertas Office\Capabilities\FileAssociations" ".${EXT}" "KertasOffice.${EXT}"
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\Kertas Office.exe\SupportedTypes" ".${EXT}" ""
!macroend

!macro customUnInstall
  DeleteRegValue SHELL_CONTEXT "Software\RegisteredApplications" "Kertas Office"
  DeleteRegKey SHELL_CONTEXT "Software\Kertas Office\Capabilities"
  DeleteRegKey /ifempty SHELL_CONTEXT "Software\Kertas Office"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\Applications\Kertas Office.exe"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
