!macro customInit
  nsExec::ExecToLog 'taskkill /F /IM "Bmail Outlook AI.exe"'
  nsExec::ExecToLog 'taskkill /F /IM "Bmail.exe"'
!macroend

!macro customInstall
  WriteRegStr HKCU "Software\Classes\Bmail.URL.mailto" "" "URL:Bmail MailTo Protocol"
  WriteRegStr HKCU "Software\Classes\Bmail.URL.mailto" "URL Protocol" ""
  WriteRegStr HKCU "Software\Classes\Bmail.URL.mailto\DefaultIcon" "" "$INSTDIR\Bmail Outlook AI.exe,0"
  WriteRegStr HKCU "Software\Classes\Bmail.URL.mailto\shell\open\command" "" '"$INSTDIR\Bmail Outlook AI.exe" "%1"'

  WriteRegStr HKCU "Software\Classes\sgw-outlook" "" "URL:Bmail SGW Outlook Protocol"
  WriteRegStr HKCU "Software\Classes\sgw-outlook" "URL Protocol" ""
  WriteRegStr HKCU "Software\Classes\sgw-outlook\DefaultIcon" "" "$INSTDIR\Bmail Outlook AI.exe,0"
  WriteRegStr HKCU "Software\Classes\sgw-outlook\shell\open\command" "" '"$INSTDIR\Bmail Outlook AI.exe" "%1"'

  WriteRegStr HKCU "Software\Clients\Mail\Bmail" "" "Bmail Outlook AI"
  WriteRegStr HKCU "Software\Clients\Mail\Bmail\DefaultIcon" "" "$INSTDIR\Bmail Outlook AI.exe,0"
  WriteRegStr HKCU "Software\Clients\Mail\Bmail\shell\open\command" "" '"$INSTDIR\Bmail Outlook AI.exe"'
  WriteRegStr HKCU "Software\Clients\Mail\Bmail\Capabilities" "ApplicationName" "Bmail Outlook AI"
  WriteRegStr HKCU "Software\Clients\Mail\Bmail\Capabilities" "ApplicationDescription" "Bmail Outlook AI — Local PST & Mailplug Mail Client"
  WriteRegStr HKCU "Software\Clients\Mail\Bmail\Capabilities" "ApplicationIcon" "$INSTDIR\Bmail Outlook AI.exe,0"
  WriteRegStr HKCU "Software\Clients\Mail\Bmail\Capabilities\URLAssociations" "mailto" "Bmail.URL.mailto"
  WriteRegStr HKCU "Software\Clients\Mail\Bmail\Capabilities\URLAssociations" "bmail" "Bmail.URL.mailto"
  WriteRegStr HKCU "Software\Clients\Mail\Bmail\Capabilities\URLAssociations" "sgw-outlook" "Bmail.URL.mailto"

  WriteRegStr HKCU "Software\RegisteredApplications" "Bmail" "Software\Clients\Mail\Bmail\Capabilities"
!macroend

!macro customUnInstall
  DeleteRegKey HKCU "Software\Classes\Bmail.URL.mailto"
  DeleteRegKey HKCU "Software\Classes\sgw-outlook"
  DeleteRegKey HKCU "Software\Clients\Mail\Bmail"
  DeleteRegValue HKCU "Software\RegisteredApplications" "Bmail"
!macroend
