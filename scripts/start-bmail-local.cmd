@echo off
setlocal
set "APP_ROOT=E:\code-project\bmail"
set "INSTALLED_EXE=%LOCALAPPDATA%\Programs\Bmail Outlook AI\Bmail Outlook AI.exe"
set "UNPACKED_EXE=%APP_ROOT%\release-exe\win-unpacked\Bmail Outlook AI.exe"
set "MAILTO_ARG=%~1"

if exist "%INSTALLED_EXE%" (
  if "%MAILTO_ARG%"=="" (
    start "" "%INSTALLED_EXE%"
  ) else (
    start "" "%INSTALLED_EXE%" "%MAILTO_ARG%"
  )
  exit /b 0
)

if exist "%UNPACKED_EXE%" (
  if "%MAILTO_ARG%"=="" (
    start "" "%UNPACKED_EXE%"
  ) else (
    start "" "%UNPACKED_EXE%" "%MAILTO_ARG%"
  )
  exit /b 0
)

set "NODE24=E:\tools\node-v24\node.exe"
if not exist "%NODE24%" set "NODE24=node"
start "" /B "%NODE24%" "%APP_ROOT%\dist-server\server.mjs"
powershell -NoProfile -Command "for ($i=0; $i -lt 30; $i++) { try { $r = Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:3000/api/config' -TimeoutSec 1; if ($r.StatusCode -eq 200) { exit 0 } } catch {}; Start-Sleep -Milliseconds 250 }"
if "%MAILTO_ARG%"=="" (
  start "" "http://127.0.0.1:3000/"
) else (
  powershell -NoProfile -Command "$u = 'http://127.0.0.1:3000/?mailto=' + [uri]::EscapeDataString('%MAILTO_ARG%'); Start-Process $u"
)
exit /b 0
