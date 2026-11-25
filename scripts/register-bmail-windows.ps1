$ErrorActionPreference = 'Stop'

$appRoot = 'E:\code-project\bmail'
$installedExe = Join-Path $env:LOCALAPPDATA 'Programs\Bmail Outlook AI\Bmail Outlook AI.exe'
$legacyInstalledExe = Join-Path $env:LOCALAPPDATA 'Programs\bmail-outlook-ai\Bmail Outlook AI.exe'
$unpackedExe = Join-Path $appRoot 'release-exe\win-unpacked\Bmail Outlook AI.exe'
$launcher = Join-Path $appRoot 'scripts\start-bmail-local.cmd'

if (Test-Path $installedExe) {
  $targetExec = $installedExe
  $iconExec = "$installedExe,0"
} elseif (Test-Path $legacyInstalledExe) {
  $targetExec = $legacyInstalledExe
  $iconExec = "$legacyInstalledExe,0"
} elseif (Test-Path $unpackedExe) {
  $targetExec = $unpackedExe
  $iconExec = "$unpackedExe,0"
} elseif (Test-Path $launcher) {
  $targetExec = $launcher
  $iconExec = 'shell32.dll,156'
} else {
  throw "Launcher not found: $launcher"
}

$cmdLine = "`"$targetExec`" `"%1`""
$openCmdLine = "`"$targetExec`""

# 1. Register Bmail.URL.mailto ProgId (HKCU only, reversible, preserves Outlook profile)
$progIdBase = 'HKCU:\Software\Classes\Bmail.URL.mailto'
New-Item -Path "$progIdBase\DefaultIcon" -Force | Out-Null
New-Item -Path "$progIdBase\shell\open\command" -Force | Out-Null
Set-ItemProperty -Path $progIdBase -Name '(default)' -Value 'URL:Bmail MailTo Protocol'
Set-ItemProperty -Path $progIdBase -Name 'URL Protocol' -Value ''
Set-ItemProperty -Path "$progIdBase\DefaultIcon" -Name '(default)' -Value $iconExec
Set-ItemProperty -Path "$progIdBase\shell\open\command" -Name '(default)' -Value $cmdLine

# 2. Register official Windows Mail Client entry under HKCU\Software\Clients\Mail\Bmail
$mailClientBase = 'HKCU:\Software\Clients\Mail\Bmail'
New-Item -Path "$mailClientBase\DefaultIcon" -Force | Out-Null
New-Item -Path "$mailClientBase\Capabilities\URLAssociations" -Force | Out-Null
New-Item -Path "$mailClientBase\shell\open\command" -Force | Out-Null
Set-ItemProperty -Path $mailClientBase -Name '(default)' -Value 'Bmail Outlook AI'
Set-ItemProperty -Path "$mailClientBase\DefaultIcon" -Name '(default)' -Value $iconExec
Set-ItemProperty -Path "$mailClientBase\shell\open\command" -Name '(default)' -Value $openCmdLine
Set-ItemProperty -Path "$mailClientBase\Capabilities" -Name 'ApplicationName' -Value 'Bmail Outlook AI'
Set-ItemProperty -Path "$mailClientBase\Capabilities" -Name 'ApplicationDescription' -Value 'Bmail Outlook AI — Local PST & Mailplug Mail Client'
Set-ItemProperty -Path "$mailClientBase\Capabilities" -Name 'ApplicationIcon' -Value $iconExec
Set-ItemProperty -Path "$mailClientBase\Capabilities\URLAssociations" -Name 'mailto' -Value 'Bmail.URL.mailto'
Set-ItemProperty -Path "$mailClientBase\Capabilities\URLAssociations" -Name 'bmail' -Value 'Bmail.URL.mailto'

# 3. Expose in Windows Default Apps (HKCU\Software\RegisteredApplications)
New-Item -Path 'HKCU:\Software\RegisteredApplications' -Force | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\RegisteredApplications' -Name 'Bmail' -Value 'Software\Clients\Mail\Bmail\Capabilities'

# 4. Create Desktop and Start Menu shortcuts
$ws = New-Object -ComObject WScript.Shell
$desktopLink = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Bmail Outlook AI.lnk'
$startMenuDir = Join-Path ([Environment]::GetFolderPath('Programs')) 'Bmail Outlook AI'
New-Item -ItemType Directory -Path $startMenuDir -Force | Out-Null
$startMenuLink = Join-Path $startMenuDir 'Bmail Outlook AI.lnk'

foreach ($lnkPath in @($desktopLink, $startMenuLink)) {
  $sc = $ws.CreateShortcut($lnkPath)
  $sc.TargetPath = $targetExec
  $sc.WorkingDirectory = (Split-Path -Parent $targetExec)
  $sc.Description = 'Bmail Outlook AI (Local Outlook Replacement)'
  $sc.IconLocation = $iconExec
  $sc.Save()
}

Write-Output "REGISTERED_BMAIL_WINDOWS_OK"
Write-Output "TargetExec=$targetExec"
Write-Output "DesktopShortcut=$desktopLink"
Write-Output "StartMenuShortcut=$startMenuLink"
