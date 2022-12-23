<# : Office.cmd
@echo off
(cd /d "%~dp0")&&(NET FILE||(powershell start-process -FilePath '%0' -verb runas)&&(exit /B)) >NUL 2>&1
powershell  -NoProfile -ExecutionPolicy Bypass "iex (${%~f0} | out-string)"
goto :EOF
: end Batch portion / begin PowerShell #>

$KMS = "kms.srv.crsoo.com"
$Configuration = @"
<Configuration>
<Add Channel="PerpetualVL2021">
<Product ID="ProPlus2021Volume"  PIDKEY="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH">
<Language ID="MatchOS" />
      <ExcludeApp ID="Access" />
     <!-- <ExcludeApp ID="Excel" /> -->
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <!-- <App ID="Outlook" /> -->
      <ExcludeApp ID="PowerPoint" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Skype" />
      <ExcludeApp ID="Teams" />
      <!-- <ExcludeApp ID="Word" /> -->
      <ExcludeApp ID="Project" />
      <ExcludeApp ID="Visio" />
</Product>
<Product ID="ProofingTools">
<Language ID="MatchOS" />
</Product>
</Add>

<Remove All="TRUE" />
<Property Name="AUTOACTIVATE" Value="1" />
<RemoveMSI All="TRUE" />
<Updates Enabled="TRUE" />
<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
<Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@

if (!(Test-Path "Configuration.xml")) {
$Configuration | Set-Content "Configuration.xml"
}
Start-Process -FilePath "notepad.exe" -Wait  -ArgumentList "Configuration.xml"
if (!(Test-Path "setup.exe")) {
(New-Object Net.WebClient).DownloadFile("https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_15629-20208.exe", "ODT.exe")
Start-Process -FilePath ".\ODT.exe" -Wait -NoNewWindow -ArgumentList "/extract:.", "/quiet"
}

Start-Process -FilePath """$([System.Environment]::SystemDirectory)\cscript.exe""" -NoNewWindow -Wait -ArgumentList "//nologo", """$([System.Environment]::SystemDirectory)\slmgr.vbs""", "/skms $KMS"
Start-Process -FilePath ".\setup.exe" -NoNewWindow -ArgumentList "/configure", "Configuration.xml" -Wait
$OSCaption = (Get-CimInstance -ClassName CIM_OperatingSystem).Caption
IF ($OSCaption -like '*Windows 7*') {
$InstallPath=$(Get-ItemPropertyValue -Path Registry::'HKLM\SOFTWARE\Microsoft\Office\ClickToRun' -Name 'InstallPath')
Set-Location -Path $InstallPath\Office*
Start-Process -FilePath """$([System.Environment]::SystemDirectory)\cscript.exe""" -NoNewWindow -Wait -ArgumentList "ospp.vbs", "/sethst:$KMS"
Start-Process -FilePath """$([System.Environment]::SystemDirectory)\cscript.exe""" -NoNewWindow -Wait -ArgumentList "ospp.vbs", "/act"
}
