. C:\oScripts\Deploy\Uninstall\Remove-PreviousOfficeInstalls.ps1
$LogFile = ($(get-date -Format yyyy-MM-dd_HH-mm-ss) + "-Remove_Previous_Office_Installs.csv")
Remove-PreviousOfficeInstalls -LogFilePath $LogFile
