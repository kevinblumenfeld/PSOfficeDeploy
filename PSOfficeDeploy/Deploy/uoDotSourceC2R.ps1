. C:\oScripts\Deploy\uninstallC2R\Remove-OfficeClickToRun.ps1
$LogFile = ($(get-date -Format yyyy-MM-dd_HH-mm-ss) + "-Remove_CLICK2RUN_Office_Installs.csv")
Remove-OfficeClickToRun -LogFilePath "C:\oScript\" + $LogFile