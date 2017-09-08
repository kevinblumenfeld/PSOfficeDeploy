cscript "C:\oScripts\Deploy\Uninstall\OffScrub03.vbs" ALL /Quiet /NoCancel
Start-Sleep -s 300
. C:\oScripts\Deploy\Uninstall\Remove-PreviousOfficeInstalls.ps1
Remove-PreviousOfficeInstalls -LogFilePath "C:\oScripts\uninstall_office.log"


