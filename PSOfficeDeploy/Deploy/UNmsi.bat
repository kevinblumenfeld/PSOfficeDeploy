@echo off
msiexec /x c:\oScripts\deploy\msi\officeproplus.msi /qb! REBOOT=ReallySuppress /L*v c:\oScripts\o365_Uninstall.log
exit %ERRORLEVEL%  
