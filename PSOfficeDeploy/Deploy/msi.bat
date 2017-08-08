@echo off
msiexec /i c:\oScripts\deploy\msi\officeproplus.msi /qb! REBOOT=ReallySuppress /L*v c:\oScripts\o365_Install.log
exit %ERRORLEVEL%  
