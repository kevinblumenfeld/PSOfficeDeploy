
Drop all the scripts and the Deploy folder into C:\oScripts folder on the SOURCE workstation.
The SOURCE workstation is from where you will "target" other workstations (to install, uninstall and report on Office versions).
Use Send-FileAsJob to copy all including the MSI (in the MSI folder) to c:\oScripts to each target workstation.

Each script needs to be dot sourced.

This will be incorporated into a module soon.
