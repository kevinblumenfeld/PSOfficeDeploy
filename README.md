
1.	Download zip file by clicking green download button here https://github.com/kevinblumenfeld/PSOfficeDeploy   
2.	On Source workstation (use Windows 10 - in PS type $psversiontable if version is lower than 5.1 install PS 5.1)  
    a.	Create directory c:\oScripts  
    b.	Open zip  
    c.	Copy contents of zip (as shown in image) to c:\oScripts  
    d.	Copy each MSI to directory c:\oScripts\Deploy\MSI  
3.	In c:\oScripts create a text file named computers.txt  (test by adding ONE computer name)  
4.	Dot Source the functions: Open PowerShell as an administrator on source workstation and run this command  
    ```Get-ChildItem c:\oScripts\  -Filter '*.ps1' | % {. $_.fullname }```

------------------------------------------
**You are ready to run the scripts**


5.	Copy files to destination computers - %C is case sensitive & be sure to include the trailing backslashes after last directory  
    ```Get-Content c:\oScripts\computers.txt | Send-FileAsJob -SourceDirsOrFiles "C:\oScripts\" -DestinationDir \\%C\C$\oScripts\```  
6.	Uninstall old Microsoft Office  
    ```Get-Content c:\oScripts\computers.txt | Uninstall-Office```  
      * Check a few workstations for EventID 1034 _Windows Installer removed the product Windows Installer installed the product_
7.	Install new Microsoft Office  
    ```Get-Content c:\oScripts\computers.txt | Install-MSI```  
      * Check a few workstations for EventID 1033 _Windows Installer installed the product_
8.	Report Microsoft Office Versions Installed – output to screen & output to CSV (excel) examples below  
    ```Get-Content c:\oScripts\computers.txt | Get-OfficeVersion```  
    ```Get-Content c:\oScripts\computers.txt | Get-OfficeVersion  | Export-Csv ./OfficeVersInstalled.csv -NoTypeInformation```  
