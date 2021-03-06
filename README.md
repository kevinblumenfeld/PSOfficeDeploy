NOTE: You must enable PS Remoting on all computers involved. A guide can be found [here](https://www.techrepublic.com/article/how-to-enable-powershell-remoting-via-group-policy/)
1.	**Download zip file** by clicking green download button here https://github.com/kevinblumenfeld/PSOfficeDeploy   
2.	**On Source workstation** (use Windows 10 - in PS type $psversiontable if version is lower than 5.1, install PS 5.1 [here](https://www.microsoft.com/en-us/download/details.aspx?id=54616))  
    a.	Create directory c:\oScripts  
    b.	Open zip & copy contents (**only pictured files and Deploy folder**) to c:\oScripts  
    ![image](https://user-images.githubusercontent.com/28877715/29216887-0b82e146-7e7e-11e7-80ce-1eceb77bfbc3.png)   
    c.	Copy officeproplus.msi to directory c:\oScripts\Deploy\MSI  
3.	**Add computer names to computers.txt in c:\oScripts** (test by adding ONE computer name)  
4.	**Dot Source the functions**: Open PowerShell as an administrator on source workstation and run this command  
    ```Get-ChildItem -Path C:\oScripts -File -Recurse | Unblock-File```  
    ```Get-ChildItem c:\oScripts\  -Filter '*.ps1' | % {. $_.fullname }```

------------------------------------------
**You are ready to run the scripts**


5.	**Copy files to destination computers** - %C is case sensitive & be sure to include the trailing backslashes after last directory  
    ```Get-Content c:\oScripts\computers.txt | Send-FileAsJob -SourceDirsOrFiles C:\oScripts\ -DestinationDir \\%C\C$\oScripts\```  
6.	**Uninstall old Microsoft Office**  
    ```Get-Content c:\oScripts\computers.txt | Uninstall-Office```  
      _Use the below command if Office 2003 also needs to be removed_  
    ```Get-Content c:\oScripts\computers.txt | Uninstall-Office -AlsoRemoveOffice2003```  
      * Use Step 8 to check progress  
7.	**Install new Microsoft Office**  
    ```Get-Content c:\oScripts\computers.txt | Install-MSI```  
      * Use Step 8 to check progress  
8.	**Report Microsoft Office Versions Installed** – output to screen & output to CSV (excel) commands below  
    ```Get-Content c:\oScripts\computers.txt | Get-OfficeVersion```  
    ```Get-Content c:\oScripts\computers.txt | Get-OfficeVersion  | Export-Csv ./OfficeVersInstalled.csv -NoTypeInformation```
    
    NOTE: Click, Install Toolkit to create a MSI/EXE that does not require a xml file. Link:  
             http://officedev.github.io/Office-IT-Pro-Deployment-Scripts/XmlEditor.html  
             Due to a bug, remove English as a language and select, MATCHOS  
             ![image](https://user-images.githubusercontent.com/28877715/32521730-afa4d79e-c3e2-11e7-8ffe-2281b578bede.png)

