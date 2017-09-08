function Uninstall-Office {   
    <#
.SYNOPSIS
Uninstalls most all types of Office

.DESCRIPTION
Designed to parallelize the uninstall of Office

.EXAMPLE
Removes any non-Click2Run version of Office - including Visio, Project etc.
get-content ./hostnames.txt | Uninstall-Office

Removes a specific install of office when the MSI (used during the install) is available
get-content ./hostnames.txt | Uninstall-Office -RemovalType MSI

Removes any Click2Run version of Office (Use when the MSI is not available)
get-content ./hostnames.txt | Uninstall-Office -RemovalType AnyC2R

#>

    Param (
        [Parameter(Mandatory = $true, 
            ValueFromPipeline = $true)]
        [string[]] $Computers,

        [Parameter(Mandatory = $false)]
        [ValidateSet("MSI", "AnyC2R")]
        [string] $RemovalType,
 
        [Parameter(Mandatory = $false)]
        [switch] $AlsoRemoveOffice2003
)
    Begin {
        Start-Transcript
        
    }
    Process {
        Start-Job -ScriptBlock {
            $comp = $using:_
            $uMSI = $using:RemovalType
            cd C:\oScripts
            If ($uMSI -eq "MSI") {
                CD C:\oScripts\Deploy\MSI
                Invoke-Command -ComputerName $comp -ScriptBlock {
                    msiexec.exe /x C:\oScripts\Deploy\MSI\OfficeProPlus.msi /qb! REBOOT=ReallySuppress /L*v c:\oScripts\o365_Uninstall_without_PSEXEC.log
                }
            }
            If ($uMSI -eq "AnyC2R") {
                CD C:\oScripts\Deploy
                Invoke-Command -ComputerName $comp -FilePath .\uoDotSourceC2R.ps1
            }
            If ($uMSI -ne "MSI" -and $uMSI -ne "AnyC2R" -and !$AlsoRemoveOffice2003) {
                CD C:\oScripts\Deploy
                Invoke-Command -ComputerName $comp -FilePath .\uoDotSource.ps1
            }
            If ($uMSI -ne "MSI" -and $uMSI -ne "AnyC2R" -and $AlsoRemoveOffice2003) {
                CD C:\oScripts\Deploy
                Invoke-Command -ComputerName $comp -FilePath .\uoDotSourceWith2003.ps1
            }
        }
    }
}