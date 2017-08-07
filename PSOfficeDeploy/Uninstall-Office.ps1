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
        [string] $RemovalType
    )
    Begin {
        Start-Transcript
        
    }
    Process {
        $_ | % {
            Start-Job -ScriptBlock {
                $comp = $using:_
                $uMSI = $using:RemovalType
                cd C:\Scripts
                If ($uMSI -eq "MSI") {
                    CD C:\Scripts\Deploy\MSI
                    Invoke-Command -ComputerName $comp -ScriptBlock {
                        msiexec.exe /x "C:\Scripts\Deploy\MSI\OfficeProPlus.msi" /qn /quiet /L*V "C:\scripts\Uninstall_MSI.log"
                    }
                }
                If ($uMSI -eq "AnyC2R") {
                    CD C:\Scripts
                    Invoke-Command -ComputerName $comp -FilePath .\uoDotSourceC2R.ps1
                }
                If ($uMSI -ne "MSI" -and $uMSI -ne "AnyC2R") {
                    CD C:\Scripts
                    Invoke-Command -ComputerName $comp -FilePath .\uoDotSource.ps1
                }
            }
        }
    }
}