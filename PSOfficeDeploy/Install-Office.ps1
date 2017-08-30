function Install-Office {   
 
    Param (
        [Parameter(Mandatory = $true, 
            ValueFromPipeline = $true)]
        [string[]] $Computers,

        [Parameter(Mandatory = $false)]
        [string] $InstallFileName

    )
    Begin {

    }
    Process {
        Start-Job -ScriptBlock {
            $comp = $using:_
            CD C:\oScripts\Deploy\MSI
            Invoke-Command -ComputerName $comp -ScriptBlock {
                msiexec.exe /i C:\oScripts\Deploy\MSI\OfficeProPlus.msi /qb! REBOOT=ReallySuppress /L*v c:\oScripts\o365_Install_without_PSEXEC.log
            }
        }
    }
    End {

    }
}
