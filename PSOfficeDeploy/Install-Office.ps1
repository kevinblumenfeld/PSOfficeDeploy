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
        $_ | % {
            Start-Job -ScriptBlock {
                $comp = $using:_
                CD C:\Scripts\Deploy\MSI
                Invoke-Command -ComputerName $comp -ScriptBlock {
                    msiexec.exe /i "C:\Scripts\Deploy\MSI\OfficeProPlus.msi" /qn /quiet /L*V "C:\scripts\Install_MSI.log"
                }
            }
        }
    }
    End {

    }
}
