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
                Start-Process -FilePath msiexec.exe -ArgumentList '/qn /quiet /i C:\Scripts\Deploy\MSI\OfficeProPlus.msi'
                }
            }
        }
    }
    End {

    }
}
