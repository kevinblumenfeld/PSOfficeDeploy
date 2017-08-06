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
                $P = Start-Process -FilePath msiexec.exe -ArgumentList "/i `"C:\Scripts\Deploy\MSI\OfficeProPlus.msi`" /qn" -Wait -NoNewWindow -PassThru
                $P.ExitCode
                }
            }
        }
    }
    End {

    }
}
