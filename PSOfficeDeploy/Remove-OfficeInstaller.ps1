function Remove-OfficeInstaller {   
 
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
                $P = Start-Process -FilePath msiexec.exe -ArgumentList "/X {46992264-DAF7-4DF9-B7BA-83B3014B9A1E} /passive /norestart /qn" -Wait -NoNewWindow -PassThru
                $P.ExitCode
                }
            }
        }
    }
    End {

    }
}
