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
                $MSIArguments = @(
                    "/i"
                    "C:\Scripts\Deploy\MSI\OfficeProPlus.msi"
                    "/qn"
                    "/norestart"
                )
                Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow 
                # Invoke-Command -ComputerName $comp –FilePath .\fooi.ps1
            }
        }
    }
    End {

    }
}
