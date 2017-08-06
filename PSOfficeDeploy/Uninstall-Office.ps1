function Uninstall-Office {   
    <#
.SYNOPSIS

.DESCRIPTION

.EXAMPLE
get-content ./hostnames.txt | Uninstall-Office

#>

    Param (
        [Parameter(Mandatory = $true, 
            ValueFromPipeline = $true)]
        [string[]] $Computers
    )
    Begin {
        Start-Transcript
        
    }
    Process {
        $_ | % {
            Start-Job -ScriptBlock {
                $comp = $using:_
                cd C:\Scripts
                Invoke-Command -ComputerName $comp -FilePath .\uoDotSourceC2R.ps1
                # Invoke-Command -ComputerName $comp -FilePath .\uoDotSource.ps1
            }
        }
    }
}