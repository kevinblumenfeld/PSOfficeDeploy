function Restart-Comp {   
 
    Param (
        [Parameter(Mandatory = $true, 
            ValueFromPipeline = $true)]
        [string[]] $Computers

    )
    Begin {

    }
    Process {
        $_ | % {
            Restart-Computer -ComputerName $_ -AsJob -Force
        }
    }
    End {

    }
}
