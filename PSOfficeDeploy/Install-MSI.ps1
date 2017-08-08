function Install-MSI {
    [CmdletBinding(SupportsShouldProcess = $True)]
    param (
        [parameter(Mandatory = $True, ValueFromPipeline = $true)]
        [string[]] $Computer

    )

    Begin {
    }
    Process {
        Write-Output $($_)
        Start-Job -ScriptBlock {
            CD c:\oscripts\deploy
            .\psexec.exe -AcceptEula -s -c \\$($args[0]) "c:\oScripts\deploy\msi.bat"
        } -ArgumentList @($_)
    }
    End {
    }
    
}