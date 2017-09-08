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
            # Start-Sleep -s 90
            # .\psexec.exe -AcceptEula -s -c \\$($args[0]) "c:\oScripts\deploy\exe.bat"
        } -ArgumentList @($_)
    }
    End {
    }
    
}