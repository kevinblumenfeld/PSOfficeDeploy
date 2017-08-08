function Install-MSI {
    [CmdletBinding(SupportsShouldProcess = $True)]
    param (
        [parameter(Mandatory = $True, ValueFromPipeline = $true)]
        [string[]] $Computer,

    )

    Begin {

    }
    Process {
        Write-Output $($_)
        CD C:\scripts
        .\psexec.exe -AcceptEula -s -c \\$($_) "c:\scripts\msi.bat"
    }
    End {
    }
}