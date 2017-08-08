function Send-FileAsJob {   
    <#
.SYNOPSIS
Function to copy files and directories recursively
Uses PowerShell jobs to help parallelize tasks

.DESCRIPTION
Function to copy directories and files recursively
Also can be used to copy a file or file(s)
When the destination is a UNC path the computers being passed at the pipeline can be passed to %C in destination path (%C is case sensitive)
ex. get-content ./hostnames.txt | Send-FileAsJob -SourceDirsOrFiles "D:\scripts\R" -DestinationDir "\\%C\C$\scripts\R"
At runtime, each computer name in the hostnames.txt will be populated in place of %C

.PARAMETER Computers
Parameter description

.PARAMETER Directory
Parameter description

.PARAMETER File
Parameter description

.PARAMETER SourceDirsOrFiles
Parameter description

.PARAMETER DestinationDir
Parameter description

.EXAMPLE
get-content ./hostnames.txt | Send-FileAsJob -SourceDirsOrFiles "C:\oScripts\" -DestinationDir "\\%C\C$\oScripts\"
IMPORTANT: INCLUDE THE TRAILING BACKSLASH AFTER FINAL DIRECTORY IN PATH!!!
#>

    Param (
        [Parameter(Mandatory = $true, 
            ValueFromPipeline = $true)]
        [string[]] $Computers,

        [Parameter(Mandatory = $false)]
        [string[]] $SourceDirsOrFiles,

        [Parameter(Mandatory = $false)]
        [string] $DestinationDir

    )
    Begin {
        $null = New-Item -Path "c:\ps\logs" -ItemType Directory -ErrorAction SilentlyContinue
        $logSuccess = ($(get-date -Format yyyy-MM-dd_HH-mm-ss) + "_CopyToComputer_SUCCESS.csv")
        $logError = ($(get-date -Format yyyy-MM-dd_HH-mm-ss) + "_CopyToComputer_FAILED.csv")
    }
    Process {
        $_ | % {
            Start-Job -ScriptBlock {
                $comp = $using:_
                $source = $using:SourceDirsOrFiles
                $destination = $using:DestinationDir
                $ulogSuccess = $using:logSuccess
                $ulogError = $using:logError
                if ($destination -match [Regex]::Escape('\\%C\')) {
                    $destination = $destination.replace('%C', $comp)
                }
                $dirs = Get-ChildItem -Path $source -Directory -Recurse -Force
                if ($dirs) {
                    $dirs | ForEach-Object {
                        $NewDir = $_.FullName.replace($source, $destination)
                        If (Test-Path $NewDir) {
                            Write-Output "$NewDir already exists"
                        }
                        Else {
                            Try {
                                $null = New-Item -Type Directory $NewDir
                                "$comp, $NewDir, Success" | Out-File "c:\ps\logs\$ulogSuccess" -Encoding ascii -Append
                            }
                            Catch {
                                "$comp, $NewDir, $_" | Out-File "c:\ps\logs\$ulogError" -Encoding ascii -Append
                            }
                        }
                    }
                    $files = Get-ChildItem -path $source -Recurse -Force -File
                    $files| ForEach-Object { 
                        Copy-item $_.fullname $_.FullName.replace($source, $destination) -force
                        Write-Output $($_.FullName.replace($source, $destination))    
                    }                           
                } 
                Else {
                    Copy-Item ($source + "*") $destination -Force
                }
            } 
        }
    }
    End {

    }
}
