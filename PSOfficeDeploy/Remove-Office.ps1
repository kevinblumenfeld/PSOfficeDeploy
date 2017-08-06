function Get-RegValue {
    param (
        [parameter(Mandatory = $False)] [string] $ComputerName = "10.20.131.92",
        [parameter(Mandatory = $True)] [string] $KeyPath,
        [parameter(Mandatory = $True)] [string] $ValueName
    )
    if ($ComputerName -ne "") {
        try {
            $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $ComputerName) 
            $RegSubKey = $Reg.OpenSubKey("$KeyPath")
            $RegSubKey.GetValue("$ValueName")
        }
        catch {
            Write-Output "error: unable to access registry key/value."
        }
    }
    else {
        <# 
        $reg = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,'Default')
        $RegSubKey = $Reg.OpenSubKey("$KeyPath")
        $RegSubKey.GetValue("$ValueName")
        #>
    }
}