[CmdletBinding()]
[OutputType([void])]
param(
    [ValidateNotNullOrWhiteSpace()]
    [ValidateScript({Test-Path $_ -IsValid})]
    [Parameter()]
    [string] $OutputDirectory = (Join-Path -Path $PSScriptRoot -ChildPath 'results' -AdditionalChildPath (Get-Date).ToString('yyyyMMddhhmmss')),

    [ValidateNotNull()]
    [Parameter()]
    [hashtable] $AdditionalData = @{},

    [Parameter()]
    [string] $Name
)

Set-StrictMode -Version 3
$ErrorActionPreference = 'Stop'

$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory, $PWD)

Start-Transcript -LiteralPath (Join-Path -Path $OutputDirectory -ChildPath 'transcript.txt') -IncludeInvocationHeader

if (-not ([string]::IsNullOrWhiteSpace($Name))) {
    $AdditionalData["Name"] = $Name
}

msinfo32.exe /nfo (Join-Path -Path $OutputDirectory -ChildPath 'msinfo32.nfo')
msinfo32.exe /report (Join-Path -Path $OutputDirectory -ChildPath 'msinfo32.txt')

function Export-Data {
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [ValidateNotNullOrWhiteSpace()]
        [Parameter(Mandatory, Position = 1)]
        [string] $Name,

        [ValidateNotNull()]
        [Parameter(Mandatory, Position = 2, ValueFromRemainingArguments)]
        [object] $Data
    )

    $Data | Export-Clixml -LiteralPath (Join-Path -Path $OutputDirectory -ChildPath "$Name.xml")
    $Data | Export-Csv -LiteralPath (Join-Path -Path $OutputDirectory -ChildPath "$Name.csv")
    $Data | ConvertTo-Json | Set-Content -LiteralPath (Join-Path -Path $OutputDirectory -ChildPath "$Name.json")
    $Data | ConvertTo-Html | Set-Content -LiteralPath (Join-Path -Path $OutputDirectory -ChildPath "$Name.html")
}

Export-Data -Name 'computerInfo'            -Data Get-ComputerInfo
Export-Data -Name 'volumes'                 -Data @(Get-Volume)
Export-Data -Name 'localUsers'              -Data @(Get-LocalUser)
Export-Data -Name 'localGroups'             -Data @(Get-LocalGroup)
Export-Data -Name 'disks'                   -Data @(Get-Disk)
Export-Data -Name 'partition'               -Data @(Get-Partition)
Export-Data -Name 'physicalDisks'           -Data @(Get-PhysicalDisk)
Export-Data -Name 'cimWin32Bios'            -Data @(Get-CimInstance -ClassName Win32_BIOS)
Export-Data -Name 'cimWin32Processor'       -Data @(Get-CimInstance -ClassName Win32_Processor)
Export-Data -Name 'cimWin32ComputerSystem'  -Data @(Get-CimInstance -ClassName Win32_ComputerSystem)
Export-Data -Name 'cimWin32OperatingSystem' -Data @(Get-CimInstance -ClassName Win32_OperatingSystem)
Export-Data -Name 'cimWin32LogicalDisk'     -Data @(Get-CimInstance -ClassName Win32_LogicalDisk)
Export-Data -Name 'cimWin32LogonSession'    -Data @(Get-CimInstance -ClassName Win32_LogonSession)
Export-Data -Name 'cimWin32LoggedOnUser'    -Data @(Get-CimInstance -ClassName Win32_LoggedOnUser)
if ($AdditionalData.Count -gt 0) {
    Export-Data -Name 'additionalData'      -Data $AdditionalData
}

Stop-Transcript