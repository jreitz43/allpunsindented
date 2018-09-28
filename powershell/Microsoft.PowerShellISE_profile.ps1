Clear-Host

# Set environment variables
$env:OctoExe = 'C:\Tools\OctopusTools.4.39.2\Octo.exe'
$env:OctopusURL = 'https://octopus.whitecloud.com'
$env:OctopusAPIKey = ''
$env:AWSUsername = ''
$env:AWSPassword = ''
$env:repo = "C:\Projects\commondata"

# Import frequently used modules
Import-Module $env:repo\build\octohelper.psm1 -Force
Import-Module $env:repo\build\CDP.psm1 -Force

function code { # Open a file with VS Code
    param([Parameter(Position=1, Mandatory=$false)] [string]$File)

    if($File){ # Pass the file to code to open
        Start-Process code $File
    }
    else { # If no file specified, just open code
        Start-Process code
    }
}

function refresh { Invoke-Expression $profile }

function pro { code $profile } # Open this profile with VS Code

function profile { pro } # Open this profile with VS Code

function cdp { Set-Location $env:repo } # Change directory to cdp repo

function scripts { Set-Location $env:repo\build } # Change directory to the PS Script folder

function setup {
    Param(
        [Parameter(Position=1, Mandatory=$false)] [alias("s")] [string]$ServerName = "localhost",
        [Parameter(Position=2, Mandatory=$false)] [alias("t")] [string]$TenantName = "Tenant",
        [Parameter(Position=3, Mandatory=$false)] [string]$SourceFilePath = "C:\CDP\",
        [Parameter(Position=4, Mandatory=$false)] [alias("n")] [switch]$NewTenant,
        [Parameter(Position=5, Mandatory=$false)] [alias("e")] [switch]$ExistingTenant,
        [Parameter(Position=6, Mandatory=$false)] [switch]$BuildME
    )
    $script = "$env:repo\build\setup.ps1 -ServerName $ServerName -TenantName $TenantName -SourceFilePath $SourceFilePath"
    if($NewTenant) {
        $script += " -NewTenant"
    }
    if($ExistingTenant) {
        $script += " -ExistingTenant"
    }
    if($BuildME) {
        $script += " -BuildME"
    }
    Invoke-Expression $script
}