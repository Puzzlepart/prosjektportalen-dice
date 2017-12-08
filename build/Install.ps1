<#

.SYNOPSIS
This script will install DICE for Prosjektportalen to a site collection

.DESCRIPTION
Use the required -Url param to specify the target site collection. You can also install assets and default data to other site collections. The script will provision all the necessary lists, files and settings necessary for Prosjektportalen to work.

.EXAMPLE
./Install.ps1 -Url https://puzzlepart.sharepoint.com/sites/prosjektportalen

.LINK
https://github.com/Puzzlepart/prosjektportalen

#>

Param(
    [Parameter(Mandatory = $true, HelpMessage = "Where do you want to install the Project Portal?")]
    [string]$Url,
    [Parameter(Mandatory = $false, HelpMessage = "Do you want to handle PnP libraries and PnP PowerShell without using bundled files?")]
    [switch]$SkipLoadingBundle,
    [Parameter(Mandatory = $false, HelpMessage = "Stored credential from Windows Credential Manager")]
    [string]$GenericCredential,
    [Parameter(Mandatory = $false, HelpMessage = "Use Web Login to connect to SharePoint. Useful for e.g. ADFS environments.")]
    [switch]$UseWebLogin,
    [Parameter(Mandatory = $false, HelpMessage = "Use the credentials of the current user to connect to SharePoint. Useful e.g. if you install directly from the server.")]
    [switch]$CurrentCredentials,
    [Parameter(Mandatory = $false, HelpMessage = "PowerShell credential to authenticate with")]
    [System.Management.Automation.PSCredential]$PSCredential,
    [Parameter(Mandatory = $false, HelpMessage = "Installation Environment. If SkipLoadingBundle is set, this will be ignored")]
    [ValidateSet('SharePointPnPPowerShell2013', 'SharePointPnPPowerShell2016', 'SharePointPnPPowerShellOnline')]
    [string]$Environment = "SharePointPnPPowerShellOnline",
    [Parameter(Mandatory = $false)]
    [ValidateSet('None', 'File', 'Host')]
    [string]$Logging = "File"
)

. ./SharedFunctions.ps1

# Loads bundle if switch SkipLoadingBundle is not present
if (-not $SkipLoadingBundle.IsPresent) {
    LoadBundle -Environment $Environment
}

# Handling credentials
if ($PSCredential -ne $null) {
    $Credential = $PSCredential
}
elseif ($GenericCredential -ne $null -and $GenericCredential -ne "") {
    $Credential = Get-PnPStoredCredential -Name $GenericCredential -Type PSCredential 
}
elseif ($Credential -eq $null -and -not $UseWebLogin.IsPresent -and -not $CurrentCredentials.IsPresent) {
    $Credential = (Get-Credential -Message "Please enter your username and password")
}

# Start installation
function Start-Install() {  
    # Prints header
    if (-not $Upgrade.IsPresent) {
        Write-Host "############################################################################" -ForegroundColor Green
        Write-Host "" -ForegroundColor Green
        Write-Host "Installing DICE for Prosjektportalen" -ForegroundColor Green
        Write-Host "Maintained by Puzzlepart @ https://github.com/Puzzlepart/prosjektportalen-dice" -ForegroundColor Green
        Write-Host "" -ForegroundColor Green
        Write-Host "Installation URL:`t`t$Url" -ForegroundColor Green
        Write-Host "Environment:`t`t`t$Environment" -ForegroundColor Green
        Write-Host "" -ForegroundColor Green
        Write-Host "############################################################################" -ForegroundColor Green
    }

    # Starts stop watch
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $ErrorActionPreference = "Stop"

    # Sets up PnP trace log
    if ($Logging -eq "File") {
        $execDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
        Set-PnPTraceLog -On -Level Debug -LogFile "pplog-$execDateTime.txt"
    }
    elseif ($Logging -eq "Host") {
        Set-PnPTraceLog -On -Level Debug
    }
    else {
        Set-PnPTraceLog -Off
    }
  

    try {
        Connect-SharePoint $Url -UseWeb
        Write-Host "Deploying required lists and files.. " -ForegroundColor Green -NoNewLine
        Apply-Template -Template "root"
        Write-Host "DONE" -ForegroundColor Green
        Disconnect-PnPOnline
    }
    catch {
        Write-Host
        Write-Host "Error installing root template to $Url" -ForegroundColor Red 
        Write-Host $error[0] -ForegroundColor Red
        exit 1 
    }

    $sw.Stop()
    if (-not $Upgrade.IsPresent) {
        Write-Host "Installation completed in $($sw.Elapsed)" -ForegroundColor Green
    }
}

Connect-SharePoint $Url -UseWeb  
$MinPPVersion = ParseVersion -VersionString "2.1.21"
$CurrentPPVersion = ParseVersion -VersionString (Get-PnPPropertyBag -Key pp_version)

if ($CurrentPPVersion -ge $MinPPVersion) {
    Start-Install
}
else {
    Write-Host 
    Write-Host "ERROR: DICE for Prosjekportalen requires Prosjektportalen 2.1.21 or newer installed." -ForegroundColor Red 
    Write-Host 
}
