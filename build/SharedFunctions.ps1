# Connect to SharePoint
function Connect-SharePoint ($Url) {
    if ($UseWebLogin.IsPresent) {
        Connect-PnPOnline $Url -UseWebLogin
    }
    elseif ($CurrentCredentials.IsPresent) {
        Connect-PnPOnline $Url -CurrentCredentials
    }
    else {
        Connect-PnPOnline $Url -Credentials $Credential
    }
}

# Apply tepmplate
function Apply-Template([string]$Template) {    
    Apply-PnPProvisioningTemplate ".\templates\$($Template).pnp"
}

function ParseVersion($VersionString) {
    if ($VersionString -like "*.*.*#*") {
        $vs = $VersionString.Split("#")[0]
        return [Version]($vs)
    }
    if ($VersionString -like "*.*.*.*") {
        $vs = $VersionString.Split(".")[0..2] -join "."
        return [Version]($vs)
    }
    if ($VersionString -like "*.*.*") {
        return [Version]($VersionString)
    }
}

function LoadBundle($Environment) {
    $BundlePath = "$PSScriptRoot\bundle\$Environment"
    Add-Type -Path "$BundlePath\Microsoft.SharePoint.Client.Taxonomy.dll" -ErrorAction SilentlyContinue
    Add-Type -Path "$BundlePath\Microsoft.SharePoint.Client.DocumentManagement.dll" -ErrorAction SilentlyContinue
    Add-Type -Path "$BundlePath\Microsoft.SharePoint.Client.WorkflowServices.dll" -ErrorAction SilentlyContinue
    Add-Type -Path "$BundlePath\Microsoft.SharePoint.Client.Search.dll" -ErrorAction SilentlyContinue
    Add-Type -Path "$BundlePath\Newtonsoft.Json.dll" -ErrorAction SilentlyContinue
    Import-Module "$BundlePath\$Environment.psd1" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
}