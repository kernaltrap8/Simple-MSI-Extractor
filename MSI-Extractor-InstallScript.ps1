# =================================== #
# Simple MSI Extractor install script #
# kernaltrap                          #
# Version 1.3                         #
# =================================== #

$ImportMSIExtractor = "`nImport-Module $HOME\Documents\PowerShell\Modules\MSI-Extractor.ps1"
$MSIExtractorPath = "$PSScriptRoot/MSI-Extractor.ps1"

if (!(Test-Path -Path $ProfilePath)) {
    mkdir $HOME\Documents\PowerShell\Modules | Out-Null
}

if (!(Test-Path -Path $PROFILE.CurrentUserAllHosts)) {
    New-Item -Type File -Path $PROFILE.CurrentUserAllHosts -Force
}

if (!(Test-Path -Path $MSIExtractorPath)) {
    Throw "MSI-Extractor script is not present in this directory!"
} else {
    Copy-Item $PSScriptRoot\MSI-Extractor.ps1 -Destination $HOME\Documents\PowerShell\Modules  
}

if (!(Test-Path -Path $ProfilePath)) {
    Throw "Cannot find Powershell profile script!"
} else {
    Write-Output $ImportMSIExtractor > $PROFILE.CurrentUserAllHosts
    Write-Host "All tests passed, MSI-Extractor is installed."
}
