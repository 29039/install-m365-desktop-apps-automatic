# Version 0.1

if (Test-Path -Path (Get-Command 7z.exe -ErrorAction SilentlyContinue).Path) {
    Write-Host "7-Zip (7z.exe) is installed and available in the PATH."
} else {
    Write-Host "7-Zip (7z.exe) is not found in the PATH or not installed (install Chocolatey or Scoop)."
    Exit
}
$LocalODTPath = "C:\IT-Disks\ODT"
$ChocoInstallUrl = "https://raw.githubusercontent.com/open-circle-ltd/chocolatey.microsoft-office-deployment/master/package/tools/chocolateyinstall.ps1"
$MSOConfigBaseUrl = "https://raw.githubusercontent.com/29039/install-m365-desktop-apps-automatic/main"
$latestMsOffice = "$LocalODTPath\latest-msoffice-choco.txt"

# Download the latest ODT choco package to scrape the URL out of
New-Item -Path $latestMsOffice -ItemType File -Force
$client = New-Object System.Net.WebClient
$client.DownloadFile($ChocoInstallUrl, $latestMsOffice)

# Scrape the URL of the latest ODT itself from that package
$urlPackageContent = Get-Content $latestMsOffice
$urlPackageLine = $urlPackageContent | Select-String -Pattern '\$urlPackage = ' | Select-Object -ExpandProperty Line
Invoke-Expression $urlPackageLine
Write-Host $urlPackage

# Cleanup
Remove-Item -Path $latestMsOffice -Force

# Split the variables between the download path and filename
$OdtFilename = $urlPackage.Split('/')[-1]
$latestOdt = "$LocalODTPath\$OdtFilename"
New-Item -Path $latestOdt -ItemType File -Force
$client.DownloadFile($urlPackage, $latestOdt)

# Extract setup.exe from the ODT
& 7z e $latestOdt setup.exe "-o$LocalODTPath" -aoa

# Get the ODT config
$client.DownloadFile("$MSOConfigBaseUrl/Configuration-EnterpriseAndBusiness-English.xml","$LocalODTPath\Configuration-EnterpriseAndBusiness-English.xml")

# Download MS Office, then customize as you see fit to actually install it
Write-Host "`nPlease wait while downloading Office...`n`n" -ForegroundColor 3
& "C:\IT-Disks\ODT\setup.exe" /download "C:\IT-Disks\ODT\Configuration-AfB-29039.xml"
Write-Host "To install, run: `
`& `"C:\IT-Disks\ODT\setup.exe`" /configure `"C:\IT-Disks\ODT\Configuration-EnterpriseAndBusiness-English.xml`"`n `
