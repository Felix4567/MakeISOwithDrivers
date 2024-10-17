#Requires -RunAsAdministrator

<#
INFO: 
https://woshub.com/integrate-drivers-to-windows-install-media/
https://w365.dk/index.php/2022/03/19/how-to-add-drivers-to-windows-installation-iso-windows-10-11/
https://github.com/wikijm/PowerShell-AdminScripts/blob/master/Miscellaneous/New-IsoFile.ps1

Requirements:
- Elevated PowerShell
- oscdimg tool, part of Deployment Tools -> Windows ADK (https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install)
- Driverpack that need to be added to the .iso image
- Windows image .iso (latest version for the most up-to-date result)
- 7-Zip https://www.7-zip.org/

Prepare the following:
- Create a folder to work in e.g. C:\Winworks
- Set working directory in new folder
- Place .iso image in the .\ root folder
- Place driverpack in .\ root folder
- Driverpack must consist of a ".exe" or ".msi" fileformat. If not, manually add the drivers to the correct folder
#>

<############### Variables ###############>

# Set working directory (optional)
cd "C:\WinWorks\Creator"

$workDir = Get-Location
$unpackedDrivers = "$workDir\DriversUnpacked"
$mountImage = "$workDir\Mount"
$newISO = "$workDir\OutputISO"
$unpackedISO = "$workDir\ISOunpacked"
$isoImageName = (Get-ChildItem "$workDir" -Filter "*.iso").Name
$diskDrivers = "$workDir\DiskDrivers"

<############### Part 1 ###############>

# Check if PowerShell is running elevated
$isAdmin = ([System.Security.Principal.WindowsPrincipal] [System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin){
    Write-Host "This script requires elevated privileges. Please run PowerShell as an administrator and run the script again." -ForegroundColor Red
    Start-Sleep 10
    exit
}

# Check if 7-Zip is installed
$zipExe = "C:\Program Files\7-Zip\7z.exe"
if (-not (Test-Path -Path $zipExe)){
    Write-Host "7-Zip is not installed, installing via Winget." -ForegroundColor Red
    winget install -e --id 7zip.7zip --force --accept-source-agreements
    Start-Sleep -Seconds 2
}

# check if WindowsADK is installed (OSCDIMG.exe)
$WindowsADK = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\x86\Oscdimg\oscdimg.exe"
if (-not (Test-Path -Path $WindowsADK)){
    Write-Host "WindowsADK is not installed, installing via Winget." -ForegroundColor Red
    winget install -e --id Microsoft.WindowsADK --force --accept-source-agreements
    Start-Sleep -Seconds 2
}


# Create folders
Write-Host "Creating required folders." -ForegroundColor Green
$requiredFolders = @(
    "$workDir\DriversUnpacked",
    "$workDir\Mount",
    "$workDir\OutputISO",
    "$workDir\ISOunpacked"
) 
foreach ($folder in $requiredFolders){
    if (-not (Test-Path -Path "$folder")){
        New-Item -ItemType Directory "$folder"
    }
}

Write-Host ""

# Extract .msi driverpack.
$MSILocation = (Get-ChildItem "$workDir" -Filter "*.msi").FullName
if ($MSILocation){
    Write-Host "Starting extraction of msi driverpack" -ForegroundColor Green
    Start-Process "MSIEXEC" -ArgumentList "/a `"$MSILocation`" /qn TARGETDIR=`"$unpackedDrivers`"" -Wait -NoNewWindow
    Write-Host "Msi driverpack extracted successfully in $unpackedDrivers" -ForegroundColor Green
    Write-Host ""
    $ExecutableName = (Get-ChildItem "$workDir" -Filter "*.msi").Name.split('.')[0]
} else {
    Write-Host "No .msi driverpack found, continuing..."
    Write-Host ""
}

# Extract .exe driverpack.
$EXELocation = (Get-ChildItem "$workDir" -Filter "*.exe").FullName
if ($EXELocation){
    Write-Host "Starting extraction of exe driverpack" -ForegroundColor Green
    & $zipExe x $EXELocation "-o$($unpackedDrivers)" -y
    Write-Host "Exe driverpack extracted successfully in $unpackedDrivers" -ForegroundColor Green
    Write-Host ""
    $ExecutableName = (Get-ChildItem "$workDir" -Filter "*.exe").Name.split('.')[0]
} else {
    Write-Host "No .exe driverpack found, continuing..."
    Write-Host ""
}

<############### Part 2 ###############>

# Extract the .iso to $ISOunpacked
Write-Host "Starting extraction .iso image to $unpackedISO." -ForegroundColor Green
$getIsoImage = Get-ChildItem .\ -Filter "*.iso"
& $zipExe x "$getIsoImage" "-o$($unpackedISO)" -Wait -PassThru

Write-Host ""

# Display all editions in the image (Home, Pro, Enterprise etc. Change the index accordingly in the next step)
Write-Host "Generating list of Windows editions." -ForegroundColor Green
Start-Sleep -Seconds 2
Dism /Get-WimInfo /WimFile:$unpackedISO\Sources\install.wim

Write-Host ""
$indexInput = $(Write-Host "Choose Windows edition (index number) " -ForegroundColor Green -NoNewline; Read-Host)
$indexOutput = $indexInput.toint32($null)
Write-Host ""

# Export the chosen Windows version
Write-Host "Exporting Windows edition (Index number $indexOutput)." -ForegroundColor Green
Dism /Export-Image /SourceImageFile:"$unpackedISO\sources\install.wim" /SourceIndex:$indexOutput /DestinationImageFile:"$workDir\install.wim"
# $winVersion = [int]3
# Dism /Export-Image /SourceImageFile:"$unpackedISO\sources\install.wim" /SourceIndex:$winVersion /DestinationImageFile:"$workDir\install.wim"
Start-Sleep -Seconds 2
Move-Item -LiteralPath "$workDir\install.wim" "$unpackedISO\Sources\install.wim" -Force
Start-Sleep -Seconds 2

# Mount the image
Write-Host "Mounting the image at $mountImage." -ForegroundColor Green
Start-Sleep -Seconds 2
Dism /Mount-Image /ImageFile:$unpackedISO\Sources\install.wim /Index:1 /MountDir:$mountImage

# Add drivers to the image
Write-Host "Adding drivers to the image." -ForegroundColor Green
Start-Sleep -Seconds 2
$inf = (get-childitem $unpackedDrivers -Filter *.inf -Recurse)
if (!$inf){echo "yes"}
Dism /Image:$mountImage /Add-Driver /Driver:$unpackedDrivers /Recurse /ForceUnsigned

# Add internal disk drivers to image
Write-Host "Adding driver for internal disks..."
Start-Sleep -Seconds 2
Dism /Image:$mountImage /Add-Driver /Driver:$diskDrivers /Recurse /ForceUnsigned

# Commit the changes and unmount
Write-Host "Committing changes and unmounting." -ForegroundColor Green
Start-Sleep -Seconds 2
Dism /Unmount-Image /MountDir:$mountImage /Commit

# Create a new .iso in $OutputISO
Write-Host "Creating a new image in $newISO." -ForegroundColor Green
Start-Sleep -Seconds 2
& "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\x86\Oscdimg\oscdimg.exe" -h -m -o -u2 -udfver102 -bootdata:2#p0,e,b$unpackedISO\boot\etfsboot.com#pEF,e,b$unpackedISO\efi\microsoft\boot\efisys.bin -lWin10 $unpackedISO "$newISO\[$ExecutableName]$isoImageName"

#& .\oscdimg.exe -h -m -o -u2 -udfver102 -bootdata:2#p0,e,b$unpackedISO\boot\etfsboot.com#pEF,e,b$unpackedISO\efi\microsoft\boot\efisys.bin -lWin10 $unpackedISO "$newISO\[Drivers]$isoImageName"

# Remove files
Write-Host "Cleaning up files in $($unpackedDrivers)" -ForegroundColor Green
Remove-Item $unpackedDrivers -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Cleaning up files in $($mountImage)" -ForegroundColor Green
Remove-Item $mountImage -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Cleaning up files in $($unpackedISO)" -ForegroundColor Green
Remove-Item $unpackedISO -Recurse -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3
Write-Host ""

# Done
Write-Host "Done. The new image is located in $newISO." -ForegroundColor Green
