<#
.Synopsis
   Import an SSL certficate into the Certificate Store
.PREREQUISITES
   Scroll down aND set the $sslBase variable to the directory that contains
   the SSL certificate files.  This is NOT the directory that contains the .zip
   archive -- that will be located by the script. 
.DESCRIPTION
   This script will:
     * Programatically look in that folder for a new folder (based on modified 
       timestamp) that contains a .zip archive.  The .zip archive should contain the 
       SSL certificate.  It's specifically going to look for a .crt file.
     * Copied the .zip archive to the $import directory where it will be extracted.
     * Attempt to locate the correct .crt file and prompt for your approval.
     * Import the .crt file into the Local Machine Personal store.
     * Obtain the thumbprint of the .crt file.
     * Repair the certificate in the Certificate Store.
     * Cleanup by moving data from $import to $old.
   You will need to implement the newly installed certificate into whatever services
   it will be used for (e.g. IIS, RDGW, etc.).
.EXAMPLE
   * Download the .zip archive that contains the SSL certificates and place it in a new
     (dated) folder under $sslBase (e.g. if $sslBase = C:\SSLcerts and the date is 
     January 14, 2022, download the .zip to a new directory named 2022-01-14 like so:
     C:\SSLcerts\2022-01-14)
   * Run this script and the rest is magic.
.INPUTS
   You'll be prompted for input if it's required.
.OUTPUTS
   This script is very chatty in what it does (to aid in debugging and knowing where
   there is an error).
.AUTHOR
   Ken Teague
   ken at onxinc dot com
.VERSION
   1.0.2022.01.14.1203
.CHANGELOG
   2022-01-14 @ 1203
     * Initial release
#>

$sslBase = "C:\SSLcerts"
$import = "$sslBase\Import"
$old = "$sslBase\Old"

#Requires -Version 3

Write-Host "Looking for: $sslBase"
if (Test-Path $sslBase) {
    Write-Host "Found: $sslBase"
    Write-Host "Checking for: $import"
    if (Test-Path $import) {
        Write-Host "Found: $import"
        Write-Host "Checking for existing data in: $import"
        if (Test-Path $import\*) {
            $oldData = $((Get-Date).ToString('yyyy-MM-dd_HHmm'))
            Write-Host "Found existing data: $import"
            Write-Host "Moving existing data: $import -> $old\$oldData"
            New-Item -ItemType Directory -Path "$old\$oldData"
            Move-Item -Path "$import\*" -Destination "$old\$oldData"
        } else {
            Write-Host "$import appears to be empty.  We can proceed with using that to work in ..."
        }
    } else {
        Write-Host "Not found: $import"
        Write-Host "Creating: $import"
        New-Item -ItemType Directory -Path "$import"
    }
    ##  Get latest folder created (excluding "Old")
    Write-Host "Locating the folder with the newly downloaded .zip file that contains the SSL certificate"
    $newFolder = (Get-ChildItem -Path "$sslBase" | Where { $_.PSIsContainer -and $_.Name -ne "Old" -and $_.Name -ne "Import" } | sort CreationTime -Descending | Select -First 1).Name
    Write-Host "Found: $newFolder"
    Write-Host "Looking for .zip file(s)"
    $newFolderFiles = (Get-ChildItem -File "$sslbase\$newFolder\*.zip")
    if ($newFolderFiles.Count -gt 1) {
        Write-Host "More than 1 .zip file found!"
        $newCertZip = @(Get-ChildItem $newFolderFiles | Out-GridView -Title 'Choose a file' -PassThru)
    } else {
        if ($newFolderFiles.Count -eq 1) {
            Write-Host "Found: $newFolderFiles"
            $newCertZip = $newFolderFiles
        } else {
            Write-Host "Unable to locate .zip file!"
            Break
        }
    }
    Copy-Item -Path "$newCertZip" -Destination "$import"
    Expand-Archive -Path "$newCertZip" -DestinationPath "$import"
    Write-Host "Looking for .crt file(s)"
    $newCertFiles = (Get-ChildItem -File "$import\*.crt")
    if ($newCertFiles.Count -gt 1) {
        Write-Host "More than 1 .crt file found!"
        $newCert = "@(Get-ChildItem $newCertFiles | Out-GridView -Title 'Choose a file' -PassThru)"
    } else {
        if ($newCertFiles.Count -eq 1) {
            Write-Host "Found: $newCertFiles"
            $newCert = "$newCertFiles"
            Read-Host -Prompt "If $newCert is the correct file to be working with, press any key to continue or CTRL+C to quit."
            Write-Host "Importing: $newCert"
            Import-Certificate -FilePath "$newCert" -CertStoreLocation Cert:\LocalMachine\My
            Write-Host "Obtaining certificate data: $newCert"
            $certData = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 "$newCert")
            Write-Host "Reparing certificate ..."
            & certutil -repairstore my ($certData).Thumbprint
            Write-Host "Cleaning up ..."
            $processedData = ((Get-Date).ToString('yyyy-MM-dd_HHmm'))
            Write-Host "Moving existing data: $import -> $old\$processedData"
            New-Item -ItemType Directory -Path "$old\$processedData"
            Move-Item -Path "$import\*" -Destination "$old\$processedData"
            Write-Host "SSL import successful!"
            Write-Host "This only completes importing the ceretificate into the Certificate Store and repairing it."
            Write-Host "You still need to implement this certificate for use in whatever services it's used in (e.g. IIS, RDGW, etc.)"
            Break
        } else {
            Write-Host "Unable to locate .crt file!"
            Break
        }
    }
} else {
    Write-Host "Base directory does not exist: $sslBase"
    Break
}