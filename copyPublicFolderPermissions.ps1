#copyPublicFolderPermissions.ps1 -- Jimmy Taylor -- 11/12/15
#
#This script copies public folder permissions from one user and applies them to another
#Exchange 2010, no idea if this works in Exchange 2013/2016
#

#Get parameters from the user when they execute the script
param([string]$user = "", [string]$folder = "", [string]$copyUser = "")
"`n"

#Attempt to be funny and informational
Write-Host "Getting the public folder permissions from the user"$user" in the public folder "$folder"." -ForegroundColor Yellow
Write-Host "Hold on to your butts!!!" -ForegroundColor Yellow

#Grab public folder permissions from the user specified via the parameters
Get-PublicFolder -Identity "\$folder" -Recurse | Get-PublicFolderClientPermission -user $user | Select-Object Identity, @{Name=’AccessRights’;Expression={[string]::join(";", ($_.AccessRights))}} | Export-Csv -Path C:\scripts\pubCopy.csv -NoTypeInformation

#Read the permission information from the CSV file in C:\scripts\
$pubFolderPerms = Import-CSV "C:\scripts\pubCopy.csv" | Select-Object `
    @{name='ident';expression={$_.'Identity'}}, `
    @{name='accessRight';expression={$_.'AccessRights'}}
"`n"
Write-Host "Grabbing permission info from the CSV file created and creating new public folder permissions for $copyUser..." -ForegroundColor Yellow

#Write the permissions for the user specified in $copyUser
foreach ($pubFolderPerm in $pubFolderPerms) {
    Add-PublicFolderClientPermission -identity $pubFolderPerm.ident -User $copyUser -AccessRights $pubFolderPerm.accessRight
}
"`n"
Write-Host "All done. Go back to reading reddit."