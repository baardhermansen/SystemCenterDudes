#############################################################################
# Author  : Benoit Lecours
# Website : www.SystemCenterDudes.com
# Twitter : @scdudes
#
# Version : 1.1
# Created : 2017/04/05
# Modified : 2024/11/24 - Baard Hermansen
#
# Purpose : This script delete collections without members and deployments
#
#############################################################################

# Load Configuration Manager PowerShell Module
Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0, $Env:SMS_ADMIN_UI_PATH.Length - 5) + '\ConfigurationManager.psd1')

# Change to Site Code
Set-Location -Path "$(Get-PSDrive -PSProvider "CMSITE"):"
Clear-Host

Write-Host "`nThis script deletes device collections without members and deployments. You will be prompted before each deletion."
Write-Host "Built-in collections are excluded.`n"
Write-Host "------------------------------------------------------------------------"
Write-Host "Building Devices Collections List. This may take a couple of minutes..."
Write-Host "------------------------------------------------------------------------`n"

[array]$CollectionList = Get-CmDeviceCollection | Where-Object { $_.CollectionID -notlike 'SMS*' -and $_.MemberCount -eq 0 } | Select-Object -Property Name, MemberCount, CollectionID#, IsReferenceCollection

"Found {0} collections without members (MemberCount = 0). `n" -f $CollectionList.Count
"Analyzing list to find collections without deployments... `n"

foreach ($Collection in $CollectionList) {
    $CollectionID = $Collection.CollectionID
    #$GetDeployment = Get-CMDeployment -CollectionName $Collection.Name

    # Delete collection if there's no members and no deployment on the collection
    #If ($null -eq $GetDeployment) {
    If ($null -eq (Get-CMDeployment -CollectionName $Collection.Name)) {
        "Collection '{0}' with ID {1} has no members and deployments." -f $Collection.Name, $CollectionID

        # User Confirmation
        If ((Read-Host -Prompt "Type `"Y`" to delete the collection, any other key to skip") -ieq "y") {
            #Check if Reference collection
            ## There is no check here...
            Try {
                # Delete the collection object
                Remove-CMCollection -Id $CollectionID -Force
                Write-Host -ForegroundColor Green "Collection '$($Collection.Name)' deleted."
            }
            Catch {
                Write-Host -ForegroundColor Red "Failed to delete '$($Collection.Name)' collection!"
                Write-Host -ForegroundColor Red "Error: " + $_.Exception.Message
            }
        }
    }
}
