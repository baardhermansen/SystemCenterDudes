#############################################################################
# Author  : Benoit Lecours
# Website : www.SystemCenterDudes.com
# Twitter : @scdudes
#
# Version : 3.7
# Created : 2014/07/17
# Modified :
# 2014/08/14 - Added Collection 34,35,36
# 2014/09/23 - Changed collection 4 to CU3 instead of CU2
# 2015/01/30 - Improve Android collection
# 2015/02/03 - Changed collection 4 to CU4 instead of CU3
# 2015/05/06 - Changed collection 4 to CU5 instead of CU4
# 2015/05/06 - Changed collection 4 to SP1 instead of CU5
#            - Add collections 37 to 42
# 2015/08/04 - Add collection 43,44
#            - Changed collection 4 to SP1 CU1 instead of SP1
# 2015/08/06 - Change collection 22 query
# 2015/08/12 - Added Windows 10 - Collection 45
# 2015/11/10 - Changed collection 4 to SP1 CU2 instead of CU1, Add collection 46
# 2015/12/04 - Changed collection 4 to CM 1511 instead of CU2, Add collection 47
# 2016/02/16 - Add collection 48 and 49. Complete Revamp of Collections naming. Comment added on all collections
# 2016/03/03 - Add collection 51
# 2016/03/14 - Add collection 52
# 2016/03/15 - Added Error handling and better output
# 2016/08/08 - Add collection 53-56. Modification to collection 4,31,32,33
# 2016/09/14 - Add collection 57
# 2016/10/03 - Add collection 58 to 63
# 2016/10/14 - Add collection 64 to 67
# 2016/10/28 - Bug fixes and updated collection 50
# 2016/11/18 - Add collection 68
# 2017/02/03 - Corrected collection 39 and 68
# 2017/03/27 - Add collection 69,70,71
# 2017/08/25 - Add collection 72
# 2017/11/21 - Add collection 73
# 2018/02/12 - Add collection 74-76. Changed "=" instead of like for OS Build Collections
# 2018/03/27 - Add collection 77-81. Corrected Collection 75,76 to limit to Workstations only. Collection 73 updated to include 1710 Hotfix
# 2018/07/04 - Version 3.0
#            - Add Collection 82-87
#            - Optimized script to run with objects, extended options for replacing existing collections, and collection folder creation when not on site server.
# 2018/08/01 - Add Collection 88
# 2019/04/04 - Add Collection 89-91
# 2019/09/17 - Add Collection 92-94, Windows 2019, Updated Windows 2016
# 2020/01/09 - Add Collection 95-100
# 2021/11/22 - Add Collection 100-133
# 2022/08/24 - Add Collection 133-148
# 2024/11/07 - Add Collection 149-155
# 2024/11/16 - Big rewrite of the script - Baard Hermansen
#
# Purpose : This script create a set of CM collections and move it in an "Operational" folder
# Special Thanks to Joshua Barnette for V3.0
#
#############################################################################

#Load Configuration Manager PowerShell Module
Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0, $Env:SMS_ADMIN_UI_PATH.Length - 5) + '\ConfigurationManager.psd1')

#Get SiteCode
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-Location $SiteCode":"

#Error Handling and output
Clear-Host
$ErrorActionPreference = 'SilentlyContinue'

#Create Default Folder
$CollectionFolder = "Operational"
New-CMFolder -ParentFolderPath "DeviceCollection" -Name $CollectionFolder
$FolderPath = Get-CMFolder -FolderPath ("DeviceCollection" + '\' + $CollectionFolder)

#Set Default limiting collections
$LimitingCollection = "All Systems"

# This query is the base for nearly all collections
$queryBase = "SELECT SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client FROM SMS_R_System"

#Refresh Schedule
$Schedule = New-CMSchedule -RecurInterval Days -RecurCount 7

#Find Existing Collections
$ExistingCollections = Get-CMDeviceCollection -Name "* | *" | Select-Object CollectionID, Name

#List of Collections Query
$Collections = @()

function New-CollectionClientVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$ClientVersion
    )

    if ($ClientVersion.Contains('%')) {
        $operator = "like"
    } else {
        $operator = "="
    }

    $clientObj = [PSCustomObject]@{
        Name               = "CM client version | $($Name)"
        Query              = "$queryBase where SMS_R_System.ClientVersion $operator '$ClientVersion'"
        LimitingCollection = $LimitingCollection
        Comment            = "All systems with CM client version $ClientVersion installed."
    }

    return $clientObj
}

function New-CollectionHardware {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$QueryAddition,

        [Parameter(Mandatory = $true)]
        [string]$LimitingCollection,

        [Parameter()]
        [string]$Comment
    )

    $stringQuery = "$queryBase inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceId = SMS_R_System.ResourceId where SMS_G_System_COMPUTER_SYSTEM"
    $tempObj = [PSCustomObject]@{
        Name               = $Name
        Query              = "$stringQuery.$QueryAddition"
        LimitingCollection = $LimitingCollection
        Comment            = $Comment
    }

    <#
    if ($Model -in "Dell", "HP", "Hewlett-Packard", "Lenovo") {
        $tempObj.Query += ".Manufacturer like '%$Model%'"
        if ($Manufacturer -eq "HP") {
            $tempObj.Query += " or SMS_G_System_COMPUTER_SYSTEM.Manufacturer like '%Hewlett-Packard%'"
        }
    } else {
        $tempObj.Query += " where SMS_G_System_COMPUTER_SYSTEM.Manufacturer = '$Manufacturer' and SMS_G_System_COMPUTER_SYSTEM.Model = '$Model'"
    }
#>
    return $tempObj
}

# All versions of the ConfigMgr client
$cmVersions = @(
    [PSCustomObject]@{
        Name = "R2 CU1"; ClientVersion = "5.00.7958.1203"
    },
    [PSCustomObject]@{
        Name = "R2 CU2"; ClientVersion = "5.00.7958.1303"
    },
    [PSCustomObject]@{
        Name = "R2 CU3"; ClientVersion = "5.00.7958.14"
    },
    [PSCustomObject]@{
        Name = "R2 CU4"; ClientVersion = "5.00.7958.1501"
    },
    [PSCustomObject]@{
        Name = "R2 CU5"; ClientVersion = "5.00.7958.1604"
    },
    [PSCustomObject]@{
        Name = "R2 CU0"; ClientVersion = "5.00.7958.1000"
    },
    [PSCustomObject]@{
        Name = "R2 SP1"; ClientVersion = "5.00.8239.1000"
    },
    [PSCustomObject]@{
        Name = "R2 SP1 CU1"; ClientVersion = "5.00.8239.1203"
    },
    [PSCustomObject]@{
        Name = "R2 SP1 CU2"; ClientVersion = "5.00.8239.1301"
    },
    [PSCustomObject]@{
        Name = "R2 SP1 CU3"; ClientVersion = "5.00.8239.1403"
    },
    [PSCustomObject]@{
        Name = "1511"; ClientVersion = "5.00.8325.1000"
    },
    [PSCustomObject]@{
        Name = "1602"; ClientVersion = "5.00.8355.1000"
    },
    [PSCustomObject]@{
        Name = "1606"; ClientVersion = "5.00.8412.1000"
    },
    [PSCustomObject]@{
        Name = "1610"; ClientVersion = "5.00.8458.1000"
    },
    [PSCustomObject]@{
        Name = "1702"; ClientVersion = "5.00.8498.1000"
    },
    [PSCustomObject]@{
        Name = "1706"; ClientVersion = "5.00.8540.1000"
    },
    [PSCustomObject]@{
        Name = "1710"; ClientVersion = "5.00.8577.1000"
    }
    [PSCustomObject]@{
        Name = "1802"; ClientVersion = '5.00.8634.10%'
    }
    [PSCustomObject]@{
        Name = "1806"; ClientVersion = '5.00.8692.10%'
    }
    [PSCustomObject]@{
        Name = "1810"; ClientVersion = '5.00.8740.10%'
    }
    [PSCustomObject]@{
        Name = "1902"; ClientVersion = '5.00.8790.10%'
    }
    [PSCustomObject]@{
        Name = "1906"; ClientVersion = '5.00.8853.10%'
    }
    [PSCustomObject]@{
        Name = "1910"; ClientVersion = '5.00.8913.10%'
    }
    [PSCustomObject]@{
        Name = "2002"; ClientVersion = '5.00.8968.10%'
    }
    [PSCustomObject]@{
        Name = "2006"; ClientVersion = '5.00.9012.10%'
    }
    [PSCustomObject]@{
        Name = "2010"; ClientVersion = '5.00.9040.10%'
    }
    [PSCustomObject]@{
        Name = "2103"; ClientVersion = '5.00.9049.10%'
    }
    [PSCustomObject]@{
        Name = "2107"; ClientVersion = '5.00.9058.10%'
    }
    [PSCustomObject]@{
        Name = "2111"; ClientVersion = '5.00.9068.10%'
    }
    [PSCustomObject]@{
        Name = "2203"; ClientVersion = '5.00.9078.10%'
    }
    [PSCustomObject]@{
        Name = "2207"; ClientVersion = '5.00.9088.10%'
    }
    [PSCustomObject]@{
        Name = "2303"; ClientVersion = '5.00.9106.10%'
    }
    [PSCustomObject]@{
        Name = "2309"; ClientVersion = '5.00.9120.10%'
    }
    [PSCustomObject]@{
        Name = "2403"; ClientVersion = '5.00.9128.10%'
    }
)

# Create client version collections
foreach ($cmVersion in $cmVersions) {
    $Collections += New-CollectionClientVersion -Name $cmVersion.Name -ClientVersion $cmVersion.ClientVersion
}

# All devices with CM client installed
$Collections += [PSCustomObject]@{
    Name               = "Clients | All"
    Query              = "$queryBase where SMS_R_System.Client = 1"
    LimitingCollection = $LimitingCollection
    Comment            = "All devices detected by CM."
}

# All devices without CM client
$Collections += [PSCustomObject]@{
    Name               = "Clients | No"
    Query              = "$queryBase where SMS_R_System.Client = 0 OR SMS_R_System.Client is NULL"
    LimitingCollection = $LimitingCollection
    Comment            = "All devices without CM client installed."
}

$Collections += [PSCustomObject]@{
    Name               = "Clients Version | Not Latest (2207)"
    Query              = "$queryBase where SMS_R_System.ClientVersion not like '5.00.9088.10%'"
    LimitingCollection = "Clients | All"
    Comment            = "All devices without CM client version 2207."
}

$Collections += [PSCustomObject]@{
    Name               = "Hardware Inventory | Clients not reporting last 14 days"
    Query              = "$queryBase where ResourceId in (select SMS_R_System.ResourceID from SMS_R_System inner join SMS_G_System_WORKSTATION_STATUS on SMS_G_System_WORKSTATION_STATUS.ResourceID = SMS_R_System.ResourceId where DATEDIFF(dd,SMS_G_System_WORKSTATION_STATUS.LastHardwareScan,GetDate()) > 14)"
    LimitingCollection = "Clients | All"
    Comment            = "All devices with CM client that have not communicated with hardware inventory over 14 days."
}

$Collections += [PSCustomObject]@{
    Name               = "Laptops | All"
    Query              = "$queryBase inner join SMS_G_System_SYSTEM_ENCLOSURE on SMS_G_System_SYSTEM_ENCLOSURE.ResourceID = SMS_R_System.ResourceId where SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes in ('8', '9', '10', '11', '12', '14', '18', '21')"
    LimitingCollection = $LimitingCollection
    Comment            = "All laptops."
}

$Collections += New-CollectionHardware -Name "Laptops | Dell" -QueryAddition "Manufacturer like '%Dell%'" -LimitingCollection "Laptops | All" -Comment "All laptops with Dell manufacturer."
$Collections += New-CollectionHardware -Name "Laptops | Lenovo" -QueryAddition "Manufacturer like '%Lenovo%'" -LimitingCollection "Laptops | All" -Comment "All laptops with Lenovo manufacturer."
$Collections += New-CollectionHardware -Name "Laptops | HP" -QueryAddition "Manufacturer like '%HP%' or SMS_G_System_COMPUTER_SYSTEM.Manufacturer like '%Hewlett-Packard%'" -LimitingCollection "Laptops | All" -Comment "All laptops with HP manufacturer."

$Collections += [PSCustomObject]@{
    Name               = "Mobile Devices | All"
    Query              = "select * from SMS_R_System where SMS_R_System.ClientType = 3"
    LimitingCollection = $LimitingCollection
    Comment            = "All Mobile Devices."
}

$Collections += [PSCustomObject]@{
    Name               = "Mobile Devices | Android"
    Query              = "$queryBase INNER JOIN SMS_G_System_DEVICE_OSINFORMATION ON SMS_G_System_DEVICE_OSINFORMATION.ResourceID = SMS_R_System.ResourceId WHERE SMS_G_System_DEVICE_OSINFORMATION.Platform like 'Android%'"
    LimitingCollection = $LimitingCollection
    Comment            = "All Android mobile devices."
}

$Collections += [PSCustomObject]@{
    Name               = "Mobile Devices | iPhone"
    Query              = "$queryBase inner join SMS_G_System_DEVICE_COMPUTERSYSTEM on SMS_G_System_DEVICE_COMPUTERSYSTEM.ResourceId = SMS_R_System.ResourceId where SMS_G_System_DEVICE_COMPUTERSYSTEM.DeviceModel like '%Iphone%'"
    LimitingCollection = $LimitingCollection
    Comment            = "All iPhone mobile devices."
}

$Collections += [PSCustomObject]@{
    Name               = "Mobile Devices | iPad"
    Query              = "$queryBase inner join SMS_G_System_DEVICE_COMPUTERSYSTEM on SMS_G_System_DEVICE_COMPUTERSYSTEM.ResourceId = SMS_R_System.ResourceId where SMS_G_System_DEVICE_COMPUTERSYSTEM.DeviceModel like '%Ipad%'"
    LimitingCollection = $LimitingCollection
    Comment            = "All iPad mobile devices"
}

$Collections += [PSCustomObject]@{
    Name               = "Mobile Devices | Windows Phone 8"
    Query              = "$queryBase inner join SMS_G_System_DEVICE_OSINFORMATION on SMS_G_System_DEVICE_OSINFORMATION.ResourceID = SMS_R_System.ResourceId where SMS_G_System_DEVICE_OSINFORMATION.Platform = 'Windows Phone' and SMS_G_System_DEVICE_OSINFORMATION.Version like '8.0%'"
    LimitingCollection = $LimitingCollection
    Comment            = "All Windows 8 mobile devices"
}

$Collections += [PSCustomObject]@{
    Name               = "Mobile Devices | Windows Phone 8.1"
    Query              = "$queryBase inner join SMS_G_System_DEVICE_OSINFORMATION on SMS_G_System_DEVICE_OSINFORMATION.ResourceID = SMS_R_System.ResourceId where SMS_G_System_DEVICE_OSINFORMATION.Platform = 'Windows Phone' and SMS_G_System_DEVICE_OSINFORMATION.Version like '8.1%'"
    LimitingCollection = $LimitingCollection
    Comment            = "All Windows 8.1 mobile devices"
}

$Collections += [PSCustomObject]@{
    Name               = "Mobile Devices | Windows Phone 10"
    Query              = "$queryBase inner join SMS_G_System_DEVICE_OSINFORMATION on SMS_G_System_DEVICE_OSINFORMATION.ResourceID = SMS_R_System.ResourceId where SMS_G_System_DEVICE_OSINFORMATION.Platform = 'Windows Phone' and SMS_G_System_DEVICE_OSINFORMATION.Version like '10%'"
    LimitingCollection = $LimitingCollection
    Comment            = "All Windows Phone 10"
}
$Collections += New-CollectionHardware -Name "Mobile Devices | Microsoft Surface" -QueryAddition "Model like '%Surface%'" -LimitingCollection $LimitingCollection -Comment "All Windows RT mobile devices."
$Collections += New-CollectionHardware -Name "Mobile Devices | Microsoft Surface 3" -QueryAddition "Model = 'Surface Pro 3' OR SMS_G_System_COMPUTER_SYSTEM.Model = 'Surface 3'" -LimitingCollection $LimitingCollection -Comment "All Microsoft Surface 3."
$Collections += New-CollectionHardware -Name "Mobile Devices | Microsoft Surface 4" -QueryAddition "Model = 'Surface Pro 4'" -LimitingCollection $LimitingCollection -Comment "All Microsoft Surface 4."

$Collections += [PSCustomObject]@{
    Name               = "Others | Linux Devices"
    Query              = "select * from SMS_R_System where SMS_R_System.ClientEdition = 13"
    LimitingCollection = $LimitingCollection
    Comment            = "All systems with Linux"
}

$Collections += [PSCustomObject]@{
    Name               = "Others | MAC OSX Devices"
    Query              = "$queryBase WHERE OperatingSystemNameandVersion LIKE 'Apple Mac OS X%'"
    LimitingCollection = $LimitingCollection
    Comment            = "All workstations with MAC OSX"
}

$Collections += [PSCustomObject]@{
    Name               = "CM | Console"
    Query              = "$queryBase inner join SMS_G_System_ADD_REMOVE_PROGRAMS on SMS_G_System_ADD_REMOVE_PROGRAMS.ResourceID = SMS_R_System.ResourceId where SMS_G_System_ADD_REMOVE_PROGRAMS.DisplayName like '%Configuration Manager Console%'"
    LimitingCollection = $LimitingCollection
    Comment            = "All systems with CM console installed"
}

$Collections += [PSCustomObject]@{
    Name               = "CM | Site Servers"
    Query              = "$queryBase where SMS_R_System.SystemRoles = 'SMS Site Server'"
    LimitingCollection = "Servers | All"
    Comment            = "All systems that is CM site server"
}

$Collections += [PSCustomObject]@{
    Name               = "CM | Site Systems"
    Query              = "$queryBase where SMS_R_System.SystemRoles = 'SMS Site System' or SMS_R_System.ResourceNames in (Select ServerName FROM SMS_DistributionPointInfo)"
    LimitingCollection = $LimitingCollection
    Comment            = "All systems that is CM site system"
}

$Collections += [PSCustomObject]@{
    Name               = "CM | Distribution Points"
    Query              = "$queryBase where SMS_R_System.ResourceNames in (Select ServerName FROM SMS_DistributionPointInfo)"
    LimitingCollection = $LimitingCollection
    Comment            = "All systems that is CM distribution point"
}

$Collections += [PSCustomObject]@{
    Name               = "Servers | All"
    Query              = "$queryBase where OperatingSystemNameandVersion like '%Server%'"
    LimitingCollection = $LimitingCollection
    Comment            = "All servers"
}

$Collections += [PSCustomObject]@{
    Name               = "Servers | Active"
    Query              = "$queryBase inner join SMS_G_System_CH_ClientSummary on SMS_G_System_CH_ClientSummary.ResourceId = SMS_R_System.ResourceId where SMS_G_System_CH_ClientSummary.ClientActiveStatus = 1 and SMS_R_System.Client = 1 and SMS_R_System.Obsolete = 0"
    LimitingCollection = "Servers | All"
    Comment            = "All servers with active state"
}

$Collections += [PSCustomObject]@{
    Name               = "Servers | Physical"
    Query              = "$queryBase where SMS_R_System.ResourceId not in (select SMS_R_SYSTEM.ResourceID from SMS_R_System inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceId = SMS_R_System.ResourceId where SMS_R_System.IsVirtualMachine = 'True') and SMS_R_System.OperatingSystemNameandVersion
 like 'Microsoft Windows NT%Server%'"
    LimitingCollection = "Servers | All"
    Comment            = "All physical servers"
}

$Collections += [PSCustomObject]@{
    Name               = "Servers | Virtual"
    Query              = "$queryBase where SMS_R_System.IsVirtualMachine = 'True' and SMS_R_System.OperatingSystemNameandVersion like 'Microsoft Windows NT%Server%'"
    LimitingCollection = "Servers | All"
    Comment            = "All virtual servers"
}

$Collections += [PSCustomObject]@{
    Name               = "Servers | Windows 2003 and 2003 R2"
    Query              = "$queryBase where OperatingSystemNameandVersion like '%Server 5.2%'"
    LimitingCollection = "Servers | All"
    Comment            = "All servers with Windows 2003 or 2003 R2 operating system"
}

$Collections += [PSCustomObject]@{
    Name               = "Servers | Windows 2008 and 2008 R2"
    Query              = "$queryBase where OperatingSystemNameandVersion like '%Server 6.0%' or OperatingSystemNameandVersion like '%Server 6.1%'"
    LimitingCollection = "Servers | All"
    Comment            = "All servers with Windows 2008 or 2008 R2 operating system"
}

$Collections += [PSCustomObject]@{
    Name               = "Servers | Windows 2012 and 2012 R2"
    Query              = "$queryBase where OperatingSystemNameandVersion like '%Server 6.2%' or OperatingSystemNameandVersion like '%Server 6.3%'"
    LimitingCollection = "Servers | All"
    Comment            = "All servers with Windows 2012 or 2012 R2 operating system"
}

$Collections += [PSCustomObject]@{
    Name               = "Servers | Windows 2016"
    Query              = "$queryBase inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceId = SMS_R_System.ResourceId where OperatingSystemNameandVersion like '%Server 10%' and SMS_G_System_OPERATING_SYSTEM.BuildNumber = '14393'"
    LimitingCollection = "Servers | All"
    Comment            = "All Servers with Windows 2016"
}

$Collections += [PSCustomObject]@{
    Name               = "Servers | Windows 2019"
    Query              = "$queryBase inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceId = SMS_R_System.ResourceId where OperatingSystemNameandVersion like '%Server 10%' and SMS_G_System_OPERATING_SYSTEM.BuildNumber = '17763'"
    LimitingCollection = "Servers | All"
    Comment            = "All Servers with Windows 2019"
}

$Collections += [PSCustomObject]@{
    Name               = "Software Inventory | Clients Not Reporting since 30 Days"
    Query              = "$queryBase where ResourceId in (select SMS_R_System.ResourceID from SMS_R_System inner join SMS_G_System_LastSoftwareScan on SMS_G_System_LastSoftwareScan.ResourceId = SMS_R_System.ResourceId where DATEDIFF(dd,SMS_G_System_LastSoftwareScan.LastScanDate,GetDate()) > 30)"
    LimitingCollection = $LimitingCollection
    Comment            = "All devices with CM client that have not communicated with software inventory over 30 days"
}

$Collections += [PSCustomObject]@{
    Name               = "System Health | Clients Active"
    Query              = "$queryBase inner join SMS_G_System_CH_ClientSummary on SMS_G_System_CH_ClientSummary.ResourceId = SMS_R_System.ResourceId where SMS_G_System_CH_ClientSummary.ClientActiveStatus = 1 and SMS_R_System.Client = 1 and SMS_R_System.Obsolete = 0"
    LimitingCollection = "Clients | All"
    Comment            = "All devices with CM client state active"
}

$Collections += [PSCustomObject]@{
    Name               = "System Health | Clients Inactive"
    Query              = "$queryBase inner join SMS_G_System_CH_ClientSummary on SMS_G_System_CH_ClientSummary.ResourceId = SMS_R_System.ResourceId where SMS_G_System_CH_ClientSummary.ClientActiveStatus = 0 and SMS_R_System.Client = 1 and SMS_R_System.Obsolete = 0"
    LimitingCollection = "Clients | All"
    Comment            = "All devices with CM client state inactive"
}

$Collections += [PSCustomObject]@{
    Name               = "System Health | Disabled"
    Query              = "$queryBase where SMS_R_System.UserAccountControl ='4098'"
    LimitingCollection = $LimitingCollection
    Comment            = "All systems with client state disabled"
}

$Collections += [PSCustomObject]@{
    Name               = "System Health | Obsolete"
    Query              = "select * from SMS_R_System where SMS_R_System.Obsolete = 1"
    LimitingCollection = $LimitingCollection
    Comment            = "All devices with CM client state obsolete"
}
$Collections += [PSCustomObject]@{
    Name               = "Systems | x86"
    Query              = "$queryBase inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_COMPUTER_SYSTEM.SystemType = 'X86-based PC'"
    LimitingCollection = $LimitingCollection
    Comment            = "All systems with 32-bit system type"
}

$Collections += [PSCustomObject]@{
    Name               = "Systems | x64"
    Query              = "$queryBase inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_COMPUTER_SYSTEM.SystemType = 'X64-based PC'"
    LimitingCollection = $LimitingCollection
    Comment            = "All systems with 64-bit system type"
}

$Collections += [PSCustomObject]@{
    Name               = "Systems | Created Since 24h"
    Query              = "select SMS_R_System.Name,SMS_R_System.CreationDate FROM SMS_R_System WHERE DateDiff(dd,SMS_R_System.CreationDate, GetDate()) <= 1"
    LimitingCollection = $LimitingCollection
    Comment            = "All systems created in the last 24 hours"
}

$Collections += [PSCustomObject]@{
    Name               = "Windows Update Agent | Outdated Version Win7 RTM and Lower"
    Query              = "$queryBase inner join SMS_G_System_WINDOWSUPDATEAGENTVERSION on SMS_G_System_WINDOWSUPDATEAGENTVERSION.ResourceID = SMS_R_System.ResourceId inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_WINDOWSUPDATEAGENTVERSION.Version
 < '7.6.7600.256' and SMS_G_System_OPERATING_SYSTEM.Version <= '6.1.7600'"
    LimitingCollection = "Workstations | All"
    Comment            = "All systems with windows update agent with outdated version Win7 RTM and lower"
}

$Collections += [PSCustomObject]@{
    Name               = "Windows Update Agent | Outdated Version Win7 SP1 and Higher"
    Query              = "$queryBase inner join SMS_G_System_WINDOWSUPDATEAGENTVERSION on SMS_G_System_WINDOWSUPDATEAGENTVERSION.ResourceID = SMS_R_System.ResourceId inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_WINDOWSUPDATEAGENTVERSION.Version
 < '7.6.7600.320' and SMS_G_System_OPERATING_SYSTEM.Version >= '6.1.7601'"
    LimitingCollection = "Workstations | All"
    Comment            = "All systems with windows update agent with outdated version Win7 SP1 and higher"
}

$Collections += [PSCustomObject]@{
    Name               = "Workstations | All"
    Query              = "$queryBase where OperatingSystemNameandVersion like 'Microsoft Windows NT Workstation%'"
    LimitingCollection = $LimitingCollection
    Comment            = "All workstations"
}

$Collections += [PSCustomObject]@{
    Name               = "Workstations | Active"
    Query              = "$queryBase inner join SMS_G_System_CH_ClientSummary on SMS_G_System_CH_ClientSummary.ResourceId = SMS_R_System.ResourceId where (SMS_R_System.OperatingSystemNameandVersion like 'Microsoft Windows NT Workstation%' or SMS_R_System.OperatingSystemNameandVersion = 'Windows 7 Enterprise 6.1') and SMS_G_System_CH_ClientSummary.ClientActiveStatus = 1 and SMS_R_System.Client = 1 and SMS_R_System.Obsolete = 0"
    LimitingCollection = "Workstations | All"
    Comment            = "All workstations with active state"
}

$Collections += [PSCustomObject]@{
    Name               = "Workstations | Windows 7"
    Query              = "$queryBase where OperatingSystemNameandVersion like 'Microsoft Windows NT Workstation 6.1%'"
    LimitingCollection = "Workstations | All"
    Comment            = "All workstations with Windows 7 operating system"
}

$Collections += [PSCustomObject]@{
    Name               = "Workstations | Windows 8"
    Query              = "$queryBase where OperatingSystemNameandVersion like 'Microsoft Windows NT Workstation 6.2%'"
    LimitingCollection = "Workstations | All"
    Comment            = "All workstations with Windows 8 operating system"
}

$Collections += [PSCustomObject]@{
    Name               = "Workstations | Windows 8.1"
    Query              = "$queryBase where OperatingSystemNameandVersion like 'Microsoft Windows NT Workstation 6.3%'"
    LimitingCollection = "Workstations | All"
    Comment            = "All workstations with Windows 8.1 operating system"
}

$Collections += [PSCustomObject]@{
    Name               = "Workstations | Windows XP"
    Query              = "$queryBase where OperatingSystemNameandVersion like 'Microsoft Windows NT Workstation 5.1%' or OperatingSystemNameandVersion like 'Microsoft Windows NT Workstation 5.2%'"
    LimitingCollection = "Workstations | All"
    Comment            = "All workstations with Windows XP operating system"
}

$Collections += [PSCustomObject]@{
    Name               = "Workstations | Windows 10"
    Query              = "$queryBase where OperatingSystemNameandVersion like 'Microsoft Windows NT Workstation 10.0%' and Build like '10.0.19%'"
    LimitingCollection = "Workstations | All"
    Comment            = "All workstations with Windows 10 operating system"
}

$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v1507" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.10240'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 operating system v1507" }
    }

$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v1511" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.10586'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 operating system v1511" }
    }

$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v1607" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.14393'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 operating system v1607" }
    }

$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v1703" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.15063'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 operating system v1703" }
    }

$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v1709" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.16299'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 operating system v1709" }
    }

$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 Current Branch (CB)" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.OSBranch = '0'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 CB" }
    }

$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 Current Branch for Business (CBB)" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.OSBranch = '1'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 CBB" }
    }

$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 Long Term Servicing Branch (LTSB)" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.OSBranch = '2'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 LTSB" }
    }

$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 Support State - Current" } },
    @{L     = "Query" ; E = { "$queryBase LEFT OUTER JOIN SMS_WindowsServicingStates ON SMS_WindowsServicingStates.Build = SMS_R_System.build01 AND SMS_WindowsServicingStates.branch = SMS_R_System.osbranch01 where SMS_WindowsServicingStates.State = '2'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "Windows 10 Support State - Current" }
    }

$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 Support State - Expires Soon" } },
    @{L     = "Query" ; E = { "$queryBase LEFT OUTER JOIN SMS_WindowsServicingStates ON SMS_WindowsServicingStates.Build = SMS_R_System.build01 AND SMS_WindowsServicingStates.branch = SMS_R_System.osbranch01 where SMS_WindowsServicingStates.State = '3'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "Windows 10 Support State - Expires Soon" }
    }

$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 Support State - Expired" } },
    @{L     = "Query" ; E = { "$queryBase LEFT OUTER JOIN SMS_WindowsServicingStates ON SMS_WindowsServicingStates.Build = SMS_R_System.build01 AND SMS_WindowsServicingStates.branch = SMS_R_System.osbranch01 where SMS_WindowsServicingStates.State = '4'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "Windows 10 Support State - Expired" }
    }


##Collection 78
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 1802" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.9029.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 1802" }
    }

##Collection 79
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 1803" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.9126.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 1803" }
    }

##Collection 80
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 1708" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.8431.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 1708" }
    }

##Collection 81
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 1705" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.8201.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 1705" }
    }

##Collection 82
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "System Health | Clients Online" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.ResourceId in (select resourceid from SMS_CollectionMemberClientBaselineStatus where SMS_CollectionMemberClientBaselineStatus.CNIsOnline = 1)" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "System Health | Clients Online" }
    }

##Collection 83
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v1803" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.17134'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "Workstations | Windows 10 v1803" }
    }

##Collection 84
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Channel | Monthly" } },
    @{L     = "Query" ; E = { "select SMS_R_System.ResourceId,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from  SMS_R_System inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.cfgUpdateChannel = 'http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Channel | Monthly" }
    }

##Collection 85
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Channel | Monthly (Targeted)" } },
    @{L     = "Query" ; E = { "select SMS_R_System.ResourceId,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from  SMS_R_System inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.cfgUpdateChannel = 'http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Channel | Monthly (Targeted)" }
    }

##Collection 86
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Channel | Semi-Annual (Targeted)" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.cfgUpdateChannel = 'http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Channel | Semi-Annual (Targeted)" }
    }

##Collection 87
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Channel | Semi-Annual" } },
    @{L     = "Query" ; E = { "select SMS_R_System.ResourceId,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from  SMS_R_System inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.cfgUpdateChannel = 'http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Channel | Semi-Annual" }
    }


##Collection 91
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "System Health | Duplicate Device Name" } },
    @{L     = "Query" ; E = { "select R.ResourceID,R.ResourceType,R.Name,R.SMSUniqueIdentifier,R.ResourceDomainORWorkgroup,R.Client from SMS_R_System as r full join SMS_R_System as s1 on s1.ResourceId = r.ResourceId full join SMS_R_System as s2 on s2.Name = s1.Name where s1.Name = s2.Name and s1.ResourceId != s2.ResourceId" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment"; E = { "All systems having a duplicate device record" } }

##Collection 93
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v1809" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.17763'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "Workstations | Windows 10 v1809" }
    }

##Collection 94
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v1903" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.18362'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "Workstations | Windows 10 v1903" }
    }

##Collection 96
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 1808" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.10730.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 1808" }
    }

##Collection 97
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 1902" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.11328.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 1902" }
    }

##Collection 98
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 1908" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.11929.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 1908" }
    }

##Collection 99
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 1912" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.12325.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 1912" }
    }

##Collection 100
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v1909" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.18363'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "Workstations | Windows 10 v1909" }
    }


################################# November 22th 2021 ###############################


##Collection 106
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v2004" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.19041'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 operating system v2004" }
    }

##Collection 107
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v20H2" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.19042'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 operating system v20H2" }
    }

##Collection 108
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v21H1" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.19043'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 operating system v21H1" }
    }

##Collection 109
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v21H2" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.19044'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 operating system v21H2" }
    }

##Collection 110
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 11" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build like '10.0.2%'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | All" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 11 operating system" }
    }

##Collection 111
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 11 v21H2" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.22000'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 11" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 11 operating system v21H2" }
    }

##Collection 112
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2001" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.12430.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2001" }
    }

##Collection 113
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2002" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.12527.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2002" }
    }

##Collection 114
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2003" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.12624.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2003" }
    }

##Collection 115
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2004" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.12730.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2004" }
    }

##Collection 116
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2005" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.12827.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2005" }
    }

##Collection 117
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2006" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.13001.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2006" }
    }

##Collection 118
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2007" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.13029.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2007" }
    }

##Collection 119
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2008" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.13127.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2008" }
    }

##Collection 120
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2009" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.13231.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2009" }
    }

##Collection 121
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2010" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.13328.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2010" }
    }

##Collection 122
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2011" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.13426.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2011" }
    }

##Collection 123
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2012" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.13530.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2012" }
    }

##Collection 124
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2101" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.13628.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2101" }
    }

##Collection 125
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2102" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.13801.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2102" }
    }

##Collection 126
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2103" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.13901.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2103" }
    }

##Collection 127
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2104" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.13929.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2104" }
    }

##Collection 128
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2105" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.14026.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2105" }
    }

##Collection 129
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2106" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.14131.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2106" }
    }

##Collection 130
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2107" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.14228.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2107" }
    }

##Collection 131
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2108" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.14326.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2108" }
    }

##Collection 132
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2109" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.14430.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2109" }
    }

##Collection 133
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2110" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.14527.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2110" }
    }

################################# August 24th 2022 ###############################

##Collection 134
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2111" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.14701.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2111" }
    }

##Collection 135
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2112" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.14729.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2112" }
    }

##Collection 136
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2201" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.14827.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2201" }
    }

##Collection 137
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2202" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.14931.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2202" }
    }

##Collection 138
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2203" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.15028.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2203" }
    }

##Collection 139
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2204" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.15128.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2204" }
    }

##Collection 140
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2205" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.15225.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2205" }
    }

##Collection 141
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2206" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.15330.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2206" }
    }

##Collection 142
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Office 365 Build Version | 2207" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS on SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport like '16.0.15427.%'" }},
    @{L     = "LimitingCollection" ; E = { $LimitingCollection }},
    @{L     = "Comment" ; E = { "Office 365 Build Version | 2207" }
    }

##Collection 146
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Co-Management Enabled" } },
    @{L     = "Query" ; E = { "$queryBase inner join SMS_Client_ComanagementState on SMS_Client_ComanagementState.ResourceId = SMS_R_System.ResourceId where SMS_Client_ComanagementState.ComgmtPolicyPresent = 1 AND SMS_Client_ComanagementState.MDMEnrolled = 1 AND MDMProvisioned = 1" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | All" }},
    @{L     = "Comment"; E = { "All workstations with CM client with Co-Management Enabled" } }

##Collection 147
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Defender ATP Onboarded" } },
    @{L     = "Query" ; E = { "select * from SMS_R_System inner join SMS_G_System_AdvancedThreatProtectionHealthStatus on SMS_G_System_AdvancedThreatProtectionHealthStatus.ResourceId = SMS_R_System.ResourceId where SMS_G_System_AdvancedThreatProtectionHealthStatus.OnboardingState = 1" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | All" }},
    @{L     = "Comment"; E = { "All workstations with CM client with Defender ATP Onboarded" } }

##Collection 148
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Defender ATP Not Onboarded" } },
    @{L     = "Query" ; E = { "select * from SMS_R_System inner join SMS_G_System_AdvancedThreatProtectionHealthStatus on SMS_G_System_AdvancedThreatProtectionHealthStatus.ResourceId = SMS_R_System.ResourceId where SMS_G_System_AdvancedThreatProtectionHealthStatus.OnboardingState = 0" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | All" }},
    @{L     = "Comment"; E = { "All workstations with CM client with Defender ATP Not Onboarded" } }


##Collection 152
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 10 v22H2" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.19045'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 10" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 10 operating system v22H2" }
    }

##Collection 153
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 11 v22H2" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.22621'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 11" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 11 operating system v22H2" }
    }

##Collection 154
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 11 v23H2" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.22631'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 11" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 11 operating system v23H2" }
    }

##Collection 155
$Collections +=
$DummyObject |
    Select-Object @{L = "Name" ; E = { "Workstations | Windows 11 v24H2" } },
    @{L     = "Query" ; E = { "$queryBase where SMS_R_System.Build = '10.0.26100'" }},
    @{L     = "LimitingCollection" ; E = { "Workstations | Windows 11" }},
    @{L     = "Comment" ; E = { "All workstations with Windows 11 operating system v24H2" }
    }

#Check Existing Collections
$Overwrite = 1
$ErrorCount = 0
$ErrorHeader = "The script has already been run. The following collections already exist in your environment:`n`r"
$ErrorCollections = @()
$ErrorFooter = "Would you like to delete and recreate the collections above? (Default = No) "
$ExistingCollections | Sort-Object Name | ForEach-Object { If ($Collections.Name -Contains $_.Name) { $ErrorCount += 1 ; $ErrorCollections += $_.Name } }

#Error
If ($ErrorCount -ge 1) {
    Write-Host $ErrorHeader $($ErrorCollections | ForEach-Object { (" " + $_ + "`n`r") }) $ErrorFooter -ForegroundColor Yellow -NoNewline
    $Overwrite = Read-Host "[Y/N]"
}

#Create Collection And Move the collection to the right folder
If ($Overwrite -ieq "y") {
    $ErrorCount = 0

    ForEach ($Collection In $($Collections | Sort-Object LimitingCollection -Descending)) {
        If ($ErrorCollections -Contains $Collection.Name) {
            Get-CMDeviceCollection -Name $Collection.Name | Remove-CMDeviceCollection -Force
            Write-Host "*** Collection $Collection.Name removed and will be recreated ***"
        }
    }

    ForEach ($Collection In $($Collections | Sort-Object LimitingCollection)) {
        Try {
            New-CMDeviceCollection -Name $Collection.Name -Comment $Collection.Comment -LimitingCollectionName $Collection.LimitingCollection -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
            Add-CMDeviceCollectionQueryMembershipRule -CollectionName $Collection.Name -QueryExpression $Collection.Query -RuleName $Collection.Name
            Write-Host "*** Collection $Collection.Name created ***"
        }
        Catch {
            Write-Host "-----------------"
            Write-Host -ForegroundColor Red ("There was an error creating the: " + $Collection.Name + " collection.")
            Write-Host -ForegroundColor Red $_.Exception.Message
            Write-Host "-----------------"
            $ErrorCount += 1
            Pause
        }

        Try {
            Move-CMObject -FolderPath $FolderPath -InputObject $(Get-CMDeviceCollection -Name $Collection.Name)
            Write-Host *** Collection $Collection.Name moved to $CollectionFolder.Name folder***
        }
        Catch {
            Write-Host "-----------------"
            Write-Host -ForegroundColor Red ("There was an error moving the: " + $Collection.Name + " collection to " + $CollectionFolder.Name + ".")
            Write-Host -ForegroundColor Red $_.Exception.Message
            Write-Host "-----------------"
            $ErrorCount += 1
            Pause
        }
    }

    If ($ErrorCount -ge 1) {
        Write-Host "-----------------"
        Write-Host -ForegroundColor Red "The script execution completed, but with errors."
        Write-Host "-----------------"
        Pause
    }
    Else {
        Write-Host "-----------------"
        Write-Host -ForegroundColor Green "Script execution completed without error. Operational Collections created successfully."
        Write-Host "-----------------"
        Pause
    }
}

Else {
    Write-Host "-----------------"
    Write-Host -ForegroundColor Red ("The following collections already exist in your environment:`n`r" + $($ErrorCollections | ForEach-Object { (" " + $_ + "`n`r") }) + "Please delete all collections manually or rename them before re-executing the script! You can also select Y to do it automatically")
    Write-Host "-----------------"
    Pause
}
