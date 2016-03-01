### Description: Script to get statistics of Exchange mailboxes and save it to SharePoint server
### Author: Andrey Nesterov
### Created: 22.05.2015
### Modified: 1.10.2015
### ver. 2.1
### Requirements ###
<#
$Requirements='
Script requires Full Access to Mailbox
- to apply Full Access execute:
Remove-MailboxPermission -Identity "UserMailbox" -User "AdminName" -AccessRights Fullaccess -Deny -InheritanceType all
Add-MailboxPermission -Identity "UserMailbox" -User "AdminName" -AccessRights Fullaccess -InheritanceType all
- for all DB:
Get-Mailbox -Database $mailboxDB -RecipientTypeDetails UserMailbox | foreach-object {
Remove-MailboxPermission -Identity $_ -User "AdminName" -AccessRights Fullaccess -Deny -InheritanceType all
Add-MailboxPermission -Identity $_ -User "AdminName" -AccessRights Fullaccess -InheritanceType all }
- var. 2 to add full permissions
Get-MailboxDatabase -Identity db01 | Add-ADPermission -User %username% -AccessRights genericall # Убираются права аналогично - командлетом Remove-ADPermission
'
#>

cls
###################################
### Import modules, confs, libs ###
###################################
Import-Module -Name “C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll”
add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction silentlycontinue
Import-Module activedirectory -ea SilentlyContinue

# Conf file
[string]$CurrentPath = Split-Path -Path $($MyInvocation.MyCommand.Path)
$FileConfParams=$args[0]  
. "$CurrentPath\..\FS\$FileConfParams"

#. "$CurrentPath\$FileConfParams" #Conf-FS-pers-fld-sizes.ps1"
#$FileConfParams="Conf-MX-pers-fld-sizes-test.ps1"

# Libraries
$libDir = "$CurrentPath\..\Libs"
. "$libDir\SaveToSP.ps1"

##########################################
### ---End Import modules, confs, libs ###
##########################################

##########################
#### Global variables ####
##########################
$Debug=$false
$SaveToSP=$true

$ErrorActionPreference = "Stop"
#[string]$CurrentPath = Split-Path -Path $($MyInvocation.MyCommand.Path)
$datetime=get-date
$fileDate = $datetime.ToString("yyyyMMdd-HHmmss")
$itemsCount = 0
$itemsSumSize = 0
$itemsCountDB = 0
$itemsSumSizeDB = 0
$MinAbsAgeLimit1 = 0
$MaxAbsAgeLimit1 = 182
$MinAbsAgeLimit2 = 183
$MaxAbsAgeLimit2 = 365
$MinAbsAgeLimit3 = 366
$MaxAbsAgeLimit3 = 730
$MinAbsAgeLimit4 = 731
$ItemSizeLimit1 = 524288
$ItemSizeLimit2 = 3145728
$MBXcount=0

$MinAgeLimit1 = -$MinAbsAgeLimit1
$MaxAgeLimit1 = -$MaxAbsAgeLimit1
$MinAgeLimit2 = -$MinAbsAgeLimit2
$MaxAgeLimit2 = -$MaxAbsAgeLimit2
$MinAgeLimit3 = -$MinAbsAgeLimit3
$MaxAgeLimit3 = -$MaxAbsAgeLimit3
$MinAgeLimit4 = -$MinAbsAgeLimit4

$ArrMailboxesDB = @()
$MbxTmp=$null
$DBTmp=$null
$folders=$null
$MBXsize=0
$MBXitems=0
$CurDateTime1=get-date
##################################
#### --End Global variables-- ####
##################################

###########################
#### Process Variables ####
###########################
### Create base variables from Exchange.WebServices API
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010)
$service.UseDefaultCredentials = $true 
$service.Url = “https://mx.test.local/ews/exchange.asmx”
$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
$itemFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100000)
### 

# Temp csv file to cache mbx statistics
$tmpFileDate = $CurDateTime1.ToString("yyyyMMdd-HHmmss")
$tmpCSVFileName="tmp-MX-report-Users-$tmpFileDate.csv"
$tmpSCVFile="$CurrentPath\$tmpCSVFileName"
New-Item -Path $CurrentPath -Name "tmp-MX-report-Users-$tmpFileDate.csv" -ItemType File
#$tmpSCVFile="$CurrentPath\tmp-MX-report-Users-20151001-060027.csv"

# Log
#New-Item -Path $CurrentPath -Name "MX-mbx-stat-attachm-$tmpFileDate.log" -ItemType File

$svcAcc=whoami
$ItemSizeLimitMB1 = [math]::round($ItemSizeLimit1 / (1024*1024), 4)
$ItemSizeLimitMB2 = [math]::round($ItemSizeLimit2 / (1024*1024), 4)

$passwdFile = "C:\Scripts\mx-mbx-stat-attachm.tmp"

########################################
### --End Process Income Variables-- ###
########################################


################
### Funtions ###
################
#####
### Fx1. Echo-Debug
#####
function Echo-Debug ($EchoString) {
    if ($Debug) {echo $EchoString}
    # Log
    #Add-Content  -Path "$CurrentPath\MX-mbx-stat-attachm-$tmpFileDate.log" -Value $EchoString

}

#####
### Fx2. Get-OURigthPart
#####
function Get-OURigthPart ($OUString) {
    $pos=$OUString.IndexOf(",") # first comma after CN
    $fxUserOU=$OUString.Substring($pos+1)
    return $fxUserOU
}

#####
### Fx3. Mirror OUs DN
#####
function Mirror-OU ($OUString) {
    $OUString.split(",") | where {$_ -notmatch "(OU=Spec|OU=Users|DC=test|DC=local)"} | foreach {
        $mirrOUfx="$_/"+$mirrOUfx
    }
    $mirrOUfx = $mirrOUfx -replace "OU=",""
    return $mirrOUfx.Substring(0,$mirrOUfx.Length-1)
}

#####
### Fx4. Read and count items from folder
#####
function doFolder($inFolder) {
    $items=$null
    echo-debug $($inFolder.DisplayName)
    try {
        $items = $inFolder.FindItems($itemFilter, $itemView)
    } catch {
        echo-debug "----- Error occured in folder -----"
        return
    }
    foreach ($item in $items.Items) {
        try {
            $ItemAge = (New-TimeSpan $datetime $item.DateTimeReceived).Days
            $ItemAgeAbs = [math]::abs($ItemAge)
            $ItemSize = [int]$item.size
            $ItemSubject = $item.Subject
            $FolderName = $inFolder.displayName
        } catch {
            echo-debug "----- Error occured in item -----"
            continue
        }
        #try {$item.Load()} catch {echo-debug "---- Item load error in folder $inFolder.DisplayName"; continue}
        #foreach($attachment in $item.Attachments) {
            #$AttachmSize = [int]$attachment.size
        if ($ItemAgeAbs -ge $MinAbsAgeLimit1 -and $ItemAgeAbs -le $MaxAbsAgeLimit1) {
            if  ($ItemSize -gt $ItemSizeLimit1) {
                $global:itemsCount1++
                #echo-debug "---" $MaxAbsAgeLimit1 $ItemSizeLimit1 $ItemSubject $ItemSize
            }
            if  ($ItemSize -gt $ItemSizeLimit2) {
                $global:itemsCount2++
                #echo-debug "---" $MaxAbsAgeLimit1 $ItemSizeLimit2 $ItemSubject $ItemSize
            }
        }
        if ($ItemAgeAbs -ge $MinAbsAgeLimit2 -and $ItemAgeAbs -le $MaxAbsAgeLimit2) {
            if  ($ItemSize -gt $ItemSizeLimit1) {
                $global:itemsCount3++
            }
            if  ($ItemSize -gt $ItemSizeLimit2) {
                $global:itemsCount4++
            }
        }
        if ($ItemAgeAbs -ge $MinAbsAgeLimit3 -and $ItemAgeAbs -le $MaxAbsAgeLimit3) {
            if  ($ItemSize -gt $ItemSizeLimit1) {
                $global:itemsCount5++
            }
            if  ($ItemSize -gt $ItemSizeLimit2) {
                $global:itemsCount6++
            }
        }
        if ($ItemAgeAbs -ge $MinAbsAgeLimit4) {
            if  ($ItemSize -gt $ItemSizeLimit1) {
                $global:itemsCount7++
            }
            if  ($ItemSize -gt $ItemSizeLimit2) {
                $global:itemsCount8++
            }
        }
        #}
    }
    try {
        $folders = $service.FindFolders($inFolder.Id, $folderView)
    } catch {
        echo-debug "----- Error occured in find folders method -----"
    }

    foreach ($folder in $folders.Folders) {
        doFolder($folder)
    }
}

#####
### Fx5. Search string in tmp file
#####
function Get-tmpString ($tmpString) {
    $tmpSCVstring=$null
    [string]$tmpSCVstring=(Select-String -Path $tmpSCVFile -Pattern $tmpString)
    $tmpSCVstring2=$tmpSCVstring.Split(":")[2]
    #echo-debug $tmpSCVstring
    #echo-debug $tmpSCVstring2
    if ($tmpSCVstring2) {
        $tmpSCVstringArr=@()
        $tmpSCVstringArr=$tmpSCVstring2.Split(";")
        echo-debug $tmpSCVstringArr
        #$tmpSCVstringArr | foreach {echo-debug "tmpSCVstringArr[i]: $_"}
        #$global:UserNameOUDescr=$tmpSCVstringArr[0]
        $global:UsermirrOU=$tmpSCVstringArr[1]
        #$global:MailboxAddress=$tmpSCVstringArr[2]
        [int]$global:MBXsize=$tmpSCVstringArr[3]
        [int]$global:MBXitems=$tmpSCVstringArr[4]
        $global:MBXDBName=$tmpSCVstringArr[5]
        [int]$global:itemsCount1=$tmpSCVstringArr[6]
        [int]$global:itemsCount2=$tmpSCVstringArr[7]
        [int]$global:itemsCount3=$tmpSCVstringArr[8]
        [int]$global:itemsCount4=$tmpSCVstringArr[9]
        [int]$global:itemsCount5=$tmpSCVstringArr[10]
        [int]$global:itemsCount6=$tmpSCVstringArr[11]
        [int]$global:itemsCount7=$tmpSCVstringArr[12]
        [int]$global:itemsCount8=$tmpSCVstringArr[13]
        return $true
    } else {
        return $false
    }
    
}

#####
### Fx6. Make CSVs
#####
function Make-CSVs ($FileInd){
    # CSV headers
    $ItemReportCSVclmns1=";;;;;;$MinAbsAgeLimit1<age<$MaxAbsAgeLimit1 (days);;$MinAbsAgeLimit2<age<$MaxAbsAgeLimit2 (days);;$MinAbsAgeLimit3<age<$MaxAbsAgeLimit3 (days);;age>$MinAbsAgeLimit4"
    $ItemReportCSVclmns2="Description;OU path;Mailbox;Size (MB); Items;DataBase;size>$ItemSizeLimitMB1 (MB);size>$ItemSizeLimitMB2 (MB);size>$ItemSizeLimitMB1 (MB);size>$ItemSizeLimitMB2 (MB);size>$ItemSizeLimitMB1 (MB);size>$ItemSizeLimitMB2 (MB);size>$ItemSizeLimitMB1 (MB);size>$ItemSizeLimitMB2 (MB)"
    $ItemReportCSVclmnsDepts="Description;OU path;Size (MB)"
    $datetime=get-date
    $fileDate = $datetime.ToString("yyyyMMdd-HHmmss")
    $global:FileNameDepts = "MX-report-$FileInd-Depts-$fileDate.csv" #+ ".csv" # + "-" + $FileInd
    $global:FileNameUsers = "MX-report-$FileInd-Users-$fileDate.csv" #+ ".csv"
    $global:ExpFileCSVDepts = $script:CurrentPath + "\" + $FileNameDepts
    $global:ExpFileCSVUsers = $script:CurrentPath + "\" + $FileNameUsers
    # Add headers to files
    Add-Content -Path  $global:ExpFileCSVUsers -Value $ItemReportCSVclmns1
    Add-Content -Path  $global:ExpFileCSVUsers -Value $ItemReportCSVclmns2
    Add-Content -Path  $global:ExpFileCSVDepts -Value $ItemReportCSVclmnsDepts
}

#####
### Fx7. Main function. Get-MX statistics
#####
function Get-PersStatMX ($ADUsersOU, $ReportSave) {
    echo-debug "ADUsersOU.Keys: $ADUsersOU"
    $hashOUsizes=@{}
    # SP report address options
    # Example. $ReportSave="https://sharepoint.test.local/it/it_inf;Reports/FS;ALL";
    $ReportSaveSplit=$ReportSave.split(";")
    $SiteURL = $ReportSaveSplit[0]
    $Library = $ReportSaveSplit[1] -replace "FS","MX"
    $FileInd = $ReportSaveSplit[2]
    
    # Make CSV reports files
    Make-CSVs $FileInd

    # Get users from OUs, get mailbox and statistics
    foreach ($OUkey in $ADUsersOU.keys) { 
        Echo-Debug "`n`nOUkey: $OUkey"
        $getADUsers=Get-ADUser -SearchBase $OUkey -f * 
        foreach ($ADuser in $getADUsers) {
            echo-debug "`n=================="
            echo-debug "`nADUser: $ADuser"
            # Get mailbox
            try {
                $mailbox = Get-Mailbox -Identity $ADuser.samaccountname -ErrorAction stop
                $MailboxAddress = $mailbox.primarySmtpAddress.ToString()
                echo-debug "Primary address: $MailboxAddress"
            } catch {
                echo-debug "No mailbox for $($ADuser.samaccountname)"
                continue
            }

            # Flush counters
            $MBXcount++
            [int]$global:itemsCount1 = 0
            [int]$global:itemsCount2 = 0
            [int]$global:itemsCount3 = 0
            [int]$global:itemsCount4 = 0
            [int]$global:itemsCount5 = 0
            [int]$global:itemsCount6 = 0
            [int]$global:itemsCount7 = 0
            [int]$global:itemsCount8 = 0
            $itemsSumSize = 0

            $UserDN=$ADuser.distinguishedname
            $UserOU=Get-OURigthPart $UserDN
            $UserNameOUDescr=(Get-ADOrganizationalUnit -Identity $UserOU -Properties Description).Description -replace "\..*",""
            
            # If string hasn't been found in tmp file then get online statistics (check cache)
            $tmpFound=$(Get-tmpString $MailboxAddress)
            echo-debug "tmpFound: $tmpFound"

            if (!$tmpFound) {
                $global:UsermirrOU = Mirror-OU $UserOU
                # Get user mailbox main statistics           
                $global:MBXsize=(Get-MailboxStatistics $mailbox).totalitemsize.value.tomb()
                $global:MBXitems=(Get-MailboxStatistics $mailbox).Itemcount
                $global:MBXDBName=$mailbox.DataBase
                $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mailbox.primarySmtpAddress.ToString())
                try {
                    $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderId)
                } catch {
                    $MailboxErrorReport=";;$MailboxAddress;Access Denied"
                    echo-debug "№: $MBXcount `n$MailboxAddress - Access Denied"
                    Add-Content -Path $ExpFileCSVUsers -Value $MailboxErrorReport
                    continue
                }
                echo-debug "`n№: $MBXcount `n$($mailbox.Name) `nFolders:"
                # Get statistics from mbx folders
                doFolder($folder)
            }
            echo-debug "Items count: $global:MBXitems `nMbxSize (MB): $global:MBXsize"

            # Add statistics to report and save
            $ItemReportCSV="$UserNameOUDescr;$global:UsermirrOU;$MailboxAddress;$global:MBXsize;$global:MBXitems;$global:MBXDBName;$global:itemsCount1;$global:itemsCount2;$global:itemsCount3;$global:itemsCount4;$global:itemsCount5;$global:itemsCount6;$global:itemsCount7;$global:itemsCount8"
            try {
                Add-Content -Path $ExpFileCSVUsers -Value $ItemReportCSV
            } catch {
                echo-debug "-----Error adding content to file $ExpFileCSVUsers"  
            }
            
            # Add statistics to temp file
            try {
                if (!$tmpFound) {Add-Content -Path  $tmpSCVFile -Value $ItemReportCSV}
            } catch {
                echo-debug "-----Error adding content to file $tmpSCVFile"  
            }
            
            # Make hash for depts statistics (UserOU=MBXsizes)
            $UserOU.split(",") | where {$_ -notmatch "(OU=Spec|OU=Users|DC=test|DC=local)"} | foreach {
                echo-debug $_
                echo-debug "UsersubOU: $UserOU"
                if ($hashOUsizes.keys -contains $UserOU) {
                    $hashOUsizes[$UserOU]+=$global:MBXsize
                } else {
                    $hashOUsizes.Add($UserOU,$global:MBXsize)
                }
                $UserOU=Get-OURigthPart $UserOU
            }
            
        }
    }
    echo-debug "End getting users"
    # Read hash with sizes and add to report
    $hashOUsizes.keys | foreach {
        $strMBsize=$hashOUsizes.Item($_)
        $UserOUDescr=(Get-ADOrganizationalUnit -Identity $_ -Properties Description).Description -replace "\..*",""
        echo-debug $UserOUDescr
        $mirrOU = Mirror-OU $($_)
        echo-debug $mirrOU
        $repStringCSV = "$UserOUDescr;$mirrOU;$strMBsize"
        echo-debug $repStringCSV
        Add-Content -Path  $ExpFileCSVDepts -Value $repStringCSV
    }
    if ($script:SaveToSP) {
        echo-debug "Files: $FileNameUsers, $FileNameDepts`nSiteURL: $SiteURL`nLibrary: $Library`nsaving to SP..."
        if ($Library -notmatch "null") {
            Save-ToSP $script:svcAcc $script:passwdFile $ExpFileCSVDepts $FileNameDepts $SiteURL $Library
            Save-ToSP $script:svcAcc $script:passwdFile $ExpFileCSVUsers $FileNameUsers $SiteURL $Library
        }
    }
    
}

$script:StatParams.keys | foreach {
    echo-debug "StatParamKey: $($_)"
    echo-debug "StatParamValue: $($script:StatParams.Item($_))"
    Get-PersStatMX $_ $($script:StatParams.Item($_))
}

Remove-Item $tmpSCVFile

# Calculate script running time
$CurDateTime2=get-date
echo-debug "StartTime: $($CurDateTime1.DateTime)"
echo-debug "EndTime: $($CurDateTime2.DateTime)"
