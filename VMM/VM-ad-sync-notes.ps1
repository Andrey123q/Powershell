# Description: Script to sync Description of PC in AD to VirtualCenter
# Author: Andrey Nesterov
# Version: 1.1
# Date: 24.04.2015
# Modified: 19.11.2015

add-pssnapin VMware.VimAutomation.Core -ea silentlycontinue
import-module activedirectory -ea silentlycontinue

# Variables
$Debug = $true
$viserver = vc.test.local"
$testserver = "test-server"

### Fx1. Echo-Debug
#####
function Echo-Debug ($EchoString) {
    if ($Debug) {echo $EchoString}
}

if ($Debug -eq $true) {
    $scope = $testserver
} else {$scope="*"}

connect-viserver $viserver -wa 0
# Get-VM -Name * - in work mode
Get-VM -Name $scope | foreach-object {
    $notes = $null
    $owner = $null
    $cmp = Get-ADComputer -f {cn -eq $_.Name} -Properties description ; 
    $descr=$cmp.description;
    try {
        $pos=$descr.IndexOf("Владелец:")
        $notes=$descr.substring(0,$pos)
        $owner=$descr.substring($pos+10)
    } catch {
        $notes=$descr
        $owner=$null
    }
    $vmowner = (Get-VM $_ | Get-Annotation -CustomAttribute "Owner").Value
    $vmnotes = (Get-VM $_).Notes
    echo-debug "1" $vmowner 
    echo-debug "2" $owner
    if ($owner -ne $vmowner) {
        Set-CustomField $_ -Name "Owner" -Value $owner #-WhatIf # whatif in debug mode    
        echo-debug "not eq owner"
    }
    if ($notes -ne $vmnotes) {
        Set-VM $_ -Notes $notes -Confirm:$false #-WhatIf # whatif in debug mode
        echo-debug "not eq notes"
    }
    <#
    if ($Debug -eq $true) {
        Set-CustomField $_ -Name "Owner" -Value $owner -WhatIf # whatif in debug mode
        Set-VM $_ -Notes $notes -Confirm:$false -WhatIf # whatif in debug mode
    } else {
        Set-CustomField $_ -Name "Owner" -Value $owner
        Set-VM $_ -Notes $notes -Confirm:$false
    }
    #>
}

# Send check to Zabbix. Must be defiened item (template) and macro {$SCRIPT1} or {$SCRIPTN}
&"C:\Program Files\Zabbix Agent\zabbix_sender.exe" -c "C:\Program Files\Zabbix Agent\zabbix_agentd.conf" -k "script.status" -o 1 | Out-Null
