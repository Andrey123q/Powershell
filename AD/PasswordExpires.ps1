## Description: script to notify about expiring passwords
## Inspired by similar *.ps1 scripts from github
## Date: 23.02.2016
## Author: Andrey.Nesterov
## ver 1.1

# Add required snapin and module
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# Variables
$ReportList = "`n"
$DaysToExpire = "5"  
$ZabbixServer = "mon.test.local"
$SourceServer = "ad"
$ItemName = "adpwdexp"
$exe = "C:\Program Files\Zabbix Agent\zabbix_sender.exe"
$SMTPServer = "mx.test.local"
$from = "support@test.local"
$subject = "Напоминание об истечении срока пароля"
$testUser = "TestUser"

$SendMail = $true
$Debug = $false
$TestMode = $false
$SendToZabbix = $true

### Functions
# Echo Debug
function Echo-Debug ($EchoString) {
    if ($Debug) {echo $EchoString}
}

# Get List of Users with emails and expires pwds lt $DaysToExpire
function Get-ListOfUsers () {
    return Get-ADUser -SearchBase (Get-ADRootDSE).defaultNamingContext `
    -Filter {(Enabled -eq "True") -and (PasswordNeverExpires -eq "False") -and (mail -like "*") -and (SamAccountName -like $scope)} -Properties * | `
    Select-Object Name,SamAccountName,DisplayName,mail,@{Name="Expires";Expression={ $MaxPassAge - ((Get-Date) - ($_.PasswordLastSet)).days}} | `
    Where-Object {$_.Expires -gt 0 -AND $_.Expires -le $DaysToExpire}
}

# Send mail to user
function send-mail ($user){
    $to = $user.mail
    $emailBody = "Уважаемый (ая), $($user.DisplayName).
До завершения срока действия Вашего пароля к аккаунту $($user.SamAccountName) осталось $($user.Expires) дн.
Просим своевременно сменить пароль, чтобы обезопасить себя от проблем с доступом к корпоративным ресурсам.
При возникновении вопросов обращайтесь в службу ИТ-поддержки.


------------------------------------------
С  уважением, 
ServiceDesk
IP тел.:911
mailto: support@test.local
"
    $mailer = new-object Net.Mail.SMTPclient($SMTPserver)
    $msg = new-object Net.Mail.MailMessage($from, $to, $subject, $emailBody)
    if ($SendMail) {
        $mailer.send($msg) 
    }
}


### Main code
$MaxPassAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.days
if($MaxPassAge -le 0) {  
   throw "Domain 'MaximumPasswordAge' password policy is not configured." 
}

if ($TestMode) {
    $scope = $testUser
} else {
    $scope = "*"
}
echo-debug "Scope: $scope"

$ListOfUsers = Get-ListOfUsers
ForEach ($user in $ListOfUsers) {
    $ReportList = $ReportList + $user.mail + "`t - $($user.Expires) days" + "`n"
    echo-debug $user
    send-mail $user
}

echo-debug $ReportList
Write-Output 1

# Send list of users to zabbix server
if ($SendToZabbix) {
    &$exe -c "C:\Program Files\Zabbix Agent\zabbix_agentd.conf" -z $ZabbixServer -s $SourceServer -k $ItemName -o $ReportList | Out-Null
}


===
