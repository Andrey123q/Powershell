
## Config file for script FS-pers-fld-sizes.ps1
## Contains: 
## subhash: OUs = User's root folder in PersonalData 
## values of main hash: options to save in SP library: "url to site; Libary; file index in report file name"
## If Library is null report is not saved to SP
## Example:
## $StatParams = @{
##    @{
##    "OU=Users,DC=test,DC=local" = "Users";
##    "OU=Dept1,OU=Users,OU=Special,DC=test,DC=local" = "Top";
##    "OU=Dept2,OU=Users,DC=test,DC=local" = "Admins"
##    } = "https://sharepoint.test.local/dept/lib;Reports/FS;ALL"
##

$StatParams = @{
    @{
    "OU=Dept,OU=Users,DC=test,DC=local" = "Users"
    } = "https://sharepoint.test.local/dept/lib;null;Dept";
}
