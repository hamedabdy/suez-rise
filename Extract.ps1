$csv = Import-Csv "C:\Users\FT5582\Desktop\DATA.csv" -Delimiter ";"

Write-Host c

[System.Collections.ArrayList]$array = $csv."user id"

$file = "C:\Users\FT5582\Desktop\LDAP_text.csv"
Remove-Item $file

$pre="BU;greenID;firstname;lastname;contract_type;manager_greenid;email;mobile;phone;user_cc;title;start_date;end_date;status;status_desc;company;site;vip"
$exists = Test-Path $file
if ($exists -ne "True") { New-Item -ItemType file -Path $file -Force }
Add-Content -Path $file -Value $pre

foreach($e in $array ) {

[String]$strFilter = "(&(objectClass=organizationalPerson)(sAMAccountName="+${e}+")(cn=*)(company=*)(givenName=*))"

#$objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://fr.green.local")
$objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://green.local")

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.PageSize = 1000
$objSearcher.Filter = $strFilter
$objSearcher.SearchScope = "Subtree"

$colProplist = "sAMAccountName","givenName","sn","company","createTimeStamp","mobile","Title","manager","cn","telephoneNumber","mail"
#foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}

$colResults = $objSearcher.FindAll()

<#This is the export part:#>
<#Check if the file exists, if not create it#>
##$file = "C:\Users\FT5582\Desktop\LDAP_text.csv"
##Remove-Item $file

foreach ($objResult in $colResults) {
    $sAMAccountName=$objResult.Properties["sAMAccountName"]
    $givenName=$objResult.Properties["givenName"]
    $company=$objResult.Properties["company"]
    $sn=$objResult.Properties["sn"]
    $created=$objResult.Properties["createTimeStamp"]
    $mail=$objResult.Properties["mail"]
    $mobile=$objResult.Properties["mobile"]
    $cn=$objResult.Properties["cn"]
    $telephoneNumber=$objResult.Properties["telephoneNumber"]
    $start_date=$objResult.Properties["start_date"]
    $end_date=$objResult.Properties["end_date"]
    $title=$objResult.Properties["title"]

    <#Here we need to extrat the green id from the CN field#>
    $s=$objResult.Properties["manager"]
    $manager = [regex]::match($s,'\(([^\)]+)\)').Groups[1].Value

    <#Here we need to check the last char ") EXT",#>
    if ($cn -like '*) EXT*') {$contract_type="External"} else {$contract_type="Internal"}

    Write-Output ";$sAMAccountName;$givenName;$sn;$contract_type;$manager;$mail;$mobile;$telephoneNumber;;$title;$start_date;$end_date;1;;$company;;0"

    $exists = Test-Path $file
    if ($exists -ne "True") { New-Item -ItemType file -Path $file -Force }
    $inputfile = ";$sAMAccountName;$givenName;$sn;$contract_type;$manager;$mail;$mobile;$telephoneNumber;;$title;$start_date;$end_date;1;;$company;;0"

    <#This is the export part:#>
    Add-Content -Path $file -Value $inputfile
}

}