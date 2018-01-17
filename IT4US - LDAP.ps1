$strFilter = "(& (objectClass=organizationalprofile)(!objectClass=computer)(company=SITA*)(createTimeStamp>=20171124000000.0Z) )"

#$objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://fr.green.local")
$objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://green.local")

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.PageSize = 1000
$objSearcher.Filter = $strFilter
$objSearcher.SearchScope = "Subtree"

$colResults = $objSearcher.FindAll()
$profile = @{}

foreach ($objResult in $colResults) {
    $profile.sAMAccountName = $objResult.Properties.samaccountname
    $profile.givenName = $objResult.Properties.givenname
    $profile.company = $objResult.Properties.company
    $profile.sn = $objResult.Properties.sn
    $profile.created = $objResult.Properties["createTimeStamp"]
    $profile.mail = $objResult.Properties.mail
    $profile.mobile = $objResult.Properties["mobile"]
    $profile.cn = $objResult.Properties["cn"]
    $profile.telephoneNumber = $objResult.Properties["telephoneNumber"]
    $profile.start_date = $objResult.Properties["start_date"]
    $profile.end_date = $objResult.Properties["end_date"]
    #$costCenter=$objResult.Properties["costCenter"]
    $profile.location = $objResult.Properties["location"]
    $profile.title = $objResult.Properties["title"]
    $profile.Disabled = $objResult.Properties["Disabled"]

    $profile.employeetype = $objResult.Properties.employeetype
    <#
    if ($employeetype -like 'I') {$employeetype="Internal"} 
    elseif ($employeetype -like 'E') {$employeetype="External"}
    else {$employeetype}
    #>
    $profile.enabled = $objResult.Properties["enabled"]

    $profile.employeeStatus = $objResult.Properties.employeetype

    # Extrat "Green ID" from the CN field
    $s = $objResult.Properties.manager
    $profile.manager = [regex]::match($s,'\(([^\)]+)\)').Groups[1].Value

    #Write-Host ";$sAMAccountName;$givenName;$sn;$employeetype;$manager;$mail;$mobile;$telephoneNumber;$costCenter;$title;$start_date;$end_date;1;;$company;;0"

}