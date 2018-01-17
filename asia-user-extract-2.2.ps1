# Date Comparison : $date1 is compared to $date2 :
# $true if $date1 > $date2; $false otherwise
function compareDate ( $date1, $date2 ) {
    #Write-Host "dates = " $date1 ",   " $date2
    if( ( $date1.GetType() -eq $date2.GetType() ) -and ( $date1.GetType() -eq "".GetType() ) )
    {
        if ( $date1 -gt $date2 ) { return $true }
        else { return $false }
    }
    else
    {
        throw "[compareDate] Please verify that date types are correct and are String"
    }
}

##############
#$objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://fr.green.local")
$objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://green.local")

$strFilter = "(&(objectClass=organizationalPerson)(!objectClass=computer) )"
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.PageSize = 1000
$objSearcher.Filter = $strFilter
$objSearcher.SearchScope = "Subtree"
$colResults = $objSearcher.FindAll()

$profile = @{BU=""; GreenID=""; FirstName=""; LastName=""; manager_greenID=""; email=""; mobile=""; phone=""; title=""; company=""; employeeType=""; whenCreated=""; }
$todayDate = Get-Date (Get-Date).AddDays(-13) -UFormat "%d/%m/%Y %H:%M:%S"
Write-Host "Date  : " $todayDate
# Write to file and replace file if already exists
#$profile.Keys -join ";" | Out-File "C:\Users\VC5622\Desktop\export.csv"
$headerFlag = $false
write-host "Total number of users to process : " $colResults.Count
foreach ($objResult in $colResults) {
    if ( $objResult.Properties.extensionattribute1 ) {
        $profile.GreenID = $objResult.Properties.samaccountname
        $profile.FirstName = $objResult.Properties.givenname
        $profile.LastName = $objResult.Properties.sn
        # Here we need to extrat the green id from the CN field
        $profile.manager_greenID = [regex]::match($objResult.Properties.manager,'\(([^\)]+)\)').Groups[1].Value
        $profile.email = $objResult.Properties.mail
        $profile.mobile = $objResult.Properties.mobile
        $profile.phone = $objResult.Properties.telephonenumber
        $profile.title = $objResult.Properties.title
        $profile.company = $objResult.Properties.company
        $profile.BU = [regex]::Split( $objResult.Properties.company, ' - ' )[0]
        $profile.whenCreated = ($objResult.Properties.whencreated)[0].ToString()
        $et = $objResult.Properties.employeetype
        if( $et -imatch "^i" ) { $profile.employeeType = "Internal" }
        if( $et -imatch "^e" ) { $profile.employeeType = "External" }
        $extensionAtt1 = $objResult.Properties.extensionattribute1.split(",")
        $tmp = ( $extensionAtt1 -match "^OU=" ).Substring(3)
        $m = @( "sita.hk", "sinofrench", "dgt.asia" )
        #if( $d = compareDate -date1 $profile.whenCreated -date2 $todayDate ) { 
        if( $true ) { 
            if( -not $headerFlag ) {
                $t = "BU;GreenID;FirstName;LastName;manager_greenID;email;mobile;phone;title;company;employeeType;whenCreated"
                $t | Out-File "C:\Users\VC5622\Desktop\export.csv"
                $headerFlag = $true
            }
            foreach ( $s in $m ) {
                if( $tmp -icontains $s ) {
                    $profile.ou = $tmp
                    #$t = ( $profile.GetEnumerator() | % { "$($_.Value)" } ) -join ";"
                    $t = "$($profile.BU);$($profile.GreenID);$($profile.FirstName);$($profile.LastName);$($profile.manager_greenID);$($profile.email);$($profile.mobile);$($profile.phone);$($profile.title);$($profile.company);$($profile.employeeType);$($profile.whenCreated)"
                    $t | Out-File -Append "C:\Users\VC5622\Desktop\export.csv"
                }
            }
        }
    }
    
}