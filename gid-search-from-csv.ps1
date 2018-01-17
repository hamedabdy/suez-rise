$users = Import-Csv "$env:HOMEDRIVE\$env:HOMEPATH\Desktop\u_export_hab_no_ref_users.csv"

$todayDate = Get-Date (Get-Date) -UFormat "%d/%m/%Y %H:%M:%S"
Write-Host "Date  : " $todayDate

##############
$objDomainArray = New-Object System.Collections.ArrayList
$objDomainArray.Add( $( New-Object System.DirectoryServices.DirectoryEntry("LDAP://fr.green.local") ) ) | Out-Null
$objDomainArray.Add( $( New-Object System.DirectoryServices.DirectoryEntry("LDAP://green.local") ) ) | Out-Null

$t = "BU|GreenID|FirstName|LastName|manager_greenID|email|mobile|phone|title|company|location|OU|employeeType|whenCreated|whenChanged|different"
$t | Out-File "$env:HOMEDRIVE\$env:HOMEPATH\Desktop\updated-users.csv"

foreach ( $u in $users ) {
    #Write-Host $($u | Out-String)
    if ( $u.u_user_name ) {
        $oe = $objDomainArray.GetEnumerator()
        while ( $oe.MoveNext() ) {
            $ue = $users.GetEnumerator()
            $strFilter = "(&((objectcategory=person) {0} ) )" -f "(samaccountname=$($u.u_user_name))"
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
            $objSearcher.SearchRoot = $oe.Current
            $objSearcher.Filter = $strFilter
            $objSearcher.PageSize = 1000
            $objSearcher.SearchScope = "Subtree"
            $colResults = $objSearcher.FindOne()
            Write-Host "Looking for $($u.u_user_name) in [$($oe.Current.path)]..." -ForegroundColor Yellow
            $p = $colResults.Properties
            $profile = @{}
            if ( $colResults.Count -eq 0  ) {
                Write-Host "Found nothing" -ForegroundColor DarkYellow
                continue
            }
            $profile.GreenID = $p.samaccountname
            $profile.FirstName = $p.givenname
            $profile.LastName = $p.sn
            # Here we need to extract the green id from the CN field
            $profile.manager_greenID = [regex]::match($p.manager,'\(([^\)]+)\)').Groups[1].Value
            $profile.email = $p.mail
            $profile.mobile = (( $p.mobile | Out-String ) -replace '(?:\r|\n)','' )
            $profile.phone = (( $p.telephonenumber | Out-String ) -replace '(?:\r|\n)','' )
            $profile.title = (( $p.title | Out-String ) -replace '(?:\r|\n)','' )
            $profile.location = (( $p.l | Out-String ) -replace '(?:\r|\n)','' )
            $tmp = [regex]::Split( $p.company, ' - ' )
            $profile.company = @{$true=$tmp[0]; $false=$tmp[1]}[ -not $tmp[1] ]
            $profile.BU = $tmp[0]
            $profile.OU = ""
            if( $objResult.extensionattribute1 ) {
                $extensionAtt1 = $objResult.extensionattribute1.split(",")
                $profile.OU = ( ( $extensionAtt1 -match "^OU=" ).Substring(3) -join ", " )
            }
            $profile.whenCreated = ( $p.whencreated )[0].ToString()
            $profile.whenchanged = ( $p.whenchanged )[0].ToString()
            $et = $p.employeetype
            if( $et -imatch "^i" ) { $profile.employeeType = "Internal" }
            if( $et -imatch "^e" ) { $profile.employeeType = "External" }
            $t = "$($profile.BU)|$($profile.GreenID)|$($profile.FirstName)|$($profile.LastName)|$($profile.manager_greenID)|$($profile.email)|$($profile.mobile)|$($profile.phone)|$($profile.title)|$($profile.company)|$($profile.location)|$($profile.OU)|$($profile.employeeType)|$($profile.whenCreated)|$($profile.whenchanged)"
            $t | Out-File -Append "$env:HOMEDRIVE\$env:HOMEPATH\Desktop\updated-users.csv" -Encoding unicode
            }
            
    }
}
Write-Host "Data written into : $env:HOMEDRIVE\$env:HOMEPATH\Desktop\updated-users.csv" -ForegroundColor Green