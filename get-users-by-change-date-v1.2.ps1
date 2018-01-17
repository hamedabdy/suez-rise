$exportPath = ".\Desktop\"
$fileName = "user-update-export"
$fileFormat = ".csv"
$todayDate = Get-Date (Get-Date).AddDays(0) -UFormat "%d-%m-%Y"
Write-Host "Date  : " $todayDate

#######################
function getUsersByDate ( $insertDate = $todayDate, $updateDate = $todayDate ) {
    $insertDate = Get-Date ($insertDate) -UFormat "%Y%m%d%H%M%S.sZ"
    $updateDate = Get-Date ($updateDate) -UFormat "%Y%m%d%H%M%S.sZ"
    # Write CSV Header to file
    $file = "$exportPath$fileName-$todayDate$fileFormat"
    $t = "BU|GreenID|FirstName|LastName|anager_greenID|email|mobile|phone|title|company|location|employeeType|whenCreated|whenChanged"
    $t | Out-File $file
    $objDomainArray = New-Object System.Collections.ArrayList
    $objDomainArray.Add( $( New-Object System.DirectoryServices.DirectoryEntry("LDAP://fr.green.local") ) ) | Out-Null
    $objDomainArray.Add( $( New-Object System.DirectoryServices.DirectoryEntry("LDAP://green.local") ) ) | Out-Null
    $oe = $objDomainArray.GetEnumerator()
    while ( $oe.MoveNext() ) {
        $upDate = "(whenchanged>=$updateDate)"
        $inDate = "(whencreated>=$insertDate)"
        $strFilter = "(&((objectcategory=person)(sn=*)(company=*)(samaccountname=*) ) (| $inDate $upDate ) )"
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
        $objSearcher.SearchRoot = $oe.Current
        $objSearcher.Filter = $strFilter
        $objSearcher.PageSize = 1000
        $objSearcher.SearchScope = "Subtree"
        $colResults = $objSearcher.FindAll()
        if ( -not $colResults ) { continue }
        $profile = @{}
        write-host "`nTotal number of users to process [in $($oe.Current.path)] : " $colResults.Count
        foreach ($objResult in $colResults) {
            $profile.GreenID = $objResult.Properties.samaccountname
            $profile.FirstName = $objResult.Properties.givenname
            $profile.LastName = $objResult.Properties.sn
            # Here we need to extrat the green id from the CN field
            $profile.manager_greenID = [regex]::match($objResult.Properties.manager,'\(([^\)]+)\)').Groups[1].Value
            $profile.email = $objResult.Properties.mail
            $profile.mobile = (( $objResult.Properties.mobile | Out-String ) -replace '(?:\r|\n)','' )
            $profile.phone = (( $objResult.Properties.telephonenumber | Out-String ) -replace '(?:\r|\n)','' )
            $profile.title = (( $objResult.Properties.title | Out-String ) -replace '(?:\r|\n)','' )
            $profile.location = (( $objResult.Properties.l | Out-String ) -replace '(?:\r|\n)','' )
            $tmp = [regex]::Split( $objResult.Properties.company, ' - ' )
            $profile.company = @{$true=$tmp[0]; $false=$tmp[1]}[ -not $tmp[1] ]
            $profile.BU = $tmp[0]
            $profile.whenCreated = ($objResult.Properties.whencreated)[0].ToString()
            $profile.whenchanged = ($objResult.Properties.whenchanged)[0].ToString()
            $et = $objResult.Properties.employeetype
            if( $et -imatch "^i" ) { $profile.employeeType = "Internal" }
            if( $et -imatch "^e" ) { $profile.employeeType = "External" }
            $t = "$($profile.BU)|$($profile.GreenID)|$($profile.FirstName)|$($profile.LastName)|$($profile.manager_greenID)|$($profile.email)|$($profile.mobile)|$($profile.phone)|$($profile.title)|$($profile.company)|$($profile.location)|$($profile.employeeType)|$($profile.whenCreated)|$($profile.whenchanged)"
            $t | Out-File -Append $file -Encoding unicode
        }
    }
    Write-Host "`nExport CSV file was written into : " $file -ForegroundColor Green
}

getUsersByDate