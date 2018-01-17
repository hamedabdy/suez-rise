# Logging
$verbose = $false
$debug = $false
# Global Date
$global:date = Get-Date ( Get-Date ) -UFormat "%d-%m-%Y %H.%M.%S"
# File path
$rootPath = "$($env:USERPROFILE)\Desktop\"
$fileName = "-from-rise"
$fileFormat = ".csv"
$global:file = ""
# Write to file flag
$global:writeToFile = $true
$global:writeHeader = $true
# File to import Users IDs from
$fromFile = "ids"
# RISE Filter
$global:strFilter = "(&((objectcategory=person) {0} ) (| (samaccountname=*){1} ) )"

$objDomainArray = New-Object System.Collections.ArrayList
$objDomainArray.Add( $( New-Object System.DirectoryServices.DirectoryEntry("LDAP://fr.green.local") ) ) | Out-Null
$objDomainArray.Add( $( New-Object System.DirectoryServices.DirectoryEntry("LDAP://green.local") ) ) | Out-Null

function writeToCsv ( $results, $myCommand ) {
    Write-Verbose $( $results.properties | Out-String ) -Verbose:$verbose
    if ( $global:writeToFile ) {
        if ( $global:writeHeader ) {
            $global:file = "$rootPath$myCommand$fileName-$global:date$fileFormat"
            $t = "BU|GreenID|FirstName|LastName|manager_greenID|email|mobile|phone|title|company|location|OU|employeeType|whenCreated|whenChanged"
            $t | Out-File $file
            $global:writeHeader = $false
        }
        $profile = @{}
        foreach ( $p in $results ) {
            Write-Verbose "Processing $($p.Properties.samaccountname)..." -Verbose:$verbose
            $profile.GreenID = $p.Properties.samaccountname
            $profile.FirstName = $p.Properties.givenname
            $profile.LastName = $p.Properties.sn
            # Here we need to extract the green id from the CN field
            $profile.manager_greenID = [regex]::match($p.Properties.manager,'\(([^\)]+)\)').Groups[1].Value
            $profile.email = $p.Properties.mail
            $profile.mobile = (( $p.Properties.mobile | Out-String ) -replace '(?:\r|\n)','' )
            $profile.phone = (( $p.Properties.telephonenumber | Out-String ) -replace '(?:\r|\n)','' )
            $profile.title = (( $p.Properties.title | Out-String ) -replace '(?:\r|\n)','' )
            $profile.location = (( $p.Properties.l | Out-String ) -replace '(?:\r|\n)','' )
            $tmp = [regex]::Split( $p.Properties.company, ' - ' )
            $profile.company = @{$true=$tmp[0]; $false=$tmp[1]}[ -not $tmp[1] ]
            $profile.BU = $tmp[0]
            $profile.OU = ""
            if( $p.properties.extensionattribute1 ) {
                $extensionAtt1 = $p.properties.extensionattribute1.split(",")
                $profile.OU = ( ( $extensionAtt1 -match "^OU=" ).Substring(3) -join ", " )
            }
            $profile.whenCreated = ( $p.Properties.whencreated )[0].ToString()
            $profile.whenchanged = ( $p.Properties.whenchanged )[0].ToString()
            $et = $p.Properties.employeetype
            if( $et -imatch "^i" ) { $profile.employeeType = "Internal" }
            if( $et -imatch "^e" ) { $profile.employeeType = "External" }
            $t = "$($profile.BU)|$($profile.GreenID)|$($profile.FirstName)|$($profile.LastName)|$($profile.manager_greenID)|$($profile.email)|$($profile.mobile)|$($profile.phone)|$($profile.title)|$($profile.company)|$($profile.location)|$($profile.OU)|$($profile.employeeType)|$($profile.whenCreated)|$($profile.whenchanged)"
            $t | Out-File -Append $global:file -Encoding unicode
        }
        Write-Host "Data written into : " $file -ForegroundColor Green
    }
}

function getUserProfile ( $filter ) {
    Write-Host $("[INFO] Search query :`n{0}`n" -f $filter) -ForegroundColor Cyan
    $oe = $objDomainArray.GetEnumerator()
    while ( $oe.MoveNext() ) {
        Write-Host "`nSearching in [$($oe.Current.path)]"
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
        $objSearcher.SearchRoot = $oe.Current
        $objSearcher.Filter = $filter
        $objSearcher.PageSize = 1000
        $objSearcher.SearchScope = "Subtree"
        $colResults = $objSearcher.FindAll()
        if ( $colResults.Count -eq 0 ) { 
            continue
        }
        Write-Host "Total number of users to process : " $colResults.Count
        writeToCsv -results $colResults -myCommand $MyInvocation.MyCommand
       
    }
}


function getUserByName ( $givenName, $sn ) {
    $nameFilter = "(givenname=$givenName)"
    $snFilter = "(sn=$sn)"
    if ( -not ( $givenName -and $sn ) ) {
        return @("{0}", "{1}")  
    }
    return @("$nameFilter$snFilter{0}", "{1}")
}

function getUsersByDate ( $insertDate, $updateDate, $operator ) {
    $insertDate = Get-Date ($insertDate) -UFormat "%Y%m%d%H%M%S.sZ"
    $inDate = "(whencreated$operator$insertDate)"
    if ( -not $updateDate ) {
        return @("(sn=*)(company=*)(samaccountname=*)$inDate{0}", "(| $inDate{1} )")
    }
    $updateDate = Get-Date ($updateDate) -UFormat "%Y%m%d%H%M%S.sZ"
    $upDate = "(whenchanged$operator$updateDate)"
    return @("(sn=*)(company=*)(samaccountname=*)$upDate{0}", "(| $inDate{1} )")
    
}

function getUserByGreenId ( $id ) {
    if( -not $id ) {
        return @("{0}", "{1}")
    }
    $nameFilter = "(samaccountname=$id)"
    return @("$nameFilter{0}", "{1}")
}

function getUsersByIds ( ) {
    $idsCsv = Import-Csv "$rootPath$fromFile$fileFormat"
    foreach ( $o in $idsCsv.id ) {
        $idFilterArr = getUserByGreenId -id $o
        $filter = $global:strFilter -f $idFilterArr[0], $idFilterArr[1]
        getUserProfile $($filter -f "", "")
    }
}

function getUsersByCompany ( $company ) {
    $companyFilter = "(company=$company)"
    if(  -not $company ) {
        return @("{0}", "{1}")
    }
    return @("$companyFilter{0}", "{1}")
}

function getActiveUsers ( $active ) {
    $activeUserFilter = "(useraccountcontrol=512)"
    if ( -not $active ) {
        return @("{0}", "{1}")
    }
    return @("$activeUserFilter{0}", "{1}")
}

function addCustomFilterField ( $and, $or ) {
    return @("$and{0}", "$or{1}")
}

function getUsers ( $p ) {
    $activeFilterArr = getActiveUsers $p.userActive
    $global:strFilter = $global:strFilter -f $activeFilterArr[0], $activeFilterArr[1]
    $idFilterArr = getUserByGreenId $p.id
    $global:strFilter = $global:strFilter -f $idFilterArr[0], $idFilterArr[1]
    $nameFilterArr = getUserByName $p.firstname $p.lastname
    $global:strFilter = $global:strFilter -f $nameFilterArr[0], $nameFilterArr[1]
    $companyFilterArr = getUsersByCompany $p.company
    $global:strFilter = $global:strFilter -f $companyFilterArr[0], $companyFilterArr[1]
    if( $p.filterByDate ) {
        $dateFilterArr = getUsersByDate $p.createDate $p.changeDate $p.dateOperator
        $global:strFilter = $global:strFilter -f $dateFilterArr[0], $dateFilterArr[1]
    }
    $customFilterArr = addCustomFilterField $p.and $p.or
    $global:strFilter = $global:strFilter -f $customFilterArr[0], $customFilterArr[1]
    getUserProfile $($global:strFilter -f "", "")
}

# By list of IDs
getUsersByIds

# For active users add "userActive=$true" inside your disctionary @{}
# For additional filter use the follwoing template:
<#
    
    @{and="(useraccountcontrol=514)"; or=""}
    
    Note 1 : These filters could be used with any other desired filter!
    Note 2 : the terms "and", "or" refer to the creation of an "and" or an "or" filter in AD
#>

# By name or ID
#getUsers @{id="IFD428";}
#getUsers @{firstname="nicolas"; lastname="bequaert"}
#getUsers @{firstname="hamed"; lastname="abdy"; id="VC5622"}

#By Company (* = wildcard)
#getUsers @{company="*NEWS*"; userActive=$true}

# By Date and Company

<#
$createDate = Get-Date (Get-Date).AddDays(-1)
$changeDate = Get-Date (Get-Date).AddDays(-1)
Write-Host "CreateDate  : " $( Get-Date ( $createDate ) -UFormat "%d-%m-%Y" )
Write-Host "CreateDate  : " $( Get-Date ( $changeDate ) -UFormat "%d-%m-%Y" )
getUsers @{company="DGT"; filterByDate=$true; createDate=$createDate; changeDate=$changeDate; dateOperator=">="}
#getUsers @{filterByDate=$true; createDate=$createDate; dateOperator=">="}
#>