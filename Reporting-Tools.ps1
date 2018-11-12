Function Get-ADUseraccess ($logon) {
    $users = $logon -split ','
    $xls = New-Object -ComObject excel.application
    $xls.Workbooks.Add() | Out-Null
    $xls.Visible = $true
    $sheet = $xls.Worksheets.Item(1)
    $sheet.Name = 'Get-Groupmemberships'
    $column = 1
    foreach ($user in $users) {
        $row = 1
        $usernm = Get-AdUser -Filter 'samAccountName -like $user' | Select-Object Name
        if ($usernm) {
            $useraccess = Get-AdPrincipalGroupMembership $user | Select-Object Name
            $sheet.Cells.Item($row, $column) = $usernm.Name
            $sheet.Cells.Item($row, $column).Font.Bold = $True
            $row++
            foreach ($group in $useraccess.Name) {
                $sheet.Cells.Item($row, $column) = $group
                $row++
            }
        }
        else {$sheet.Cells.Item($row, $column) = "Could not find $user"; Write-Host "Could not find $user"}
        $column++
    }
    $usedRange = $sheet.UsedRange						
    $usedRange.EntireColumn.AutoFit() | Out-Null
}

Function Get-ADGroupMemberof ($inputgroup) {
    $groups = $inputgroup -split ','
    $xls = New-Object -ComObject excel.application
    $xls.Workbooks.Add() | Out-Null
    $xls.Visible = $true
    $sheet = $xls.Worksheets.Item(1)
    $sheet.Name = 'Get-Groupmemberships'

    $column = 1
    foreach ($group in $groups) {
        $row = 1
        $groupcn = Get-AdGroup $group | Select-Object Name
        if ($groupcn) {
            $groupmemberof = Get-AdPrincipalGroupMembership $group | Select-Object Name
            $sheet.Cells.Item($row, $column) = $groupcn.Name
            $sheet.Cells.Item($row, $column).Font.Bold = $True
            $row++
            foreach ($membergroup in $groupmemberof.Name) {
                $sheet.Cells.Item($row, $column) = $membergroup
                $row++
            }
        }
        else {$sheet.Cells.Item($row, $column) = "Could not find $group"; Write-Host "Could not find $group"}
        $column++
    }
    $usedRange = $sheet.UsedRange						
    $usedRange.EntireColumn.AutoFit() | Out-Null
    #$xls.Visible = $true
}


Function Get-RecursiveGroupmembers ($group) {
    $groups = $group -split ','

    $xls = New-Object -ComObject excel.application
    $xls.Workbooks.Add() | Out-Null
    $xls.Visible = $true
    $sheet = $xls.Worksheets.Item(1)
    $sheet.Name = 'Get-RecursiveGroupmembers'
    $column = 1
    foreach ($grp in $groups) {
        $sheet.Cells.Item(1,$column) = 'Group'
        $sheet.Cells.Item(1, $column).Font.Bold = $True
        $sheet.Cells.Item(1,$column+1) = 'Member'
        $sheet.Cells.Item(1, $column+1).Font.Bold = $True
        $sheet.Cells.Item(1,$column+2) = 'Title'
        $sheet.Cells.Item(1, $column+2).Font.Bold = $True
        $row = 2
        $adgroup = Get-ADGroup $grp | Select-Object Name
        if ($adgroup) {
            $sheet.Cells.Item($row, $column) = $adgroup.Name
            #$sheet.Cells.Item($row,$column).Font.ColorIndex = 1 # color index
            $row++
            $groupmembers = Get-ADGroupMember $adgroup.Name
            if ($groupmembers | Where-Object objectClass -eq 'user') {
                foreach ($user in ($groupmembers | Where-Object objectClass -eq 'user')) {
                    $sheet.Cells.Item($row, $column+1) = $user.Name
                    $title = Get-Aduser $user -Properties Title | Select-Object Title
                    if ($title) {
                        $sheet.Cells.Item($row, $column+2) = $title.Title
                    }
                    else {
                        $sheet.Cells.Item($row, $column+2) = "No Job title!"
                        $sheet.Cells.Item($row,$column+2).Interior.ColorIndex = 3
                        Write-Host "Could not find $user title!"
                    }
                    $row++
                }
            }
            if ($groupmembers | Where-Object objectClass -eq 'group') {
                foreach ($nestedadgroup in ($groupmembers | Where-Object objectClass -eq 'group')) {
                    $sheet.Cells.Item($row, $column) = $nestedadgroup.Name
                    $nestedadgroupmembers = Get-ADGroupMember $nestedadgroup.Name
                    if ($nestedadgroupmembers | Where-Object objectClass -eq 'user') {
                        foreach ($nesteduser in ($nestedadgroupmembers | Where-Object objectClass -eq 'user')) {
                            $sheet.Cells.Item($row, $column+1) = $nesteduser.Name
                            $nestedtitle = Get-Aduser $nesteduser -Properties Title | Select-Object Title
                            if ($nestedtitle) {
                                $sheet.Cells.Item($row, $column+2) = $nestedtitle.Title
                            }
                            else {
                                $sheet.Cells.Item($row, $column+2) = "No Job title!"
                                $sheet.Cells.Item($row,$column+2).Interior.ColorIndex = 3
                                Write-Host "Could not find $user title!"
                            }
                            $row++
                        }
                    }
                    #AD Groups in Nested AD Groups
                    if ($nestedadgroupmembers | Where-Object objectClass -eq 'group') {
                        $doublenested = $true
                        $sheet.Cells.Item(1,$column+1) = 'Member (Group+1)'
                        $sheet.Cells.Item(1,$column+2) = 'Title (Member+1)'
                        $sheet.Cells.Item(1,$column+3) = 'Title+1'
                        $sheet.Cells.Item(1,$column+3).Font.Bold = $True
                        $column += 1
                        $doublenestedgroupmembers = Get-ADGroupMember $nestedadgroup.name
                        foreach ($nestedadgroup in ($doublenestedgroupmembers | Where-Object objectClass -eq 'group')) {
                            $sheet.Cells.Item($row, $column) = $nestedadgroup.Name
                            $tripplenestedadgroupmembers = Get-ADGroupMember $nestedadgroup.Name
                            if ($tripplenestedadgroupmembers | Where-Object objectClass -eq 'user') {
                                foreach ($nesteduser in ($tripplenestedadgroupmembers | Where-Object objectClass -eq 'user')) {
                                    $sheet.Cells.Item($row, $column+1) = $nesteduser.Name
                                    $nestedtitle = Get-Aduser $nesteduser -Properties Title | Select-Object Title
                                    if ($nestedtitle) {
                                        $sheet.Cells.Item($row, $column+2) = $nestedtitle.Title
                                    }
                                    else {
                                        $sheet.Cells.Item($row, $column+2) = "No Job title!"
                                        $sheet.Cells.Item($row,$column+2).Interior.ColorIndex = 3
                                        Write-Host "Could not find $user title!"
                                    }
                                    $row++
                                }
                            } else {
                                $row++
                            }   
                        }
                        $column -= 1
                    }
                    if (!$nestedadgroupmembers) {
                        $row++
                    }
                }
            }
        } else {
            $sheet.Cells.Item($row, $column) = "Could not find $grp"
            $sheet.Cells.Item($row,$column).Interior.ColorIndex = 3
            Write-Host "Could not find $grp!"
        }
        $adgroup = $null
        $column += 3
        if ($doublenested) {
            $column += 1
        }
    }
    $usedRange = $sheet.UsedRange						
    $usedRange.EntireColumn.AutoFit() | Out-Null
}

Function Get-ADGroupUserAccess ($group,$appcode) {
    $xls = New-Object -ComObject excel.application
    $xls.Workbooks.Add() | Out-Null
    $xls.Visible = $true
    $sheet = $xls.Worksheets.Item(1)
    $sheet.Name = $appcode
    #Titles
    $sheet.Cells.Item(1, 1) = 'USERID'
    $sheet.Cells.Item(1, 2) = 'VOLLEDIGE_NAAM'
    $sheet.Cells.Item(1, 3) = 'APPLICATIE_CODE'
    $sheet.Cells.Item(1, 4) = 'AD_GROEPSNAAM'
    $sheet.Cells.Item(1, 5) = 'FUNCTIE'
    $sheet.Cells.Item(1, 6) = 'DIRECTIE'
    $sheet.Cells.Item(1, 7) = 'NAAM_VERANTWOORDELIJKE'
    #Entries
    $row = 2
    $groups = $group -split ','
    foreach ($grp in $groups) {
        $adgroup = Get-ADGroup $grp | Select-Object Name
        if ($adgroup) {
            $groupmembers = Get-ADGroupMember $adgroup.Name
            foreach ($user in $groupmembers) {
                $userinfo = Get-Aduser $user -Properties Title,Department,Manager | Select-Object SamAccountName,Name,Title,Department,Manager
                $manager = Get-Aduser $userinfo.Manager | Select-Object Name
                #Fill in Excel sheet user by user
                $sheet.Cells.Item($row, 1) = $user.samAccountName
                $sheet.Cells.Item($row, 2) = $userinfo.Name
                $sheet.Cells.Item($row, 3) = $appcode
                $sheet.Cells.Item($row, 4) = $adgroup.Name
                $sheet.Cells.Item($row, 5) = $userinfo.Title
                $sheet.Cells.Item($row, 6) = $userinfo.Department
                $sheet.Cells.Item($row, 7) = $manager.Name
                $row++
            }
        }
        else {
            $sheet.Cells.Item($row, $column) = "Could not find $grp"
            Write-Host "Could not find $grp!"
        }
    }
    $usedRange = $sheet.UsedRange						
    $usedRange.EntireColumn.AutoFit() | Out-Null
}

Function Get-SBPUserAccess {
    $xls = New-Object -ComObject excel.application
    $xls.Workbooks.Add() | Out-Null
    $xls.Visible = $true
    $sheet = $xls.Worksheets.Item(1)
    $sheet.Name = 'SBP'
    #Titles
    $sheet.Cells.Item(1, 1) = 'USERID'
    $sheet.Cells.Item(1, 2) = 'VOLLEDIGE_NAAM'
    $sheet.Cells.Item(1, 3) = 'APPLICATIE_CODE'
    $sheet.Cells.Item(1, 4) = 'SBP_GROEPSNAAM'
    $sheet.Cells.Item(1, 5) = 'FUNCTIE'
    $sheet.Cells.Item(1, 6) = 'DIRECTIE'
    $sheet.Cells.Item(1, 7) = 'NAAM_VERANTWOORDELIJKE'
    #Entries
    $row = 2
    $cur = import-excel '.\sbp.xlsx'
    foreach ($user in $cur) {
        #Write-Host $user
        $userinfo = Get-Aduser $user.Logon -Properties Title,Department,Manager | Select-Object SamAccountName,Name,Title,Department,Manager
        if (!$userinfo) {
            Write-Error 'User {0} does not exist, enter new logon!' -ForegroundColor Red -f $user
            $user = Read-Host 'Logon: '
        }
        $manager = Get-Aduser $userinfo.Manager | Select-Object Name
        #Fill in Excel sheet user by user
        $sheet.Cells.Item($row, 1) = $user.Logon
        $sheet.Cells.Item($row, 2) = $userinfo.Name
        $sheet.Cells.Item($row, 3) = 'SBP'
        $sheet.Cells.Item($row, 4) = $user.Group
        $sheet.Cells.Item($row, 5) = $userinfo.Title
        $sheet.Cells.Item($row, 6) = $userinfo.Department
        $sheet.Cells.Item($row, 7) = $manager.Name
        $row++
    }
    $usedRange = $sheet.UsedRange						
    $usedRange.EntireColumn.AutoFit() | Out-Null
}

<#function Out-Excel ($arr, $title, $header) {
    $xls = New-Object -ComObject excel.application
    $xls.Workbooks.Add() | Out-Null
    $xls.Visible = $true
    $sheet = $xls.Worksheets.Item(1)
    $sheet.Name = 'Out-Excel'
    $column = 1
    $row = 1
    foreach ($entry in $arr) {
        $sheet.Cells.Item($row, $column) = $entry
        $row++
    }
}#>
<#
$addressbook = $outlook.Session.GetGlobalAddressList().addressentries
$entries = foreach ($address in $addressbook) {write-output $address.GetExchangeDistributionList()}
$belang = $entries | where name -eq 'belangenconflicten'
#>

Function Out-Pivot ($path) {
    $xls = $path | Import-Excel
    foreach ($directie in ($xls | Select-Object DIRECTIE -Unique).DIRECTIE) {
        #New-Item -ItemType Directory -Path $directie #run once
        write-host 'Creating pivot for:' $directie'...'
        #$xls | Where-Object DIRECTIE -Match $directie | Export-Excel -Path $directie'\Groepen_'$directie'.xlsx' -IncludePivotTable -PivotRows DIRECTIE,NAAM_VERANTWOORDELIJKE,FUNCTIE,VOLLEDIGE_NAAM -PivotData LAST_LOGON -PivotColumns APPLICATIE_CODE,SADM_GROEPSNAAM -NoTotalsInPivot #Groepen
        #$xls | Where-Object DIRECTIE -Match $directie | Export-Excel -Path $directie'\Taken_'$directie'.xlsx' -IncludePivotTable -PivotRows APPLICATIE_NAAM,TAAK_NAAM -PivotData PRESENT -PivotColumns SADM_GROEPSNAAM -NoTotalsInPivot #Taken
        #$xls | Where-Object DIRECTIE -Match $directie | Export-Excel -Path "\\Pdc6601\data\Prd\hoofdzetel\IAM HTMP\$directie\AD_$directie.xlsx" -IncludePivotTable -PivotRows DIRECTIE,NAAM_VERANTWOORDELIJKE,FUNCTIE,VOLLEDIGE_NAAM -PivotData PRESENT -PivotColumns APPLICATIE_CODE,AD_GROEPSNAAM -NoTotalsInPivot #AD
        $xls | Where-Object DIRECTIE -Match $directie | Export-Excel -Path "$directie\SBP_$directie.xlsx" -IncludePivotTable -PivotRows DIRECTIE,NAAM_VERANTWOORDELIJKE,FUNCTIE,VOLLEDIGE_NAAM -PivotData PRESENT -PivotColumns APPLICATIE_CODE,SBP_GROEPSNAAM -NoTotalsInPivot #SBP
        write-host 'Created pivot for:' $directie
	}
}
