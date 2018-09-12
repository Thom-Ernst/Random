function get-test ($logon) {

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

function Out-Excel ($arr, $title, $header) {
    $xls = New-Object -ComObject excel.application
    $xls.Workbooks.Add() | Out-Null
    $xls.Visible = $true
    $sheet = $xls.Worksheets.Item(1)
    $sheet.Name = 'Out-Excel'
    $column = 1
    if (!$header) {
        $row = 1
        foreach ($entry in $arr) {
            $sheet.Cells.Item($row, $column) = $entry
            $row++
        }
    }
}
