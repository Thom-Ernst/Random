<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$ReportingTools                  = New-Object system.Windows.Forms.Form
$ReportingTools.ClientSize       = '376,406'
$ReportingTools.text             = "AD Reporting Tools"
$ReportingTools.BackColor        = "#d2ebc6"
$ReportingTools.TopMost          = $false
$ReportingTools.icon             = ".\img\useraccess.ico"

$InputBox                        = New-Object system.Windows.Forms.TextBox
$InputBox.multiline              = $true
$InputBox.width                  = 344
$InputBox.height                 = 299
$InputBox.Anchor                 = 'top,right,bottom,left'
$InputBox.location               = New-Object System.Drawing.Point(16,95)
$InputBox.Font                   = 'Microsoft Sans Serif,10'

$LogoBox                         = New-Object system.Windows.Forms.PictureBox
$LogoBox.width                   = 75
$LogoBox.height                  = 57
$LogoBox.Anchor                  = 'top,right'
$LogoBox.location                = New-Object System.Drawing.Point(285,23)
$LogoBox.imageLocation           = ".\img\useraccess.png"
$LogoBox.SizeMode                = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$StartButton                     = New-Object system.Windows.Forms.Button
$StartButton.BackColor           = "#7ed321"
$StartButton.text                = "Submit"
$StartButton.width               = 89
$StartButton.height              = 40
$StartButton.location            = New-Object System.Drawing.Point(170,44)
$StartButton.Font                = 'Microsoft Sans Serif,10,style=Bold'

$SelectionBox                    = New-Object system.Windows.Forms.Groupbox
$SelectionBox.height             = 73
$SelectionBox.width              = 135
$SelectionBox.BackColor          = "#dddddd"
$SelectionBox.location           = New-Object System.Drawing.Point(16,13)

$InfoLabel                       = New-Object system.Windows.Forms.Label
$InfoLabel.text                  = "Lookup Type:"
$InfoLabel.AutoSize              = $true
$InfoLabel.width                 = 25
$InfoLabel.height                = 10
$InfoLabel.location              = New-Object System.Drawing.Point(7,7)
$InfoLabel.Font                  = 'Microsoft Sans Serif,10,style=Bold'

$LookupOption1                   = New-Object system.Windows.Forms.RadioButton
$LookupOption1.text              = "Group Report"
$LookupOption1.AutoSize          = $true
$LookupOption1.width             = 104
$LookupOption1.height            = 20
$LookupOption1.location          = New-Object System.Drawing.Point(11,30)
$LookupOption1.Font              = 'Microsoft Sans Serif,10'

$LookupOption2                   = New-Object system.Windows.Forms.RadioButton
$LookupOption2.text              = "User Report"
$LookupOption2.AutoSize          = $true
$LookupOption2.width             = 104
$LookupOption2.height            = 20
$LookupOption2.location          = New-Object System.Drawing.Point(11,48)
$LookupOption2.Font              = 'Microsoft Sans Serif,10'

$ReportingTools.controls.AddRange(@($InputBox,$LogoBox,$StartButton,$SelectionBox))
$SelectionBox.controls.AddRange(@($InfoLabel,$LookupOption1,$LookupOption2))

#region gui events {
$StartButton.Add_Click({ Get-UserInput })
#endregion events }

#endregion GUI }


#Write your logic code here

Function Get-UserInput {
    if ($InputBox.Text) {
        $inputlist = $InputBox.Text -split '\r\n'
        if ($LookupOption1.Checked) {
            Get-GroupReport $inputlist
        }
        elseif ($LookupOption2.Checked) {
            Get-UserReport $inputlist
        }
        else {
            Write-Error 'No option Selected!'
        }
    }
    else {
        Write-Error 'No Input given!'
    }
}

Function Get-GroupReport ($groups) {
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

Function Get-UserReport ($users) {
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

[void]$ReportingTools.ShowDialog()
