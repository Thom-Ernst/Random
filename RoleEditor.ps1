<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    IAM Role Editor
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$RoleEditor                      = New-Object system.Windows.Forms.Form
$RoleEditor.ClientSize           = '1100,170'
$RoleEditor.text                 = "IAM Role Editor"
$RoleEditor.BackColor            = "#d2ebc6"
$RoleEditor.TopMost              = $false

$BrowsePath                      = New-Object system.Windows.Forms.TextBox
$BrowsePath.multiline            = $false
$BrowsePath.width                = 470
$BrowsePath.height               = 20
$BrowsePath.location             = New-Object System.Drawing.Point(15,40)
$BrowsePath.Font                 = 'Microsoft Sans Serif,9'

$FilePathLabel                   = New-Object system.Windows.Forms.Label
$FilePathLabel.text              = "File Path"
$FilePathLabel.AutoSize          = $true
$FilePathLabel.width             = 25
$FilePathLabel.height            = 10
$FilePathLabel.location          = New-Object System.Drawing.Point(15,10)
$FilePathLabel.Font              = 'Microsoft Sans Serif,12,style=Bold'

$BrowseButton                    = New-Object system.Windows.Forms.Button
$BrowseButton.BackColor          = "#7ed321"
$BrowseButton.text               = "Browse..."
$BrowseButton.width              = 78
$BrowseButton.height             = 30
$BrowseButton.location           = New-Object System.Drawing.Point(500,31)
$BrowseButton.Font               = 'Microsoft Sans Serif,10,style=Bold'
$BrowseButton.ForeColor          = ""

$XMLViewer                       = New-Object system.Windows.Forms.DataGridView
$XMLViewer.width                 = 1060
$XMLViewer.height                = 75
$XMLViewerData = @(@("","","","",""))
$XMLViewer.ColumnCount = 5
$XMLViewer.ColumnHeadersVisible = $true
$XMLViewer.Columns[0].Name = "CN"
$XMLViewer.Columns[1].Name = "Name"
$XMLViewer.Columns[2].Name = "Description"
$XMLViewer.Columns[3].Name = "Owner"
$XMLViewer.Columns[4].Name = "Entitlements"
foreach ($row in $XMLViewerData){
          $XMLViewer.Rows.Add($row)
      }
$XMLViewer.Anchor                = 'top,right,bottom,left'
$XMLViewer.location              = New-Object System.Drawing.Point(12,80)

$Commit                          = New-Object system.Windows.Forms.Button
$Commit.BackColor                = "#7ed321"
$Commit.text                     = "Commit"
$Commit.width                    = 78
$Commit.height                   = 30
$Commit.location                 = New-Object System.Drawing.Point(600,31)
$Commit.Font                     = 'Microsoft Sans Serif,10,style=Bold'

$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 60
$PictureBox1.height              = 60
$PictureBox1.Anchor              = 'top,right'
$PictureBox1.location            = New-Object System.Drawing.Point(925,10)
$PictureBox1.imageLocation       = "useraccess.png"
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Statusbar                       = New-Object system.Windows.Forms.ProgressBar
$Statusbar.width                 = 230
$Statusbar.height                = 25
$Statusbar.Anchor                = 'top,right'
$Statusbar.location              = New-Object System.Drawing.Point(688,34)

$StatusBarLabel                  = New-Object system.Windows.Forms.Label
$StatusBarLabel.AutoSize         = $true
$StatusBarLabel.width            = 25
$StatusBarLabel.height           = 10
$StatusBarLabel.Anchor           = 'top,right'
$StatusBarLabel.location         = New-Object System.Drawing.Point(740,18)
$StatusBarLabel.Font             = 'Microsoft Sans Serif,10,style=Bold'

$RoleEditor.controls.AddRange(@($BrowsePath,$FilePathLabel,$BrowseButton,$XMLViewer,$Commit,$PictureBox1,$Statusbar,$StatusBarLabel))

#region gui events {
$BrowseButton.Add_Click({ Get-XML })
$Commit.Add_Click({ Set-XML })
#endregion events }

#endregion GUI }

$XMLViewer.Columns[0].Width = 220
$XMLViewer.Columns[1].Width = 220
$XMLViewer.Columns[2].Width = 240
$XMLViewer.Columns[3].Width = 50
$XMLViewer.Columns[4].Width = 280
$XMLViewer.Rows.RemoveAt(0)
$XMLViewer.Rows.RemoveAt(0)
$XMLViewer.Rows.RemoveAt(0)
$XMLViewer.Rows.RemoveAt(0)
$XMLViewer.Rows.RemoveAt(0)

#Write your logic code here

#pre-vars
$dir = "$HOME\designer_workspace\IAM PRD\Model\Provisioning\AppConfig\RoleConfig\RoleDefs\Level10\"

#public functions

Function Get-XML {
    $Statusbar.Value = 0
    $pt = Get-FullIAMRoles
    if (!$pt) {
        $StatusBarLabel.text = "No new Role selected"
        return
    }
    $StatusBarLabel.text = "Searching Role..."
    $global:path = $pt
    $BrowsePath.text = $path
    [xml]$global:cursor = Get-Content $path
    $Statusbar.Value = 50
    Update-RoleData ($cursor)
    $Statusbar.Value = 100
    $StatusBarLabel.text = "Role Selected"
}

Function Set-XML {
    $Statusbar.Value = 0
    $StatusBarLabel.text = "Saving..."
    $data = ($XMLViewer.Rows[0].Cells | Select value).value
    $rcname = $data[0]
    $rlname = $data[1]
    $rldesc = $data[2]
    $rowner = $data[3]
    $rentitlements = $data[4]
    $Statusbar.Value = 20
    $cursor.Role.id = $rcname
    $cursor.Role.localizedName[0].label,$cursor.Role.localizedName[1].label = $rlname,$rlname
    $cursor.Role.localizedDescription[0].label,$cursor.Role.localizedDescription[1].label = $rldesc,$rldesc
    if ($rowner) {
        $cursor.Role.owner.dn = "$rowner.Actives.Users.IDENTITIES"
    }
    $Statusbar.Value = 60
    Set-Entitlements($rentitlements)
    $cursor.save($path)
    $Statusbar.Value = 100
    $StatusBarLabel.text = "Saved!"
}

#private functions

Function Update-RoleData ($cur) {
    $cname = $cur.Role.id
    $lname = $cur.Role.localizedName.label[0]
    $ldesc = $cur.Role.localizedDescription.label[0]
    $owner = ($cur.Role.owner.dn -split '\.')[0]
    $entitlements = $cur.Role.entitlement.parameter -join ';'
    $global:refxml = $cur.Role.entitlement.ReferenceXML
    if ($XMLViewer.rows.count -gt 1) {
        $XMLViewer.rows.RemoveAt(0)   
    }
    $XMLViewer.rows.add($cname,$lname,$ldesc,$owner,$entitlements)
}

<#Function Open-FileBrowser {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $dir
    $OpenFileDialog.filter = "IAM Role (*.role10)| *.role10"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}#>

function Get-FullIAMRoles {
    $q = Get-ChildItem -Path $global:dir *.role10  | Select-Object Basename | Out-GridView -PassThru
    if (!$q) {
        Write-Host "no role selected"
        $StatusBarLabel.text = "No role selected"
        return
    }
    return $global:dir+$q.BaseName+".role10"
}

Function Set-Entitlements ($rent) {
    $arent = $rent -split ';'
    if ($arent.count -gt 1) {
        $i = 0
        foreach ($ent in $arent) {
            $cursor.Role.entitlement[$i].parameter = $ent
            $rxml = $cursor.Role.entitlement[$i].ReferenceXML -replace "am>.*</pa", "am>$ent</pa"
            $cursor.Role.entitlement[$i].ReferenceXML = $rxml
            $i++
        }
    }
    elseif ($arent.count -eq 1) {
        $cursor.Role.entitlement.parameter = $rent
        $rxml = $cursor.Role.entitlement.ReferenceXML -replace "am>.*</pa", "am>$rent</pa"
        $cursor.Role.entitlement.ReferenceXML = $rxml
    }
    else {
        Write-Host "No entitlements changed"
    }
}

[void]$RoleEditor.ShowDialog()
