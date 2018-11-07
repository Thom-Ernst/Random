<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Role Viewer
.SYNOPSIS
    Using existing CSV files - lets you browse L10 and L30 Roles
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$RoleViewer                      = New-Object system.Windows.Forms.Form
$RoleViewer.ClientSize           = '499,430'
$RoleViewer.text                 = "Role Viewer"
$RoleViewer.BackColor            = "#d2ebc6"
$RoleViewer.TopMost              = $false
$RoleViewer.icon                 = ".\img\useraccess.ico"

$ParseBar                        = New-Object system.Windows.Forms.ProgressBar
$ParseBar.text                   = "Parsing Roles"
$ParseBar.width                  = 264
$ParseBar.height                 = 28
$ParseBar.location               = New-Object System.Drawing.Point(139,46)

$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 60
$PictureBox1.height              = 57
$PictureBox1.Anchor              = 'top,right'
$PictureBox1.location            = New-Object System.Drawing.Point(422,17)
$PictureBox1.imageLocation       = ".\img\useraccess.png"
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$ParseAllButton                  = New-Object system.Windows.Forms.Button
$ParseAllButton.BackColor        = "#7ed321"
$ParseAllButton.text             = "Reload Roles"
$ParseAllButton.width            = 112
$ParseAllButton.height           = 45
$ParseAllButton.location         = New-Object System.Drawing.Point(13,29)
$ParseAllButton.Font             = 'Microsoft Sans Serif,10,style=Bold'

$Level10Box                      = New-Object system.Windows.Forms.Groupbox
$Level10Box.height               = 153
$Level10Box.width                = 241
$Level10Box.Anchor               = 'top,right,left'
$Level10Box.text                 = "Role Data (Level10)"
$Level10Box.location             = New-Object System.Drawing.Point(242,99)

$Level30Box                      = New-Object system.Windows.Forms.Groupbox
$Level30Box.height               = 136
$Level30Box.width                = 241
$Level30Box.Anchor               = 'top,right,left'
$Level30Box.text                 = "Function Data (Level30)"
$Level30Box.location             = New-Object System.Drawing.Point(242,282)

$ShowRolesBtn                    = New-Object system.Windows.Forms.Button
$ShowRolesBtn.BackColor          = "#7ed321"
$ShowRolesBtn.text               = "Show Roles"
$ShowRolesBtn.width              = 112
$ShowRolesBtn.height             = 45
$ShowRolesBtn.location           = New-Object System.Drawing.Point(13,91)
$ShowRolesBtn.Font               = 'Microsoft Sans Serif,10,style=Bold'

$ShowFunctionsBtn                = New-Object system.Windows.Forms.Button
$ShowFunctionsBtn.BackColor      = "#7ed321"
$ShowFunctionsBtn.text           = "Show Functions"
$ShowFunctionsBtn.width          = 112
$ShowFunctionsBtn.height         = 45
$ShowFunctionsBtn.location       = New-Object System.Drawing.Point(14,151)
$ShowFunctionsBtn.Font           = 'Microsoft Sans Serif,10,style=Bold'

$Searchbar                       = New-Object system.Windows.Forms.TextBox
$Searchbar.multiline             = $false
$Searchbar.width                 = 210
$Searchbar.height                = 20
$Searchbar.location              = New-Object System.Drawing.Point(13,207)
$Searchbar.Font                  = 'Microsoft Sans Serif,10'

$SearchBtn                       = New-Object system.Windows.Forms.Button
$SearchBtn.BackColor             = "#7ed321"
$SearchBtn.text                  = "Search"
$SearchBtn.width                 = 60
$SearchBtn.height                = 30
$SearchBtn.location              = New-Object System.Drawing.Point(14,233)
$SearchBtn.Font                  = 'Microsoft Sans Serif,10,style=Bold'

$SearchDropdown                  = New-Object system.Windows.Forms.ComboBox
$SearchDropdown.text             = "Lookup type"
$SearchDropdown.width            = 142
$SearchDropdown.height           = 66
@('Role Name','Role Entitlement','Role Owner','Function Name','Function Entitlement') | ForEach-Object {[void] $SearchDropdown.Items.Add($_)}
$SearchDropdown.location         = New-Object System.Drawing.Point(82,233)
$SearchDropdown.Font             = 'Microsoft Sans Serif,10'

$Level10BoxLabelCount            = New-Object system.Windows.Forms.Label
$Level10BoxLabelCount.text       = "roles"
$Level10BoxLabelCount.AutoSize   = $true
$Level10BoxLabelCount.width      = 25
$Level10BoxLabelCount.height     = 10
$Level10BoxLabelCount.location   = New-Object System.Drawing.Point(13,28)
$Level10BoxLabelCount.Font       = 'Microsoft Sans Serif,10'

$Level30BoxLabelCount            = New-Object system.Windows.Forms.Label
$Level30BoxLabelCount.text       = "functions"
$Level30BoxLabelCount.AutoSize   = $true
$Level30BoxLabelCount.width      = 25
$Level30BoxLabelCount.height     = 10
$Level30BoxLabelCount.location   = New-Object System.Drawing.Point(13,22)
$Level30BoxLabelCount.Font       = 'Microsoft Sans Serif,10'

$Level10BoxLabelOwners           = New-Object system.Windows.Forms.Label
$Level10BoxLabelOwners.text      = "owners"
$Level10BoxLabelOwners.AutoSize  = $true
$Level10BoxLabelOwners.width     = 25
$Level10BoxLabelOwners.height    = 10
$Level10BoxLabelOwners.location  = New-Object System.Drawing.Point(13,48)
$Level10BoxLabelOwners.Font      = 'Microsoft Sans Serif,10'

$Level10BoxLabelRequests         = New-Object system.Windows.Forms.Label
$Level10BoxLabelRequests.text    = "requests"
$Level10BoxLabelRequests.AutoSize  = $true
$Level10BoxLabelRequests.width   = 25
$Level10BoxLabelRequests.height  = 10
$Level10BoxLabelRequests.location  = New-Object System.Drawing.Point(13,68)
$Level10BoxLabelRequests.Font    = 'Microsoft Sans Serif,10'

$Level30BoxLabelThaler           = New-Object system.Windows.Forms.Label
$Level30BoxLabelThaler.text      = "thaler"
$Level30BoxLabelThaler.AutoSize  = $true
$Level30BoxLabelThaler.width     = 25
$Level30BoxLabelThaler.height    = 10
$Level30BoxLabelThaler.location  = New-Object System.Drawing.Point(13,42)
$Level30BoxLabelThaler.Font      = 'Microsoft Sans Serif,10'

$Level30BoxLabelTenforce         = New-Object system.Windows.Forms.Label
$Level30BoxLabelTenforce.text    = "tenforce"
$Level30BoxLabelTenforce.AutoSize  = $true
$Level30BoxLabelTenforce.width   = 25
$Level30BoxLabelTenforce.height  = 10
$Level30BoxLabelTenforce.location  = New-Object System.Drawing.Point(13,62)
$Level30BoxLabelTenforce.Font    = 'Microsoft Sans Serif,10'

$RoleViewer.controls.AddRange(@($ParseBar,$PictureBox1,$ParseAllButton,$Level10Box,$Level30Box,$ShowRolesBtn,$ShowFunctionsBtn,$Searchbar,$SearchBtn,$SearchDropdown))
$Level10Box.controls.AddRange(@($Level10BoxLabelCount,$Level10BoxLabelOwners,$Level10BoxLabelRequests))
$Level30Box.controls.AddRange(@($Level30BoxLabelCount,$Level30BoxLabelThaler,$Level30BoxLabelTenforce))

#region gui events {
$ParseAllButton.Add_Click({ Get-Roles })
$ShowRolesBtn.Add_Click({ Out-Roles })
$ShowFunctionsBtn.Add_Click({ Out-Functions })
$SearchBtn.Add_Click({ Out-Search })
$RoleViewer.Add_Load({ Get-Roles })
#endregion events }

#endregion GUI }

Add-Type -AssemblyName PresentationFramework
$Searchbar.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Out-Search
    }
})

#Write your logic code here

#Privates
Function Parse-Level10Roles {
    $path = '.\level10.csv'
    $q = import-csv $path
    foreach ($role in $q) {
        #name
        #show only nl version
        if ($role.nrfLocalizedNames) {
            $name = $role.nrfLocalizedNames
            $Matches = ''
            $name -Match '~(.+)\|' | out-null
            $newname = $Matches[1]
            $role.nrfLocalizedNames = $newname
        }
        #description
        #show only nl version
        if ($role.nrfLocalizedDescrs) {
            $description = $role.nrfLocalizedDescrs
            $Matches = ''
            $description -Match '~(.+)\|' | out-null
            $newdescription = $Matches[1]
            $role.nrfLocalizedDescrs = $newdescription
        }
        #owner
        #show cn and first ou
        if ($role.owner) {
            $owner = $role.owner
            $Matches = ''
            $owner -Match 'cn=(.+),ou=(.+),ou.+' | out-null
            $newowner = $Matches[1,2] -join ','
            $role.owner = $newowner
        }
        #entitlements
        #show type and parameter of entitlement
        if ($role.nrfEntitlementRef) {
            $entitlements = $role.nrfEntitlementRef -split '\|'
            $arr = New-Object System.Collections.ArrayList
            foreach ($entitlement in $entitlements) {
                $Matches = ''
                $entitlement -Match 'cn=(.+),cn=.*,cn=.*' | out-null
                $entitlementtype = $Matches[1]
                $Matches = ''
                $entitlement -Match '<param>(.+)</param>' | out-null
                $entitlementparam = $Matches[1]
                $entitlementfull = "$entitlementtype,$entitlementparam"
                $arr.Add($entitlementfull) | out-null
            }
            # Row for each nrfentitlement type
            $newentitlements = $arr -join "`r`n"
            $role.nrfEntitlementRef = $newentitlements
        }
        #request def
        if ($role.nrfRequestDef){
            $requestdef = $role.nrfRequestDef
            $Matches = ''
            $requestdef -Match 'cn=(.+),cn=RequestDefs,cn=.+,ou.+' | out-null
            $newreqdef = $Matches[1]
            $role.nrfRequestDef = $newreqdef
        }
    }
    $global:Level10RoleList = $q
}

Function Parse-Level30Roles {
    $path = '.\level30.csv'
    $q = import-csv $path
    foreach ($role in $q) {
        #entitlements
        #show type and parameter of entitlement
        if ($role.nrfEntitlementRef) {
            $entitlements = $role.nrfEntitlementRef -split '\|'
            $arr = New-Object System.Collections.ArrayList
            foreach ($entitlement in $entitlements) {
                $Matches = ''
                $entitlement -Match 'cn=(.+),cn=.*,cn=.*' | out-null
                $entitlementtype = $Matches[1]
                $Matches = ''
                $entitlement -Match '<param>(.+)</param>' | out-null
                $entitlementparam = $Matches[1]
                $entitlementfull = "$entitlementtype,$entitlementparam"
                $arr.Add($entitlementfull) | out-null
            }
            # Row for each nrfentitlement type
            $newentitlements = $arr -join "`r`n"
            $role.nrfEntitlementRef = $newentitlements
        }
    }
    $global:Level30RoleList = $q
}

Function Update-RoleInfo {
    $Level10BoxLabelCount.Text = '{0} roles' -f $Level10RoleList.Count
    $Level10BoxLabelOwners.Text = '{0} roles without an owner' -f ($Level10RoleList | Where-Object owner -eq '').Count
    $Level10BoxLabelRequests.Text = '{0} roles without approval' -f ($Level10RoleList | Where-Object nrfRequestDef -eq '').Count
    #Add various metadata about roles?
    #Roles without owners, ...
    $Level30BoxLabelCount.Text = '{0} functions' -f $Level30RoleList.Count
    $Level30BoxLabelThaler.Text = '{0} functions with Thaler' -f ($Level30RoleList | Where-Object nrfEntitlementRef -like '*C12_Thaler*').Count
    $Level30BoxLabelTenforce.Text = '{0} functions with TenForce' -f ($Level30RoleList | Where-Object nrfEntitlementRef -like '*C23_TF*').Count
}

#Publics

Function Get-Roles {
    $ParseBar.Value = 10
    Parse-Level10Roles
    $ParseBar.Value = 50
    Parse-Level30Roles
    $ParseBar.Value = 80
    Update-RoleInfo
    $ParseBar.Value = 100
}

Function Out-Roles {
    $Level10RoleList | Out-GridView -Title 'Level10 Status'
}

Function Out-Functions {
    $Level30RoleList | Out-GridView -Title 'Level30 Status'
}

Function Out-Search {
    if ($Searchbar.Text) {
        $ParseBar.Value = 0
        $search = $Searchbar.Text
        Switch ($SearchDropdown.SelectedIndex) {
            -1 {
                Write-Error 'No search option selected!'
            }
            0 {
                #Search Role Name
                $q = $Level10RoleList | Where-Object nrfLocalizedNames -like "*$search*"
                if ($q) {
                    $q | Out-GridView -Title "Search result for Role Name: $search"
                } else {
                    return [System.Windows.MessageBox]::Show('No roles found')
                }
            }
            1 {
                #Search Role Entitlement
                $q = $Level10RoleList | Where-Object nrfEntitlementRef -like "*$search*"
                if ($q) {
                    $q | Out-GridView -Title "Search result for Role Entitlement: $search"
                } else {
                    return [System.Windows.MessageBox]::Show('No roles found')
                }
            }
            2 {
                #Search Role Owner
                $q = $Level10RoleList | Where-Object Owner -like "*$search*"
                if ($q) {
                    $q | Out-GridView -Title "Search result for Role Owner: $search"
                } else {
                    return [System.Windows.MessageBox]::Show('No roles found')
                }
            }
            3 {
                #Search Function Name
                $q = $Level30RoleList | Where-Object nrfLocalizedNames -like "*$search*"
                if ($q) {
                    $q | Out-GridView -Title "Search result for Function Name: $search"
                } else {
                    return [System.Windows.MessageBox]::Show('No functions found')
                }    
            }
            4 {
                #Search Function Entitlement
                $q = $Level30RoleList | Where-Object nrfEntitlementRef -like "*$search*"
                if ($q) {
                    $q | Out-GridView -Title "Search result for Function Entitlement: $search"
                } else {
                    return [System.Windows.MessageBox]::Show('No functions found')
                }
            }
        }
        $ParseBar.Value = 100
    }
}

[void]$RoleViewer.ShowDialog()