Function Open-RoleGridLevel10 ($list) {
    #$date = (Get-ChildItem .\roles.csv).LastWriteTime
    $date = 'TODO - update this code'
    $roleselected = $list[0]

        $selectedrole = $list | Out-GridView -PassThru -Title "Level10/30 status - $date"
        if ($selectedrole) {
            $roleselected = $selectedrole.nrfEntitlementRef -split "`r`n" | Out-GridView -PassThru -Title $selectedrole.cn
        }
    }


function Parse-Level10Roles {
    $path = '.\level10.csv'
    $q = import-csv $path
    foreach ($role in $q) {
        #name
        #show only nl version
        if ($role.nrfLocalizedNames) {
            $name = $role.nrfLocalizedNames
            $Matches = ''
            $name -Match '~(.+)\|'
            $newname = $Matches[1]
            $role.nrfLocalizedNames = $newname
        }
        #description
        #show only nl version
        if ($role.nrfLocalizedDescrs) {
            $description = $role.nrfLocalizedDescrs
            $Matches = ''
            $description -Match '~(.+)\|'
            $newdescription = $Matches[1]
            $role.nrfLocalizedDescrs = $newdescription
        }
        #owner
        #show cn and first ou
        if ($role.owner) {
            $owner = $role.owner
            $Matches = ''
            $owner -Match 'cn=(.+),ou=(.+),ou.+'
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
                $entitlement -Match 'cn=(.+),cn=.*,cn=.*'
                $entitlementtype = $Matches[1]
                $Matches = ''
                $entitlement -Match '<param>(.+)</param>'
                $entitlementparam = $Matches[1]
                $entitlementfull = "$entitlementtype,$entitlementparam"
                $arr.Add($entitlementfull)
            }
            # Row for each nrfentitlement type
            $newentitlements = $arr -join "`r`n"
            $role.nrfEntitlementRef = $newentitlements
        }
        #request def
        if ($role.nrfRequestDef){
            $requestdef = $role.nrfRequestDef
            $Matches = ''
            $requestdef -Match 'cn=(.+),cn=RequestDefs,cn=.+,ou.+'
            $newreqdef = $Matches[1]
            $role.nrfRequestDef = $newreqdef
        }
    }
    $global:Level10RoleList = $q
    Open-RoleGridLevel10 $Level10RoleList
}

function Parse-Level30Roles {
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
                $entitlement -Match 'cn=(.+),cn=.*,cn=.*'
                $entitlementtype = $Matches[1]
                $Matches = ''
                $entitlement -Match '<param>(.+)</param>'
                $entitlementparam = $Matches[1]
                $entitlementfull = "$entitlementtype,$entitlementparam"
                $arr.Add($entitlementfull)
            }
            # Row for each nrfentitlement type
            $newentitlements = $arr -join "`r`n"
            $role.nrfEntitlementRef = $newentitlements
        }
    }
    $global:Level30RoleList = $q
    Open-RoleGridLevel10 $Level30RoleList
}

function Get-IamRole ($search){
    $FilteredRoleList = $IAMRoleList | Where-Object cn -Like *$search*
    Open-RoleGridLevel10 $FilteredRoleList
}
function Get-IamEntitlement ($search){
    $FilteredRoleList = $IAMRoleList | Where-Object nrfEntitlementRef -Like *$search*
    Open-RoleGridLevel10 $FilteredRoleList
}

## ↑ samenvoegen
Function Out-EntitlementList ($rolelist) {
    $processedroles = @()
    foreach ($role in ($rolelist | where nrfEntitlementRef)) {
        foreach ($entitlement in ($role.nrfEntitlementRef -split "`r`n")) {
            $driver,$parameter = $entitlement -split ','
            $roleobject = [PSCustomObject]@{
                CN = $role.cn
                Driver = $driver
                Parameter = $parameter
            }
            $processedroles += $roleobject
        }
    }
    return $processedroles
}
