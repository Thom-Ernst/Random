Function Open-RoleGridLevel10 ($list) {
    $selectedrole = $list | Out-GridView -PassThru
    if ($selectedrole) {
        $ent = $selectedrole.nrfEntitlementRef -split "`r`n" | Out-GridView -PassThru
    }
    if ($ent) {
        Open-RoleGridLevel10
    }
}

function Show-IamRoles {
    $path = '.\roles.csv'
    $q = import-csv $path
    foreach ($role in $q) {
    Write-Host $role.cn
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
    $global:IAMRoleList = $q
    Open-RoleGridLevel10 $IAMRoleList
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
