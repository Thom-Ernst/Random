Function Open-RoleGrid ($list,$level) {
    if ($level -eq 10) {
        $text = 'Level10 Status'
        $date = (Get-ChildItem .\level10.csv).LastWriteTime
    }
    elseif ($level -eq 30) {
        $text = 'Level30 Status'
        $date = (Get-ChildItem .\level30.csv).LastWriteTime
    }
    $roleselected = $list[0]
        $selectedrole = $list | Out-GridView  -Title  "$text - $date"
        <#if ($selectedrole) {
            $roleselected = $selectedrole.nrfEntitlementRef -split "`r`n" | Out-GridView -Title $selectedrole.cn
        }#>
    }


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
    Open-RoleGrid $Level10RoleList 10
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
    Open-RoleGrid $Level30RoleList 30
}

function Get-IamRole ($search){
    $FilteredRoleList = $Level10RoleList | Where-Object cn -Like *$search*
    Open-RoleGrid $FilteredRoleList 10
}
function Get-IamEntitlement ($search){
    $FilteredRoleList = $Level10RoleList | Where-Object nrfEntitlementRef -Like *$search*
    Open-RoleGrid $FilteredRoleList 10
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
