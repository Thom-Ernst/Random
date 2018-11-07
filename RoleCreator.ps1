Function Read-ExcelInput ($role) {
    $driverlist = $role.Driver -split ','
    $typelist = $role.Type -split ','
    $grouplist = $role.Group -split ','
    if ($role.Driver) {
        $entitlements = @()
        for ($i=0; $i -lt $driverlist.count; $i++) {
            $entitlement = @{
                Driver = $driverlist[$i]
                Type = $typelist[$i]
                Group = $grouplist[$i]
            }
            $entitlements += $entitlement
        }
    }
    else {
        $entitlements = $null
    }
    $RoleObject = @{
        CName = $role.CName
        LName = $role.Name
        LDesc = $role.Description
        Owner = $role.Owner
        Entitlements = $entitlements
        RequestType = $role.RequestType
    }
    return $RoleObject
}

Function Out-BatchRequestXML ($rolesitems) {
    [xml]$batch = New-Object System.Xml.XmlDocument
    $batch.AppendChild($batch.CreateXmlDeclaration("1.0","UTF-8",$null)) #adds metadata for xml file
    $req = $batch.CreateNode("element","batchRequest",$null)
    $req.SetAttribute("xmlns:xsd","http://www.w3c.org/2001/XMLSchema")
    $req.SetAttribute("xmlns:xsi","http://www.w3c.org/2001/XMLSchema-instance")
    foreach ($role in $roleitems) {
        $rolerequest = $batch.CreateNode("element","addRequest",$null)
        $rolerequest.SetAttribute('dn','cn='+$role.CName+',cn=Level10,cn=RoleDefs,cn=RoleConfig,cn=AppConfig,cn=UserApplication,cn=DriverSet,ou=RESOURCES,o=SYSTEM') #location of role with absolute path
        ###attributes
        #rolelevel
        $rolelevelattr = $batch.CreateNode("element","attr",$null)
        $rolelevelattr.SetAttribute("name","nrfrolelevel")
        $rolelevelattr.AppendChild($batch.CreateElement("value")).InnerText = "10" #make variable for this in future
        $rolerequest.AppendChild($rolelevelattr)
        #nrfstatus
        $statusattr = $batch.CreateNode("element","attr",$null)
        $statusattr.SetAttribute("name","nrfstatus")
        $statusattr.AppendChild($batch.CreateElement("value")).InnerText = "50"
        $rolerequest.AppendChild($statusattr)
        #rolecategorykey
        $rolecategorykeyattr = $batch.CreateNode("element","attr",$null)
        $rolecategorykeyattr.SetAttribute("name","nrfrolecategorykey")
        $rolecategorykeyattr.AppendChild($batch.CreateElement("value")).InnerText = "default"
        $rolerequest.AppendChild($rolecategorykeyattr)
        #ojectclass
        $objectclassattr = $batch.CreateNode("element","attr",$null)
        $objectclassattr.SetAttribute("name","objectclass")
        $objectclassattr.AppendChild($batch.CreateElement("value")).InnerText = "nrfRole"
        $objectclassattr.AppendChild($batch.CreateElement("value")).InnerText = "Top"
        $rolerequest.AppendChild($objectclassattr)
        #acl
        $aclattr = $batch.CreateNode("element","attr",$null)
        $aclattr.SetAttribute("name","acl")
        $aclattr.AppendChild($batch.CreateElement("value")).InnerText = '1#subtree#ou=Actives,ou=Users,o=IDENTITIES#[Entry Rights]'
        $aclattr.AppendChild($batch.CreateElement("value")).InnerText = '3#subtree#ou=Actives,ou=Users,o=IDENTITIES#[All Attributes Rights]'
        $rolerequest.AppendChild($aclattr)
        ##Variable attributes
        #cn
        $cnattr = $batch.CreateNode("element","attr",$null)
        $cnattr.SetAttribute("name","cn")
        $cnattr.AppendChild($batch.CreateElement("value")).InnerText = $role.CName
        $rolerequest.AppendChild($cnattr)
        #localizednames
        $localizednamesattr = $batch.CreateNode("element","attr",$null)
        $localizednamesattr.SetAttribute("name","nrflocalizednames")
        $localizednamesattr.AppendChild($batch.CreateElement("value")).InnerText = 'en~'+$role.LName+'|nl~'+$role.LName
        $rolerequest.AppendChild($localizednamesattr)
        #localizeddescrs
        $localizeddescrsattr = $batch.CreateNode("element","attr",$null)
        $localizeddescrsattr.SetAttribute("name","nrflocalizeddescrs")
        $localizeddescrsattr.AppendChild($batch.CreateElement("value")).InnerText = 'en~'+$role.LDesc+'|nl~'+$role.LDesc
        $rolerequest.AppendChild($localizeddescrsattr)
        #owner
        if ($role.owner){
        $ownerattr = $batch.CreateNode("element","attr",$null)
        $ownerattr.SetAttribute("name","owner")
        $ownerattr.AppendChild($batch.CreateElement("value")).InnerText = 'cn='+$role.Owner+',ou=Actives,ou=Users,o=IDENTITIES'
        $rolerequest.AppendChild($ownerattr)
        }
        #entitlementref
        if ($role.Entitlements) {
        $entitlementrefattr = $batch.CreateNode("element","attr",$null)
        $entitlementrefattr.SetAttribute("name","nrfentitlementref")
        foreach ($entitlement in $role.Entitlements) {
            $entitlementxml = 'cn={1},cn={0},cn=DriverSet,ou=RESOURCES,o=SYSTEM#0#<ref><id/><param>{2}</param></ref>
        '-f $entitlement.Driver,$entitlement.Type,$entitlement.Group
            $base64entitlement = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($entitlementxml))
            $entvalue = $batch.CreateElement("value")
            $entvalue.InnerText = $base64entitlement
            $xsi_type = $batch.CreateNode("attribute", "xsi", "type", "http://www.w3c.org/2001/XMLSchema-instance")
            $xsi_type.Value = "xsd:base64Binary"
            $entvalue.SetAttributeNode($xsi_type)
            $entitlementrefattr.AppendChild($entvalue)
            $rolerequest.AppendChild($entitlementrefattr)
        }
        }
        #requestdef
        if ($role.RequestType) {
            $requestdefattr = $batch.CreateNode("element","attr",$null)
            $requestdefattr.SetAttribute("name","nrfrequestdef")
            $requestdefattr.AppendChild($batch.CreateElement("value")).InnerText = 'cn='+$role.RequestType+',cn=RequestDefs,cn=AppConfig,cn=UserApplication,cn=DriverSet,ou=RESOURCES,o=SYSTEM' #make variable for this in future
            $rolerequest.AppendChild($requestdefattr)
        }
        ###Append Role to Batchrequest
        $req.AppendChild($rolerequest)
    }
    $batch.AppendChild($req)
    $batch.Save("newroles.xml")
}

Function Create-Roles ($file) {
    if (!$file) {
        $file = '.\rolexcel.xlsx'
    }
    $xls = Import-Excel $file
    $roleitems = @()
    foreach ($row in $xls) {
        $role = Read-ExcelInput $row
        $roleitems += $role
    }
    Out-BatchRequestXML $roleitems | Out-Null
    
}
