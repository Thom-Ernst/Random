import-module ActiveDirectory

# Functions that would be seen as 'private'
function Get-Username ($id) { #Use in other functions like (function $parameter)[1]
    $q = Get-Aduser -Filter {SAMaccountname -Like $id} | Select-Object Name
    $q -match "@{Name=(.*)}" #select names returns a weird string so we must escape it.
    return $matches[1]
}

function Get-Userlogon ($id) { #Use in other functions like (function $parameter)[1]
    $q = Get-Aduser -Filter {Name -Like $id} | Select-Object SAMaccountname
    $q -match "@{SAMaccountname=(.*)}" #select names returns a weird string so we must escape it.
    return $matches[1]
}

function Get-Idmail ($i, $p) { #Returns users principal email
    switch ($p) {
        0 {
            $q = Get-Aduser -Filter {Name -Like $i} | Select-Object userprincipalname
        }
        1 {
            $q = Get-Aduser -Filter {SAMaccountname -Like $i} | Select-Object userprincipalname
        }
    }
    return $q
}

function Get-Usermail ($username){ #Returns users' email from AD
    $r = (Read-Host "Input id or name? i/n").ToLower()
    if ($r -eq "i") {#from id look for name
        Write-Host "Lookup using id..."
        foreach ($i in $arrid) {
        $q = (Get-Adusermail $i 0)[1]
        }  
    } elseif ($r -eq "n") { #from name look for id
        Write-Host "Lookup using name..."
        foreach ($i in $arrid) {
        $firstname, $lastnames = $i -split ' ' #lastnames is [lastnames samen]
        $lastname = $lastnames -join ' '
        $invname = "{0} {1}" -f ($lastname, $firstname)
        $q = (Get-Adusermail $invname 1)[1]
        }
    } else {
        Write-Host "Error: Bad input!" -ForegroundColor Red
    }
    return $q
}


# Functions that would be seen as 'public'

# Clipboard operations
function Output-Folderpath ($path) {
    $output = "Naam Nieuwe AD groep(en) incl. domeinnaam (max. 15): Alles in CORPARG`nDLG_{0}_M`nGG_PROJECT_{0}_M`nNesting in groep(en): GG_PROJECT_{0}_M member maken van DLG_{0}_M`nDLG_{0}_M schrijfrechten geven op gedeelde folder GRPARG\#GRPARG\{0}" -f ($path)
    Write-Output $output | clip
}
Set-Alias fol Output-Folderpath

function Output-Mailboxname ($path) {
    $output = "DLG_MBX_{0}_M`nGG_PROJECT_MBX_{0}_M`nNesting in groep(en): GG_PROJECT_MBX_{0}_M member maken van DLG_MBX_{0}_M`n`nGelieve DLG_MBX_{0}_M te koppelen aan emailbox {0}@argenta.be" -f ($path)
    Write-Output $output | clip
}
Set-Alias mbx Output-Mailboxname

function Output-Folderchange ($path, $group=$path) {
    $output = "Server-naam: in domein CORPARG`nFolder-naam (gelieve het volledige pad te vermelden): \\GRPARG\#GRPARG\{0}`nGewenste aanpassing (access rights vermelden, reference users worden genegeerd): De groep {1} moet worden toegevoegd aan de folder {0}" -f ($path, $group)
    Write-Output $output | clip
}
Set-Alias folc Output-Folderchange

function Output-Sfttext ($name) {
    $mail = Get-Usermail $name
    $output = "new; {0}; {1}" -f ($name, $mail)
    Write-Output $output | clip
}
Set-Alias sft Output-Sfttext

#AD Interactions
Function Get-Allmembers ($group) { #Get all members in an AD group and inside groups in the group
    Write-Host "Fetching AD Members..."
    $q = Get-Adgroupmember $group -Recursive | Select-Object Name
    Write-Host "Found" $q.count -ForegroundColor Green
    $r = (Read-Host "Export as csv? y/n").ToLower()
    if ($r -eq "y") {
        Write-Host "Exporting csv..."
        $q | Export-Csv -Path .\export.csv
    } elseif ($r -eq "n") {
        Write-Host "Outputting in a grid..."
        $q | Out-Gridview
    } else {
        Write-Host "Error: Bad input!" -ForegroundColor Red
    }
}
Set-Alias gam Get-Allmembers

Function Get-Localmembers ($group) { #Get only the members inside the AD + get the nested groups
    Write-Host "Fetching AD Members..."
    $q = Get-Adgroupmember $group | Select-Object Name
    Write-Host "Found" $q.count -ForegroundColor Green
    $r = (Read-Host "Export as csv? y/n").ToLower()
    if ($r -eq "y") {
        Write-Host "Exporting csv..."
        $q | Export-Csv -Path .\export.csv
    } elseif ($r -eq "n") {
        Write-Host "Outputting in a grid..."
        $q | Out-Gridview
    } else {
        Write-Host "Error: Bad input!" -ForegroundColor Red
    }
}
Set-Alias glm Get-Localmembers

Function Get-Groupmemberships ($id) { # Get groups of ad user based on ID
    Write-Host "Fetching Group Members..."
    $q = Get-AdPrincipalGroupMembership $id | Select-Object Name
    Write-Host "Found" $q.count -ForegroundColor Green
    $r = (Read-Host "Export as csv? y/n").ToLower()
    if ($r -eq "y") {
        Write-Host "Exporting csv..."
        $q | Export-Csv -Path .\export.csv
    } elseif ($r -eq "n") {
        Write-Host "Outputting in a grid..."
        $q | Out-Gridview
    } else {
        Write-Host "Error: Bad input!" -ForegroundColor Red
    }
}
Set-Alias ggm Get-Groupmemberships

Function Find-Adgroup ($name) { # Finds all AD groups that contain these characters
    $i = "*{0}*" -f ($name)
    $q = Get-Adgroup -Filter {name -like $i} | Select-Object Name
    Write-Host "Found" $q.count -ForegroundColor Green
    $r = (Read-Host "Export as csv? y/n").ToLower()
    if ($r -eq "y") {
        Write-Host "Exporting csv..."
        $q | Export-Csv -Path .\export.csv
    } elseif ($r -eq "n") {
        Write-Host "Outputting in a grid..."
        $q | Out-Gridview
    } else {
        Write-Host "Error: Bad input!" -ForegroundColor Red
    }
}
Set-Alias fag Find-Adgroup

function Get-Multipleusers ($ids) {
    $arrid = $ids -split "," #Split csv into array
    $usrnames = @()
    $r = (Read-Host "Input id or name? i/n").ToLower()
    if ($r -eq "i") { #from id look for name
        Write-Host "Lookup using id..."
        foreach ($i in $arrid) {
            $q = (Get-Username $i)[1]
            $usrnames += $q
        }
    } elseif ($r -eq "n") { #from name look for id
        Write-Host "Lookup using name..."
        foreach ($i in $arrid) {
            $firstname, $lastnames = $i -split ' ' #lastnames is [lastnames samen]
            $lastname = $lastnames -join ' '
            $invname = "{0} {1}" -f ($lastname, $firstname)
            $q = (Get-Userlogon $invname)[1]
            $usrnames += $q
        }
    } else {
        Write-Host "Error: Bad input!" -ForegroundColor Red
    }
    return $usrnames
}
Set-Alias gmu Get-Multipleusers
