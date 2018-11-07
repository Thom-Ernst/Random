#import-module ActiveDirectory

##########################################Private

#Function Print-Input ($input, $text = "Your input ") { ###Huge WIP
#    $i = $input
#    while (!$i) {
#    
#    Write-Host  "Please enter the needed input..." -ForegroundColor Green
#        $i = Read-Host $text
#    }
#    return $i
#}

Function Out-Query ($query, $search="Query Output"){
    Function Get-Input ($query = $query) {
        $r = (Read-Host "Export as csv, excel, grid or interactive grid?`n`nc/e/g/i").ToLower()
        #$r = 'g'
        switch ($r){ #c,g,i,e,default
            c {
                $f = Read-Host "Name of the file?`nDefault: $search.csv"
                if (!$f) {
                    Write-Host "Defaulting name."
                    $f = $search
                }
                Write-Host "Exporting data as $f.csv..."
                $query | Export-Csv -Path "./filedump/$f.csv"
            }
            g {
                Write-Host "Outputting in a grid..."
                $query | Out-Gridview -Title $search
            }
            i {
                Write-Host "Outputting in an interactive grid..."
                $out = $query | Out-Gridview -Title $search -PassThru #<-needs user input
                if ($out) {
                    Write-Host "Running new search!" -ForegroundColor Yellow
                    Get-Groupmembers (($out -split '=')[1] -split '}')[0] r
                }
            }
            e {
                $f = Read-Host "Name of the file?`nDefault: $search.xlsx"
                    if (!$f) {
                        Write-Host "Defaulting name."
                        $f = $search
                    }
                    Write-Host "Exporting data as $f.xlsx..."
                    $query | Export-Excel -Path "./filedump/$f.xlsx" -Show
                }
                default {
                    Write-Host "Error: Bad input!" -ForegroundColor Red
                    Get-Input
            }
        }
    }
    $c = $query.count #TODO!
    #$c = 2
    Write-Host #Line Break
    if ($c -gt 1 -or $c -eq 0) {
        Write-Host "Found $c" -ForegroundColor Green #How many items?
        if ($c -ne 0) { #If nothing is found, no need to output anything
            Get-Input #Output dialogue
        }
    } else {
        Write-Host "Found 1" -ForegroundColor Green #One item returns customPSobject which is nullvalued
        if (($query -split '=').count -gt 1) {
            $result = (($query -split '=')[1] -split '}')[0]
        } else {
            $result = $query
        }
        Write-Host $result -ForegroundColor Yellow
    }
}

Function Export-XmlDoc ($logon, $path) {
    Write-Host "Writing xml for $logon... `nFind the results in $path"
    $out = '<?xml version="1.0"?><CommandLineFile xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">  <QueryString>&lt;?xml version="1.0" encoding="utf-16"?&gt;&lt;!--AD Info Query - Application Version 1.7.9.0--&gt;&lt;Query&gt;&lt;QueryType&gt;7&lt;/QueryType&gt;&lt;Name&gt;User with specified username&lt;/Name&gt;&lt;IconIndex&gt;72&lt;/IconIndex&gt;&lt;AllowModify&gt;False&lt;/AllowModify&gt;&lt;CreatedOn&gt;23/01/2011 21:37:28&lt;/CreatedOn&gt;&lt;Author&gt;Cjwdev&lt;/Author&gt;&lt;Parameters&gt;&lt;Parameter&gt;&lt;Attribute&gt;UsernamePre2000&lt;/Attribute&gt;&lt;Operator&gt;is&lt;/Operator&gt;&lt;Prompt&gt;False&lt;/Prompt&gt;&lt;NoValue&gt;False&lt;/NoValue&gt;&lt;Value&gt;{0}&lt;/Value&gt;&lt;/Parameter&gt;&lt;/Parameters&gt;&lt;/Query&gt;</QueryString>  <Domain>CORPARG.LAN</Domain>  <CustomAttributes />  <IncludedAttributeIds>    <string>AllGroupMembership</string>    <string>Cn</string>  </IncludedAttributeIds>  <ContainerPath />  <IncludeSubContainers>true</IncludeSubContainers>  <ExportPath>{1}</ExportPath>  <LdapPort>389</LdapPort>  <GcPort>3268</GcPort></CommandLineFile>' -f ($logon,$path)
    $out | Out-File -FilePath config.xml
}

Function Get-InvertedName ($name) { #put first name last or vice versa
    $namearr = $name -split ' '
    $lastname = $namearr[1..($namearr.count-1)] -join ' '
    return $lastname + ' ' + $namearr[0] 
}

Function Get-InvertedLastName ($name) { #put first name last or vice versa
    $namearr = $name -split ' '
    $lastname = $namearr[0..($namearr.count-2)] -join ' '
    return $namearr[$namearr.count-1] + ' ' + $lastname
}

Function Get-Name ($logon, $swap) { #add 's' to change to firstname name
    Write-Host "Getting username..."
    $q = Get-ADUser -Filter {SAMAccountName -Like $logon} | Select-Object Name #query for name from samaccountname
    if ($swap) {
        #return Get-Invertedname $s
        return $q.Name
    } else {
        return $q.Name
    }
}

Function Get-Logon ($name) { #Input firstname lastname
    #query for samaccountname
    $name = Get-Invertedname $name
    Write-Host "Getting userid..."
    $q = Get-ADUser -Filter {Name -Like $name} | Select-Object SAMAccountName
    return $q.SAMAccountName
}

Function Get-Email ($logon) { #get the email using logon, returns named address.
    #query for email
    Write-Host "Getting email..."
    $q = Get-ADUser -Filter {SAMAccountName -Like $logon} -Properties EmailAddress | Select-Object EmailAddress
    return $q.EmailAddress
}

Function Get-SamEmail ($logon) { #get the email using logon, returns technical address, not named address. ##WIP
    #query for email
    Write-Host "Getting email..."
    $q = Get-ADUser -Filter {SAMAccountName -Like $logon} | Select-Object UserPrincipalName
    return $q.UserPrincipalName
}

Function Get-Fulluserinfo ($logon) {
    Write-Host "Getting ad user info..."
    $q = Get-ADUser $logon -Properties EmailAddress,Title | Select-Object Name,SamAccountName,EmailAddress,Enabled,GivenName,Surname,UserPrincipalName,DistinguishedName,Title
    return $q
}

Function Get-UserNestedGroups ($logon, $path) {
    Write-Host "Generate $path"
    Export-XmlDoc $logon $path
    .\ADInfoCmd.exe /config "config.xml" | Out-Null
}

Function Get-AdUserMemberships ($logon) {
    while (!$logon) {
        Write-Host "Please enter the needed input..." -ForegroundColor Green
        $logon = Read-Host "User login "
    }
    #$f = Read-Host "Name of the file?`nDefault: $logon.csv"
    if (!$f) {
        Write-Host "Defaulting name."
        $f = "$logon.csv"
    }
    Get-UserNestedGroups $logon $f
    $q = import-csv $f
    return $q
}

Function Get-OrganizationInfo ($logon) {
    while (!$logon) {
        Write-Host "Please enter the needed input..." -ForegroundColor Green
        $logon = Read-Host "User login "
    }
    $userinfo = Get-ADUser $logon -Properties Title,Department,Manager | Select-Object Title,Department,Manager
    $manager = Get-Aduser $userinfo.Manager | Select-Object Name
    $out = "{0}`t{1}`t{2}" -f $userinfo.Title,$userinfo.Department,$manager.Name
    return $out
}


##########################################Public

#Get

Function Get-Groups ($name, $sv) {
    while (!$name) { #Print-Input $name "Groep name "
        Write-Host  "Please enter the needed input..." -ForegroundColor Green
        $name = Read-Host "Name "
    }
    if ($sv -eq "d") {
        $svr = "argenta.be"
    }
    elseif (!$sv) {
        $svr = "CORPARG.LAN"
    }
    else {
        Write-Host "Something in your search went wrong, look at your syntax highlighting!" -ForegroundColor Red
        break
    }
    Write-Host $svr -ForegroundColor DarkGreen
    $identifier = "Get-Groups $name"
    #query the groups
    $i = "*$name*"
    $q = Get-Adgroup -Filter {name -like $i} -Server $svr | Select-Object Name
   Out-Query $q $identifier
}

Function Get-MultipleGroups ($names) {
    while (!$names) { #Print-Input $name "Groep name "
        Write-Host  "Please enter the needed input..." -ForegroundColor Green
        $names = Read-Host "Names "
    }
    $oarr = New-Object System.Collections.ArrayList
    $identifier = 'Get-MultipleGroups'
    foreach ($name in $names -split ',') {
        $q = Get-Adgroup -Filter {name -like $name} | Select-Object Name
        if (!$q) {
        $q = "Could not find result for $i"
        }
        $oarr.Add($q) | Out-Null
    }
    Out-Query $oarr $identifier
}

Function Get-User ($in, $type <#, $multiple#>) {
    while (!$in) { #Print-Input $name "Groep name "
        Write-Host  "Please enter the needed input..." -ForegroundColor Green
        $in = Read-Host "Logon or Name "
    }
    while (!$type) { #Print-Input $name "Groep name "
        Write-Host  "Please enter the needed input..." -ForegroundColor Green
        $type = Read-Host "What data do you need? i/n/e "
    }
    $identifier = "Get-User $in"
    $iarr = $in -split ","
    $oarr = New-Object System.Collections.ArrayList
    foreach ($i in $iarr) {
        switch ($type) {
            i {
                $q = Get-Logon $i s
            }
            n {
                $q = Get-Name $i s
            }
            e {
                $q = Get-Email $i
            }
            t {
                $q = Get-InvertedLastName $i
            }
            f {
                $q = Get-Fulluserinfo $i
            }
            default {
                Write-Host "Error: Bad input!" -ForegroundColor Red
            }
            o {
                #organizational output
                $q = Get-OrganizationInfo $i
                Write-Host $q
                [Windows.Clipboard]::SetText($q.ToString())
            }
        }
        if (!$q) {
        $q = "Could not find result for $i"
        }
        $oarr.Add($q) | Out-Null
    }
    Out-Query $oarr $identifier
}

Function Get-GroupMembers ($group, $rec) {
    while (!$group) {
        Write-Host  "Please enter the needed input..." -ForegroundColor Green
        $group = Read-Host "AD Group "
    }
    $identifier = "Get-GroupMembers $group"
    switch ($rec) {
        r {
            $q = Get-Adgroupmember $group -Recursive | Select-Object Name
        }
        default {
            $q = Get-Adgroupmember $group | Select-Object Name
        }
    }
    Out-Query $q $identifier
}

Function Get-GroupMemberships ($logon, $rec) {
    while (!$logon) {
        Write-Host  "Please enter the needed input..." -ForegroundColor Green
        $logon = Read-Host "User login "
    }
    $identifier = "Get-GroupsMemberships $logon"
    switch ($rec) {
        r {
            $f = Read-Host "Name of the file?`nDefault: $logon.csv"
                if (!$f) {
                    Write-Host "Defaulting name."
                    $f = "$logon.csv"
                }
                Get-UserNestedGroups $logon $f
                $q = import-csv $f
                return $q.'Group Membership (All)' -split ';  ' | Out-Gridview -Title $identifier

        }
        m {
            $users = $logon -split ','
            $q = @()
            foreach ($user in $users) {
                $q += Get-AdUserMemberships $user
            }
        }
        default {
            Write-Host "Fetching users groups... "
            $q = Get-AdPrincipalGroupMembership $logon | Select-Object Name
        }
        t {
            $users = $logon -split ','
            $q = @()
            foreach ($user in $users) {
                $usernm = get-aduser -Filter 'samAccountName -like $user' | Select-Object Name
                $useraccess = Get-AdPrincipalGroupMembership $user | Select-Object Name
                $userobj = New-Object PSObject
                $userobj | Add-Member Noteproperty $usernm.name $useraccess.name
                $q += $userobj
            }
        }
    }
    Out-Query $q $identifier
}

#Clip

function Out-NewFolder ($path) {
    while (!$path) {
		Write-Host "Please enter the needed input..." -ForegroundColor Green
        $path = Read-Host "Folder name "
	}
	$output = "Naam Nieuwe AD groep(en) incl. domeinnaam (max. 15): Alles in CORPARG`nDLG_{0}_M`nGG_PROJECT_{0}_M`nNesting in groep(en): GG_PROJECT_{0}_M member maken van DLG_{0}_M`nDLG_{0}_M schrijfrechten geven op gedeelde folder GRPARG\#GRPARG\{0}" -f ($path)
    Write-Host $output -ForegroundColor Green
	Write-Output $output | clip
}

function Out-NewMailbox ($name, $group) {
    while (!$name) {
		Write-Host "Please enter the needed input..." -ForegroundColor Green
        $name = Read-Host "Mailbox name "
	}
    while (!$group) {
		Write-Host "Please enter the needed input..." -ForegroundColor Green
        $group = Read-Host "AD Group "
	}
    $output = "Aanmaken in exchange:`n{0}@argenta.be`nAanmaken in AD:`nCORPARG\DLG_MBX_{1}, `nOU IAM Groups - CORPARG\GG_PROJECT_MBX_{1}`nConfiguratie:`nDLG_MBX_{1} schrijfrechten geven op {0}@argenta.be`nGG_PROJECT_MBX_{1} member maken van DLG_MBX_{1}" -f ($name,$group)
    Write-Host $output -ForegroundColor Green
	Write-Output $output | clip
}

Function Out-Sft ($name) {
    while (!$name) {
		Write-Host "Please enter the needed input..." -ForegroundColor Green
        $name = Read-Host "Name "
	}
    #$name = Get-User $name n #lookup with logon instead of name
    $id = Get-Logon $name
    $email = Get-Email $id
    $output = "Pseudo Code:`nnew; $name; $email"
    Write-Host $output -ForegroundColor Green
    Write-Output $output | clip
}

Function Out-Commas ($i) {
    Set-Clipboard -Value (($i -split "`r`n") -join ',').ToString()
}

#Aliases

Set-Alias gg Get-Groups
Set-Alias gus Get-User
Set-Alias ggm Get-GroupMembers
Set-Alias gum Get-GroupMemberships

Set-Alias cnf Out-NewFolder
Set-Alias cnm Out-NewMailbox
Set-Alias sft Out-Sft
Set-Alias ocm Out-Commas
