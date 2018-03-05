import-module ActiveDirectory

##########################################Private

#Function Print-Input ($input, $text = "Your input ") { ###Huge WIP
#    $i = $input
#    while (!$i) {
#        Write-Host  "Please enter the needed input..." -ForegroundColor Green
#        $i = Read-Host $text
#    }
#    return $i
#}

Function Print-Output ($query){
    Function Get-Input ($query = $query) {
        $r = (Read-Host "Export as csv? y/n").ToLower()
        switch ($r){ #y,n,default
            y {
                $f = Read-Host "Name of the file?"
                Write-Host "Exporting csv..."
                $query | Export-Csv -Path "$f.csv"
            }
            n {
                Write-Host "Outputting in a grid..."
                $query | Out-Gridview
            }
            default {
                Write-Host "Error: Bad input!" -ForegroundColor Red
                Get-Input
            }
        }
    }
    $c = $query.count
    if ($c -or $c -eq 0) {
        Write-Host "Found $c" -ForegroundColor Green #How many items?
    } else {
        Write-Host "Found 1" -ForegroundColor Green #One item returns customPSobject which is nullvalued
    }
    if ($c -ne 0) { #If nothing is found, no need to output anything
        Get-Input #Output dialogue
    }
}

Function Get-InvertedName ($name) { #put first name last or vice versa
    $namearr = $name -split ' '
    $lastname = $namearr[1..($namearr.count-1)] -join ' '
    return $lastname + ' ' + $namearr[0] 
}

Function Get-Name ($logon, $swap = '') { #add 's' to change to firstname name
    Write-Host "Getting username..."
    $q = Get-ADUser -Filter {SAMAccountName -Like $logon} | Select-Object Name #query for name from samaccountname
    $s = (($q -split '=')[1] -split '}')[0] #no more regex, just splitting on '=' and '}'
    switch ($swap) {
        s {
            Write-Host "Reversed"
            return Get-Invertedname ($s)
        }
        default {
            return $s
        }
    }
}

Function Get-Logon ($name) { #Input name firstname
    #query for samaccountname
    Write-Host "Getting userid..."
    $q = Get-ADUser -Filter {Name -Like $name} | Select-Object SAMAccountName
    return $q
}

Function Get-Email ($logon) { #get the email using logon, returns technical address, not named address.
    #query for email
    $q = Get-ADUser -Filter 
}

##########################################Public

#Get

Function Get-Groups ($name = '') {
    while ($name -eq '') { #Print-Input $name "Groep name "
        Write-Host  "Please enter the needed input..." -ForegroundColor Green
        $name = Read-Host "Name "
    }
    #query the groups
    $i = "*$name*"
    $q = Get-Adgroup -Filter {name -like $i} | Select-Object Name
   Print-Output $q
}

#Clip



#Aliases

Set-Alias gg Get-Groups
Set-Alias gu Get-User
Set-Alias ggm Get-GroupMembers
Set-Alias gum Get-UserMemberships
Set-Alias cnf Clip-NewFolder
Set-Alias cnm Clip-NewMailbox
Set-Alias sft Clip-Sft
