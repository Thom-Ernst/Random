import-module ActiveDirectory

#Private

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
    Write-Host "Found $c." -ForegroundColor Green #How many items?
    Get-Input #Output dialogue
}

#Public

#Get



#Clip



#Aliases

Set-Alias gg Get-Groups
Set-Alias gu Get-User
Set-Alias ggm Get-GroupMembers
Set-Alias gum Get-UserMemberships
Set-Alias cnf Clip-NewFolder
Set-Alias cnm Clip-NewMailbox
Set-Alias sft Clip-Sft
