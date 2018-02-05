#Automatically put mailbox request in clipboard.
function Output-Mailboxname ($path) {
    $output = "DLG_MBX_{0}_M`nGG_PROJECT_MBX_{0}_M`nNesting in groep(en): GG_PROJECT_MBX_{0}_M member maken van DLG_MBX_{0}_M`n`nGelieve DLG_MBX_{0}_M te koppelen aan emailbox {0}@argenta.be" -f ($path)
    Write-Output $output | clip
}
Set-Alias mailbox Output-Mailboxname

#Automatically put sharefolder request in clipboard.
function Output-Folderpath ($path) {
    $output = "Naam Nieuwe AD groep(en) incl. domeinnaam (max. 15): Alles in CORPARG`nDLG_{0}_M`nGG_PROJECT_{0}_M`nNesting in groep(en): GG_PROJECT_{0}_M member maken van DLG_{0}_M`nDLG_{0}_M schrijfrechten geven op gedeelde folder GRPARG\#GRPARG\{0}" -f ($path)
    Write-Output $output | clip
}
Set-Alias folder Output-Folderpath

#System uptime without win8 powershell
((get-date) - ([Management.managementdatetimeconverter]::ToDateTime((get-wmiobject win32_operatingsystem).lastbootuptime)))
