Set-Alias ns neoscmd

function Delete-LocalProfiles ($uid, $hostname){
    ns userprofiles /Del:INT\$uid /s:$hostname
}
Set-Alias dlp Delete-LocalProfiles

function Set-DefaultDomain ($hostname){
    ns SetDefaultLogonDomain /s:$hotsname
}
Set-Alias sdm Set-DefaultDomain

function Remove-PrinterRefs ($hostname){
    ns RemovePDFPrinterReferences /s:$hostname
}
Set-Alias rpr Remove-PrinterRefs

function Update-Printers ($hostname) {
    ns PrintersUpdateDrivers /S:$hostname
}
Set-Alias udp Update-Printers

function Get-Startmenu ($hostname) {
    ns RefreshStartMenuIcons /S:$hostname
}

Set-Alias gsm Get-Startmenu

function Neos-Reboot ($hostname) {
    ns Reboot /s:$hostname
}
Set-Alias nrb Neos-Reboot

function Find-Group ($uid, $group) {
    groupview -gm -u INT\$uid | select-string "$group"
}
Set-Alias fg Find-Group

function Remote-Assist ($ip){
    msra /offerra $ip
}
Set-Alias ras Remote-Assist

Set-Alias pn ping
function Run-Pingt ($adress){
    pn -t $adress
}
Set-Alias pt Run-Pingt
