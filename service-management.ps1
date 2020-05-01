<#
PowerShell 5.1

Notes:
    For $matches variable, we don't want to declare variable type because it will be different if 1 match is found or if more-than-1 is found
#>
Clear-Host
Write-Host
Import-Module ActiveDirectory

# have 1 or more AD OUs listed in a text file named "ou-list.txt", in the current working directory, and they get imported here
[array]$ou_list = $(Get-Content .\ou-list.txt)

# create array containing all windows servers from each OU
$ou_list | ForEach-Object {
    [array]$win_servers += Get-ADComputer -SearchBase $_ -Filter {OperatingSystem -like "*Windows Server*"} -Properties OperatingSystem,Description
}

# search & save server name
[string]$search_string = ""
while($search_string -eq ""){
    [string]$search_string = Read-Host "Search servers by keyword (matches Name or Description)"
}

# get a list of servers that match the search string so user can choose later
$matches = $search_string | Where-Object {($_.Name -like "*$search_string*") -or ($_.Description -like "*$search_string*")}

$matches | Out-Host
