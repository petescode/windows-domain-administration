<#
PowerShell 5.1

Notes:
    ou-list.txt should have Active Directory Organizational Unit paths like this:
        OU=Windows,OU=Servers,OU=Canada,DC=FABRIKAM,DC=COM
        OU=Windows,OU=Servers,OU=India,DC=FABRIKAM,DC=COM
        OU=Windows,OU=Servers,OU=Australia,DC=FABRIKAM,DC=COM

    For $matches variable, we don't want to declare variable type because it will be different if 1 match is found or if more-than-1 is found

    Use $SCRIPT: instead of $GLOBAL: to avoid the variables persisting in the shell after the script exits
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

# move this to top of script?
function Get-SearchString{
    [string]$SCRIPT:search_string = ""
    Do{
        [string]$SCRIPT:search_string = Read-Host "Search servers by keyword (matches Name or Description)"
    } Until ($SCRIPT:search_string -ne "")
}

# get a list of servers that match the search string so user can choose later
$matches = $win_servers | Where-Object {($_.Name -like "*$SCRIPT:search_string*") -or ($_.Description -like "*$SCRIPT:search_string*")} | Sort-Object Name

#$matches | Out-Host

# catch if no matches are found, give another opportunity to search
while ($NULL -eq $matches){
    Write-Host "`nNo matches for '$SCRIPT:search_string' found! Try again." -BackgroundColor Black -ForegroundColor Yellow; Start-Sleep 3
    Clear-Host; Write-Host
    Get-SearchString
    $matches = $win_servers | Where-Object {($_.Name -like "*$SCRIPT:search_string*") -or ($_.Description -like "*$SCRIPT:search_string*")} | Sort-Object Name
}

#$matches | Out-Host

# re-sorting $matches into a new array allows us to add a number column to be used later for selection
$array = @()
[int]$count = 0
$matches | ForEach-Object{
    $count++
    $object = [PSCustomObject]@{
        Num = $count
        Name = $_.Name
        Description = $_.Description
    }
    $array += $object
}
$array | Out-Host