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

Clear-Host
Write-Host "`nPossible matches for '$SCRIPT:search_string':" -BackgroundColor DarkGray
$array | Out-Host


# this do/until loop makes it so invalid input is ignored - the user can only input a real index number of the array
[int]$select = 0
Do{
    Try{ $select = Read-Host "Select a server" }
    Catch{
        # do nothing, including "do nothing" with any error messages
    }
} Until (($select -gt 0) -and ($select -le $array.Length))


# getting name and description from the selected server - they need to be at script level so we can use inside a function later
$SCRIPT:server_name = $array[$select-1].Name
$SCRIPT:server_desc = $array[$select-1].Description

# test that server is powered on and responsive
if(!(Test-Connection -ComputerName $SCRIPT:server_name -Count 2 -Quiet)){
    Write-Host "`n$SCRIPT:server_name ($SCRIPT:server_desc) did not respond to ping. Exiting." -BackgroundColor Black -ForegroundColor Red
    Start-Sleep 5
    EXIT
}

function Invoke-MainMenu{
    Clear-Host
    Write-Host "`nServer: $SCRIPT:server_name ($SCRIPT:server_desc)" -BackgroundColor DarkGray
    Write-Host "1. Show all services"
    Write-Host "2. Show running services"
    Write-Host "3. Show stopped services"
    Write-Host "4. Filter services by keyword"
    Write-Host "Q. Quit"

    [string]$select = Read-Host "`nSelect an option"
    while("1","2","3","4","q" -notcontains $select){ [string]$select = Read-Host "Select an option" }

    [int]$count = 0
    $SCRIPT:services = @()

    switch($select){
        '1'{
            Clear-Host
            Get-Service -ComputerName $SCRIPT:server_name | ForEach-Object{
                $count++
                $object = [PSCustomObject]@{
                    Num         = $count
                    Status      = $_.Status
                    Name        = $_.Name
                    DisplayName = $_.DisplayName
                }
                $SCRIPT:services += $object
            }
        }
        '2'{
            Clear-Host
            Get-Service -ComputerName $SCRIPT:server_name | Where-Object {$_.Status -eq 'Running'} | ForEach-Object{
                $count++
                $object = [PSCustomObject]@{
                    Num         = $count
                    Status      = $_.Status
                    Name        = $_.Name
                    DisplayName = $_.DisplayName
                }
                $SCRIPT:services += $object
            }
        }
        '3'{
            Clear-Host
            Get-Service -ComputerName $SCRIPT:server_name | Where-Object {$_.Status -eq 'Stopped'} | ForEach-Object{
                $count++
                $object = [PSCustomObject]@{
                    Num         = $count
                    Status      = $_.Status
                    Name        = $_.Name
                    DisplayName = $_.DisplayName
                }
                $SCRIPT:services += $object
            }
        }
        '4'{
            Clear-Host
            Write-Host
            $keyword = ""
            while($keyword -eq ""){ $keyword = Read-Host "Enter a search keyword" }

            Get-Service -ComputerName $SCRIPT:server_name | Where-Object {$_.DisplayName -like "*$keyword*"} | ForEach-Object{
                $count++
                $object = [PSCustomObject]@{
                    Num         = $count
                    Status      = $_.Status
                    Name        = $_.Name
                    DisplayName = $_.DisplayName
                }
                $SCRIPT:services += $object
            }
        }
        'q'{ EXIT }
    }
    $SCRIPT:services | Out-Host
}

Invoke-MainMenu