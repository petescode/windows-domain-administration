<#
Author: Pete Wood
Purpose: Connect to remote Active Directory Windows Server and start/stop/restart a service
Notes:
    Developed using PowerShell 5.1

    In the same directory there should be a text file in the naming convention "name-of-script-INFO.txt"
        It should have Active Directory Organizational Unit paths like this:
            OU=Windows,OU=Servers,OU=Canada,DC=FABRIKAM,DC=COM
            OU=Windows,OU=Servers,OU=India,DC=FABRIKAM,DC=COM
            OU=Windows,OU=Servers,OU=Australia,DC=FABRIKAM,DC=COM

    For $matches variable, we don't want to declare the type because it will be different if 1 match is found or if more-than-1 is found

    Use $SCRIPT: instead of $GLOBAL: to avoid the variables persisting in the shell after the script exits

    In Invoke-ServiceMenu...
        The use of -PassThru and Out-Host on the same line was necessary to get the shell output desired

    Although Invoke-MainMenu and Invoke-ServiceMenu reference each other neither function is actually called until both have already been declared
#>
Clear-Host
Write-Host
Import-Module ActiveDirectory

function Get-SearchString{
    [string]$SCRIPT:search_string = ""
    Do{
        [string]$SCRIPT:search_string = Read-Host "Search servers by keyword (matches Name or Description)"
    } Until ($SCRIPT:search_string -ne "")
}

[string]$info_file = $MyInvocation.MyCommand.Name.Split(".")[0] + "-INFO.txt"
[array]$ou_list = $(Get-Content .\$info_file)

# create array containing all windows servers from each OU
$ou_list | ForEach-Object {
    [array]$win_servers += Get-ADComputer -SearchBase $_ -Filter {OperatingSystem -like "*Windows Server*"} -Properties OperatingSystem,Description
}

Get-SearchString

# get a list of servers that match the search string so user can choose later
$matches = $win_servers | Where-Object {($_.Name -like "*$SCRIPT:search_string*") -or ($_.Description -like "*$SCRIPT:search_string*")} | Sort-Object Name

# catch if no matches are found, give another opportunity to search
while ($NULL -eq $matches){
    Write-Host "`nNo matches for '$SCRIPT:search_string' found! Try again." -BackgroundColor Black -ForegroundColor Yellow; Start-Sleep 3
    Clear-Host; Write-Host
    Get-SearchString
    $matches = $win_servers | Where-Object {($_.Name -like "*$SCRIPT:search_string*") -or ($_.Description -like "*$SCRIPT:search_string*")} | Sort-Object Name
}

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
    Catch{} # do nothing, including "do nothing" with any error messages
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

# declare function for the main menu (sorting through services)
function Invoke-MainMenu{
    Clear-Host
    Write-Host "`nServer: $SCRIPT:server_name ($SCRIPT:server_desc)`n" -BackgroundColor DarkGray
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
            Do{
                Clear-Host
                Write-Host
                $keyword = ""
                while($keyword -eq ""){ $keyword = Read-Host "Enter a search keyword" }

                $test = Get-Service -ComputerName $SCRIPT:server_name | Where-Object {$_.DisplayName -like "*$keyword*"}
                if($NULL -eq $test){
                    Write-Host "`nNo services match the keyword! Try again." -BackgroundColor Black -ForegroundColor Yellow
                    Start-Sleep 3
                }
            } Until ($NULL -ne $test) # test is not null (a service was found that matched keyword)

            $test | ForEach-Object{
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

    # now make a selection of which service you want to act on
    [int]$select = 0
    Do{
        Try{ $select = Read-Host "Select a service" }
        Catch{} # do nothing, including "do nothing" with any error messages
    } Until (($select -gt 0) -and ($select -le $SCRIPT:services.Length))

    # select the service and save to a variable we can use later
    $SCRIPT:the_service = Get-Service -Name $SCRIPT:services[$select-1].Name -ComputerName $SCRIPT:server_name

    # start the next menu!
    Invoke-ServiceMenu
}


# menu for what to do with selected service; allow option to go back to main menu
function Invoke-ServiceMenu{
    Clear-Host
    Write-Host "`n$($SCRIPT:the_service.Name) @ $SCRIPT:server_name`n" -BackgroundColor DarkGray
    Write-Host "1. Start this service"
    Write-Host "2. Stop this service"
    Write-Host "3. Restart this service"
    Write-Host "4. Go back to previous menu"
    Write-Host "Q. Quit"
    Write-Host
    [string]$select = Read-Host "Select an option"
    while("1","2","3","4","q" -notcontains $select){ [string]$select = Read-Host "Select an option" }

    # option 1 uses a different method than the rest just to show this is possible - either way works
    switch($select){
        '1'{ $SCRIPT:the_service | Start-Service -PassThru -Verbose | Out-Host }
        '2'{ Stop-Service -InputObject $SCRIPT:the_service -PassThru -Verbose -Force | Out-Host }
        '3'{ Restart-Service -InputObject $SCRIPT:the_service -PassThru -Verbose -Force | Out-Host }
        '4'{ Invoke-MainMenu }
        'q'{ EXIT }
    }
}

Invoke-MainMenu

Read-Host "`nPress any key to exit"