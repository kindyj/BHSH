### Get the Time Script Started
$starttime = Get-Date

### Username of whoever is running script
$username = [Environment]::UserName

### The path of the New Hires file
$newhirepath = "C:\Users\$username\Beaumont Health\ITOS - General\AD Group Project\Input Files\New Hires.csv"

### import the New Hires spreadsheet, be sure to add a column to the sheet called login with their login ID
$newhires = Import-Csv -path $newhirepath | Select-Object 'Login', 'job title', 'Location Descr'

### The path of the Templates file
$templatespath = "C:\Users\$username\Beaumont Health\ITOS - General\AD Group Project\Input Files\Job Titles and AD Groups Master Spreadsheet.csv"

### import the templates CSV
$templates = Import-Csv -path $templatespath | Select-Object 'groups', 'JobTitle', 'Location'
# $split = $templates.Groups -split ';' -replace ' ','' | Group-Object | Select-Object -Property name 

### a counter to loop through the templates
$i = 0

### a Counter to display how many users have been processed
$f = 0

### the output arrays for missing groups and templates
$MissingTemplates = @()
$missinggroups = @()

### Estimates the amount of time to complete the script
$minutes = 1.34 * $newhires.Length / 60 
Write-Host 'Estimated time to completion:' $minutes 'Minutes'

### Iterate over the new hire sheet
foreach ($newhire in $newhires) { 
    $f += 1
    write-host "User:" $newhire.Login $f "of" $newhires.Length
    
    ### fetches the user's current groups from AD
    $groups = Get-ADPrincipalGroupMembership -identity $newhire.login | select-object -property samaccountname | Group-Object -Property samaccountname | select-Object -Property name

    ### tries to find a matching template based on job title and location
    while ($i -le $templates.Length) {

        if ($newhire.'Job Title' -eq $templates[$i].JobTitle -and $newhire.'Location Descr' -eq $templates[$i].Location) {
            write-host 'Template Found'
            
            ### formats the templates to be manipulated
            $template = $templates[$i].Groups -split ';' -replace ' ;','' | Group-Object | Select-Object -Property name
            
            ### compares the user's current groups and the matching template
            $comparegroups = Compare-Object -ReferenceObject $template -DifferenceObject $groups -Property name | Where-Object sideindicator -EQ "<="
            
            ### breaks the while loop
            $i = $templates.Length + 1

            ### an empty string to place the missing groups in
            $missing = ''

            ### adds the missing groups to the missing string
            foreach ($group in $comparegroups) {
                $Missing += $group.name + ';'
            }
            
            ### adds the user's login id and the missing groups to an array
            if ($missing -ne '') {
                Write-Host 'AD Groups Missing:'
                $comparegroups | Format-Table

                $missinggroups += New-Object psobject -Property @{
                    LoginID = $newhire.Login
                    JobTitle = $newhire.'Job Title'
                    Location = $newhire.'Location Descr'
                    Groups = $missing
                }
            }
        }

        ### if the template does not match, adds to the counter
        else {
            $i += 1

            ### if no template matches the user, add the user's job title and location to an array
            if ($i -ge $templates.Length) {
                Write-host 'No Template Found'
                $i += 1
                $MissingTemplates += New-Object psobject -Property @{
                    NewHire = $newhire.Login
                    JobTitle = $newhire.'Job Title'
                    Location = $newhire.'Location Descr'
                }
            }
        }
    }
    ### resets the counter
    $i = 0
    Write-Host '--------------------------------------------------'
}

### Checks if the user is disabled, if the user is disabled, add them to the DisabledUsers Array
$missinggroups2 = @()
$disabledusers = @()
write-host "Checking if user's are disabled"
foreach ($missinggroup in $missinggroups) {
    $enabled = get-aduser -identity $missinggroup.LoginID | Select-Object Enabled
    if ($enabled -like "*True*") {
        $missinggroups2 += New-Object psobject -Property @{
            LoginID = $missinggroup.LoginID
            JobTitle = $missinggroup.jobTitle
            Location = $missinggroup.location
            Groups = $missinggroup.groups
        }
        write-host "User" $missinggroup.loginID "is not disabled"
    }
    else {
        $disabledusers += New-Object psobject -Property @{
            LoginID = $missinggroup.LoginID
        }
        write-host "User $missinggroup.loginID IS Disabled"
    }
}

### Display Disabled Users and export the users to a file
Write-Host 'Disabled users:'
$disabledusers | Format-Table
Remove-Item -path "C:\Users\$username\Beaumont Health\ITOS - General\AD Group Project\Output Files\Disabled Users.csv"
$disabledusers | export-csv -path "C:\Users\$username\Beaumont Health\ITOS - General\AD Group Project\Output Files\Disabled Users.csv"

### display the missing templates in the console
Write-Host 'Missing Templates:'
$MissingTemplates | Format-Table

### deletes the Missing template file and re-creates it with missing templates
Remove-Item -path "C:\Users\$username\Beaumont Health\ITOS - General\AD Group Project\Output Files\Missing Templates.csv"
$missingtemplates | Select-Object jobTitle, location, NewHire | export-csv -Path "C:\Users\$username\Beaumont Health\ITOS - General\AD Group Project\Output Files\Missing Templates.csv" -NoTypeInformation

### line break
Write-Host ''

### Displays the Missing groups in the console
Write-Host 'Missing Groups'
$missinggroups2 | Select-Object LoginID, jobTitle, location, groups | Format-Table

### deletes the Missing group file and re-creates it with missing groups
Remove-Item -path "C:\Users\$username\Beaumont Health\ITOS - General\AD Group Project\Output Files\Missing Groups.csv"
$missinggroups2 | Select-Object LoginID, jobTitle, location, groups | export-csv -Path "C:\Users\$username\Beaumont Health\ITOS - General\AD Group Project\Output Files\Missing Groups.csv" -NoTypeInformation

### Get the time the script ended
$endtime = Get-Date

### Display the start and End times
Write-Host 'Estimated time:' $minutes 'Minutes'
Write-Host 'Actual Time:'
$starttime
$endtime