### Record the time the script started
$starttime = get-date

### Username of whoever is running script
$username = [Environment]::UserName

### Ask the user for the Job Title of the Template they want to create
# $Job_Title =  Read-Host 'Job Title'
$Job_Title = 'Staff Nurse Anesthetist'

### Ask the user for the Location of the template they want to create
# $location = Read-Host 'Location'
$location = ''

### An empty array to place the matching users in
$matchingusers = @()

### Import list of Dead AD Groups
$DeadGroups = Import-Csv -path "C:\Users\$username\Beaumont Health\ITOS - General\AD Group Project\Input Files\Dead AD Groups.csv"

### Import the All System Users CSV
Write-Host 'Importing All System Users...'
$allsystemusers = Import-Csv -Path "C:\Users\$username\Beaumont Health\ITOS - General\AD Group Project\Input Files\AllSystemUsers.csv"

### A Timestamp to capture how long it took to do the imports
$postimporttime = Get-Date

### Find matching users and add them to the Matching Users Array
Write-Host 'Finding Matching users...'
$a = 0
foreach ($user in $allsystemusers) {
    if ($user.jobTitle -eq $Job_Title -and $user.officelocation -like "*$location*") {
        $matchingusers += New-Object PSobject -Property @{
            UserID = $user.accountName
            JobTitle = $user.jobTitle
            Location = $user.officelocation
        }
        $a += 1
    }
}

write-host '...Matching users:' $a

### an empty array for the AD groups of matching users
$matchinggroups = @()

### Sets the number of users to use
$numusers = $matchingusers.Length

### Roughly estimates the amount of time to finish
# $minutes = $numusers * 1.44 / 60
$minutes = (5 + ($numusers * 1.47)) / 60

### Fetch the AD groups for each user
write-host "Fetching the Groups of $numusers users..."
Write-Host 'Estimated time to completion:' $minutes "Minutes"
$b = 0
Foreach ($match in $matchingusers) {
    if ($b -lt $numusers) {
        $matchinggroups += Get-ADPrincipalGroupMembership -identity $match.UserID | Select-Object -Property samaccountname
        $b += 1
        write-host "$b of $numusers users added"
    }
}

### sets the percentage of users a group has to be part of
$majority = $numusers * 0.8

### Groups the AD groups and keeps the ones in the majority of users
$groups = $matchinggroups | Group-Object -Property samaccountname | Where-Object -Property count -GE $majority

### Removes any Dead/Automated groups
$Sortedgroups = Compare-Object -ReferenceObject $groups -DifferenceObject $DeadGroups -Property name | Where-Object sideindicator -EQ "<="

### display the Groups in the console
Write-Host 'All Matching Groups'
$groups
Write-Host ''
write-host 'Matching Groups Minus Dead/Automated Groups'
$Sortedgroups | Format-Table

### The empty string to place the template in
$template = ''

### adds the groups to the template string
foreach ($group in $Sortedgroups) {
    $template += $group.name + ';'
}
$template = $template.TrimEnd(';')

$newtemplate += New-Object psobject -Property @{
    JobTitle = $Job_Title
    Location = $location
    Template = $template
}

### output the new template to the Console and to a file
write-host 'New Template'
$template
$newtemplate | select-Object jobTitle, location, Template | Export-Csv -Path "C:\Users\$username\Beaumont Health\ITOS - General\AD Group Project\Output Files\New Template.csv"
set-clipboard -value $template

### Record the time the script ended
$endtime = get-date

### Output the start amd end time of the script
$starttime | Select-Object -Property DateTime
$postimporttime | Select-Object -Property DateTime
$endtime | Select-Object -Property DateTime