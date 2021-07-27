#Requires -RunAsAdministrator

# See if they want to check against O365
Write-Host "Would you like to check against Office 365 as well? (Default is No)" -ForegroundColor 'Yellow'
$Readhost = Read-Host " (Y / N)"
$CheckO365 = $false
Switch ($Readhost) {
	Y {$CheckO365 = $true}
	default {$CheckO365 = $false}
}

# Check against O365
if ($CheckO365) {
	Connect-AzureAD
	Connect-MsolService

	if (!Get-MsolUser) {
		Write-Host "Could not connect to O365." -ForegroundColor "red"
		$CheckO365 = $false
	}
}

# Get the csv file with the usernames
Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
	InitialDirectory = [Environment]::GetFolderPath('Desktop') 
	Filter = "CSV (*.csv) | *.csv"
}
$OpenFileDialog.ShowDialog() | Out-Null
$UsersCSV = Import-CSV $OpenFileDialog.FileName

# Check against AD
$ADMatches = @()
$NoMatch = @()
$i = 0
foreach ($User in $UsersCSV) {
	$FullName = $User.FullName -replace '[^a-zA-Z0-9 ]', ''
	$Username = $User.Username -replace '[\/\\\[\]\:;\|=,\+\*\?<>\s]', ''
	$Email = $User.Email -replace '\s',''
	$Filters = @()
	if ($FullName) { $Filters += "Name -like '*$FullName*'" }
	if ($Username) { $Filters += "SamAccountName -like '*$Username*'" }
	if ($Email) { $Filters += "EmailAddress -eq '$Email'" }

	$ADMatch = $false
	if ($Filters) {
		$ADMatch = Get-ADUser -Filter ($Filters -join " -or ") -Properties * | 
			Select-Object -Property @{Name="CSVOrder"; E={$i}}, Name, @{Name="Status"; E={$false}}, @{Name="Username"; E={$_.SamAccountName}}, EmailAddress, Enabled, 
									Description, LastLogonDate, @{Name="OU"; E={[regex]::matches($_.DistinguishedName, '\b(OU=)([^,]+)')[0].Groups[2]}}, 
									City, Department, Division, Title
	}
	
	if ($ADMatch) {
		if ($ADMatch.Name -like '*Disabled*' -or $ADMatch.Enabled -eq $false -or $ADMatch.OU -like '*Disabled*') {
			$ADMatch.Status = 'Disabled'
		} elseif ($ADMatch.Description -like '*Disabled*') {
			$ADMatch.Status = 'Improperly Disable'
		} elseif ($ADMatch.LastLogonDate -and $ADMatch.LastLogonDate -lt (Get-Date).AddDays(-150)) {
			$ADMatch.Status = 'Maybe Disable?'
		} else {
			$ADMatch.Status = 'Enabled'
		}
		$ADMatches += $ADMatch
	} else {
		$NoMatch += [pscustomobject]@{FullName = $User.FullName; Username = $User.Username; Email = $User.Email}
	}
	$i += 1
}

# Display a grid with all found users and their properties
$ADMatches | Sort-Object -Property CSVOrder -Unique | Out-GridView -Title "AD Matches"

# Write users who weren't found at all
Write-Host "The following users were not found: " -ForegroundColor 'black' -BackgroundColor 'red'
$NoMatch | Format-Table 