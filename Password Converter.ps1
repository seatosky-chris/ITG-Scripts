####################################################################
$APIKEy =  "<ITG API KEY>"
$APIEndpoint = "https://api.itglue.com"
$orgID = "<ITG ORG ID>"
$Limit = 50 # Will convert a max of X amount of passwords at a time
$Blacklist = @("Encryption Password", "Pre-Shared Key", "Security Code / Pin / Spoken Password", "Database Password") # Will not convert any passwords with this name (these should be password fields embedded in flexible assets e.g. a wifi password)
####################################################################

# Ensure they are using the latest TLS version
$CurrentTLS = [System.Net.ServicePointManager]::SecurityProtocol
if ($CurrentTLS -notlike "*Tls12" -and $CurrentTLS -notlike "*Tls13") {
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	Write-Host "This device is using an old version of TLS. Temporarily changed to use TLS v1.2."
}

# Settings IT-Glue logon information
Add-ITGlueBaseURI -base_uri $APIEndpoint
Add-ITGlueAPIKey $APIKEy
Write-Host "Configured the ITGlue API"

# Get all the existing passwords
$Passwords = Get-ITGluePasswords -organization_id $orgID -page_size 1000
if ($Passwords.meta.'total-count' -gt 1000) {
	$TotalPages = $Passwords.meta.'total-pages'
	for ($i = 2; $i -le $TotalPages; $i++) {
		$Passwords.data += (Get-ITGluePasswords -organization_id $orgID -page_size 1000 -page_number $i).data
	}
}

# Filter out general passwords, computer bios/local admin passwords, and anything in the blacklist
$Passwords.data = $Passwords.data | Where-Object { $_.attributes.'resource-id' -ne $null }
$Passwords.data = $Passwords.data | Where-Object { $_.attributes.name -notlike 'BIOS - *' -and $_.attributes.name -notlike 'Local Admin - *' }
$Passwords.data = $Passwords.data | Where-Object { $_.attributes.name -notin $Blacklist }

# Limit to $Limit passwords
$Passwords.data = $Passwords.data | Select-Object -First $Limit

function formatDestinationType($OriginalType) {
	if ($OriginalType -eq "ssl-certificates") {
		$FormattedType = "SSL Certificate"
	} else {
		$FormattedType = (Get-Culture).TextInfo.ToTitleCase(($OriginalType.trimend('s') -replace '-', ' ' -replace '_', ' '))
	}
	$FormattedType
	return
}

$ManualFixes = @()

$i = 0
$TotalPasswords = ($Passwords.data | Measure-Object).Count
Write-Progress -Activity 'Converting Passwords' -Status 'Starting' -PercentComplete 0
foreach ($Password in $Passwords.data) {
	$i++
	[int]$PercentComplete = $i / $TotalPasswords * 100
	Write-Progress -Activity "Converting Passwords" -PercentComplete $PercentComplete -Status ("Working - " + $PercentComplete + "%  (converting password '$($Password.attributes.name)')")

	$PasswordDetails = Get-ITGluePasswords -id $Password.id -show_password $true -include related_items
	$PAttributes = $PasswordDetails.data[0].attributes

	if ($PAttributes.'otp-enabled' -eq $true) {
		Write-Host "Manual Fix Required (OTP Enabled): $($PAttributes.name) - $($PAttributes.'resource-url')" -ForegroundColor Red
		$ManualFixes += [PSCustomObject]@{
			Password = $PAttributes.name
			Issue = "Manual Password Fix (OTP Enabled)"
			ToFix = $PAttributes.'resource-url'
			NewPassword = ''
		}
		continue;
	}

	# Create a replacement general password
	$PasswordAssetBody = 
	@{
		type = 'passwords'
		attributes = @{
			"name" = $PAttributes.name
			"username" = $PAttributes.username
			"password" = $PAttributes.password
			"url" = $PAttributes.url
			"notes" = $PAttributes.notes
			"password-category-id" = $PAttributes."password-category-id"
			"password-updated-at" = $PAttributes."password-updated-at"
			"restricted" = $PAttributes.restricted
			"archived" = $PAttributes.archived
			"password-folder-id" = $PAttributes."password-folder-id"
		}
	}

	try {
		$ReplacementPassword = New-ITGluePasswords -organization_id $orgID -data $PasswordAssetBody
		$Success = $true
	} catch {
		$Success = $false
	}

	# Add related items
	$RelatedItemsBody = @(
		@{
			type = 'related_items'
			attributes = @{
				'destination-id' = $PAttributes.'resource-id'
				'destination-type' = formatDestinationType -OriginalType $PAttributes.'resource-type'
				'notes' = "Password is for this asset."
			}
		}
	)
	$Related = $PasswordDetails.included | Where-Object { $_.type -eq 'related-items' }
	foreach ($RelatedItem in $Related) {
		if ($RelatedItem.attributes.'resource-id' -in $RelatedItemsBody.attributes.'destination-id') {
			continue;
		}
		$RelatedItemsBody += @{
			type = 'related_items'
			attributes = @{
				'destination-id' = $RelatedItem.attributes.'resource-id'
				'destination-type' = formatDestinationType -OriginalType $RelatedItem.attributes.'asset-type'
				'notes' = $RelatedItem.attributes.notes
			}
		}
	}
	try {
		New-ITGlueRelatedItems -resource_type 'passwords' -resource_id $ReplacementPassword.data.id -data $RelatedItemsBody | Out-Null
	} catch {
		$Success = $false
		Write-Host "Failed to add related items for password: $($PAttributes.name). Please manually fix this. Old: $($PAttributes.'resource-url') New: $($ReplacementPassword.data[0].attributes.'resource-url')" -ForegroundColor Red
		$ManualFixes += [PSCustomObject]@{
			Password = $PAttributes.name
			Issue = "Could not add related items"
			ToFix = $PAttributes.'resource-url'
			NewPassword = $ReplacementPassword.data[0].attributes.'resource-url'
		}
	}

	$Tags = $PasswordDetails.included | Where-Object { $_.type -eq 'tags' }

	if ($Tags) {
		Write-Host "The password '$($PAttributes.name)' is tagged from 1 or more other assets. Edit those assets directly to update to the new password. Asset(s) to fix: " -ForegroundColor Yellow
		foreach ($Tag in $Tags) {
			Write-Host $Tag.attributes.'resource-url'
		}
		$ManualFixes += [PSCustomObject]@{
			Password = $PAttributes.name
			Issue = "Tagged from other asset(s)"
			ToFix = $Tags.attributes.'resource-url' -join ' '
			NewPassword = $ReplacementPassword.data[0].attributes.'resource-url'
		}
	}

	# Delete the old password
	if ($Success) {
		Remove-ITGluePasswords -id $Password.id
		Write-Host "Replaced password: $($PAttributes.name)" -ForegroundColor Green
	} else {
		Write-Host "!!!! FAILED to convert password: $($PAttributes.name)" -ForegroundColor Red
	}
}
Write-Progress -Activity 'Converting Passwords' -Status 'Completed' -PercentComplete 100

# Show any manual fixes that are required
if ($ManualFixes) {
	$ManualFixes | Out-GridView -Title "Manual Fixes are required" -PassThru
}
