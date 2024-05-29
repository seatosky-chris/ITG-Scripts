###
# File: \Password Converter copy.ps1
# Project: Scripts
# Created Date: Friday, November 3rd 2023, 11:40:33 am
# Author: Chris Jantzen
# -----
# Last Modified: Tue May 28 2024
# Modified By: Chris Jantzen
# -----
# Copyright (c) 2023 Sea to Sky Network Solutions
# License: MIT License
# -----
# 
# HISTORY:
# Date      	By	Comments
# ----------	---	----------------------------------------------------------
###

####################################################################
$APIKEy =  "<ITG API KEY>"
$APIEndpoint = "https://api.itglue.com"
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

$ITGlueCompanies = Get-ITGlueOrganizations -page_size 1000

if ($ITGlueCompanies -and $ITGlueCompanies.data) {
	$ITGlueCompanies = $ITGlueCompanies.data | Where-Object { $_.attributes.'organization-status-name' -eq "Active" }
}

if (!$ITGlueCompanies) {
	exit
}

$AllConverted = [System.Collections.ArrayList]::new()
$AllManualFixes = [System.Collections.ArrayList]::new()

foreach ($Company in $ITGlueCompanies) {
	Write-Host "Auditing: $($Company.attributes.name)" -ForegroundColor Green
	# Get all the existing passwords
	$Passwords = Get-ITGluePasswords -organization_id $Company.id -page_size 1000
	if ($Passwords.meta.'total-count' -gt 1000) {
		$TotalPages = $Passwords.meta.'total-pages'
		for ($i = 2; $i -le $TotalPages; $i++) {
			$Passwords.data += (Get-ITGluePasswords -organization_id $Company.id -page_size 1000 -page_number $i).data
		}
	}

	if (!$Passwords) {
		continue
	}

	# Filter out general passwords, computer bios/local admin passwords, and anything in the blacklist
	$Passwords.data = $Passwords.data | Where-Object { $_.attributes.'resource-id' -ne $null }
	$Passwords.data = $Passwords.data | Where-Object { $_.attributes.'resource-type' -eq 'contacts' }
	$Passwords.data = $Passwords.data | Where-Object { $_.attributes.name -notlike 'BIOS - *' -and $_.attributes.name -notlike 'Local Admin - *' }
	$Passwords.data = $Passwords.data | Where-Object { $_.attributes.name -notin $Blacklist }

	# Limit to $Limit passwords
	$TotalEmbeddedCount = ($Passwords.data | Measure-Object).Count
	if ($TotalEmbeddedCount -lt 1) {
		continue
	}
	$Passwords.data = $Passwords.data | Select-Object -First $Limit
	$ToChangeEmbeddedCount = ($Passwords.data | Measure-Object).Count
	if ($ToChangeEmbeddedCount -lt $TotalEmbeddedCount) {
		Write-Host "$TotalEmbeddedCount embedded contact passwords found but only updating $ToChangeEmbeddedCount due to the set limit."
	}

	function formatDestinationType($OriginalType) {
		if ($OriginalType -eq "ssl-certificates") {
			$FormattedType = "SSL Certificate"
		} else {
			$FormattedType = (Get-Culture).TextInfo.ToTitleCase(($OriginalType.trimend('s') -replace '-', ' ' -replace '_', ' '))
		}
		$FormattedType
		return
	}

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
			$AllManualFixes.Add([PSCustomObject]@{
				Company = $Company.attributes.name
				Password = $PAttributes.name
				Issue = "Manual Password Fix (OTP Enabled)"
				ToFix = $PAttributes.'resource-url'
				NewPassword = ''
			})
			$AllConverted.Add([PSCustomObject]@{
				Company = $Company.attributes.name
				OldID = $Password.id
				NewID = ""
				Name = $PAttributes.name
				OldLink = $PAttributes.'resource-url'
				NewLink = ""
				Type = "Manual Password Fix (OTP Enabled)"
			})
			continue;
		}

		$NewName = $PAttributes.name
		if ($NewName -notlike "*$($PAttributes.username)*" -and $NewName -notlike "*$($PAttributes.'cached-resource-name')*") {
			if ($NewName -notlike "*-*" -and ($NewName -like "*AD*" -or $NewName -like "*O365*" -or $NewName -like "*0365*" -or $NewName -like "*M365*" -or 
				$NewName -like "*Email*" -or $NewName -like "*Azure AD*" -or $NewName -like "*Office 365*" -or $NewName -like "*Active Directory*" -or $NewName -like "*Azure AD*")) 
			{
				$NewName = $NewName + " - $($PAttributes.'cached-resource-name')"
			} else {
				$NewName = $NewName + " ($($PAttributes.'cached-resource-name'))"
			}
		}

		if ($PAttributes.name -like $NewName) {
			$LogName = $PAttributes.name
		} else {
			$LogName = "$($PAttributes.name) (New Name: $NewName)"
		}

		# Create a replacement general password
		$PasswordAssetBody = 
		@{
			type = 'passwords'
			attributes = @{
				"name" = $NewName
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
			$ReplacementPassword = New-ITGluePasswords -organization_id $Company.id -data $PasswordAssetBody
			if ($ReplacementPassword -and $ReplacementPassword.data -and ($ReplacementPassword.data | Measure-Object).Count -gt 0) {
				$Success = $true
			}
		} catch {
			$Success = $false
		}

		# Checked for tagged assets
		if ($Success) {
			$Tags = $PasswordDetails.included | Where-Object { $_.type -eq 'tags' }

			if ($Tags) {
				Write-Host "The password '$($PAttributes.name)' is tagged from 1 or more other assets. Edit those assets directly to update to the new password. Asset(s) to fix: " -ForegroundColor Yellow
				foreach ($Tag in $Tags) {
					Write-Host $Tag.attributes.'resource-url'
				}
				$AllManualFixes.Add([PSCustomObject]@{
					Company = $Company.attributes.name
					Password = $PAttributes.name
					Issue = "Tagged from other asset(s)"
					ToFix = $Tags.attributes.'resource-url' -join ' '
					NewPassword = $ReplacementPassword.data[0].attributes.'resource-url'
				})
			}
		}

		# Add related items
		if ($Success) {
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
				$AllManualFixes.Add([PSCustomObject]@{
					Company = $Company.attributes.name
					Password = $PAttributes.name
					Issue = "Could not add related items"
					ToFix = $PAttributes.'resource-url'
					NewPassword = $ReplacementPassword.data[0].attributes.'resource-url'
				})

				$AllConverted.Add([PSCustomObject]@{
					Company = $Company.attributes.name
					OldID = $Password.id
					NewID = $ReplacementPassword.data[0].id
					Name = $LogName
					OldLink = $PAttributes.'resource-url'
					NewLink = $ReplacementPassword.data[0].attributes.'resource-url'
					Type = "Auto Converted - Failed to Delete Old and Update Related Items"
				})
			}
		}

		if ($Success) {
			# Delete the old password
			Remove-ITGluePasswords -id $Password.id
			Write-Host "Replaced password: $($LogName)" -ForegroundColor DarkCyan

			$AllConverted.Add([PSCustomObject]@{
				Company = $Company.attributes.name
				OldID = $Password.id
				NewID = $ReplacementPassword.data[0].id
				Name = $LogName
				OldLink = $PAttributes.'resource-url'
				NewLink = $ReplacementPassword.data[0].attributes.'resource-url'
				Type = "Auto Converted"
			})
		} else {
			Write-Host "!!!! FAILED to convert password: $($PAttributes.name)" -ForegroundColor Red
		}
	}
	Write-Progress -Activity 'Converting Passwords' -Status 'Completed' -PercentComplete 100

	# Show any manual fixes that are required
	$ManualFixes = $AllManualFixes | Where-Object { $_.Company -eq $Company.id }
	if ($ManualFixes) {
		$ManualFixes | Out-GridView -Title "Manual Fixes are required"
	}

	Read-Host "Completed '$($Company.attributes.name)', press any key to continue..."
}

if ($AllConverted) {
	$AllConverted | Export-Csv -Path "./PasswordsConverted-All.csv" -NoTypeInformation
}
if ($AllManualFixes) {
	$AllManualFixes | Export-Csv -Path "./PasswordsConverted-ManualFixes.csv" -NoTypeInformation
}
