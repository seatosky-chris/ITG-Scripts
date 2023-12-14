 ###
# File: \Password Cleanup - Automatic.ps1
# Project: Scripts
# Created Date: Tuesday, July 11th 2023, 3:50:36 pm
# Author: Chris Jantzen
# -----
# Last Modified: Fri Dec 08 2023
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
$ITGAPIKey = @{
	Url = "https://api.itglue.com"
	Key = ""
}
# The global APIKey for the email forwarder. The key should give access to all organizations.
$Email_APIKey = @{
	Url = ""
	Key = ""
}

$EmailFrom = @{
	Email = ''
	Name = "Password Cleanup"
}
$EmailTo = @(
	@{
		Email = ''
		Name = ""
	}
)

$PasswordCategoryIDs = @{
	AD = 41693
	ADAdmin = 179527
	AzureAD = 369671
	Email = 244224
	EmailAdmin = 369672
	OldEmail = 179318
	OldEmailAdmin = 41685
	LocalUser = 41689
	LocalAdmin = 163059
}

$ITG_ADFlexAsset = "Active Directory"
$ITG_EmailFlexAsset = "Email"

$PrefixTypes = @{
	AD = @("AD", "Active Directory")
	O365 = @("AAD", "O365", "0365", "M365", "Email", "Azure AD", "AzureAD", "O365 Email", "Office 365", "Office365", "Microsoft 365")
	LocalUser = @("Local", "Local User")
	LocalAdmin = @("Local Admin")
}
$AllPrefixTypes = $PrefixTypes.GetEnumerator() | ForEach-Object { $_.Value }
$AllLocalPrefixTypes = $PrefixTypes.GetEnumerator() | Where-Object { $_.Name -like "Local*" } | ForEach-Object { $_.Value }

$PasswordCatTypes_ForMatching = @{
	ByContactName = @(
		"Active Directory", "Active Directory - Administrator", "Active Directory - Service Account", "Active Directory - Vendor", "Application / Software",
		"Azure AD", "Configurations - Local Account (Workstation)", "Email Account", "Microsoft 365", "Network - Remote Access", "Office 365", "Other - Apple ID",
		"Vendor"
	)
	ByContactEmail = @(
		"Active Directory", "Active Directory - Vendor", "Application / Software", "Azure AD", "Cloud Management / Licensing Portal", "Email Account", "Microsoft 365",
		"Microsoft 365 - Global Admin", "Network - Remote Access", "Office 365", "Other - Apple ID"
	)
	ByConfigName = @(
		"Application / Software", "Application / Software - Office Key", "BDR - Backup & Disaster Recovery", "Configurations - BIOS", "Configurations - BitLocker",
		"Configurations - DVR & CCTV", "Configurations - FileVault", "Configurations - Local Account (Workstation)", "Configurations - Local Admin (Workstation / Server)",
		"Configurations - Other", "Configurations - Printers", "Configurations - Server / VM Management", "Network - File Share / FTP", "Network - Infrastructure Equipment", 
		"Network - Other Devices", "Network - Printing / Scanning", "Network - Remote Access", "Network - Storage"
	)
} # These password category types will be matched with contacts (list categories by name)
$AllPasswordCatTypes_ForMatching = $PasswordCatTypes_ForMatching.GetEnumerator() | ForEach-Object { $_.Value } | Sort-Object -Unique

$ContactTypes_Employees = @(
	"Approver", "Champion", "Contractor", "Decision Maker", "Employee", "End User", "External User"
	"Employee - Email Only", "Employee - Part Time", "Employee - Temporary", "Employee - Multi User",
	"Influencer", "Internal IT", "Management", "Owner", "Shared Account", "Terminated", "Employee - On Leave",
	"Internal / Shared Mailbox"
)
$ContactTypes_Service = @("Service Account")
$ContactTypes_Vendor = @("Vendor Support")
$ConfigurationTypes_Primary = @("Desktop", "Laptop", "Server", "Workstation") # Workstation and server types, the types most passwords will be associated with
$ConfigurationTypes_Backup = @("Backup Device", "Datto Device")
$ConfigurationTypes_Camera = @("Camera", "Camera System")
$ConfigurationTypes_Printer = @("Printer")
$ConfigurationTypes_Infrastructure = @("Firewall", "Network Device", "Router", "Switch")
$ConfigurationTypes_Storage = @("NAS")
####################################################################

$ScriptPath = $PSScriptRoot

### This code is common for every company and can be ran before looping through multiple companies
$CurrentTLS = [System.Net.ServicePointManager]::SecurityProtocol
if ($CurrentTLS -notlike "*Tls12" -and $CurrentTLS -notlike "*Tls13") {
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	Write-Output "This device is using an old version of TLS. Temporarily changed to use TLS v1.2."
	Write-PSFMessage -Level Warning -Message "Temporarily changed TLS to TLS v1.2."
}

# Setup logging
If (Get-Module -ListAvailable -Name "PSFramework") {Import-module PSFramework} Else { install-module PSFramework -Force; import-module PSFramework}
$logFile = Join-Path -path "$ScriptPath\PasswordCleanupLogs" -ChildPath "log-$(Get-date -f 'yyyyMMddHHmmss').txt";
Set-PSFLoggingProvider -Name logfile -FilePath $logFile -Enabled $true;
Write-PSFMessage -Level Verbose -Message "Starting password cleanup."

If (Get-Module -ListAvailable -Name "ITGlueAPI") {Import-module ITGlueAPI -Force} Else { install-module ITGlueAPI -Force; import-module ITGlueAPI -Force}

# Connect to IT Glue
if ($ITGAPIKey.Key) {
	Add-ITGlueBaseURI -base_uri $ITGAPIKey.Url
	Add-ITGlueAPIKey $ITGAPIKey.Key
}

$ITGlueCompanies = Get-ITGlueOrganizations -page_size 1000

if ($ITGlueCompanies -and $ITGlueCompanies.data) {
	$ITGlueCompanies = $ITGlueCompanies.data | Where-Object { $_.attributes.'organization-status-name' -eq "Active" }
	Write-PSFMessage -Level Verbose -Message "Cleaning up passwords in $($ITGlueCompanies.Count) companies."
}

if (!$ITGlueCompanies) {
	exit
}

# Get password matching last updated time
$PasswordMatchingLastUpdated = $false
if ((Test-Path -Path ("$ScriptPath\passwordMatching_lastUpdated.txt"))) {
	$PasswordMatchingLastUpdated = Get-Content -Path "$ScriptPath\passwordMatching_lastUpdated.txt" -Raw
	if ([string]$PasswordMatchingLastUpdated -as [DateTime])   {
		$PasswordMatchingLastUpdated = Get-Date $PasswordMatchingLastUpdated
	}
}

# Get the flexible asset IDs
$FlexAssetID_AD = (Get-ITGlueFlexibleAssetTypes -filter_name $ITG_ADFlexAsset).data
$FlexAssetID_Email = (Get-ITGlueFlexibleAssetTypes -filter_name $ITG_EmailFlexAsset).data

$PasswordCategories = (Get-ITGluePasswordCategories).data

$QPMatchingFixes = [System.Collections.ArrayList]::new()

# Takes a password category ID and returns the name of the category
function Get-PasswordCategoryNameByID ($CategoryID) {
	$Category = $PasswordCategories | Where-Object { $_.id -eq $CategoryID }

	if (!$Category) {
		return "Unknown"
	}

	$CategoryName = $Category.attributes.name

	if (!$CategoryName) {
		return "Unknown"
	}

	return $CategoryName
}

# This function will return an ITG password category ID
# It bases this off of the naming of the password
#
# $PasswordName - the name of the password
# $DefaultCategory - the default category ID to use, if not set it will be AD
# $ForceType - if you know the type (AD, O365, Local User, Local Admin), you can force it here, then just the admin checks will be performed
# $ReturnCategoryName - if set, the function will return the categories name instead of ID. Used for emailing suggested changes.
function Get-PasswordCategory ($PasswordName, $DefaultCategory = $false, $ForceType = $false, $ReturnCategoryName = $false) {
	if ($DefaultCategory) {
		$Category = $DefaultCategory
	} else {
		$Category = $PasswordCategoryIDs.AD
	}

	if (!$ForceType -and $PasswordName -like "*-*") {
		$Prefix = $PasswordName.Substring(0, $PasswordName.IndexOf("-")).Trim()
		if ($Prefix -like "*and*") {
			$Prefix = $Prefix.Replace("and", "&")
		}
		$CategoriesFound = $Prefix.split(",&/\".ToCharArray())
		$CategoriesFound = $CategoriesFound | ForEach-Object { $_.Trim() } | Where-Object { $_ -in $AllPrefixTypes }
		if ($DefaultCategory) {
			if ($DefaultCategory -in @($PasswordCategoryIDs.AD, $PasswordCategoryIDs.ADAdmin) -and ($CategoriesFound | Where-Object {$PrefixTypes.AD -contains $_}).Count -gt 0) {
				$ForceType = "AD"
			} elseif ($DefaultCategory -in @($PasswordCategoryIDs.AzureAD, $PasswordCategoryIDs.EmailAdmin, $PasswordCategoryIDs.Email) -and ($CategoriesFound | Where-Object {$PrefixTypes.O365 -contains $_}).Count -gt 0) {
				$ForceType = "O365"
			} elseif ($DefaultCategory -in @($PasswordCategoryIDs.LocalAdmin, $PasswordCategoryIDs.LocalUser) -and ($CategoriesFound | Where-Object {$AllLocalPrefixTypes -contains $_}).Count -gt 0) {
				if ($PasswordName -like "*Admin*" -or ($CategoriesFound | Where-Object {$PrefixTypes.LocalAdmin -contains $_}).Count -gt 0) {
					$ForceType = "Local Admin"
				} else {
					$ForceType = "Local User"
				}
			}
		}
		if (!$ForceType) {
			if (($CategoriesFound | Where-Object {$PrefixTypes.AD -contains $_}).Count -gt 0) {
				$ForceType = "AD"
			} elseif (($CategoriesFound | Where-Object {$PrefixTypes.O365 -contains $_}).Count -gt 0) {
				$ForceType = "O365"
			} elseif (($CategoriesFound | Where-Object {$AllLocalPrefixTypes -contains $_}).Count -gt 0) {
				if ($PasswordName -like "*Admin*" -or ($CategoriesFound | Where-Object {$PrefixTypes.LocalAdmin -contains $_}).Count -gt 0) {
					$ForceType = "Local Admin"
				} else {
					$ForceType = "Local User"
				}
			}
		}
	}

	if ($PasswordName -like "Local User -*" -or $PasswordName -like "Local -*" -or $ForceType -eq "Local User") {
		$Category = $PasswordCategoryIDs.LocalUser
	} elseif ($PasswordName -like "Local Admin -*" -or $PasswordName -like "Local Administrator -*" -or $ForceType -eq "Local Admin") {
		$Category = $PasswordCategoryIDs.LocalAdmin
	} elseif (
		$PasswordName -like "O365 -*" -or $PasswordName -like "O365,*" -or $PasswordName -like "O365 &*" -or $PasswordName -like "M365 -*" -or $PasswordName -like "AAD -*" -or $PasswordName -like "Azure AD -*" -or $PasswordName -like "O365 Email -*" -or $PasswordName -like "Email -*" -or $PasswordName -like "*Office 365*" -or $ForceType -eq "O365"
	) {
		if ($PasswordName -like "*Global Admin*") {
			$Category = $PasswordCategoryIDs.EmailAdmin
		} elseif ($PasswordName -like "AAD -*" -or $PasswordName -like "Azure AD -*") {
			$Category = $PasswordCategoryIDs.AzureAD
		} else {
			$Category = $PasswordCategoryIDs.Email
		}
	} elseif ($PasswordName -like "AD -*" -or $PasswordName -like "*Active Directory*" -or $DefaultCategory -eq $PasswordCategoryIDs.AD -or $ForceType -eq "AD") {
		if ($PasswordName -like "*Admin*" -or $PasswordName -like "*Sea to Sky*") {
			$Category = $PasswordCategoryIDs.ADAdmin
		} else {
			$Category = $PasswordCategoryIDs.AD
		}
	} elseif ($DefaultCategory -eq $PasswordCategoryIDs.Email) {
		if ($PasswordName -like "*Global Admin*") {
			$Category = $PasswordCategoryIDs.EmailAdmin
		}
	}

	if ($ReturnCategoryName) {
		$Category = Get-PasswordCategoryNameByID -CategoryID $Category
	}
	
	return $Category
}

# Run through each company and get/compare passwords for cleanup
$MatchingErrors = [System.Collections.ArrayList]::new()
$UpdatedPasswordMatching = $false
foreach ($Company in $ITGlueCompanies) {
	Write-PSFMessage -Level Verbose -Message "Auditing: $($Company.attributes.name)"

	$ITGContacts = @()
	$i = 1
	while ($i -le 10 -and ($ITGContacts | Measure-Object).Count -eq (($i-1) * 1000)) {
		$ITGContacts += (Get-ITGlueContacts -page_size 1000 -page_number $i -organization_id $Company.id).data
		$i++
	}

	if (!$ITGContacts) {
		Write-PSFMessage -Level Verbose -Message "Skipped Company. Could not find any contacts in ITG."
		continue
	}

	$ITGPasswords = @()
	$i = 1
	while ($i -le 10 -and ($ITGPasswords | Measure-Object).Count -eq (($i-1) * 1000)) {
		$ITGPasswords += (Get-ITGluePasswords -page_size 1000 -page_number $i -organization_id $Company.id).data
		$i++
	}

	if (!$ITGPasswords) {
		Write-PSFMessage -Level Verbose -Message "Skipped Company. Could not find any passwords in ITG."
		continue
	}

	$ADFlexAsset = @()
	$AADFlexAsset = @()
	if ($FlexAssetID_AD) {
		$ADFlexAsset = (Get-ITGlueFlexibleAssets -filter_flexible_asset_type_id $FlexAssetID_AD.id -filter_organization_id $Company.id).data
		if ($ADFlexAsset) {
			$AADFlexAsset = $ADFlexAsset | Where-Object { $_.attributes.traits.'ad-level' -eq "Azure AD" -and !$_.attributes.archived } 
			$ADFlexAsset = $ADFlexAsset | Where-Object { $_.attributes.traits.'ad-level' -ne "Azure AD" -and !$_.attributes.archived } 

			if (($ADFlexAsset | Measure-Object).Count -gt 1) {
				$ADFlexAsset_Filtered = $ADFlexAsset | Where-Object { 
					($_.attributes.traits.'primary-domain-controller' -and $_.attributes.traits.'primary-domain-controller'.values.count -gt 0) -or
					($_.attributes.traits.'ad-servers' -and $_.attributes.traits.'ad-servers'.values.count -gt 0)
				}
				if (($ADFlexAsset_Filtered | Measure-Object).Count -gt 0) {
					$ADFlexAsset = $ADFlexAsset_Filtered
				}
			}
			if (($ADFlexAsset | Measure-Object).Count -gt 1) {
				$ADFlexAsset = $ADFlexAsset | Sort-Object -Property { $_.attributes.'updated-at' } -Descending | Select-Object -First 1
			}
			if (($AADFlexAsset | Measure-Object).Count -gt 1) {
				$AADFlexAsset = $AADFlexAsset | Sort-Object -Property { $_.attributes.'updated-at' } -Descending | Select-Object -First 1
			}
		}
	}

	$EmailFlexAsset = @()
	if ($FlexAssetID_Email) {
		$EmailFlexAsset = (Get-ITGlueFlexibleAssets -filter_flexible_asset_type_id $FlexAssetID_Email.id -filter_organization_id $Company.id).data
		if ($EmailFlexAsset) {
			$EmailFlexAsset = $EmailFlexAsset | Where-Object { $_.attributes.traits.type -like "Office 365" -and $_.attributes.traits.status -eq "Active" -and !$_.attributes.archived }

			if (($EmailFlexAsset | Measure-Object).Count -gt 1) {
				$EmailFlexAsset = $EmailFlexAsset | Sort-Object -Property { $_.attributes.'updated-at' } -Descending | Select-Object -First 1
			}
		}
	}

	if ((!$ADFlexAsset -or !$AADFlexAsset) -and !$EmailFlexAsset) {		
		Write-PSFMessage -Level Verbose -Message "Skipped Company. Could not find an AD or Email asset in ITG."
		continue
	}

	$PasswordNaming_CatPrepend = "AD"
	$DefaultCategory = $PasswordCategoryIDs.AD
	if ($EmailFlexAsset -and $EmailFlexAsset.attributes.traits.'azure-ad-connect' -like 'Yes*') {
		$PasswordNaming_CatPrepend = "AD & O365"
		$DefaultCategory = $PasswordCategoryIDs.AD
	} elseif ($EmailFlexAsset -and $EmailFlexAsset.attributes.traits.'azure-ad-connect' -notlike 'Yes*') {
		if ($ADFlexAsset) {
			$PasswordNaming_CatPrepend = "AD"
			$DefaultCategory = $PasswordCategoryIDs.AD
		} elseif ($AADFlexAsset) {
			$PasswordNaming_CatPrepend = "AAD & O365"
			$DefaultCategory = $PasswordCategoryIDs.AzureAD
		} else {
			$PasswordNaming_CatPrepend = "O365"
			$DefaultCategory = $PasswordCategoryIDs.Email
		}
	} elseif ($AADFlexAsset) {
		$PasswordNaming_CatPrepend = "AAD"
		$DefaultCategory = $PasswordCategoryIDs.AzureAD
	} elseif ($ADFlexAsset) {
		$PasswordNaming_CatPrepend = "AD"
		$DefaultCategory = $PasswordCategoryIDs.AD
	}
	Write-PSFMessage -Level Verbose -Message "Using the default settings: Password Naming Category - $($PasswordNaming_CatPrepend), Default Category - $(Get-PasswordCategoryNameByID -CategoryID $DefaultCategory) ($($DefaultCategory))"

	# Find passwords with an email category and no "@" in the username
	$BadEmailPasswords = $ITGPasswords | Where-Object { $_.attributes.'password-category-name' -in @("Office 365", "Microsoft 365", "Microsoft 365 - Global Admin", "Azure AD", "Email Account") -and $_.attributes.'username' -notlike "*@*" }
	if ($BadEmailPasswords) {
		Write-PSFMessage -Level Warning -Message "Bad Email Passwords found without an @ in the email. $($BadEmailPasswords.Count) bad passwords found."
	}
	$BadEmailPasswords | ForEach-Object {
		$QPMatchingFixes.Add([PSCustomObject]@{
			Company = $Company.attributes.name
			id = $_.id
			Name = $_.attributes.name
			Link = $_.attributes.'resource-url'
			Related = ""
			FixType = "Email Username - No @"
		})

		Write-PSFMessage -Level Verbose -Message "Emailed Suggestion:: Bad email password found (no @): '$($_.attributes.name)' (Link: $($_.attributes.'resource-url')) - suggestion: fix username"
	}


	# Check each password for duplicates, fix naming, link to contacts, and export QP fix list
	foreach ($Password in $ITGPasswords) {
		$ForceCategoryType = $false
		if ($Password.attributes.'password-category-name') {
			switch ($Password.attributes.'password-category-name') {
				"Active Directory" { $ForceCategoryType = "AD" }
				"Active Directory - Administrator" { $ForceCategoryType = "AD" }
				"Active Directory - Service Account" { $ForceCategoryType = "AD" }
				"Active Directory - Vendor" { $ForceCategoryType = "AD" }
				"Office 365" { $ForceCategoryType = "O365" }
				"Microsoft 365" { $ForceCategoryType = "O365" }
				"Microsoft 365 - Global Admin" { $ForceCategoryType = "O365" }
				"Azure AD" { $ForceCategoryType = "O365" }
			}
		}
		
		if ($Password.attributes.name -like "0365 *" -or $Password.attributes.name -like "* 0365 *" -or $Password.attributes.name -like "0365-*" -or $Password.attributes.name -like "* 0365-*") {
			# Fix this god awful naming...
			$NewName = $Password.attributes.name -replace "0365", "O365"

			$UpdatedPassword = @{
				type = "passwords"
				attributes = @{
					name = $NewName
				}
			}

			$UpdateResult = Set-ITGluePasswords -organization_id $Password.attributes.'organization-id' -id $Password.id -data $UpdatedPassword
			if ($UpdateResult -and $UpdateResult.data -and ($UpdateResult.data | Measure-Object).count -gt 0) {
				Write-PSFMessage -Level Verbose -Message "Updated password: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) - updated: 0365 to O365"
				$Password.attributes = $UpdateResult.data[0].attributes
			}
		}

		# Quickpass created passwords, look for duplicates or cleanup naming
		if ($Password.attributes.notes -like "*Created by Quickpass*" -and $Password.attributes.notes -notlike "*Created by Quickpass - Cleaned*") {
			$RelatedPasswords = $ITGPasswords | Where-Object { $_.attributes.username -like $Password.attributes.username -and $_.id -ne $Password.id }
			$NameMatch = $false

			if (($RelatedPasswords | Measure-Object).count -eq 0) {
				$RelatedPasswords = $ITGPasswords | Where-Object { ($_.attributes.username -like "$($Password.attributes.username)@*" -or $_.attributes.username -like "*\$($Password.attributes.username)") -and $_.id -ne $Password.id }
			}
			if (($RelatedPasswords | Measure-Object).count -eq 0 -and $Password.attributes.name.Trim() -like "* *") {
				$RelatedPasswords = $ITGPasswords | Where-Object { $_.attributes.name -like "*$($Password.attributes.name)*" -and $_.id -ne $Password.id }
				$NameMatch = $true
			}

			# Narrow down
			if ($PasswordNaming_CatPrepend -ne "AD & O365") {
				if ($Password.attributes.'password-category-name' -in @("Office 365", "Microsoft 365")) {
					$RelatedPasswords = $RelatedPasswords | Where-Object {
						$_.attributes.name -like "O365 *" -or $_.attributes.name -like "* O365 *" -or $_.attributes.name -like "O365-*" -or $_.attributes.name -like "* O365-*" -or
						$_.attributes.name -like "M365 *" -or $_.attributes.name -like "* M365 *" -or $_.attributes.name -like "M365-*" -or $_.attributes.name -like "* M365-*" -or
						$_.attributes.name -like "AAD *" -or $_.attributes.name -like "* AAD *" -or $_.attributes.name -like "AAD-*" -or $_.attributes.name -like "* AAD-*" -or
						$_.attributes.name -like "*Azure AD *" -or $_.attributes.name -like "*Office 365*"	-or $_.attributes.name -like "*Email*"	-or
						$_.attributes.'password-category-name' -like "Email Account / O365 User" -or $_.attributes.'password-category-name' -like "Microsoft 365*" -or $_.attributes.'password-category-name' -like "Office 365" -or $_.attributes.'password-category-name' -like "Cloud Management*" -or $_.attributes.'password-category-name' -like "Azure AD"
					}
				} else {
					$RelatedPasswords = $RelatedPasswords | Where-Object { 
						($_.attributes.name -like "AD *" -or $_.attributes.name -like "* AD *" -or $_.attributes.name -like "AD-*" -or $_.attributes.name -like "* AD-*" -or
						$_.attributes.'password-category-name' -like "Active Directory" -or $_.attributes.'password-category-name' -like "Active Directory -*") -and
						$_.attributes.name -notlike "*Azure AD*"
					}
				}
			}

			if (($RelatedPasswords | Measure-Object).count -gt 1) {
				$RelatedPasswords_Filtered = $RelatedPasswords | Where-Object { $_.attributes.name -like "*AD*" -or $_.attributes.name -like "*AAD*" -or $_.attributes.name -like "*Azure AD*" -or $_.attributes.name -like "*O365*" -or $_.attributes.name -like "*M365*" -or $_.attributes.name -like "*Office 365*" }
				if (($RelatedPasswords_Filtered | Measure-Object).count -gt 0) {
					$RelatedPasswords = $RelatedPasswords_Filtered
				}
			}
			if (($RelatedPasswords | Measure-Object).count -gt 1) {
				$RelatedPasswords_Filtered = $RelatedPasswords | Where-Object { !$_.attributes.'password-category-name' -or $_.attributes.'password-category-name' -like "Active Directory" -or $_.attributes.'password-category-name' -like "Active Directory -*" -or $_.attributes.'password-category-name' -like "Email Account / O365 User" -or $_.attributes.'password-category-name' -like "Microsoft 365*" -or $_.attributes.'password-category-name' -like "Office 365" -or $_.attributes.'password-category-name' -like "Cloud Management*" -or $_.attributes.'password-category-name' -like "Azure AD" }
				if (($RelatedPasswords_Filtered | Measure-Object).count -gt 0) {
					$RelatedPasswords = $RelatedPasswords_Filtered
				}
			}
			if (($RelatedPasswords | Measure-Object).count -gt 1) {
				if ($Password.attributes.'password-category-name' -in @("Office 365", "Microsoft 365")) {
					$RelatedPasswords_Filtered = $RelatedPasswords | Where-Object { 
						$_.attributes.name -like "O365 *" -or $_.attributes.name -like "* O365 *" -or $_.attributes.name -like "O365-*" -or $_.attributes.name -like "* O365-*" -or
						$_.attributes.name -like "M365 *" -or $_.attributes.name -like "* M365 *" -or $_.attributes.name -like "M365-*" -or $_.attributes.name -like "* M365-*" -or
						$_.attributes.name -like "AAD *" -or $_.attributes.name -like "* AAD *" -or $_.attributes.name -like "AAD-*" -or $_.attributes.name -like "* AAD-*" -or
						$_.attributes.name -like "*Azure AD *" -or $_.attributes.name -like "*Office 365*"		
					}
				} else {
					$RelatedPasswords_Filtered = $RelatedPasswords | Where-Object { ($_.attributes.name -like "AD *" -or $_.attributes.name -like "* AD *" -or $_.attributes.name -like "AD-*" -or $_.attributes.name -like "* AD-*") -and $_.attributes.name -notlike "*Azure AD*" }
				}
				if (($RelatedPasswords_Filtered | Measure-Object).count -gt 0) {
					$RelatedPasswords = $RelatedPasswords_Filtered
				}
			}

			# If we found matches (and not too many), make some changes
			if (($RelatedPasswords | Measure-Object).count -gt 0 -and ($RelatedPasswords | Measure-Object).count -le 3) {
				$UsernameMatches = $false
				if ($Password.attributes.username) {
					$UsernameMatches = $RelatedPasswords | Where-Object { $_.attributes.username -like "$($Password.attributes.username.Trim())" }
				}
				$RelatedUpdates = 0

				if (!$UsernameMatches) {
					# The related password appears to be using an incorrect or old username, update it
					if ($NameMatch) {
						$RelatedPasswords | ForEach-Object {
							$NewCategory = Get-PasswordCategory -PasswordName $_.attributes.name -DefaultCategory $DefaultCategory -ForceType $ForceCategoryType -ReturnCategoryName $true
							$QPMatchingFixes.Add([PSCustomObject]@{
								Company = $Company.attributes.name
								id = $_.id
								Name = $_.attributes.name
								Link = $_.attributes.'resource-url'
								Related = $Password.attributes.'resource-url'
								FixType = "Update Username to '$($Password.attributes.username)' and Category to '$($NewCategory)' (No Auto-Update due to Name-based match)"
							})

							Write-PSFMessage -Level Verbose -Message "Emailed Suggestion:: Update password: '$($_.attributes.name)' (Link: $($_.attributes.'resource-url')) (Related: $($Password.attributes.'resource-url')) - suggestion: update username to $($Password.attributes.username)' and category to '$($NewCategory)'"
						}
					} else {
						foreach ($RelatedPassword in $RelatedPasswords) {							
							$NewNotes = $RelatedPassword.attributes.notes

							if ($RelatedPassword.attributes.username -and $RelatedPassword.attributes.username -notlike "*$($Password.attributes.username.Trim())*") {
								if ($NewNotes) {
									$NewNotes = $NewNotes.Trim()
									$NewNotes += "`n"
								}
								$NewNotes += "Old/Other Username: $($RelatedPassword.attributes.username)"
							}

							$NewCategoryID = Get-PasswordCategory -PasswordName $RelatedPassword.attributes.name -DefaultCategory $DefaultCategory -ForceType $ForceCategoryType
							$UpdatedPassword = @{
								type = "passwords"
								attributes = @{
									"username" = $Password.attributes.username
									'password-category-id' = $NewCategoryID
									'notes' = $NewNotes
								}
							}

							$UpdateResult = Set-ITGluePasswords -organization_id $RelatedPassword.attributes.'organization-id' -id $RelatedPassword.id -data $UpdatedPassword

							if ($UpdateResult -and $UpdateResult.data -and ($UpdateResult.data | Measure-Object).count -gt 0) {
								$LogChanges = "Username - $($RelatedPassword.attributes.username) to $($Password.attributes.username)"
								if ($NewCategoryID -ne $RelatedPassword.attributes.'password-category-id') {
									$LogChanges += ", Category - $($RelatedPassword.attributes.'password-category-name') to $(Get-PasswordCategoryNameByID -CategoryID $NewCategoryID)"
								}
								Write-PSFMessage -Level Verbose -Message "Updated password: '$($RelatedPassword.attributes.name)' (Link: $($RelatedPassword.attributes.'resource-url')) - updated: $($LogChanges)"
								$RelatedPassword.attributes = $UpdateResult.data[0].attributes
								$RelatedUpdates++
							}
						}
					}
				}
				
				if ($Password.attributes.'password-category-name' -like "Active Directory*") {
					$ADTypeMatches = $RelatedPasswords | Where-Object { $_.attributes.'password-category-name' -like 'Active Directory*' }

					if (!$ADTypeMatches) {
						# The related passwords appear to not be using an AD category type, update it
						foreach ($RelatedPassword in $RelatedPasswords) {
							$NewName = $RelatedPassword.attributes.name
							if ($NewName -like "O365 -*" -or $NewName -like "M365 -*" -or $NewName -like "AAD -*") {
								if ($NewName -like "O365 -*" -or $NewName -like "M365 -*") {
									$NewName = $NewName.Substring(6).Trim()
								} elseif ($NewName -like "AAD -*") {
									$NewName = $NewName.Substring(5).Trim()
								}
								$NewName = "$($PasswordNaming_CatPrepend) - $($NewName)"
							} elseif ($NewName -notlike "*-*") {
								$NewName = "$($PasswordNaming_CatPrepend) - $($NewName)"
							}

							if ($NameMatch) {
								$NewCategory = Get-PasswordCategory -PasswordName $NewName -DefaultCategory $DefaultCategory -ForceType 'AD' -ReturnCategoryName $true
								$QPMatchingFixes.Add([PSCustomObject]@{
									Company = $Company.attributes.name
									id = $RelatedPassword.id
									Name = $RelatedPassword.attributes.name
									Link = $RelatedPassword.attributes.'resource-url'
									Related = $Password.attributes.'resource-url'
									FixType = "Update Category to '$($NewCategory)' and Name to '$($NewName)' (No Auto-Update due to Name-based match)"
								})

								Write-PSFMessage -Level Verbose -Message "Emailed Suggestion:: Update password: '$($RelatedPassword.attributes.name)' (Link: $($RelatedPassword.attributes.'resource-url')) (Related: $($Password.attributes.'resource-url')) - suggestion: update category to '$($NewCategory)' and name to '$($NewName)'"
							} else {
							
								$NewCategoryID = Get-PasswordCategory -PasswordName $NewName -DefaultCategory $DefaultCategory -ForceType 'AD'
								$UpdatedPassword = @{
									type = "passwords"
									attributes = @{
										'name' = $NewName
										'password-category-id' = $NewCategoryID
									}
								}

								$UpdateResult = Set-ITGluePasswords -organization_id $RelatedPassword.attributes.'organization-id' -id $RelatedPassword.id -data $UpdatedPassword

								if ($UpdateResult -and $UpdateResult.data -and ($UpdateResult.data | Measure-Object).count -gt 0) {
									$LogChanges = "Category - $($RelatedPassword.attributes.'password-category-name') to $(Get-PasswordCategoryNameByID -CategoryID $NewCategoryID)"
									if ($NewName -ne $RelatedPassword.attributes.name) {
										$LogChanges += ", Name - $($RelatedPassword.attributes.name) to $($NewName)"
									}

									Write-PSFMessage -Level Verbose -Message "Updated password: '$($RelatedPassword.attributes.name)' (Link: $($RelatedPassword.attributes.'resource-url')) - updated: $($LogChanges)"
									$RelatedPassword.attributes = $UpdateResult.data[0].attributes
									$RelatedUpdates++
								}
							}
						}
					}
				} elseif ($Password.attributes.'password-category-name' -like "Office 365*" -or $Password.attributes.'password-category-name' -like "Microsoft 365*" -or $Password.attributes.'password-category-name' -like "Azure AD*") {
					$EmailTypeMatches = $RelatedPasswords | Where-Object { $_.attributes.'password-category-name' -like 'Office 365*' -or $_.attributes.'password-category-name' -like 'Email*' -or $_.attributes.'password-category-name' -like 'Cloud Management*' -or $_.attributes.'password-category-name' -like 'Azure AD*' -or $_.attributes.'password-category-name' -like 'Microsoft 365*' }

					if (!$EmailTypeMatches) {
						# The related passwords appear to not be using an AD category type, notify to update it
						foreach ($RelatedPassword in $RelatedPasswords) {
							$NewName = $RelatedPassword.attributes.name
							if ($NewName -like "AD -*") {
								$NewName = $NewName.Substring(4).Trim()
								if ($AADFlexAsset) {
									$NewName = "AAD & O365 - $($NewName)"
								} else {
									$NewName = "O365 - $($NewName)"
								}
							} elseif ($NewName -notlike "*-*") {
								if ($AADFlexAsset) {
									$NewName = "AAD & O365 - $($NewName)"
								} else {
									$NewName = "O365 - $($NewName)"
								}
							}

							if ($NameMatch) {
								$NewCategory = Get-PasswordCategory -PasswordName $NewName -DefaultCategory $DefaultCategory -ForceType 'O365' -ReturnCategoryName $true
								$QPMatchingFixes.Add([PSCustomObject]@{
									Company = $Company.attributes.name
									id = $RelatedPassword.id
									Name = $RelatedPassword.attributes.name
									Link = $RelatedPassword.attributes.'resource-url'
									Related = $Password.attributes.'resource-url'
									FixType = "Update Category to '$($NewCategory)' and Name to '$($NewName)' (No Auto-Update due to Name-based match)"
								})

								Write-PSFMessage -Level Verbose -Message "Emailed Suggestion:: Update password: '$($RelatedPassword.attributes.name)' (Link: $($RelatedPassword.attributes.'resource-url')) (Related: $($Password.attributes.'resource-url')) - suggestion: update category to '$($NewCategory)' and name to '$($NewName)'"
							} else {
								$NewCategoryID = Get-PasswordCategory -PasswordName $NewName -DefaultCategory $DefaultCategory -ForceType 'O365'
								$UpdatedPassword = @{
									type = "passwords"
									attributes = @{
										'name' = $NewName
										'password-category-id' = $NewCategoryID
									}
								}

								$UpdateResult = Set-ITGluePasswords -organization_id $RelatedPassword.attributes.'organization-id' -id $RelatedPassword.id -data $UpdatedPassword

								if ($UpdateResult -and $UpdateResult.data -and ($UpdateResult.data | Measure-Object).count -gt 0) {
									$LogChanges = "Category - $($RelatedPassword.attributes.'password-category-name') to $(Get-PasswordCategoryNameByID -CategoryID $NewCategoryID)"
									if ($NewName -ne $RelatedPassword.attributes.name) {
										$LogChanges += ", Name - $($RelatedPassword.attributes.name) to $($NewName)"
									}

									Write-PSFMessage -Level Verbose -Message "Updated password: '$($RelatedPassword.attributes.name)' (Link: $($RelatedPassword.attributes.'resource-url')) - updated: $($LogChanges)"
									$RelatedPassword.attributes = $UpdateResult.data[0].attributes
									$RelatedUpdates++
								} 
							}
						}
					}
				}

				if ($RelatedUpdates -eq 0) {
					$QPMatchingFixes.Add([PSCustomObject]@{
						Company = $Company.attributes.name
						id = $Password.id
						Name = $Password.attributes.name
						Link = $Password.attributes.'resource-url'
						Related = @($RelatedPasswords.attributes.'resource-url') -join " "
						FixType = "Possible Duplicate - Not Deleted. No related passwords were updated. Consider deleting, and if so, update QP matching. ($(($RelatedPasswords | Measure-Object).count) related passwords found.)"
					})

					Write-PSFMessage -Level Verbose -Message "Emailed Suggestion:: Delete duplicate password: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) (Related: $(@($RelatedPasswords.attributes.'resource-url') -join " ")) - suggestion: possible duplicate, consider deleting, if so, update QP matching"
					continue
				}

				# Delete duplicate quickpass password and add to the QP Matching Fixes list
				try {
					$null = Remove-ITGluePasswords -id $Password.id -ErrorAction Stop
					$ITGPasswords = $ITGPasswords | Where-Object { $_.id -ne $Password.id }

					$QPMatchingFixes.Add([PSCustomObject]@{
						Company = $Company.attributes.name
						id = $Password.id
						Name = $Password.attributes.name
						Link = $Password.attributes.'resource-url'
						Related = @($RelatedPasswords.attributes.'resource-url') -join " "
						FixType = "Duplicate - Deleted. Update QP Matching for password."
					})
					Write-PSFMessage -Level Verbose -Message "Deleted duplicate password: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) - in favour of: $(@($RelatedPasswords.attributes.'resource-url') -join ", ")"
					Write-PSFMessage -Level Verbose -Message "Emailed Suggestion:: Duplicate Deleted, Update QP Matching: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) (Related: $(@($RelatedPasswords.attributes.'resource-url') -join " ")) - suggestion: update QP matching"
				} catch {
					$QPMatchingFixes.Add([PSCustomObject]@{
						Company = $Company.attributes.name
						id = $Password.id
						Name = $Password.attributes.name
						Link = $Password.attributes.'resource-url'
						Related = @($RelatedPasswords.attributes.'resource-url') -join " "
						FixType = "Duplicate - Not Deleted. Manually delete and update QP Matching for password."
					})

					Write-PSFMessage -Level Verbose -Message "Emailed Suggestion:: Delete duplicate password: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) (Related: $(@($RelatedPasswords.attributes.'resource-url') -join " ")) - suggestion: delete duplicate and update QP matching"
				}
			} elseif (($RelatedPasswords | Measure-Object).count -gt 3) {
				# More than 3 matches found, lets manually fix this
				$QPMatchingFixes.Add([PSCustomObject]@{
					Company = $Company.attributes.name
					id = $Password.id
					Name = $Password.attributes.name
					Link = $Password.attributes.'resource-url'
					Related = @($RelatedPasswords.attributes.'resource-url') -join " "
					FixType = ">3 Matches, Manually Fix"
				})

				Write-PSFMessage -Level Verbose -Message "Emailed Suggestion:: >3 Matches, Manually Fix: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) (Related: $(@($RelatedPasswords.attributes.'resource-url') -join " ")) - suggestion: too many matches, manually audit"
				continue
			} else {
				# No matches found, just update the name and category of this quickpass added password
				$NewName = $Password.attributes.name

				if ($Password.attributes.username -like "*@*" -or $Password.attributes.'password-category-name' -in @("Office 365", "Microsoft 365", "Azure AD")) {
					if ($NewName -notlike "O365 -*" -and $NewName -notlike "M365 -*" -and $NewName -notlike "AAD -*" -and $NewName -notlike "Azure AD -*") {
						if ($AADFlexAsset) {
							$NewName = "AAD & O365 - $($Password.attributes.name)"
						} else {
							$NewName = "O365 - $($Password.attributes.name)"
						}
					}
				} else {
					if ($NewName -notlike "$($PasswordNaming_CatPrepend) -*") {
						if ($NewName -like "AD -*") {
							$NewName = $NewName.Substring(4).Trim()
						}
						$NewName = "$($PasswordNaming_CatPrepend) - $($NewName)"
					}
				}

				if ($NewName) {
					$NewCategoryID = Get-PasswordCategory -PasswordName $NewName -DefaultCategory $DefaultCategory -ForceType $ForceCategoryType
					$UpdatedPassword = @{
						type = "passwords"
						attributes = @{
							"name" = $NewName
							'password-category-id' = $NewCategoryID
							'notes' = ($Password.attributes.notes -replace "Quickpass", "Quickpass - Cleaned")
						}
					}

					$UpdateResult = Set-ITGluePasswords -organization_id $Company.id -id $Password.id -data $UpdatedPassword
					if ($UpdateResult -and $UpdateResult.data -and ($UpdateResult.data | Measure-Object).count -gt 0) {
						$LogChanges = "Name - $($Password.attributes.name) to $($NewName)"
						if ($NewCategoryID -ne $Password.attributes.'password-category-id') {
							$LogChanges += ", Category - $($Password.attributes.'password-category-name') to $(Get-PasswordCategoryNameByID -CategoryID $NewCategoryID)"
						}

						Write-PSFMessage -Level Verbose -Message "Updated password: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) - updated: $($LogChanges)"
						$Password.attributes = $UpdateResult.data[0].attributes
					}
				}
			}


			
		} else {
			# Not quickpass created, check naming and category
			$UpdatedPassword = $false
			$LogChanges = ""

			if (($Password.attributes.name -like "AD *" -or $Password.attributes.name -like "AD-*" -or $Password.attributes.name -like "* AD *") -and $Password.attributes.name -notlike "*Azure AD*") {
				if (!$Password.attributes.'password-category-name' -or $Password.attributes.'password-category-name' -like "Email Account / O365 User" -or $Password.attributes.'password-category-id' -eq $PasswordCategoryIDs.Email) {
					# Update AD category
					$NewCategoryID = Get-PasswordCategory -PasswordName $Password.attributes.name -DefaultCategory $DefaultCategory

					if ($Password.attributes.'password-category-id' -ne $NewCategoryID) {
						$UpdatedPassword = @{
							type = "passwords"
							attributes = @{
								'password-category-id' = $NewCategoryID
							}
						}

						$LogChanges = "Category - $($Password.attributes.'password-category-name') to $(Get-PasswordCategoryNameByID -CategoryID $NewCategoryID)"
					}
				}
			} elseif (
				($Password.attributes.name -like "O365 *" -or $Password.attributes.name -like "* O365 *" -or $Password.attributes.name -like "O365-*" -or $Password.attributes.name -like "* O365-*" -or
				$Password.attributes.name -like "M365 *" -or $Password.attributes.name -like "* M365 *" -or $Password.attributes.name -like "M365-*" -or $Password.attributes.name -like "* M365-*" -or
				$Password.attributes.name -like "AAD *" -or $Password.attributes.name -like "* AAD *" -or $Password.attributes.name -like "AAD-*" -or $Password.attributes.name -like "* AAD-*" -or
				$Password.attributes.name -like "*Azure AD *" -or $Password.attributes.name -like "*Office 365*"-or $Password.attributes.name -like "*Email*") -and $Password.attributes.name -notlike "* Sync*"
			) {
				if (!$Password.attributes.'password-category-name' -or $Password.attributes.'password-category-name' -like "Active Directory*" -or $Password.attributes.'password-category-name' -like "Email Account / O365 User" -or $Password.attributes.'password-category-id' -eq $PasswordCategoryIDs.OldEmail) {
					# Update O365 category
					$NewCategoryID = Get-PasswordCategory -PasswordName $Password.attributes.name -DefaultCategory $DefaultCategory

					if ($Password.attributes.'password-category-id' -ne $NewCategoryID) {
						$UpdatedPassword = @{
							type = "passwords"
							attributes = @{
								'password-category-id' = $NewCategoryID
							}
						}

						$LogChanges = "Category - $($Password.attributes.'password-category-name') to $(Get-PasswordCategoryNameByID -CategoryID $NewCategoryID)"
					}
				}
			} elseif ($Password.attributes.'password-category-name' -like "Active Directory *") {
				if ($Password.attributes.name -notlike "AD *" -and $Password.attributes.name -notlike "`**") {
					# Update AD Name
					$NewName = $Password.attributes.name
					$NewName = "$($PasswordNaming_CatPrepend) - $($Password.attributes.name)"

					$UpdatedPassword = @{
						type = "passwords"
						attributes = @{
							name = $NewName
						}
					}

					$LogChanges = "Name - $($Password.attributes.name) to $($NewName)"
				}
			} elseif ($Password.attributes.'password-category-name' -like "Email Account / O365 User" -or $Password.attributes.'password-category-name' -like "Microsoft 365*" -or $Password.attributes.'password-category-name' -like "Office 365" -or $_.attributes.'password-category-name' -like "Azure AD" -or $Password.attributes.'password-category-id' -in @($PasswordCategoryIDs.Email, $PasswordCategoryIDs.EmailAdmin, $PasswordCategoryIDs.OldEmail, $PasswordCategoryIDs.AzureAD)) {
				$UpdatedPassword = @{
					type = "passwords"
					attributes = @{
					}
				}
				if ($Password.attributes.name -notlike "O365 *" -and $Password.attributes.name -notlike "M365 *" -and $Password.attributes.name -notlike "AAD *" -and $Password.attributes.name -notlike "`**") {
					# Update O365 Name and category if its one of the improper ones
					$NewName = $Password.attributes.name

					if ($Password.attributes.'password-category-id' -eq $PasswordCategoryIDs.AzureAD) {
						if ($PasswordNaming_CatPrepend -like "*AAD*") {
							$NewName = "$($PasswordNaming_CatPrepend) - $($Password.attributes.name)"
						} else {
							$NewName = "AAD - $($Password.attributes.name)"
						}
					} else {
						$NewName = "O365 - $($Password.attributes.name)"
					}
					
					$UpdatedPassword.attributes.name = $NewName
					$LogChanges = "Name - $($Password.attributes.name) to $($NewName)"
				}

				if ($Password.attributes.'password-category-id' -notin @($PasswordCategoryIDs.Email, $PasswordCategoryIDs.EmailAdmin, $PasswordCategoryIDs.AzureAD)) {
					$NewCategoryID = Get-PasswordCategory -PasswordName $Password.attributes.name -DefaultCategory $DefaultCategory -ForceType "O365"
					$UpdatedPassword.attributes.'password-category-id' = $NewCategoryID

					if ($LogChanges) { $LogChanges += ", " }
					$LogChanges += "Category - $($Password.attributes.'password-category-name') to $(Get-PasswordCategoryNameByID -CategoryID $NewCategoryID)"
				}
				
				if ($UpdatedPassword.attributes.count -eq 0) {
					$UpdatedPassword = $false
				}
			} elseif ($PasswordName -like "*Global Admin*" -and $Password.attributes.'password-category-name' -like "Cloud Management*") {
				if ($Password.attributes.'password-category-id' -notin @($PasswordCategoryIDs.Email, $PasswordCategoryIDs.EmailAdmin, $PasswordCategoryIDs.AzureAD)) {
					$NewCategoryID = Get-PasswordCategory -PasswordName $Password.attributes.name -DefaultCategory $DefaultCategory -ForceType "O365"
					$UpdatedPassword = @{
						type = "passwords"
						attributes = @{
							'password-category-id' = $NewCategoryID
						}
					}

					$LogChanges = "Category - $($Password.attributes.'password-category-name') to $(Get-PasswordCategoryNameByID -CategoryID $NewCategoryID)"
				}
			}

			if ($UpdatedPassword) {
				$UpdateResult = Set-ITGluePasswords -organization_id $Password.attributes.'organization-id' -id $Password.id -data $UpdatedPassword
				if ($UpdateResult -and $UpdateResult.data -and ($UpdateResult.data | Measure-Object).count -gt 0) {
					Write-PSFMessage -Level Verbose -Message "Updated password (non-QP): '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) - updated: $($LogChanges)"
					$Password.attributes = $UpdateResult.data[0].attributes
				}
			}
		}
	}

	# Match passwords to contacts / configurations (only newly created passwords since we last did this)
	$ToMatch_Passwords = $ITGPasswords
	if ($PasswordMatchingLastUpdated) {
		$ToMatch_Passwords = $ITGPasswords | Where-Object { (Get-Date $_.attributes.'created-at') -ge $PasswordMatchingLastUpdated }
	}
	$ToMatch_Passwords = $ToMatch_Passwords | Where-Object { $_.attributes.'password-category-name' -and $_.attributes.'password-category-name' -in $AllPasswordCatTypes_ForMatching }

	if ($ToMatch_Passwords -and ($ToMatch_Passwords | Measure-Object).Count -gt 0 ) {
		$ITGConfigurations = @()
		if (($ToMatch_Passwords | Where-Object { $_.attributes.'password-category-name' -in $PasswordCatTypes_ForMatching.ByConfigName } | Measure-Object).Count -gt 0) {
			# There are config name matches to make, grab all configurations
			$i = 1
			while ($i -le 10 -and ($ITGConfigurations | Measure-Object).Count -eq (($i-1) * 1000)) {
				$ITGConfigurations += (Get-ITGlueConfigurations -page_size 1000 -page_number $i -organization_id $Company.id).data
				$i++
			}
		}

		foreach ($Password in $ToMatch_Passwords) {
			$PasswordDetails = Get-ITGluePasswords -id $Password.id -include "related_items"
			$CurRelatedItems = $PasswordDetails.included | Where-Object { $_.type -eq "related-items" -and $_.attributes.archived -eq $false }
			$PasswordCategory = $Password.attributes.'password-category-name'

			# Config matching
			if ($PasswordCategory -in $PasswordCatTypes_ForMatching.ByConfigName) {
				if (($CurRelatedItems | Where-Object { $_.attributes.'asset-type' -eq "configuration" } | Measure-Object).Count -gt 0) {
					# Already has a configuration related item, skip
					continue
				}

				$ConfigMatches = $ITGConfigurations | Where-Object { $Password.attributes.name -like "*$($_.attributes.name)*" }
				if (!$ConfigMatches) {
					$ConfigMatches = $ITGConfigurations | Where-Object { $_.attributes.hostname -and $Password.attributes.name -like "*$($_.attributes.hostname)*" }
				}
				if (!$ConfigMatches) {
					$ConfigMatches = $ITGConfigurations | Where-Object { 
						if ($_.attributes.name -notlike "*-*") {
							return $false;
						} 
						$ContactParts = $_.attributes.name -split "-";
						if ($ContactParts[1] -and $ContactParts[1] -match "^\d+$" -and $ContactParts[1].length -gt 3) {
							$AssetNumber = $ContactParts[1] 
						} elseif ($ContactParts[2] -and $ContactParts[2] -match "^\d+$" -and $ContactParts[2].length -gt 3) {
							$AssetNumber = $ContactParts[2] 
						} else {
							return $false;
						}
						if ($Password.attributes.name -like "*$($AssetNumber.Trim())*") {
							return $true;
						}
						return $false;
					}
				}

				if (($ConfigMatches | Measure-Object).Count -gt 1) {
					$ConfigMatches_Temp = $ConfigMatches | Where-Object { $_.attributes.archived -eq $false }
					if (($ConfigMatches_Temp | Measure-Object).Count -gt 0) {
						$ConfigMatches = $ConfigMatches_Temp
					}
				}
				if (($ConfigMatches | Measure-Object).Count -gt 1) {
					$ConfigMatches_Temp = $ConfigMatches | Where-Object { $_.attributes.'configuration-status-name' -ne "Inactive" }
					if (($ConfigMatches_Temp | Measure-Object).Count -gt 0) {
						$ConfigMatches = $ConfigMatches_Temp
					}
				}
				if (($ConfigMatches | Measure-Object).Count -gt 1) {
					if ($Password.attributes.'password-category-name' -like '*Backup*') {
						$ConfigMatches_Temp = $ConfigMatches | Where-Object { $_.attributes.'configuration-type-name' -in $ConfigurationTypes_Backup }
					} elseif ($Password.attributes.'password-category-name' -like '*Camera*' -or $Password.attributes.'password-category-name' -like '*CCTV*' -or $Password.attributes.'password-category-name' -like '*DVR*') {
						$ConfigMatches_Temp = $ConfigMatches | Where-Object { $_.attributes.'configuration-type-name' -in $ConfigurationTypes_Camera }
					} elseif ($Password.attributes.'password-category-name' -like '*Printer*' -or $Password.attributes.'password-category-name' -like '*Printing*') {
						$ConfigMatches_Temp = $ConfigMatches | Where-Object { $_.attributes.'configuration-type-name' -in $ConfigurationTypes_Printer }
					} elseif ($Password.attributes.'password-category-name' -like '*Infrastructure*') {
						$ConfigMatches_Temp = $ConfigMatches | Where-Object { $_.attributes.'configuration-type-name' -in $ConfigurationTypes_Infrastructure }
					} elseif ($Password.attributes.'password-category-name' -like '*Storage*') {
							$ConfigMatches_Temp = $ConfigMatches | Where-Object { $_.attributes.'configuration-type-name' -in $ConfigurationTypes_Storage }
					} else {
						$ConfigMatches_Temp = $ConfigMatches | Where-Object { $_.attributes.'configuration-type-name' -in $ConfigurationTypes_Primary }
					}
					if (($ConfigMatches_Temp | Measure-Object).Count -gt 0) {
						$ConfigMatches = $ConfigMatches_Temp
					}
				}

				$ConfigMatches = $ConfigMatches | Sort-Object -Property id -Unique
				if (($ConfigMatches | Measure-Object).Count -gt 3) {
					$ConfigMatches = @()
				}

				if ($ConfigMatches -and ($ConfigMatches | Measure-Object).Count -gt 0) {
					# Add related items in ITG
					$UpdatedPasswordMatching = $true
					foreach ($ConfigMatch in $ConfigMatches) {
						$RelatedItems = @{
							type = 'related_items'
							attributes = @{
								destination_id = $ConfigMatch.id
								destination_type = "Configuration"
								notes = "Auto-Mapped by password cleanup"
							}
						}
						Write-PSFMessage -Level Verbose -Message "New Password Match Made: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) - Related to Configuration: '$($ConfigMatch.attributes.name)' (Link: $($ConfigMatch.attributes.'resource-url'))"

						try {
							New-ITGlueRelatedItems -resource_type passwords -resource_id $Password.id -data $RelatedItems | Out-Null
						} catch {
							$MatchingErrors.Add([PSCustomObject]@{
								Company = $Company.attributes.name
								PasswordID = $Password.id
								Name = $Password.attributes.name
								Link = $Password.attributes.'resource-url'
								RelatedItemType = "configuration"
								RelatedItemID = $ConfigMatch.id
								RelatedItemName = $ConfigMatch.attributes.name
								RelatedItemLink = $ConfigMatch.attributes.'resource-url'
							})
							Write-PSFMessage -Level Warning -Message "Emailed Error:: Password Match could not be made: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) - Related to Configuration: '$($ConfigMatch.attributes.name)' (Link: $($ConfigMatch.attributes.'resource-url'))"
						}
					}
				}
			}

			# Contact matching
			if ($PasswordCategory -in $PasswordCatTypes_ForMatching.ByContactName -or $PasswordCategory -in $PasswordCatTypes_ForMatching.ByContactEmail) {
				if (($CurRelatedItems | Where-Object { $_.attributes.'asset-type' -eq "contact" } | Measure-Object).Count -gt 0) {
					# Already has a contact related item, skip
					continue
				}

				# Find contact matches by name
				$AllContactMatches = @()
				if ($PasswordCategory -in $PasswordCatTypes_ForMatching.ByContactName) {
					$ContactMatches = $ITGContacts | Where-Object { $Password.attributes.name -like "*$($_.attributes.name)*" }
					if (!$ContactMatches) {
						$ContactMatches = $ITGContacts | Where-Object { $Password.attributes.name -like "*$($_.attributes.'first-name')*" -and $Password.attributes.name -like "*$($_.attributes.'last-name')*" }
					}
					if (($ContactMatches | Measure-Object).Count -gt 1) {
						$ContactMatches_Temp = $ContactMatches | Where-Object { $_.attributes.'contact-type-name' -ne "Terminated" }
						if (($ContactMatches_Temp | Measure-Object).Count -gt 0) {
							$ContactMatches = $ContactMatches_Temp
						}
					}
					if (($ContactMatches | Measure-Object).Count -gt 1) {
						if ($Password.attributes.'password-category-name' -like '*Service*') {
							$ContactMatches_Temp = $ContactMatches | Where-Object { $_.attributes.'contact-type-name' -in $ContactTypes_Service }
						} elseif ($Password.attributes.'password-category-name' -like '*Vendor*') {
							$ContactMatches_Temp = $ContactMatches | Where-Object { $_.attributes.'contact-type-name' -in $ContactTypes_Vendor }
						} else {
							$ContactMatches_Temp = $ContactMatches | Where-Object { $_.attributes.'contact-type-name' -in $ContactTypes_Employees }
						}
						if (($ContactMatches_Temp | Measure-Object).Count -gt 0) {
							$ContactMatches = $ContactMatches_Temp
						}
					}
					if (($ContactMatches | Measure-Object).Count -gt 3) {
						$ContactMatches = @()
					}
					$AllContactMatches += $ContactMatches
				}
				
				# Find contact matches by email/username
				if ($PasswordCategory -in $PasswordCatTypes_ForMatching.ByContactEmail) {
					$ContactMatches = $ITGContacts | Where-Object { ($_.attributes.'contact-emails' | Where-Object { $_.primary -eq $true }).value -like $Password.attributes.username }
					if (!$ContactMatches) {
						$ContactMatches = $ITGContacts | Where-Object { ($_.attributes.'contact-emails' | Where-Object { $_.primary -eq $true }).value -like "$($Password.attributes.username)*" }
					}
					if (!$ContactMatches -and $Password.attributes.name -like "*Global Admin*" -and ($Password.attributes.name -like "*365*" -or $Password.attributes.name -like "*Office*")) {
						$ContactMatches = $ITGContacts | Where-Object { ($_.attributes.'contact-emails' | Where-Object { $_.primary -eq $true }).value -like "it@*" -or $_.attributes.name -like "*Sea to Sky*" }
					}
					if (!$ContactMatches) {
						$ContactMatches = $ITGContacts | Where-Object { ($_.attributes.'contact-emails' | Where-Object { $_.primary -eq $true }).value -like "*$($Password.attributes.username)*" }
					}
					if (!$ContactMatches) {
						$ContactMatches = $ITGContacts | Where-Object { ($_.attributes.'contact-emails'.value -join " ") -like "*$($Password.attributes.username)*" }
					}
					if (!$ContactMatches -and $Password.attributes.username -like "*@*") {
						$Username = ($Password.attributes.username -split "@")[0]
						if ($Username.length -le 3) {
							# This helps prevent acronyms like "it" & "hr" from being matched to letters in employee names
							$Username = "$($Username)@"
						}
						$ContactMatches = $ITGContacts | Where-Object { ($_.attributes.'contact-emails' | Where-Object { $_.primary -eq $true }).value -like "*$($Username)*" }
						if (!$ContactMatches) {
							$ContactMatches = $ITGContacts | Where-Object { ($_.attributes.'contact-emails'.value -join " ") -like "*$Username*" }
						}
					}

					if (($ContactMatches | Measure-Object).Count -gt 1) {
						$ContactMatches_Temp = $ContactMatches | Where-Object { $_.attributes.'contact-type-name' -ne "Terminated" }
						if (($ContactMatches_Temp | Measure-Object).Count -gt 0) {
							$ContactMatches = $ContactMatches_Temp
						}
					}
					if (($ContactMatches | Measure-Object).Count -gt 1) {
						if ($Password.attributes.'password-category-name' -like '*Service*') {
							$ContactMatches_Temp = $ContactMatches | Where-Object { $_.attributes.'contact-type-name' -in $ContactTypes_Service }
						} elseif ($Password.attributes.'password-category-name' -like '*Vendor*') {
							$ContactMatches_Temp = $ContactMatches | Where-Object { $_.attributes.'contact-type-name' -in $ContactTypes_Vendor }
						} else {
							if ($Password.attributes.username -like "it@*") {
								$ContactMatches_Temp = $ContactMatches | Where-Object { $_.attributes.'contact-type-name' -in $ContactTypes_Employees -or $_.attributes.'contact-type-name' -like "Other" }
							} else {
								$ContactMatches_Temp = $ContactMatches | Where-Object { $_.attributes.'contact-type-name' -in $ContactTypes_Employees }
							}
						}
						if (($ContactMatches_Temp | Measure-Object).Count -gt 0) {
							$ContactMatches = $ContactMatches_Temp
						}
					}
					if (($ContactMatches | Measure-Object).Count -gt 3) {
						$ContactMatches = @()
					}
					$AllContactMatches += $ContactMatches
				}

				if ($AllContactMatches -and ($AllContactMatches | Measure-Object).Count -gt 0) {
					$AllContactMatches = $AllContactMatches | Sort-Object -Property id -Unique

					if (($AllContactMatches | Measure-Object).Count -gt 3) {
						$AllContactMatches = @()
					}
				}

				if ($AllContactMatches -and ($AllContactMatches | Measure-Object).Count -gt 0) {
					# Add related items in ITG
					$UpdatedPasswordMatching = $true
					foreach ($ContactMatch in $AllContactMatches) {
						$RelatedItems = @{
							type = 'related_items'
							attributes = @{
								destination_id = $ContactMatch.id
								destination_type = "Contact"
								notes = "Auto-Mapped by password cleanup"
							}
						}
						Write-PSFMessage -Level Verbose -Message "New Password Match Made: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) - Related to Contact: '$($ContactMatch.attributes.name)' (Link: $($ContactMatch.attributes.'resource-url'))"

						try {
							New-ITGlueRelatedItems -resource_type passwords -resource_id $Password.id -data $RelatedItems | Out-Null
						} catch {
							$MatchingErrors.Add([PSCustomObject]@{
								Company = $Company.attributes.name
								PasswordID = $Password.id
								Name = $Password.attributes.name
								Link = $Password.attributes.'resource-url'
								RelatedItemType = "contact"
								RelatedItemID = $ContactMatch.id
								RelatedItemName = $ContactMatch.attributes.name
								RelatedItemLink = $ContactMatch.attributes.'resource-url'
							})
							Write-PSFMessage -Level Warning -Message "Emailed Error:: Password Match could not be made: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) - Related to Contact: '$($ContactMatch.attributes.name)' (Link: $($ContactMatch.attributes.'resource-url'))"
						}
					}
				}
			}
		}
	}
}

if ($UpdatedPasswordMatching) {
	(Get-Date).ToString() | Out-File -FilePath "$ScriptPath\passwordMatching_lastUpdated.txt"
	Write-PSFMessage -Level Verbose -Message "Updated password matching last updated to: $((Get-Date).ToString())"
}

if ($QPMatchingFixes -and $Email_APIKey -and $Email_APIKey.Key) {
	# Send email with any manual fixes to be looked into
	$EmailTemplate = '
	<!doctype html>
	<html>
	<head>
		<meta name="viewport" content="width=device-width">
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
		<title>Simple Transactional Email</title>
		<style>
		/* -------------------------------------
			INLINED WITH htmlemail.io/inline
		------------------------------------- */
		.mobile_table_fallback {{
			display: none;
		}}
		/* -------------------------------------
			RESPONSIVE AND MOBILE FRIENDLY STYLES
		------------------------------------- */
		@media only screen and (max-width: 620px) {{
		table[class=body] h1 {{
			font-size: 28px !important;
			margin-bottom: 10px !important;
		}}
		table[class=body] p,
				table[class=body] ul,
				table[class=body] ol,
				table[class=body] td,
				table[class=body] span,
				table[class=body] a {{
			font-size: 16px !important;
		}}
		table[class=body] .wrapper,
				table[class=body] .article {{
			padding: 10px !important;
		}}
		table[class=body] .content {{
			padding: 0 !important;
		}}
		table[class=body] .container {{
			padding: 0 !important;
			width: 100% !important;
		}}
		table[class=body] .main {{
			border-left-width: 0 !important;
			border-radius: 0 !important;
			border-right-width: 0 !important;
		}}
		table[class=body] .btn table {{
			width: 100% !important;
		}}
		table[class=body] .btn a {{
			width: 100% !important;
		}}
		table[class=body] .img-responsive {{
			height: auto !important;
			max-width: 100% !important;
			width: auto !important;
		}}
		table.desktop_only_table {{
			display: none;
		}}
		.mobile_table_fallback {{
			display: block !important;
		}}
		}}

		/* -------------------------------------
			PRESERVE THESE STYLES IN THE HEAD
		------------------------------------- */
		@media all {{
		.ExternalClass {{
			width: 100%;
		}}
		.ExternalClass,
				.ExternalClass p,
				.ExternalClass span,
				.ExternalClass font,
				.ExternalClass td,
				.ExternalClass div {{
			line-height: 100%;
		}}
		.apple-link a {{
			color: inherit !important;
			font-family: inherit !important;
			font-size: inherit !important;
			font-weight: inherit !important;
			line-height: inherit !important;
			text-decoration: none !important;
		}}
		#MessageViewBody a {{
			color: inherit;
			text-decoration: none;
			font-size: inherit;
			font-family: inherit;
			font-weight: inherit;
			line-height: inherit;
		}}
		}}
		</style>
	</head>
	<body class="" style="background-color: #f6f6f6; font-family: sans-serif; -webkit-font-smoothing: antialiased; font-size: 14px; line-height: 1.4; margin: 0; padding: 0; -ms-text-size-adjust: 100%; -webkit-text-size-adjust: 100%;">
		<span class="preheader" style="color: transparent; display: none; height: 0; max-height: 0; max-width: 0; opacity: 0; overflow: hidden; mso-hide: all; visibility: hidden; width: 0;">This is preheader text. Some clients will show this text as a preview.</span>
		<table border="0" cellpadding="0" cellspacing="0" class="body" style="border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%; background-color: #f6f6f6;">
		<tr>
			<td style="font-family: sans-serif; font-size: 14px; vertical-align: top;">&nbsp;</td>
			<td class="container" style="font-family: sans-serif; font-size: 14px; vertical-align: top; display: block; Margin: 0 auto; max-width: 580px; padding: 10px; width: 580px;">
			<div class="content" style="box-sizing: border-box; display: block; Margin: 0 auto; max-width: 580px; padding: 10px;">

				<!-- START CENTERED WHITE CONTAINER -->
				<table class="main" style="border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%; background: #ffffff; border-radius: 3px;">

				<!-- START MAIN CONTENT AREA -->
				<tr>
					<td class="wrapper" style="font-family: sans-serif; font-size: 14px; vertical-align: top; box-sizing: border-box; padding: 20px;">
					<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%;">
						<tr>
						<td style="font-family: sans-serif; font-size: 14px; vertical-align: top;">
							<p style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 0; Margin-bottom: 15px;">{0}</p>
							<br />
							<p style="font-family: sans-serif; font-size: 18px; font-weight: normal; margin: 0; Margin-bottom: 15px;"><strong>{1}</strong></p>
							{2}
							<br />
							<p style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 0; Margin-bottom: 15px;">{3}</p>
						</td>
						</tr>
					</table>
					</td>
				</tr>

				<!-- END MAIN CONTENT AREA -->
				</table>

				<!-- START FOOTER -->
				<div class="footer" style="clear: both; Margin-top: 10px; text-align: center; width: 100%;">
				<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%;">
					<tr>
					<td class="content-block" style="font-family: sans-serif; vertical-align: top; padding-bottom: 10px; padding-top: 10px; font-size: 12px; color: #999999; text-align: center;">
						<span class="apple-link" style="color: #999999; font-size: 12px; text-align: center;">Sea to Sky Network Solutions, 2554 Vine Street, Vancouver BC V6K 3L1</span>
					</td>
					</tr>
				</table>
				</div>
				<!-- END FOOTER -->

			<!-- END CENTERED WHITE CONTAINER -->
			</div>
			</td>
			<td style="font-family: sans-serif; font-size: 14px; vertical-align: top;">&nbsp;</td>
		</tr>
		</table>
	</body>
	</html>'

	$HTMLBody = '
					<table class="desktop_only_table" cellpadding="0" cellspacing="0" style="border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: auto;">
						<tbody>
						<tr>
							<th>Company</th>
							<th>Password ID</th>
							<th>Password Name</th>
							<th>Link</th>
							<th style="width: 20%; max-width: 40px;">Related</th>
							<th>Issue</th>
						</tr>'
	$QPMatchingFixes | ForEach-Object {
		$HTMLBody += '
						<tr>
							<td style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 5px; border: 1px solid #000000;">{0}</td>
							<td style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 5px; border: 1px solid #000000;">{1}</td>
							<td style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 5px; border: 1px solid #000000;">{2}</td>
							<td style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 5px; border: 1px solid #000000;"><a href="{3}">{3}</a></td>
							<td style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 5px; border: 1px solid #000000; width: 20%; max-width: 40px;"><p style="word-break: break-all">{4}</p></td>
							<td style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 5px; border: 1px solid #000000;">{5}</td>
						</tr>' -f $_.Company, $_.id, $_.Name, $_.Link, ($_.Related -replace " ", "<br>"), $_.FixType
	}
	$HTMLBody += '
						</tbody>
					</table>
					<div class="mobile_table_fallback" style="display: none;">
						Table version hidden. You can view a tabular version of the above data on a desktop.
					</div><br />'

	$HTMLEmail = $EmailTemplate -f `
					"Issues were found during the password cleanup that need to be manually fixed.", 
					"Password Issues Found", 
					$HTMLBody, 
					"Please resolve these issues manually."

	$Path = "C:\Temp\QPMatchingFixes.csv"
	if (Test-Path $Path) {
        Remove-Item $Path
    }
	$QPMatchingFixes | Export-CSV -NoTypeInformation -Path $Path
	$QPMatchingFixes_Encoded = [System.Convert]::ToBase64String([IO.File]::ReadAllBytes($Path))

	$mailbody = @{
		"From" = $EmailFrom
		"To" = $EmailTo
		"Subject" = "Password Cleanup - Manual Fixes Required"
		"HTMLContent" = $HTMLEmail
		"Attachments" = @(
			@{
				Base64Content = $QPMatchingFixes_Encoded
				Filename = "QPMatchingFixes.csv"
				ContentType = "text/csv"
			}
		)
	} | ConvertTo-Json -Depth 6

	$headers = @{
		'x-api-key' = $Email_APIKey.Key
	}
	
	Invoke-RestMethod -Method Post -Uri $Email_APIKey.Url -Body $mailbody -Headers $headers -ContentType application/json
	Write-PSFMessage -Level Verbose -Message "Sent QPMatchingFixes manual cleanup email with $($QPMatchingFixes.Count) issues."
}

Write-PSFMessage -Level Verbose -Message "Script Complete."
Wait-PSFMessage