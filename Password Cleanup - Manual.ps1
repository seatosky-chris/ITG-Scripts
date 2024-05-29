###
# File: \Password Cleanup - Manual.ps1
# Project: Scripts
# Created Date: Tuesday, July 11th 2023, 3:50:36 pm
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
$ITGAPIKey = @{
	Url = "https://api.itglue.com"
	Key = ""
}
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

$QPMatchingFix_Export = $true

$PrefixTypes = @{
	AD = @("AD", "Active Directory")
	O365 = @("AAD", "O365", "0365", "M365", "Email", "Azure AD", "AzureAD", "O365 Email", "Office 365", "Office365", "Microsoft 365")
	LocalUser = @("Local", "Local User")
	LocalAdmin = @("Local Admin")
}
$AllPrefixTypes = $PrefixTypes.GetEnumerator() | ForEach-Object { $_.Value }
$AllLocalPrefixTypes = $PrefixTypes.GetEnumerator() | Where-Object { $_.Name -like "Local*" } | ForEach-Object { $_.Value }
####################################################################

### This code is common for every company and can be ran before looping through multiple companies
$CurrentTLS = [System.Net.ServicePointManager]::SecurityProtocol
if ($CurrentTLS -notlike "*Tls12" -and $CurrentTLS -notlike "*Tls13") {
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	Write-Output "This device is using an old version of TLS. Temporarily changed to use TLS v1.2."
	Write-PSFMessage -Level Warning -Message "Temporarily changed TLS to TLS v1.2."
}

If (Get-Module -ListAvailable -Name "ITGlueAPI") {Import-module ITGlueAPI -Force} Else { install-module ITGlueAPI -Force; import-module ITGlueAPI -Force}

# Connect to IT Glue
if ($ITGAPIKey.Key) {
	Add-ITGlueBaseURI -base_uri $ITGAPIKey.Url
	Add-ITGlueAPIKey $ITGAPIKey.Key
}

$ITGlueCompanies = Get-ITGlueOrganizations -page_size 1000

if ($ITGlueCompanies -and $ITGlueCompanies.data) {
	$ITGlueCompanies = $ITGlueCompanies.data | Where-Object { $_.attributes.'organization-status-name' -eq "Active" }
}

if (!$ITGlueCompanies) {
	exit
}

# Get the flexible asset IDs
$FlexAssetID_AD = (Get-ITGlueFlexibleAssetTypes -filter_name $ITG_ADFlexAsset).data
$FlexAssetID_Email = (Get-ITGlueFlexibleAssetTypes -filter_name $ITG_EmailFlexAsset).data
if (!$FlexAssetID_AD -or !$FlexAssetID_Email) {
	Write-Error "Could not get the AD or Email flex asset type ID. Exiting..."
	exit 1
}

$QPMatchingFixes = [System.Collections.ArrayList]::new()

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
foreach ($Company in $ITGlueCompanies) {
	Write-Host "Auditing: $($Company.attributes.name)" -ForegroundColor Green

	$ITGContacts = @()
	$i = 1
	while ($i -le 10 -and ($ITGContacts | Measure-Object).Count -eq (($i-1) * 1000)) {
		$ITGContacts += (Get-ITGlueContacts -page_size 1000 -page_number $i -organization_id $Company.id).data
		Write-Host "- Got Contact group set $i"
		$TotalGroups = ($ITGContacts | Measure-Object).Count
		Write-Host "- Total: $($ITGContacts.count)"
		$i++
	}

	if (!$ITGContacts) {
		continue
	}

	$ITGPasswords = @()
	$i = 1
	while ($i -le 10 -and ($ITGPasswords | Measure-Object).Count -eq (($i-1) * 1000)) {
		$ITGPasswords += (Get-ITGluePasswords -page_size 1000 -page_number $i -organization_id $Company.id).data
		Write-Host "- Got Password group set $i"
		$TotalGroups = ($ITGPasswords | Measure-Object).Count
		Write-Host "- Total: $($ITGPasswords.count)"
		$i++
	}

	if (!$ITGPasswords) {
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
		Write-Host "Skipping Company... Could not find an AD or Email asset in ITG for: $($Company.attributes.name)" -ForegroundColor Yellow
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

	$BadEmailPasswords = $ITGPasswords | Where-Object { $_.attributes.'password-category-name' -in @("Office 365", "Microsoft 365", "Microsoft 365 - Global Admin", "Azure AD", "Email Account") -and $_.attributes.'username' -notlike "*@*" }
	$BadEmailPasswords | ForEach-Object {
		$QPMatchingFixes.Add([PSCustomObject]@{
			Company = $Company.attributes.name
			id = $_.id
			Name = $_.attributes.name
			Link = $_.attributes.'resource-url'
			Related = ""
			FixType = "Email Username - No @"
		})
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
				Write-Host "Updated password: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) - updated: 0365 to O365" -ForegroundColor DarkGreen
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
			# Narrow down, specifically for O365 cloud admin accounts
			<# if (($RelatedPasswords | Measure-Object).count -gt 1 -and $Password.attributes.name -like "*Admin*" -and ($Password.attributes.name -like "*Office 365*" -or $Password.attributes.name -like "*O365*"  -or $Password.attributes.name -like "*M365*")) {
				$RelatedPasswords_Filtered = $RelatedPasswords | Where-Object {  }
				if (($RelatedPasswords_Filtered | Measure-Object).count -gt 0) {
					$RelatedPasswords = $RelatedPasswords_Filtered
				}
			} #>

			# If we found matches (and not too many), make some changes
			if (($RelatedPasswords | Measure-Object).count -gt 0 -and ($RelatedPasswords | Measure-Object).count -le 3) {
				if ($NameMatch) {
					Write-Host "Resorted to a name based search for: '$($Password.attributes.name)' (rather than by username)" -ForegroundColor Yellow
				}
				
				$UsernameMatches = $false
				if ($Password.attributes.username) {
					$UsernameMatches = $RelatedPasswords | Where-Object { $_.attributes.username -like "$($Password.attributes.username.Trim())" }
				}
				$RelatedUpdates = 0

				if (!$UsernameMatches) {
					# The related password appears to be using an incorrect or old username, notify to update it
					Write-Host "Related Usernames need updating: $($Password.attributes.name) (Link: $($Password.attributes.'resource-url')) - No related passwords use this username" -ForegroundColor Cyan
					foreach ($RelatedPassword in $RelatedPasswords) {
						$UpdateUsername = Read-Host "Update username of '$($RelatedPassword.attributes.name)' to '$($Password.attributes.username)'? Y or N? (Current Username: $($RelatedPassword.attributes.username)) ($(($RelatedPasswords | Measure-Object).count) related passwords found) (Link: $($RelatedPassword.attributes.'resource-url')))"
						if ($UpdateUsername -eq 'Y' -or $UpdateUsername -eq 'Yes') { 
							# Yes, update
							$NewNotes = $RelatedPassword.attributes.notes

							if ($RelatedPassword.attributes.username -and $RelatedPassword.attributes.username -notlike "*$($Password.attributes.username.Trim())*") {
								if ($NewNotes) {
									$NewNotes = $NewNotes.Trim()
									$NewNotes += "`n"
								}
								$NewNotes += "Old/Other Username: $($RelatedPassword.attributes.username)"
							}

							$UpdatedPassword = @{
								type = "passwords"
								attributes = @{
									"username" = $Password.attributes.username
									'password-category-id' = Get-PasswordCategory -PasswordName $RelatedPassword.attributes.name -DefaultCategory $DefaultCategory -ForceType $ForceCategoryType
									'notes' = $NewNotes
								}
							}

							$UpdateResult = Set-ITGluePasswords -organization_id $RelatedPassword.attributes.'organization-id' -id $RelatedPassword.id -data $UpdatedPassword

							if ($UpdateResult -and $UpdateResult.data -and ($UpdateResult.data | Measure-Object).count -gt 0) {
								Write-Host "Updated password: '$($RelatedPassword.attributes.name)' (Link: $($RelatedPassword.attributes.'resource-url')) - updated: username, category" -ForegroundColor DarkBlue
								$RelatedPassword.attributes = $UpdateResult.data[0].attributes
								$RelatedUpdates++
							}
						} else {
							# No, don't update
							continue
						} 
					}
				}
				
				if ($Password.attributes.'password-category-name' -like "Active Directory*") {
					$ADTypeMatches = $RelatedPasswords | Where-Object { $_.attributes.'password-category-name' -like 'Active Directory*' }

					if (!$ADTypeMatches) {
						# The related passwords appear to not be using an AD category type, notify to update it
						Write-Host "Related Passwords need updating: $($Password.attributes.name) (Link: $($Password.attributes.'resource-url')) - No related passwords have an AD category type" -ForegroundColor DarkCyan
						foreach ($RelatedPassword in $RelatedPasswords) {
							$UpdateCategory = Read-Host "Update category type of '$($RelatedPassword.attributes.name)' to 'Active Directory'? Y or N? (Current Category: $($RelatedPassword.attributes.'password-category-name')) ($(($RelatedPasswords | Measure-Object).count) related passwords found) (Link: $($RelatedPassword.attributes.'resource-url')))"
							if ($UpdateCategory -eq 'Y' -or $UpdateCategory -eq 'Yes') { 
								# Yes, update
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
								
								$UpdatedPassword = @{
									type = "passwords"
									attributes = @{
										'name' = $NewName
										'password-category-id' = Get-PasswordCategory -PasswordName $NewName -DefaultCategory $DefaultCategory -ForceType 'AD'
									}
								}

								$UpdateResult = Set-ITGluePasswords -organization_id $RelatedPassword.attributes.'organization-id' -id $RelatedPassword.id -data $UpdatedPassword

								if ($UpdateResult -and $UpdateResult.data -and ($UpdateResult.data | Measure-Object).count -gt 0) {
									Write-Host "Updated password: '$($RelatedPassword.attributes.name)' (Link: $($RelatedPassword.attributes.'resource-url')) - updated: category" -ForegroundColor DarkBlue
									$RelatedPassword.attributes = $UpdateResult.data[0].attributes
									$RelatedUpdates++
								}
							} else {
								# No, don't update
								continue
							} 
						}
					}
				} elseif ($Password.attributes.'password-category-name' -like "Office 365*" -or $Password.attributes.'password-category-name' -like "Microsoft 365*" -or $Password.attributes.'password-category-name' -like "Azure AD*") {
					$EmailTypeMatches = $RelatedPasswords | Where-Object { $_.attributes.'password-category-name' -like 'Office 365*' -or $_.attributes.'password-category-name' -like 'Email*' -or $_.attributes.'password-category-name' -like 'Cloud Management*' -or $_.attributes.'password-category-name' -like 'Azure AD*' -or $_.attributes.'password-category-name' -like 'Microsoft 365*' }

					if (!$EmailTypeMatches) {
						# The related passwords appear to not be using an AD category type, notify to update it
						Write-Host "Related Passwords need updating: $($Password.attributes.name) (Link: $($Password.attributes.'resource-url')) - No related passwords have an Email category type" -ForegroundColor DarkCyan
						foreach ($RelatedPassword in $RelatedPasswords) {
							$UpdateCategory = Read-Host "Update category type of '$($RelatedPassword.attributes.name)' to 'Microsoft 365'? Y or N? (Current Category: $($RelatedPassword.attributes.'password-category-name')) ($(($RelatedPasswords | Measure-Object).count) related passwords found) (Link: $($RelatedPassword.attributes.'resource-url')))"
							if ($UpdateCategory -eq 'Y' -or $UpdateCategory -eq 'Yes') { 
								# Yes, update
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

								$UpdatedPassword = @{
									type = "passwords"
									attributes = @{
										'name' = $NewName
										'password-category-id' = Get-PasswordCategory -PasswordName $NewName -DefaultCategory $DefaultCategory -ForceType 'O365'
									}
								}

								$UpdateResult = Set-ITGluePasswords -organization_id $RelatedPassword.attributes.'organization-id' -id $RelatedPassword.id -data $UpdatedPassword

								if ($UpdateResult -and $UpdateResult.data -and ($UpdateResult.data | Measure-Object).count -gt 0) {
									Write-Host "Updated password: '$($RelatedPassword.attributes.name)' (Link: $($RelatedPassword.attributes.'resource-url')) - updated: category" -ForegroundColor DarkBlue
									$RelatedPassword.attributes = $UpdateResult.data[0].attributes
									$RelatedUpdates++
								}
							} else {
								# No, don't update
								continue
							} 
						}
					}
				}

				if ($RelatedUpdates -eq 0) {
					$AllowDeletion = Read-Host "No related passwords were updated for '$($Password.attributes.name)' and its a possible duplicate. Still Delete? Y or N? ($(($RelatedPasswords | Measure-Object).count) related passwords found) (Link: $($Password.attributes.'resource-url')))"
					if ($AllowDeletion -ne 'Y' -and $AllowDeletion -ne 'Yes') {
						if ($QPMatchingFix_Export) {
							$QPMatchingFixes.Add([PSCustomObject]@{
								Company = $Company.attributes.name
								id = $Password.id
								Name = $Password.attributes.name
								Link = $Password.attributes.'resource-url'
								Related = @($RelatedPasswords.attributes.'resource-url') -join " "
								FixType = "Duplicate - Not Deleted"
							})
						}
						continue
					}
				}

				# Delete duplicate quickpass password and add to the QP Matching Fixes list
				try {
					$null = Remove-ITGluePasswords -id $Password.id -ErrorAction Stop
					$ITGPasswords = $ITGPasswords | Where-Object { $_.id -ne $Password.id }

					if ($QPMatchingFix_Export) {
						$QPMatchingFixes.Add([PSCustomObject]@{
							Company = $Company.attributes.name
							id = $Password.id
							Name = $Password.attributes.name
							Link = $Password.attributes.'resource-url'
							Related = @($RelatedPasswords.attributes.'resource-url') -join " "
							FixType = "Duplicate - Deleted"
						})
					}
				} catch {
					if ($QPMatchingFix_Export) {
						$QPMatchingFixes.Add([PSCustomObject]@{
							Company = $Company.attributes.name
							id = $Password.id
							Name = $Password.attributes.name
							Link = $Password.attributes.'resource-url'
							Related = @($RelatedPasswords.attributes.'resource-url') -join " "
							FixType = "Duplicate - Not Deleted"
						})
					}
					Write-Host "Could not remove ITG Password: $($Password.attributes.name) (ID: $($Password.id))" -ForegroundColor Red
				}
			} elseif (($RelatedPasswords | Measure-Object).count -gt 3) {
				# More than 3 matches found, lets manually fix this
				Write-Host "Manually Fix: $($Password.attributes.name) (Link: $($Password.attributes.'resource-url')) - More than 3 related passwords found" -ForegroundColor Magenta

				if ($QPMatchingFix_Export) {
					$QPMatchingFixes.Add([PSCustomObject]@{
						Company = $Company.attributes.name
						id = $Password.id
						Name = $Password.attributes.name
						Link = $Password.attributes.'resource-url'
						Related = @($RelatedPasswords.attributes.'resource-url') -join " "
						FixType = ">3 Matches, Manual Fix"
					})
				}
				continue
			} else {
				# No matches found, just update the name and category of this quickpass added password
				$NewName = $Password.attributes.name

				if ($Password.attributes.username -like "*@*" -or $Password.attributes.'password-category-name' -in @("Office 365", "Microsoft 365")) {
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
					$UpdatedPassword = @{
						type = "passwords"
						attributes = @{
							"name" = $NewName
							'password-category-id' = Get-PasswordCategory -PasswordName $NewName -DefaultCategory $DefaultCategory -ForceType $ForceCategoryType
							'notes' = ($Password.attributes.notes -replace "Quickpass", "Quickpass - Cleaned")
						}
					}

					$UpdateResult = Set-ITGluePasswords -organization_id $Company.id -id $Password.id -data $UpdatedPassword
					if ($UpdateResult -and $UpdateResult.data -and ($UpdateResult.data | Measure-Object).count -gt 0) {
						Write-Host "Updated password: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) - updated: name, category" -ForegroundColor DarkCyan
						$Password.attributes = $UpdateResult.data[0].attributes
					}
				}
			}


			
		} else {
			# Not quickpass created, check naming and category
			$UpdatedPassword = $false

			if (($Password.attributes.name -like "AD *" -or $Password.attributes.name -like "AD-*" -or $Password.attributes.name -like "* AD *") -and $Password.attributes.name -notlike "*Azure AD*") {
				if (!$Password.attributes.'password-category-name' -or $Password.attributes.'password-category-name' -like "Email Account / O365 User" -or $Password.attributes.'password-category-id' -eq $PasswordCategoryIDs.Email) {
					# Update AD category
					$NewCategory = Get-PasswordCategory -PasswordName $Password.attributes.name -DefaultCategory $DefaultCategory

					if ($Password.attributes.'password-category-id' -ne $NewCategory) {
						$UpdatedPassword = @{
							type = "passwords"
							attributes = @{
								'password-category-id' = $NewCategory
							}
						}
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
					$NewCategory = Get-PasswordCategory -PasswordName $Password.attributes.name -DefaultCategory $DefaultCategory

					if ($Password.attributes.'password-category-id' -ne $NewCategory) {
						$UpdatedPassword = @{
							type = "passwords"
							attributes = @{
								'password-category-id' = $NewCategory
							}
						}
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
				}

				if ($Password.attributes.'password-category-id' -notin @($PasswordCategoryIDs.Email, $PasswordCategoryIDs.EmailAdmin, $PasswordCategoryIDs.AzureAD)) {
					$UpdatedPassword.attributes.'password-category-id' = Get-PasswordCategory -PasswordName $Password.attributes.name -DefaultCategory $DefaultCategory -ForceType "O365"
				}
				
				if ($UpdatedPassword.attributes.count -eq 0) {
					$UpdatedPassword = $false
				}
			} elseif ($PasswordName -like "*Global Admin*" -and $Password.attributes.'password-category-name' -like "Cloud Management*") {
				if ($Password.attributes.'password-category-id' -notin @($PasswordCategoryIDs.Email, $PasswordCategoryIDs.EmailAdmin, $PasswordCategoryIDs.AzureAD)) {
					$UpdatedPassword = @{
						type = "passwords"
						attributes = @{
							'password-category-id' = Get-PasswordCategory -PasswordName $Password.attributes.name -DefaultCategory $DefaultCategory -ForceType "O365"
						}
					}
				}
			}

			if ($UpdatedPassword) {
				$UpdateResult = Set-ITGluePasswords -organization_id $Password.attributes.'organization-id' -id $Password.id -data $UpdatedPassword
				if ($UpdateResult -and $UpdateResult.data -and ($UpdateResult.data | Measure-Object).count -gt 0) {
					Write-Host "Updated password: '$($Password.attributes.name)' (Link: $($Password.attributes.'resource-url')) - updated: name or category" -ForegroundColor DarkGray
					$Password.attributes = $UpdateResult.data[0].attributes
				}
			}
		}
	}

	Read-Host "Completed '$($Company.attributes.name)', press any key to continue..."
}

if ($QPMatchingFixes) {
	$QPMatchingFixes | Export-Csv -Path "./QPMatchingFixes.csv" -NoTypeInformation
}
