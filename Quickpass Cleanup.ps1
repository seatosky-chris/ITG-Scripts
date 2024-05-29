###
# File: \Quickpass Cleanup.ps1
# Project: Scripts
# Created Date: Tuesday, December 5th 2023, 3:56:52 pm
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

#####################################
# QP login info
# This account must not use SSO, and can have MFA setup. If SSO is setup then to bypass SSO, this account must have the Super or Owner login role
$QP_Login = @{
	Email = ""
	Password = ""
	MFA_Secret = ""
}
$QP_BaseURI = "https://admin.getquickpass.com/api/"

$ITGAPIKey = @{
	Url = "https://api.itglue.com"
	Key = ""
}

$AutotaskAPIKey = @{
	Url = "https://webservicesX.autotask.net/atservicesrest"
	Username = ""
	Key = ''
	IntegrationCode = ""
}

# The global APIKey for the email forwarder. The key should give access to all organizations.
$Email_APIKey = @{
	Url = ""
	Key = ""
}

$EmailFrom = @{
	Email = ''
	Name = ""
}
$EmailTo = @(
	@{
		Email = ''
		Name = ""
	}
)

$RateLimit = 20 # The max amount of iterations to process in Quickpass ($RateLimit x 50 = max customers/users/etc.)
$EmailCategories = @("Azure AD", "Email Account", "Microsoft 365", "Office 365", "Microsoft 365 - Global Admin") # Email password categories

$ITG_ADFlexAssetName = "Active Directory"
#####################################


### This code is common for every company and can be ran before looping through multiple companies
$CurrentTLS = [System.Net.ServicePointManager]::SecurityProtocol
if ($CurrentTLS -notlike "*Tls12" -and $CurrentTLS -notlike "*Tls13") {
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	Write-Output "This device is using an old version of TLS. Temporarily changed to use TLS v1.2."
}

if ($QP_Login.MFA_Secret) {
	Unblock-File -Path ".\GoogleAuthenticator.psm1"
	Import-Module ".\GoogleAuthenticator.psm1"
}

function IsValidEmail { 
    param([string]$EmailAddress)

    try {
        $null = [mailaddress]$EmailAddress
        return $true
    }
    catch {
        return $false
    }
}

If (Get-Module -ListAvailable -Name "ITGlueAPI") {Import-module ITGlueAPI -Force} Else { install-module ITGlueAPI -Force; import-module ITGlueAPI -Force}
If (Get-Module -ListAvailable -Name "AutotaskAPI") {Import-module AutotaskAPI -Force} Else { install-module AutotaskAPI -Force; import-module AutotaskAPI -Force}

# Connect to IT Glue
if ($ITGAPIKey.Key) {
	Add-ITGlueBaseURI -base_uri $ITGAPIKey.Url
	Add-ITGlueAPIKey $ITGAPIKey.Key
}

# Connect to Autotask
$AutotaskConnected = $false
if ($AutotaskAPIKey.Key) {
	$Secret = ConvertTo-SecureString $AutotaskAPIKey.Key -AsPlainText -Force
	$Creds = New-Object System.Management.Automation.PSCredential($AutotaskAPIKey.Username, $Secret)
	Add-AutotaskAPIAuth -ApiIntegrationcode $AutotaskAPIKey.IntegrationCode -credentials $Creds
	Add-AutotaskBaseURI -BaseURI $AutotaskAPIKey.Url

	# Verify the Autotask API key works
	$AutotaskConnected = $true
	try { 
		$Test = Get-AutotaskAPIResource -Resource Companies -ID 0 -ErrorAction Stop 
	} catch { 
		$CleanError = ($_ -split "/n")[0]
		if ($_ -like "*(401) Unauthorized*") {
			$CleanError = "API Key Unauthorized. ($($CleanError))"
		}
		Write-Host $CleanError -ForegroundColor Red
		$AutotaskConnected = $false
	}
}

# Try to auth with QuickPass
$attempt = 3
$AuthResponse = $false
while ($attempt -ge 0 -and !$AuthResponse) {
	if ($attempt -eq 0) {
		# Already tried 10x, lets give up and exit the script
		Write-Error "Could not authenticate with QuickPass. Please verify the credentials and try again."
		exit
	}

	if ($QP_Login.MFA_Secret) {
		$MFACode = Get-GoogleAuthenticatorPin -Secret $QP_Login.MFA_Secret
		if ($MFACode.'Seconds Remaining' -le 5) {
			# If the current code is about to expire, lets wait until a new one is ready to be generated to grab the code and try to login
			Start-Sleep -Seconds ($MFACode.'Seconds Remaining' + 1)
			$MFACode = Get-GoogleAuthenticatorPin -Secret $QP_Login.MFA_Secret
		}

		$FormBody = @{
			email = $QP_Login.Email
			password = $QP_Login.Password
			accessCode = $MFACode."PIN Code" -replace " ", ""
			isIframe = $false
		} | ConvertTo-Json
	} else {
		$FormBody = @{
			email = $QP_Login.Email
			password = $QP_Login.Password
			isIframe = $false
		} | ConvertTo-Json
	}

	try {
		$AuthResponse = Invoke-WebRequest "$($QP_BaseURI)auth/login" -SessionVariable 'QPWebSession' -Body $FormBody -Method 'POST' -ContentType 'application/json; charset=utf-8'
	} catch {
		$attempt--
		Write-Host "Failed to connect to: QuickPass"
		Write-Host "Status Code: $($_.Exception.Response.StatusCode.Value__)"
		Write-Host "Message: $($_.Exception.Message)"
		Write-Host "Status Description: $($_.Exception.Response.StatusDescription)"
		start-sleep (get-random -Minimum 10 -Maximum 100)
		continue
	}
	if (!$AuthResponse) {
		$attempt--
		Write-Host "Failed to connect to: QuickPass"
		start-sleep (get-random -Minimum 10 -Maximum 100)
		continue
	}
}

if (!$QPWebSession) {
	Write-Host "Failed to connect to: QuickPass. No session found."
	exit
}

# Authentication successful, get the full list of customers
$Response = Invoke-WebRequest "$($QP_BaseURI)customer?page=1&rowsPerPage=50&searchText=&adStatus=ALL&localStatus=ALL&o365Status=ALL&filterUpdated=false" -WebSession $QPWebSession
$QP_Customers = $Response.Content | ConvertFrom-Json

if (!$QP_Customers) {
	Write-Host "Failed to connect to: QuickPass. No customers found."
	exit
}

$i = 1
while ($QP_Customers -and $QP_Customers.maxCount -gt $QP_Customers.clients.count -and $i -le $RateLimit) {
	$i++
	$Response = Invoke-WebRequest "$($QP_BaseURI)customer?page=$i&rowsPerPage=50&searchText=&adStatus=ALL&localStatus=ALL&o365Status=ALL&filterUpdated=false" -WebSession $QPWebSession
	$QP_Customers_ToAdd = $Response.Content | ConvertFrom-Json
	if ($QP_Customers_ToAdd -and $QP_Customers_ToAdd.clients) {
		$QP_Customers.clients += $QP_Customers_ToAdd.clients
	}
}

# Get QP to ITG matching info from QuickPass
$Headers = @{
	Integration = "itglue"
}
$Response = Invoke-WebRequest "$($QP_BaseURI)integrations/matched-customers?page=1&rowsPerPage=50&searchText=&integrationType=itglue" -WebSession $QPWebSession -Headers $Headers
$QP_ITGMatching = $Response.Content | ConvertFrom-Json

if ($QP_ITGMatching -and $QP_ITGMatching.maxCount -gt $QP_ITGMatching.companies.Count) {
	$i = 1
	while ($QP_ITGMatching -and $QP_ITGMatching.maxCount -gt $QP_ITGMatching.companies.count -and $i -le $RateLimit) {
		$i++
		$Response = Invoke-WebRequest "$($QP_BaseURI)integrations/matched-customers?page=$i&rowsPerPage=50&searchText=&integrationType=itglue" -WebSession $QPWebSession -Headers $Headers
		$QP_ITGMatching_ToAdd = $Response.Content | ConvertFrom-Json
		if ($QP_ITGMatching_ToAdd -and $QP_ITGMatching_ToAdd.companies) {
			$QP_ITGMatching.companies += $QP_ITGMatching_ToAdd.companies
		}
	}
}

# Get QP to Autotask matching info from QuickPass
$Headers = @{
	Integration = "dattoAutoTask"
}
$Response = Invoke-WebRequest "$($QP_BaseURI)integrations/matched-customers?page=1&rowsPerPage=50&searchText=&integrationType=dattoAutoTask" -WebSession $QPWebSession -Headers $Headers
$QP_ATMatching = $Response.Content | ConvertFrom-Json

if ($QP_ATMatching -and $QP_ATMatching.maxCount -gt $QP_ATMatching.companies.Count) {
	$i = 1
	while ($QP_ATMatching -and $QP_ATMatching.maxCount -gt $QP_ATMatching.companies.count -and $i -le $RateLimit) {
		$i++
		$Response = Invoke-WebRequest "$($QP_BaseURI)integrations/matched-customers?page=$i&rowsPerPage=50&searchText=&integrationType=dattoAutoTask" -WebSession $QPWebSession -Headers $Headers
		$QP_ATMatching_ToAdd = $Response.Content | ConvertFrom-Json
		if ($QP_ATMatching_ToAdd -and $QP_ATMatching_ToAdd.companies) {
			$QP_ATMatching.companies += $QP_ATMatching_ToAdd.companies
		}
	}
}

# Get all ITG and Autotask customers
$ITGlueCompanies = Get-ITGlueOrganizations -page_size 1000
if ($ITGlueCompanies -and $ITGlueCompanies.data) {
	$ITGlueCompanies = $ITGlueCompanies.data | Where-Object { $_.attributes.'organization-status-name' -eq "Active" }
}

if (!$ITGlueCompanies) {
	exit
}

$Autotask_Companies = @()
if ($AutotaskConnected) {
	$Autotask_Companies = Get-AutotaskAPIResource -Resource Companies -SimpleSearch "isActive eq 1"
}

$QPFixes = [System.Collections.ArrayList]::new()
$Organizations = [System.Collections.ArrayList]::new()

# Match QP customers to ITG organizations
foreach ($Customer in $QP_Customers.clients) {
	$ITGMatch = $false
	
	# First look for any itg matches made in QP
	$ITGMatches = $QP_ITGMatching.companies | Where-Object { $_.customers.id -contains $Customer.id }

	if ($ITGMatches) {
		# Narrow down
		if (($ITGMatches | Measure-Object).Count -gt 1) {
			$ITGMatches_Temp = $ITGMatches | Where-Object { $_.status -eq "Active" }
			if (($ITGMatches_Temp | Measure-Object).Count -gt 0) {
				$ITGMatches = $ITGMatches_Temp
			}
		}
		if (($ITGMatches | Measure-Object).Count -gt 1) {
			$ITGMatches_Temp = $ITGMatches | Where-Object { $_.name -like $Customer.name }
			if (($ITGMatches_Temp | Measure-Object).Count -gt 0) {
				$ITGMatches = $ITGMatches_Temp
			}
		}
		if (($ITGMatches | Measure-Object).Count -gt 1) {
			$ITGMatches_Temp = $ITGMatches | Where-Object { $_.name -like "*$($Customer.name)*" }
			if (($ITGMatches_Temp | Measure-Object).Count -gt 0) {
				$ITGMatches = $ITGMatches_Temp
			}
		}
		if (($ITGMatches | Measure-Object).Count -gt 1) {
			$ITGMatches_Temp = $ITGMatches | Where-Object { $Customer.name -like "*$($_.name)*" }
			if (($ITGMatches_Temp | Measure-Object).Count -gt 0) {
				$ITGMatches = $ITGMatches_Temp
			}
		}
		if (($ITGMatches | Measure-Object).Count -gt 1) {
			$ITGMatches_Temp = $ITGMatches | Where-Object { $_.type -eq "Customer" }
			if (($ITGMatches_Temp | Measure-Object).Count -gt 0) {
				$ITGMatches = $ITGMatches_Temp
			}
		}
		$ITGMatch = $ITGMatches | Select-Object -First 1

		if ($ITGMatch) {
			$ITGCompany = $ITGlueCompanies | Where-Object { $_.id -eq $ITGMatch.id }
			if ($ITGCompany) {
				$Organizations.Add(@{
					QP = $Customer
					ITG = $ITGCompany
				})
			}
			continue
		}
	}

	# If no matching with ITG is setup, try to find a match just by name, or create a warning
	if (!$ITGMatch -and $Customer.integrations.itglue.active -eq $true) {
		$ITGMatches = $ITGlueCompanies | Where-Object { $_.attributes.name -like $Customer.Name }
		if (!$ITGMatches) {
			$ITGMatches = $ITGlueCompanies | Where-Object { $_.attributes.name -like  "*$($Customer.name)*" }
		}
		if (!$ITGMatches) {
			$ITGMatches = $ITGlueCompanies | Where-Object { $Customer.name -like  "*$($_.attributes.name)*" }
		}

		if ($ITGMatches) {
			# Narrow down
			if (($ITGMatches | Measure-Object).Count -gt 1) {
				$ITGMatches_Temp = $ITGMatches | Where-Object { $_.attributes."organization-status-name" -eq "Active" }
				if (($ITGMatches_Temp | Measure-Object).Count -gt 0) {
					$ITGMatches = $ITGMatches_Temp
				}
			}
			if (($ITGMatches | Measure-Object).Count -gt 1) {
				$ITGMatches_Temp = $ITGMatches | Where-Object { $_.attributes."organization-type-name" -eq "Customer" }
				if (($ITGMatches_Temp | Measure-Object).Count -gt 0) {
					$ITGMatches = $ITGMatches_Temp
				}
			}
			$ITGMatch = $ITGMatches | Select-Object -First 1

			if ($ITGMatch) {
				$Organizations.Add(@{
					QP = $Customer
					ITG = $ITGMatch
				})
			}
		} else {
			$QPFixes.Add([PSCustomObject]@{
				Company = $Customer.Name
				id = $Customer.id
				Name = $Customer.Name
				FixType = "Could not find a customer match in ITG"
			})
			Write-Host "Could not find a customer match for $($Customer.Name) in ITG." -ForegroundColor Red
		}
	} else {
		$QPFixes.Add([PSCustomObject]@{
			Company = $Customer.Name
			id = $Customer.id
			Name = $Customer.Name
			FixType = "No ITG customer match is setup"
		})
		Write-Warning "No ITG customer match is setup for: $($Customer.Name)"
	}
}

# Match QP customers to Autotask organizations
foreach ($Customer in $QP_Customers.clients) {
	$ATMatch = $false
	
	# First look for any Autotask matches made in QP
	$ATMatches = $QP_ATMatching.companies | Where-Object { $_.customers.id -contains $Customer.id }

	if ($ATMatches) {
		# Narrow down
		if (($ATMatches | Measure-Object).Count -gt 1) {
			$ATMatches_Temp = $ATMatches | Where-Object { $_.status -eq "Active" }
			if (($ATMatches_Temp | Measure-Object).Count -gt 0) {
				$ATMatches = $ATMatches_Temp
			}
		}
		if (($ATMatches | Measure-Object).Count -gt 1) {
			$ATMatches_Temp = $ATMatches | Where-Object { $_.name -like $Customer.name }
			if (($ATMatches_Temp | Measure-Object).Count -gt 0) {
				$ATMatches = $ATMatches_Temp
			}
		}
		if (($ATMatches | Measure-Object).Count -gt 1) {
			$ATMatches_Temp = $ATMatches | Where-Object { $_.name -like "*$($Customer.name)*" }
			if (($ATMatches_Temp | Measure-Object).Count -gt 0) {
				$ATMatches = $ATMatches_Temp
			}
		}
		if (($ATMatches | Measure-Object).Count -gt 1) {
			$ATMatches_Temp = $ATMatches | Where-Object { $Customer.name -like "*$($_.name)*" }
			if (($ATMatches_Temp | Measure-Object).Count -gt 0) {
				$ATMatches = $ATMatches_Temp
			}
		}
		if (($ATMatches | Measure-Object).Count -gt 1) {
			$ATMatches_Temp = $ATMatches | Where-Object { $_.type -eq "Customer" }
			if (($ATMatches_Temp | Measure-Object).Count -gt 0) {
				$ATMatches = $ATMatches_Temp
			}
		}
		$ATMatch = $ATMatches | Select-Object -First 1

		if ($ATMatch) {
			$ATCompany = $Autotask_Companies | Where-Object { $_.id -eq $ATMatch.id }
			if ($ATCompany) {
				$ExistingOrg = $Organizations | Where-Object { $_.QP.id -eq $Customer.id }
				if ($ExistingOrg) {
					$ExistingOrg.AT = $ATCompany
				} else {
					$Organizations.Add(@{
						QP = $Customer
						AT = $ATCompany
					})
				}
			}
			continue
		}
	}

	# If no matching with Autotask is setup, try to find a match just by name, or create a warning
	if (!$ATMatch -and $Customer.integrations.dattoAutoTask.active -eq $true) {
		$ATMatches = $Autotask_Companies | Where-Object { $_.companyName -like $Customer.Name }
		if (!$ATMatches) {
			$ATMatches = $Autotask_Companies | Where-Object { $_.companyName -like  "*$($Customer.name)*" }
		}
		if (!$ATMatches) {
			$ATMatches = $Autotask_Companies | Where-Object { $Customer.name -like  "*$($_.companyName)*" }
		}
		if (!$ATMatches) {
			$ExistingOrg = $Organizations | Where-Object { $_.QP.id -eq $Customer.id }
			if ($ExistingOrg -and $ExistingOrg.ITG) {
				$ATMatches = $Autotask_Companies | Where-Object { $_.companyNumber -like $ExistingOrg.ITG.attributes.'short-name' }
			}
		}

		if ($ATMatches) {
			# Narrow down
			if (($ATMatches | Measure-Object).Count -gt 1) {
				$ExistingOrg = $Organizations | Where-Object { $_.QP.id -eq $Customer.id }
				if ($ExistingOrg -and $ExistingOrg.ITG) {
					$ATMatches_Temp = $ATMatches | Where-Object { $_.companyNumber -like $ExistingOrg.ITG.attributes.'short-name' }
					if (($ATMatches_Temp | Measure-Object).Count -gt 0) {
						$ATMatches = $ATMatches_Temp
					}
				}
			}
			if (($ATMatches | Measure-Object).Count -gt 1) {
				$ATMatches_Temp = $ATMatches | Where-Object { $_.companyType -eq 1 }
				if (($ATMatches_Temp | Measure-Object).Count -gt 0) {
					$ATMatches = $ATMatches_Temp
				}
			}
			$ATMatch = $ATMatches | Select-Object -First 1

			if ($ATMatch) {
				$ExistingOrg = $Organizations | Where-Object { $_.QP.id -eq $Customer.id }
				if ($ExistingOrg) {
					$ExistingOrg.AT = $ATCompany
				} else {
					$Organizations.Add(@{
						QP = $Customer
						AT = $ATCompany
					})
				}
			}
		} else {
			$QPFixes.Add([PSCustomObject]@{
				Company = $Customer.Name
				id = $Customer.id
				Name = $Customer.Name
				FixType = "Could not find a customer match in Autotask"
			})
			Write-Host "Could not find a customer match for $($Customer.Name) in Autotask." -ForegroundColor Red
		}
	} else {
		$QPFixes.Add([PSCustomObject]@{
			Company = $Customer.Name
			id = $Customer.id
			Name = $Customer.Name
			FixType = "No Autotask customer match is setup"
		})
		Write-Warning "No Autotask customer match is setup for: $($Customer.Name)"
	}
}

$Organizations | ConvertTo-Json -Depth 6 | Out-File -PSPath "./QPCustomerMatching.json" -Force

# Get QP to ITG match cache
$QPPasswordMatch_Cache = [PSCustomObject]@{}
if ((Test-Path -Path "./QPPasswordMatchingCache.json")) {
	$QPPasswordMatch_Cache = Get-Content -Raw -Path "./QPPasswordMatchingCache.json" | ConvertFrom-Json
} else {
	$QPPasswordMatch_Cache = [PSCustomObject]@{
		lastUpdated = $false
		customers = @{}
		firstTimeUpdates = [PSCustomObject]@{}
		matchAttempted = [PSCustomObject]@{}
	}
}

# Get ITG flex asset type id
$ITG_ADFilterID = (Get-ITGlueFlexibleAssetTypes -filter_name $ITG_ADFlexAssetName).data

if (!$ITG_ADFilterID) {
	Write-Error "Could not get the AD flex asset type ID. Exiting..."
	exit 1
}

# Get event lists since the cache was last updated so we can query it for ITG/Autotask match changes
if ($QPPasswordMatch_Cache.lastUpdated) {
	$CacheLastUpdated = Get-Date $QPPasswordMatch_Cache.lastUpdated
	$StartDate = Get-Date($CacheLastUpdated) -Format 'yyyy-MM-dd'
	$EndDate = Get-Date -Format 'yyyy-MM-dd'

	# ITG
	$Response = Invoke-WebRequest "$($QP_BaseURI)events?page=1&rowsPerPage=50&startDate=$($StartDate)&endDate=$($EndDate)&status=all&eventType=IT-GLUE&customer=all&timezoneOffset=480" -WebSession $QPWebSession
	$QP_ITG_EventLog = $Response.Content | ConvertFrom-Json

	if ($QP_ITG_EventLog -and $QP_ITG_EventLog.maxCount -gt $QP_ITG_EventLog.events.Count) {
		$i = 1
		while ($QP_ITG_EventLog -and $QP_ITG_EventLog.maxCount -gt $QP_ITG_EventLog.users.count -and $i -le $RateLimit) {
			$i++
			$Response = Invoke-WebRequest "$($QP_BaseURI)events?page=$i&rowsPerPage=50&startDate=$($StartDate)&endDate=$($EndDate)&status=all&eventType=IT-GLUE&customer=all&timezoneOffset=480" -WebSession $QPWebSession
			$QP_ITG_EventLog_ToAdd = $Response.Content | ConvertFrom-Json
			if ($QP_ITG_EventLog_ToAdd -and $QP_ITG_EventLog_ToAdd.events) {
				$QP_ITG_EventLog.events += $QP_ITG_EventLog_ToAdd.events
			}
		}
	}
	$QP_ITG_EventLog.events = $QP_ITG_EventLog.events | Where-Object { $_.details -like "*match*" }

	# Autotask
	$Response = Invoke-WebRequest "$($QP_BaseURI)events?page=1&rowsPerPage=50&startDate=$($StartDate)&endDate=$($EndDate)&status=all&eventType=DATTO-AUTOTASK&customer=all&timezoneOffset=480" -WebSession $QPWebSession
	$QP_AT_EventLog = $Response.Content | ConvertFrom-Json

	if ($QP_AT_EventLog -and $QP_AT_EventLog.maxCount -gt $QP_AT_EventLog.events.Count) {
		$i = 1
		while ($QP_AT_EventLog -and $QP_AT_EventLog.maxCount -gt $QP_AT_EventLog.users.count -and $i -le $RateLimit) {
			$i++
			$Response = Invoke-WebRequest "$($QP_BaseURI)events?page=$i&rowsPerPage=50&startDate=$($StartDate)&endDate=$($EndDate)&status=all&eventType=DATTO-AUTOTASK&customer=all&timezoneOffset=480" -WebSession $QPWebSession
			$QP_AT_EventLog_ToAdd = $Response.Content | ConvertFrom-Json
			if ($QP_AT_EventLog_ToAdd -and $QP_AT_EventLog_ToAdd.events) {
				$QP_AT_EventLog.events += $QP_AT_EventLog_ToAdd.events
			}
		}
	}
	$QP_AT_EventLog.events = $QP_AT_EventLog.events | Where-Object { $_.details -like "*match*" }
}

# Loop through each QP organization and audit users & settings
foreach ($Organization in $Organizations) {
	Write-Output "Auditing: $($Organization.QP.Name)"

	# Get customer details
	$Response = Invoke-WebRequest "$($QP_BaseURI)customer/$($Organization.QP.id)" -WebSession $QPWebSession
	$OrgDetails = $Response.Content | ConvertFrom-Json

	if ($OrgDetails) {
		$Organization.QPDetails = $OrgDetails
	}

	# Get rotation settings
	$Response = Invoke-WebRequest "$($QP_BaseURI)password-rotation/configuration?clientId=$($Organization.QP.id)&isServiceAccounts=false" -WebSession $QPWebSession
	$RotateSettings_Admin = $Response.Content | ConvertFrom-Json

	$Response = Invoke-WebRequest "$($QP_BaseURI)password-rotation/configuration?clientId=$($Organization.QP.id)&isServiceAccounts=true" -WebSession $QPWebSession
	$RotateSettings_Service = $Response.Content | ConvertFrom-Json

	# Update Rotation Settings to defaults if not set yet
	if ($RotateSettings_Admin -and !$RotateSettings_Admin.defaultFrequency) {
		$Body = @{
			assignSamePassword = $false
			clientId = $Organization.QP.id
			defaultFrequency = 30
			isServiceAccounts = $false
			minPasswordLenght = 8
			passphraseLength = "4L"
			passwordComplexity = "COMPLEX_PASSWORD"
			passwordLength = 20
			time = "01:00"
			timeZone = "America/Vancouver"
			usePassphrase = $false
		}

		$Params = @{
			Method = "Post"
			Uri = "$($QP_BaseURI)password-rotation/customer-configuration"
			Body = ($Body | ConvertTo-Json)
			ContentType = "application/json"
			WebSession = $QPWebSession
		}

		$RotateSettingsUpdated = $false
		try {
			Write-Output "Setting default Admin Rotation Settings for: $($Organization.QP.Name)"
			$Response = Invoke-RestMethod @Params
			Start-Sleep -Seconds 1
			if ($Response -and $Response.message -like "Configuration applied successfully.") {
				$RotateSettingsUpdated = $true
			}
		} catch {
			Start-Sleep -Seconds 3
		}

		if (!$RotateSettingsUpdated) {
			$QPFixes.Add([PSCustomObject]@{
				Company = $Organization.QP.name
				id = $Organization.QP.id
				Name = $Organization.QP.name
				FixType = "Rotation Settings - Admin need to be Configured, auto configuration failed."
			})
		}
	}

	if ($RotateSettings_Service -and !$RotateSettings_Service.defaultFrequency) {
		$Body = @{
			assignSamePassword = $false
			clientId = $Organization.QP.id
			defaultFrequency = 30
			isServiceAccounts = $true
			minPasswordLenght = 8
			passphraseLength = "4L"
			passwordComplexity = "COMPLEX_PASSWORD"
			passwordLength = 20
			time = "01:00"
			timeZone = "America/Vancouver"
			usePassphrase = $false
		}

		$Params = @{
			Method = "Post"
			Uri = "$($QP_BaseURI)password-rotation/customer-configuration"
			Body = ($Body | ConvertTo-Json)
			ContentType = "application/json"
			WebSession = $QPWebSession
		}

		$RotateSettingsUpdated = $false
		try {
			Write-Output "Setting default Service Account Rotation Settings for: $($Organization.QP.Name)"
			$Response = Invoke-RestMethod @Params
			Start-Sleep -Seconds 1
			if ($Response -and $Response.message -like "Configuration applied successfully.") {
				$RotateSettingsUpdated = $true
			}
		} catch {
			Start-Sleep -Seconds 3
		}

		if (!$RotateSettingsUpdated) {
			$QPFixes.Add([PSCustomObject]@{
				Company = $Organization.QP.name
				id = $Organization.QP.id
				Name = $Organization.QP.name
				FixType = "Rotation Settings - Service Accounts need to be Configured, auto configuration failed."
			})
		}
	}

	if ($Organization.ITG) {
		# Get AD servers (from ITG) and check for any without a QP agent
		$ITG_ADAssets = (Get-ITGlueFlexibleAssets -filter_flexible_asset_type_id $ITG_ADFilterID.id -filter_organization_id $Organization.ITG.id).data
		if ($ITG_ADAssets) {
			$ITG_ADAssets = $ITG_ADAssets | Where-Object { !$_.attributes.archived }
		}
		$ITG_ADServers = $ITG_ADAssets | ForEach-Object { $_.attributes.traits.'ad-servers' } | ForEach-Object { $_.values }

		if ($ITG_ADServers) {
			$Response = Invoke-WebRequest "$($QP_BaseURI)customer/agents/$($Organization.QP.id)/?page=1&rowsPerPage=50&searchText=" -WebSession $QPWebSession
			$QPAgents = $Response.Content | ConvertFrom-Json

			foreach ($ADServer in $ITG_ADServers) {
				if ($ADServer.name -notin $QPAgents.agents.serverName) {
					# Found server missing a QP agent
					$ServerAsset = (Get-ITGlueConfigurations -id $ADServer.id -organization_id $Organization.ITG.id).data

					if ($ServerAsset -and !$ServerAsset.attributes.archived -and $ServerAsset.attributes.notes -notlike "*# Ignore QP Agent Installs*") {
						Write-Output "Found AD Server in ITG without a QP Agent: $($ADServer.name)"
						$QPFixes.Add([PSCustomObject]@{
							Company = $Organization.QP.name
							id = $Organization.QP.id
							Name = $Organization.QP.name
							FixType = "Found AD Server Missing the QP Agent: $($ADServer.name)  (To Ignore, add '# Ignore QP Agent Installs' to the config notes in ITG)"
						})
					}
				}
			}
		}
	}

	# Get QP end user list
	$Response = Invoke-WebRequest "$($QP_BaseURI)customer/$($Organization.QP.id)/users?status=allActive&accountStatus=all&userAppType=all&page=1&type=standard&rowsPerPage=50&qpStatus=all&searchText=" -WebSession $QPWebSession
	$QP_EndUsers = $Response.Content | ConvertFrom-Json

	if ($QP_EndUsers -and $QP_EndUsers.maxCount -gt $QP_EndUsers.users.Count) {
		$i = 1
		while ($QP_EndUsers -and $QP_EndUsers.maxCount -gt $QP_EndUsers.users.count -and $i -le $RateLimit) {
			$i++
			$Response = Invoke-WebRequest "$($QP_BaseURI)customer/$($Organization.QP.id)/users?status=allActive&accountStatus=all&userAppType=all&page=$i&type=standard&rowsPerPage=50&qpStatus=all&searchText=" -WebSession $QPWebSession
			$QP_EndUsers_ToAdd = $Response.Content | ConvertFrom-Json
			if ($QP_EndUsers_ToAdd -and $QP_EndUsers_ToAdd.users) {
				$QP_EndUsers.users += $QP_EndUsers_ToAdd.users
			}
		}
	}

	# Get Autotask contacts
	if ($AutotaskConnected -and $Organization.AT -and $Organization.AT.id) {
		$Autotask_Contacts = Get-AutotaskAPIResource -Resource Contacts -SimpleSearch "companyID eq $($Organization.AT.id)"
	}

	# Add customer to the cache if missing
	if (!$QPPasswordMatch_Cache.customers.($Organization.QP.id)) {
		$QPPasswordMatch_Cache.customers | Add-Member -MemberType NoteProperty -Name $Organization.QP.id -Value $null -ErrorAction Ignore
		$QPPasswordMatch_Cache.customers.($Organization.QP.id) = [PSCustomObject]@{}
	}
	if (!$QPPasswordMatch_Cache.firstTimeUpdates.($Organization.QP.id)) {
		$QPPasswordMatch_Cache.firstTimeUpdates | Add-Member -MemberType NoteProperty -Name $Organization.QP.id -Value $null -ErrorAction Ignore
		$QPPasswordMatch_Cache.firstTimeUpdates.($Organization.QP.id) = @()
	}
	if (!$QPPasswordMatch_Cache.matchAttempted.($Organization.QP.id)) {
		$QPPasswordMatch_Cache.matchAttempted | Add-Member -MemberType NoteProperty -Name $Organization.QP.id -Value $null -ErrorAction Ignore
		$QPPasswordMatch_Cache.matchAttempted.($Organization.QP.id) = [PSCustomObject]@{}
	}

	# Look for any passwords not connected to ITG passwords or Autotask contacts and try to match them
	$No_ITGIntegration = $QP_EndUsers.users | Where-Object { !$_.integrations.itglue.active }
	$No_ATIntegration = $QP_EndUsers.users | Where-Object { !$_.integrations.dattoAutoTask.active }

	$No_ITGIntegration = $No_ITGIntegration | Where-Object { $_.qpID -notin $QPPasswordMatch_Cache.matchAttempted.($Organization.QP.id).PSObject.Properties.Name -or 'ITG' -notin $QPPasswordMatch_Cache.matchAttempted.($Organization.QP.id).($_.qpID) }
	$No_ATIntegration = $No_ATIntegration | Where-Object { $_.qpID -notin $QPPasswordMatch_Cache.matchAttempted.($Organization.QP.id).PSObject.Properties.Name -or 'AT' -notin $QPPasswordMatch_Cache.matchAttempted.($Organization.QP.id).($_.qpID) }

	$Updates = 0
	if ($Organization.ITG -and ($No_ITGIntegration | Measure-Object).Count -gt 0) {
		# Found passwords without an ITG match

		# Get ITG passwords
		$ITGPasswords = Get-ITGluePasswords -page_size 1000 -organization_id $Organization.ITG.id 
		$i = 1
		while ($ITGPasswords.links.next) {
			$i++
			$Passwords_Next = Get-ITGluePasswords -page_size 1000 -page_number $i -organization_id $Organization.ITG.id
			if (!$Passwords_Next -or $Passwords_Next.Error) {
				# We got an error querying passwords, wait and try again
				Start-Sleep -Seconds 2
				$Passwords_Next = Get-ITGluePasswords -page_size 1000 -page_number $i -organization_id $Organization.ITG.id
		
				if (!$Passwords_Next -or $Passwords_Next.Error) {
					Write-Error "An error occurred trying to get the existing passwords from ITG."
					Write-Error $Passwords_Next.Error
				}
			}
			$ITGPasswords.data += $Passwords_Next.data
			$ITGPasswords.links = $Passwords_Next.links
		}
		if ($ITGPasswords -and $ITGPasswords.data) {
			$ITGPasswords = $ITGPasswords.data
		}

		if (!$ITGPasswords) {
			Write-Error "Could not find any passwords in ITG. Skipping..."
			continue
		}

		# Find best match (if one exists)
		foreach ($QPUser in $No_ITGIntegration) {
			$QPEmail = $QPUser.email.Replace("*", "?").Replace("[", "?").Replace("]", "?")
			$QPSamAccountName = $QPUser.samAccountName.Replace("*", "?").Replace("[", "?").Replace("]", "?")
			$QPUserPrincipalName = $QPUser.userPrincipalName.Replace("*", "?").Replace("[", "?").Replace("]", "?")

			if (!$QPEmail) {
				$QPEmail = "THIS WILL NOT MATCH"
			}
			$QPEmailStart = "THIS WILL NOT MATCH"
			if ($QPEmail -and $QPEmail -like "*@*") {
				$QPEmailStart = ($QPEmail -split "@")[0]
				if (!$QPEmailStart) {
					$QPEmailStart = "THIS WILL NOT MATCH"
				}
			}
			if (!$QPSamAccountName) {
				$QPSamAccountName = "THIS WILL NOT MATCH"
			}
			if (!$QPUserPrincipalName) {
				$QPUserPrincipalName = "THIS WILL NOT MATCH"
			}

			$Related_ITGPasswords = $ITGPasswords | Where-Object { 
				($_.attributes.username -and $_.attributes.username -like $QPEmail) -or
				($_.attributes.username -and ($_.attributes.username -like $QPEmailStart -or $_.attributes.username -like "$($QPEmailStart)@*" -or $_.attributes.username -like "*\$($QPEmailStart)")) -or
				($_.attributes.username -and ($_.attributes.username -like $QPSamAccountName -or $_.attributes.username -like "$($QPSamAccountName)@*" -or $_.attributes.username -like "*\$($QPSamAccountName)")) -or
				($_.attributes.username -and ($_.attributes.username -like $QPUserPrincipalName -or $_.attributes.username -like "$($QPUserPrincipalName)@*" -or $_.attributes.username -like "*\$($QPUserPrincipalName)")) -or
				($_.attributes.name -and ($_.attributes.name -like $QPUser.displayName -or $_.attributes.name -like "*$($QPUser.displayName)*"))
			}
			$Related_ITGPasswords = $Related_ITGPasswords | Where-Object { !$_.attributes.archived }

			$FilterTypes = @()
			if (($Related_ITGPasswords | Measure-Object).Count -gt 0) {
				# Filter by password type
				if ($QPUser.userType -eq "AD" -or $QPUser.integrations.activeDirectory.active) {
					$FilterTypes += "AD"
				} elseif ($QPUser.userType -eq "OFFICE" -or $QPUser.integrations.office365.active) {
					$FilterTypes += "O365"
				}

				$Related_ITGPasswords = $Related_ITGPasswords | Where-Object {
					$Success = $false
					if ($FilterTypes -contains "AD" -and !$Success) {
						$Success = $_.attributes.'password-category-name' -like "Active Directory*" -or $_.attributes.name -like "AD *" -or $_.attributes.name -like "* AD *" -or $_.attributes.name -like "* Active Directory *"
					}
					if ($FilterTypes -contains "O365" -and !$Success) {
						$Success = $_.attributes.'password-category-name' -in $EmailCategories -or $_.attributes.name -like "AAD *" -or $_.attributes.name -like "* AAD *" -or
						$_.attributes.name -like "O365 *" -or $_.attributes.name -like "* O365 *" -or
						$_.attributes.name -like "M365 *" -or $_.attributes.name -like "* M365 *" -or
						$_.attributes.name -like "Email *" -or $_.attributes.name -like "* Email *" -or
						$_.attributes.name -like "Azure AD*" -or $_.attributes.name -like "*Azure AD*" -or
						$_.attributes.name -like "AzureAD *" -or $_.attributes.name -like "* AzureAD *" -or
						$_.attributes.name -like "O365 Email*" -or $_.attributes.name -like "*O365 Email*" -or
						$_.attributes.name -like "Office 365*" -or $_.attributes.name -like "*Office 365*" -or
						$_.attributes.name -like "Office365 *" -or $_.attributes.name -like "* Office365 *" -or
						$_.attributes.name -like "Microsoft 365*" -or $_.attributes.name -like "*Microsoft 365*"
					}

					return $Success
				}
			}

			# Narrow down if necessary
			if (($Related_ITGPasswords | Measure-Object).Count -gt 1) {
				$Related_ITGPasswords_Temp = @()
				if ($FilterTypes -contains "AD" -and $FilterTypes -contains "O365") {
					$Related_ITGPasswords_Temp = $Related_ITGPasswords | Where-Object {
						($_.attributes.username -and (
							$_.attributes.username -like $QPEmail -or $_.attributes.username -like $QPEmailStart -or $_.attributes.username -like "$($QPEmailStart)@*" -or 
							$_.attributes.username -like $QPUserPrincipalName -or $_.attributes.username -like "$($QPUserPrincipalName)@*" -or $_.attributes.username -like "*\$($QPUserPrincipalName)" -or
							$_.attributes.username -like $QPSamAccountName -or $_.attributes.username -like "$($QPSamAccountName)@*" -or $_.attributes.username -like "*\$($QPSamAccountName)"
						))
				}
				} elseif ($FilterTypes -contains "AD") {
					$Related_ITGPasswords_Temp = $Related_ITGPasswords | Where-Object {
						($_.attributes.username -and (
							$_.attributes.username -like $QPEmailStart -or $_.attributes.username -like "$($QPEmailStart)@*" -or 
							$_.attributes.username -like $QPUserPrincipalName -or $_.attributes.username -like "$($QPUserPrincipalName)@*" -or $_.attributes.username -like "*\$($QPUserPrincipalName)" -or
							$_.attributes.username -like $QPSamAccountName -or $_.attributes.username -like "$($QPSamAccountName)@*" -or $_.attributes.username -like "*\$($QPSamAccountName)"
						))
					}
				} elseif ($FilterTypes -contains "O365") {
					$Related_ITGPasswords_Temp = $Related_ITGPasswords | Where-Object {
						($_.attributes.username -and (
							$_.attributes.username -like $QPEmail -or $_.attributes.username -like $QPEmailStart -or $_.attributes.username -like "$($QPEmailStart)@*" -or 
							$_.attributes.username -like $QPUserPrincipalName -or $_.attributes.username -like "$($QPUserPrincipalName)@*"
						))
					}
				}
				if (($Related_ITGPasswords_Temp | Measure-Object).Count -gt 0) {
					$Related_ITGPasswords = $Related_ITGPasswords_Temp
				}

				if (($Related_ITGPasswords | Measure-Object).Count -gt 1) {
					$Related_ITGPasswords_Temp = $Related_ITGPasswords | Where-Object {
						$_.attributes.name -and $_.attributes.name -like "*$($QPUser.displayName)*"
					}
					if (($Related_ITGPasswords_Temp | Measure-Object).Count -gt 0) {
						$Related_ITGPasswords = $Related_ITGPasswords_Temp
					}
				}

				if (($Related_ITGPasswords | Measure-Object).Count -gt 1) {
					if ($FilterTypes -contains "AD") {
						$Related_ITGPasswords_Temp = $Related_ITGPasswords | Where-Object {
							$_.attributes.'password-category-name' -like "Active Directory*" 
						}
					} else {
						$Related_ITGPasswords_Temp = $Related_ITGPasswords | Where-Object {
							$_.attributes.'password-category-name' -in $EmailCategories
						}
					}
					if (($Related_ITGPasswords_Temp | Measure-Object).Count -gt 0) {
						$Related_ITGPasswords = $Related_ITGPasswords_Temp
					}
				}

				if (($Related_ITGPasswords | Measure-Object).Count -gt 1) {
					if ($FilterTypes -contains "AD") {
						$Related_ITGPasswords_Temp = $Related_ITGPasswords | Where-Object {
							$_.attributes.name -like "AD *" -or $_.attributes.name -like "* AD *"
						}
					} elseif ($FilterTypes -contains "O365") {
						$Related_ITGPasswords_Temp = $Related_ITGPasswords | Where-Object {
							$_.attributes.name -like "O365 *" -or $_.attributes.name -like "* O365 *" -or
							$_.attributes.name -like "M365 *" -or $_.attributes.name -like "* M365 *"
						}
					} elseif ($FilterTypes -contains "AD" -and $FilterTypes -contains "O365") {
						$Related_ITGPasswords_Temp = $Related_ITGPasswords | Where-Object {
							($_.attributes.name -like "AD *" -or $_.attributes.name -like "* AD *") -and
							($_.attributes.name -like "O365 *" -or $_.attributes.name -like "* O365 *" -or $_.attributes.name -like "M365 *" -or $_.attributes.name -like "* M365 *")
						}
					}
					if (($Related_ITGPasswords_Temp | Measure-Object).Count -gt 0) {
						$Related_ITGPasswords = $Related_ITGPasswords_Temp
					}
				}
			}

			# Either add the new match, or send an email with a matching suggestion (if a subpar match)
			if (($Related_ITGPasswords | Measure-Object).Count -gt 1) {
				# Too many matches found, send an email with suggestions
				$Suggestions = ($Related_ITGPasswords | ForEach-Object { "$($_.attributes.name) ($($_.id))" }) -join ", "
				$QPFixes.Add([PSCustomObject]@{
					Company = $Organization.QP.name
					id = $QPUser.qpId
					Name = $QPUser.displayName
					FixType = "Manual ITG Match Required, found multiple suggestions: $($Suggestions)"
				})
			} elseif (($Related_ITGPasswords | Measure-Object).Count -gt 0) {
				# Only 1 match found and it has a name match, looks safe, auto update match!
				$Related_ITGPassword = $Related_ITGPasswords | Select-Object -First 1

				$Headers = @{
					Integration = "itglue"
				}
				$Body = @{
					allSelected = $false
					customerId = $Organization.QP.id
					excludedIds = @()
					includedIds = @()
					integrationType = "itglue"
					matchType = "ALL"
					matches = @{
						$QPUser.qpId = @{
							autoMatch = $false
							displayName = $QPUser.displayName
							id = "$($Related_ITGPassword.id)"
							name = $Related_ITGPassword.attributes.name
							_id = $QPUser.qpId
						}
					}
					searchText = ""
					userType = "standard"
				}

				$Params = @{
					Method = "Post"
					Uri = "$($QP_BaseURI)integrations/accounts/matches/createjob"
					Body = ($Body | ConvertTo-Json -Depth 10)
					ContentType = "application/json"
					Headers = $Headers
					WebSession = $QPWebSession
				}

				$AutoUpdated = $false
				try {
					Write-Output "Creating new ITG match: $($QPUser.displayName) to $($Related_ITGPassword.attributes.name)"
					$Response = Invoke-RestMethod @Params
					Start-Sleep -Seconds 1
					if ($Response -and $Response.message -like "Matching Changes Submitted") {
						$Updates++
						$AutoUpdated = $true
					}
				} catch {
					$AutoUpdated = $false
					Start-Sleep -Seconds 3
				}

				if (!$AutoUpdated) {
					# Auto update did not work, send an email
					$Suggestion = "$($Related_ITGPassword.attributes.name) ($($Related_ITGPassword.id))"
					$QPFixes.Add([PSCustomObject]@{
						Company = $Organization.QP.name
						id = $QPUser.qpId
						Name = $QPUser.displayName
						FixType = "Manual ITG Match Required, automatch failed: $($Suggestion)"
					})
				}
			}

			if (($Related_ITGPasswords | Measure-Object).Count -gt 0) {
				$QPPasswordMatch_Cache.matchAttempted.($Organization.QP.id) | Add-Member -MemberType NoteProperty -Name $QPUser.qpId -Value @() -ErrorAction Ignore
				$QPPasswordMatch_Cache.matchAttempted.($Organization.QP.id).($QPUser.qpId) += "ITG"
			}
		}
	}

	if (($No_ATIntegration | Measure-Object).Count -gt 0) {
		# Found passwords without an Autotask match

		# Find best match (if one exists)
		:userLoop foreach ($QPUser in $No_ATIntegration) {	
			$QPEmail = $QPUser.email.Replace("*", "?").Replace("[", "?").Replace("]", "?")
			$QPSamAccountName = $QPUser.samAccountName.Replace("*", "?").Replace("[", "?").Replace("]", "?")
			$QPUserPrincipalName = $QPUser.userPrincipalName.Replace("*", "?").Replace("[", "?").Replace("]", "?")

			if (!$QPEmail) {
				$QPEmail = "THIS WILL NOT MATCH"
			}
			$QPEmailStart = "THIS WILL NOT MATCH"
			if ($QPEmail -and $QPEmail -like "*@*") {
				$QPEmailStart = ($QPEmail -split "@")[0]
				if (!$QPEmailStart) {
					$QPEmailStart = "THIS WILL NOT MATCH"
				}
			}
			if (!$QPSamAccountName) {
				$QPSamAccountName = "THIS WILL NOT MATCH"
			}
			if (!$QPUserPrincipalName) {
				$QPUserPrincipalName = "THIS WILL NOT MATCH"
			}

			$Related_ATUsers_InitSearch = $Autotask_Contacts | Where-Object { 
				($_.emailAddress -and ($_.emailAddress -like $QPEmail -or $_.emailAddress -like "$($QPSamAccountName)@*" -or $_.emailAddress -like "$($QPUserPrincipalName)@*" -or $_.emailAddress -like "$($QPEmailStart)@*")) -or
				($_.emailAddress2 -and ($_.emailAddress2 -like $QPEmail -or $_.emailAddress2 -like "$($QPSamAccountName)@*" -or $_.emailAddress2 -like "$($QPUserPrincipalName)@*" -or $_.emailAddress2 -like "$($QPEmailStart)@*")) -or
				($_.emailAddress3 -and ($_.emailAddress3 -like $QPEmail -or $_.emailAddress3 -like "$($QPSamAccountName)@*" -or $_.emailAddress3 -like "$($QPUserPrincipalName)@*" -or $_.emailAddress3 -like "$($QPEmailStart)@*")) -or
				(($_.firstName -or $_.lastName) -and (($QPUser.displayName -like "*$($_.firstName)*" -and $QPUser.displayName -like "*$($_.lastName)*") -or $QPUser.displayName -like "*$($_.firstName) $($_.lastName)*" -or $QPUser.displayName -like $_.firstName))
			}
			$Related_ATUsers = $Related_ATUsers_InitSearch | Where-Object { $_.isActive }

			# Narrow down if necessary
			if (($Related_ATUsers | Measure-Object).Count -gt 1) {
				$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
					(($_.firstName -or $_.lastName) -and (($QPUser.displayName -like "*$($_.firstName)*" -and $QPUser.displayName -like "*$($_.lastName)*") -or $QPUser.displayName -like "*$($_.firstName) $($_.lastName)*"))
				}
				if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
					$Related_ATUsers = $Related_ATUsers_Temp
				}

				if (($Related_ATUsers | Measure-Object).Count -gt 1) {
					$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
						$_.firstName -and $QPUser.displayName -like $_.firstName -and (!$_.lastName -or $_.lastName -like ".")
					}
					if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
						$Related_ATUsers = $Related_ATUsers_Temp
					}
				}

				if (($Related_ATUsers | Measure-Object).Count -gt 1) {
					$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
						$_.lastName -and $QPUser.displayName -like "*$($_.lastName)*"
					}
					if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
						$Related_ATUsers = $Related_ATUsers_Temp
					}
				}

				if (($Related_ATUsers | Measure-Object).Count -gt 1) {
					$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
						$_.lastName -and $_.lastName -notlike "*(Old)*"
					}
					if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
						$Related_ATUsers = $Related_ATUsers_Temp
					}
				}

				if (($Related_ATUsers | Measure-Object).Count -gt 1) {
					$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
						$_.lastName -and $_.lastName -notlike "*Disabled*"
					}
					if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
						$Related_ATUsers = $Related_ATUsers_Temp
					}
				}

				if (($Related_ATUsers | Measure-Object).Count -gt 1) {
					$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
						($_.emailAddress -and $_.emailAddress -like $QPEmail)
					}
					if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
						$Related_ATUsers = $Related_ATUsers_Temp
					}
				}

				if (($Related_ATUsers | Measure-Object).Count -gt 1) {
					$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
						($_.emailAddress -and ($_.emailAddress -like "$($QPSamAccountName)@*" -or $_.emailAddress -like "$($QPUserPrincipalName)@*" -or $_.emailAddress -like "$($QPEmailStart)@*"))
					}
					if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
						$Related_ATUsers = $Related_ATUsers_Temp
					}
				}

				if (($Related_ATUsers | Measure-Object).Count -gt 1) {
					$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
						($_.emailAddress2 -and $_.emailAddress2 -like $QPEmail) -or
						($_.emailAddress3 -and $_.emailAddress3 -like $QPEmail)
					}
					if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
						$Related_ATUsers = $Related_ATUsers_Temp
					}
				}

				if (($Related_ATUsers | Measure-Object).Count -gt 1) {
					$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
						($_.emailAddress2 -and ($_.emailAddress2 -like "$($QPSamAccountName)@*" -or $_.emailAddress2 -like "$($QPUserPrincipalName)@*" -or $_.emailAddress2 -like "$($QPEmailStart)@*")) -or
						($_.emailAddress3 -and ($_.emailAddress3 -like "$($QPSamAccountName)@*" -or $_.emailAddress3 -like "$($QPUserPrincipalName)@*" -or $_.emailAddress3 -like "$($QPEmailStart)@*"))
					}
					if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
						$Related_ATUsers = $Related_ATUsers_Temp
					}
				}
			}

			# Either add the new match, send an email with a matching suggestion (if a subpar match), or create a new contact
			if (($Related_ATUsers | Measure-Object).Count -gt 1) {
				# Too many matches found, send an email with suggestions
				$Suggestions = ($Related_ATUsers | ForEach-Object { "$($_.firstName) $($_.lastName) ($($_.id))" }) -join ", "
				$QPFixes.Add([PSCustomObject]@{
					Company = $Organization.QP.name
					id = $QPUser.qpId
					Name = $QPUser.displayName
					FixType = "Manual AT Match Required, found multiple suggestions: $($Suggestions)"
				})
			} elseif (($Related_ATUsers | Measure-Object).Count -gt 0) {
				$Related_ATUser = $Related_ATUsers | Select-Object -First 1
				$UnsafeMatch = $false

				if ($Related_ATUser.firstName -and $QPUser.displayName -notlike "*$($Related_ATUser.firstName)*") {
					$UnsafeMatch = $true
				} elseif ($Related_ATUser.lastName -and $QPUser.displayName -notlike "*$($Related_ATUser.lastName)*") {
					$UnsafeMatch = $true
				}
				if ($Related_ATUser.firstName -and $Related_ATUser.lastName -and $QPUser.displayName -like "*$($Related_ATUser.firstName) $($Related_ATUser.lastName)*") {
					$UnsafeMatch = $false
				}
				if ($Related_ATUser.firstName -and $QPUser.displayName -like $Related_ATUser.firstName -and (!$Related_ATUser.lastName -or $Related_ATUser.lastName -like ".")) {
					$UnsafeMatch = $false
				}

				if ($UnsafeMatch) {
					# No name match, send an email with the suggestion to be safe
					$Suggestion = "$($Related_ATUser.firstName) $($Related_ATUser.lastName) ($($Related_ATUser.id))"
					$QPFixes.Add([PSCustomObject]@{
						Company = $Organization.QP.name
						id = $QPUser.qpId
						Name = $QPUser.displayName
						FixType = "Manual AT Match Required, name doesn't match: $($Suggestion)"
					})
				} else {
					# Only 1 match found and it has a name match, looks safe, auto update match!
					$Headers = @{
						Integration = "dattoAutoTask"
					}
					$Body = @{
						allSelected = $false
						customerId = $Organization.QP.id
						excludedIds = @()
						includedIds = @()
						integrationType = "dattoAutoTask"
						matchType = "ALL"
						matches = @{
							$QPUser.qpId = @{
								autoMatch = $false
								displayName = $QPUser.displayName
								id = "$($Related_ATUser.id)"
								name = "$($Related_ATUser.firstName) $($Related_ATUser.lastName)"
								_id = $QPUser.qpId
							}
						}
						searchText = ""
						userType = "standard"
					}

					$Params = @{
						Method = "Post"
						Uri = "$($QP_BaseURI)integrations/accounts/matches/createjob"
						Body = ($Body | ConvertTo-Json -Depth 10)
						ContentType = "application/json"
						Headers = $Headers
						WebSession = $QPWebSession
					}

					$AutoUpdated = $false
					try {
						Write-Output "Creating new AT match: $($QPUser.displayName) to $($Related_ATUser.firstName) $($Related_ATUser.lastName)"
						$Response = Invoke-RestMethod @Params
						Start-Sleep -Seconds 1
						if ($Response -and $Response.message -like "Matching Changes Submitted") {
							$Updates++
							$AutoUpdated = $true
						}
					} catch {
						$AutoUpdated = $false
						Start-Sleep -Seconds 3
					}

					if (!$AutoUpdated) {
						# Auto update did not work, send an email
						$Suggestion = "$($Related_ATUser.firstName) $($Related_ATUser.lastName) ($($Related_ATUser.id))"
						$QPFixes.Add([PSCustomObject]@{
							Company = $Organization.QP.name
							id = $QPUser.qpId
							Name = $QPUser.displayName
							FixType = "Manual AT Match Required, automatch failed: $($Suggestion)"
						})
					}
				}
			} elseif (($Related_ATUsers | Measure-Object).Count -eq 0 -and $QPUser.enabled) {
				# No related matches found at all, consider making a new contact in AT if this QP contact is enabled

				$CreateNew = $true
				if (($Related_ATUsers_InitSearch | Measure-Object).Count -gt 0) {
					# Check if there is an existing inactive AT contact
					$Related_ATUsers = $Related_ATUsers_InitSearch | Where-Object { !$_.isActive }

					# Narrow down if necessary
					if (($Related_ATUsers | Measure-Object).Count -gt 1) {
						$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
							(($_.firstName -or $_.lastName) -and (($QPUser.displayName -like "*$($_.firstName)*" -and $QPUser.displayName -like "*$($_.lastName)*") -or $QPUser.displayName -like "*$($_.firstName) $($_.lastName)*"))
						}
						if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
							$Related_ATUsers = $Related_ATUsers_Temp
						}

						if (($Related_ATUsers | Measure-Object).Count -gt 1) {
							$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
								$_.firstName -and $QPUser.displayName -like $_.firstName -and (!$_.lastName -or $_.lastName -like ".")
							}
							if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
								$Related_ATUsers = $Related_ATUsers_Temp
							}
						}

						if (($Related_ATUsers | Measure-Object).Count -gt 1) {
							$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
								$_.lastName -and $QPUser.displayName -like "*$($_.lastName)*"
							}
							if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
								$Related_ATUsers = $Related_ATUsers_Temp
							}
						}

						if (($Related_ATUsers | Measure-Object).Count -gt 1) {
							$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
								$_.lastName -and $_.lastName -notlike "*(Old)*"
							}
							if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
								$Related_ATUsers = $Related_ATUsers_Temp
							}
						}

						if (($Related_ATUsers | Measure-Object).Count -gt 1) {
							$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
								$_.lastName -and $_.lastName -notlike "*Disabled*"
							}
							if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
								$Related_ATUsers = $Related_ATUsers_Temp
							}
						}

						if (($Related_ATUsers | Measure-Object).Count -gt 1) {
							$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
								($_.emailAddress -and $_.emailAddress -like $QPEmail)
							}
							if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
								$Related_ATUsers = $Related_ATUsers_Temp
							}
						}

						if (($Related_ATUsers | Measure-Object).Count -gt 1) {
							$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
								($_.emailAddress -and ($_.emailAddress -like "$($QPSamAccountName)@*" -or $_.emailAddress -like "$($QPUserPrincipalName)@*" -or $_.emailAddress -like "$($QPEmailStart)@*"))
							}
							if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
								$Related_ATUsers = $Related_ATUsers_Temp
							}
						}

						if (($Related_ATUsers | Measure-Object).Count -gt 1) {
							$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
								($_.emailAddress2 -and $_.emailAddress2 -like $QPEmail) -or
								($_.emailAddress3 -and $_.emailAddress3 -like $QPEmail)
							}
							if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
								$Related_ATUsers = $Related_ATUsers_Temp
							}
						}

						if (($Related_ATUsers | Measure-Object).Count -gt 1) {
							$Related_ATUsers_Temp = $Related_ATUsers | Where-Object {
								($_.emailAddress2 -and ($_.emailAddress2 -like "$($QPSamAccountName)@*" -or $_.emailAddress2 -like "$($QPUserPrincipalName)@*" -or $_.emailAddress2 -like "$($QPEmailStart)@*")) -or
								($_.emailAddress3 -and ($_.emailAddress3 -like "$($QPSamAccountName)@*" -or $_.emailAddress3 -like "$($QPUserPrincipalName)@*" -or $_.emailAddress3 -like "$($QPEmailStart)@*"))
							}
							if (($Related_ATUsers_Temp | Measure-Object).Count -gt 0) {
								$Related_ATUsers = $Related_ATUsers_Temp
							}
						}
					}

					if (($Related_ATUsers | Measure-Object).Count -gt 1) {
						# Too many inactive matches found, send an email with suggestions
						$Suggestions = ($Related_ATUsers | ForEach-Object { "$($_.firstName) $($_.lastName) ($($_.id))" }) -join ", "
						$QPFixes.Add([PSCustomObject]@{
							Company = $Organization.QP.name
							id = $QPUser.qpId
							Name = $QPUser.displayName
							FixType = "Manual AT Match Required, found multiple suggestions for inactive AT contacts: $($Suggestions) [consider enabling & matching one]"
						})
						$CreateNew = $false
					} elseif (($Related_ATUsers | Measure-Object).Count -gt 0) {
						$Related_ATUser = $Related_ATUsers | Select-Object -First 1
						$UnsafeMatch = $false

						if ($Related_ATUser.firstName -and $QPUser.displayName -notlike "*$($Related_ATUser.firstName)*") {
							$UnsafeMatch = $true
						} elseif ($Related_ATUser.lastName -and $QPUser.displayName -notlike "*$($Related_ATUser.lastName)*") {
							$UnsafeMatch = $true
						}
						if ($Related_ATUser.firstName -and $Related_ATUser.lastName -and $QPUser.displayName -like "*$($Related_ATUser.firstName) $($Related_ATUser.lastName)*") {
							$UnsafeMatch = $false
						}
						if ($Related_ATUser.firstName -and $QPUser.displayName -like $Related_ATUser.firstName -and (!$Related_ATUser.lastName -or $Related_ATUser.lastName -like ".")) {
							$UnsafeMatch = $false
						}

						if ($UnsafeMatch) {
							# No name match, send an email with the suggestion to be safe
							$Suggestion = "$($Related_ATUser.firstName) $($Related_ATUser.lastName) ($($Related_ATUser.id))"
							$QPFixes.Add([PSCustomObject]@{
								Company = $Organization.QP.name
								id = $QPUser.qpId
								Name = $QPUser.displayName
								FixType = "Manual AT Match Required, name doesn't match for inactive AT contact: $($Suggestion) [consider enabling & matching]"
							})
							$CreateNew = $false
						} else {
							# Only 1 match found and it has a name match, looks safe, re-activated the AT contact and auto update match!

							# Reactivate
							$Updated_ATUser = $Related_ATUser.PsObject.Copy()
							$Updated_ATUser.isActive = 1
							$Updated_ATUser.PSObject.Properties.Remove('id')
							Set-AutotaskAPIResource -Resource CompanyContactsChild -ParentID $Organization.AT.id -ID $Related_ATUser.id -body $Updated_ATUser

							# Matching
							$Headers = @{
								Integration = "dattoAutoTask"
							}
							$Body = @{
								allSelected = $false
								customerId = $Organization.QP.id
								excludedIds = @()
								includedIds = @()
								integrationType = "dattoAutoTask"
								matchType = "ALL"
								matches = @{
									$QPUser.qpId = @{
										autoMatch = $false
										displayName = $QPUser.displayName
										id = "$($Related_ATUser.id)"
										name = "$($Related_ATUser.firstName) $($Related_ATUser.lastName)"
										_id = $QPUser.qpId
									}
								}
								searchText = ""
								userType = "standard"
							}
		
							$Params = @{
								Method = "Post"
								Uri = "$($QP_BaseURI)integrations/accounts/matches/createjob"
								Body = ($Body | ConvertTo-Json -Depth 10)
								ContentType = "application/json"
								Headers = $Headers
								WebSession = $QPWebSession
							}
		
							$AutoUpdated = $false
							try {
								Write-Output "Creating new AT match: $($QPUser.displayName) to $($Related_ATUser.firstName) $($Related_ATUser.lastName)"
								$Response = Invoke-RestMethod @Params
								Start-Sleep -Seconds 1
								if ($Response -and $Response.message -like "Matching Changes Submitted") {
									$Updates++
									$AutoUpdated = $true
								}
							} catch {
								$AutoUpdated = $false
								Start-Sleep -Seconds 3
							}
		
							if (!$AutoUpdated) {
								# Auto update did not work, send an email
								$Suggestion = "$($Related_ATUser.firstName) $($Related_ATUser.lastName) ($($Related_ATUser.id))"
								$QPFixes.Add([PSCustomObject]@{
									Company = $Organization.QP.name
									id = $QPUser.qpId
									Name = $QPUser.displayName
									FixType = "Manual AT Match Required, automatch failed: $($Suggestion) [inactive AT account]"
								})
							}
							$CreateNew = $false
						}
					}
				}
				
				if ($CreateNew -and (Get-Date $QPUser.createdAt) -lt (Get-Date).AddDays(-7) -and $QPUser.displayName) {
					# If the contact is older than 7 days and we haven't found any inactive matches, make a new contact
					$NewATContact = [PSCustomObject]@{
						firstName = ""
						lastName = ""
						emailAddress = ""
						mobilePhone = ""
						companyID = $Organization.AT.id
						isActive = 1
					}

					$NameParts = $QPUser.displayName -split " "
					$FirstName =  ($NameParts[0..([math]::Max($NameParts.Count - 2, 0))] -join " ")
					$LastName = if ($NameParts.Count -gt 1) { $NameParts[-1] } else { "." }
					if ($LastName -eq "." -and $FirstName.Length -gt 20) {
						$LastName = $FirstName.substring(20)
					}
					$FirstName = $FirstName.substring(0, [System.Math]::Min(20, $FirstName.Length))
					$LastName = $LastName.substring(0, [System.Math]::Min(20, $LastName.Length))

					$NewATContact.firstName = $FirstName
					$NewATContact.lastName = $LastName
					$NewATContact.emailAddress = $QPUser.email
					$NewATContact.mobilePhone = $QPUser.phoneNumber

					if ($QPUser.email -and $QPUser.email.Length -gt 50) {
						continue :userLoop
					}

					$Result = New-AutotaskAPIResource -Resource CompanyContactsChild -ParentID $Organization.AT.id -Body $NewATContact

					if ($Result -and $Result.itemId) {
						# Link new AT Contact to QP user
						$Headers = @{
							Integration = "dattoAutoTask"
						}
						$Body = @{
							allSelected = $false
							customerId = $Organization.QP.id
							excludedIds = @()
							includedIds = @()
							integrationType = "dattoAutoTask"
							matchType = "ALL"
							matches = @{
								$QPUser.qpId = @{
									autoMatch = $false
									displayName = $QPUser.displayName
									id = "$($Result.itemId)"
									name = "$($QPUser.displayName)"
									_id = $QPUser.qpId
								}
							}
							searchText = ""
							userType = "standard"
						}
	
						$Params = @{
							Method = "Post"
							Uri = "$($QP_BaseURI)integrations/accounts/matches/createjob"
							Body = ($Body | ConvertTo-Json -Depth 10)
							ContentType = "application/json"
							Headers = $Headers
							WebSession = $QPWebSession
						}
	
						$AutoUpdated = $false
						try {
							$Response = Invoke-RestMethod @Params
							Start-Sleep -Seconds 1
							if ($Response -and $Response.message -like "Matching Changes Submitted") {
								Write-Output "Created new AT contact and matched: $($QPUser.displayName)"
								$Updates++
								$AutoUpdated = $true
							}
						} catch {
							$AutoUpdated = $false
							Start-Sleep -Seconds 3
						}
	
						if (!$AutoUpdated) {
							# Auto update did not work, send an email
							$Suggestion = "$($FirstName) $($LastName) ($($Result.itemId))"
							$QPFixes.Add([PSCustomObject]@{
								Company = $Organization.QP.name
								id = $QPUser.qpId
								Name = $QPUser.displayName
								FixType = "Created new AT Contact but QP matching failed: $($Suggestion)"
							})
						}
					}
				}
			}

			if (($Related_ATUsers | Measure-Object).Count -gt 0) {
				$QPPasswordMatch_Cache.matchAttempted.($Organization.QP.id) | Add-Member -MemberType NoteProperty -Name $QPUser.qpId -Value @() -ErrorAction Ignore
				$QPPasswordMatch_Cache.matchAttempted.($Organization.QP.id).($QPUser.qpId) += "AT"
			}
		}
	}

	# If we updated ITG/AT matching, re-query the QP users list
	if ($Updates -gt 0) {
		$Response = Invoke-WebRequest "$($QP_BaseURI)customer/$($Organization.QP.id)/users?status=allActive&accountStatus=all&userAppType=all&page=1&type=standard&rowsPerPage=50&qpStatus=all&searchText=" -WebSession $QPWebSession
		$QP_EndUsers = $Response.Content | ConvertFrom-Json

		if ($QP_EndUsers -and $QP_EndUsers.maxCount -gt $QP_EndUsers.users.Count) {
			$i = 1
			while ($QP_EndUsers -and $QP_EndUsers.maxCount -gt $QP_EndUsers.users.count -and $i -le $RateLimit) {
				$i++
				$Response = Invoke-WebRequest "$($QP_BaseURI)customer/$($Organization.QP.id)/users?status=allActive&accountStatus=all&userAppType=all&page=$i&type=standard&rowsPerPage=50&qpStatus=all&searchText=" -WebSession $QPWebSession
				$QP_EndUsers_ToAdd = $Response.Content | ConvertFrom-Json
				if ($QP_EndUsers_ToAdd -and $QP_EndUsers_ToAdd.users) {
					$QP_EndUsers.users += $QP_EndUsers_ToAdd.users
				}
			}
		}
	}

	# Filter event logs for this customer
	$OrgITGEvents = $false
	$OrgATEvents = $false
	if ($QP_ITG_EventLog -and $QP_ITG_EventLog.events) {
		$OrgITGEvents = $QP_ITG_EventLog.events | Where-Object { $_.customerName -eq $Organization.QP.name }
	}
	if ($QP_AT_EventLog -and $QP_AT_EventLog.events) {
		$OrgATEvents = $QP_AT_EventLog.events | Where-Object { $_.customerName -eq $Organization.QP.name }
	}

	# Match QP users to ITG passwords and update cached list
	$PasswordMatches = @{}
	foreach ($QPUser in $QP_EndUsers.users) {
		$ITGPasswordMatch = $false

		# Check if this password was recently updated by checking the event log, if so, clear the match cache for this password
		if ($OrgITGEvents) {
			$RelatedEvents = $OrgITGEvents | Where-Object { $_.details -like "*Password Entry Name: $($QPUser.displayName)*" }

			if (!$RelatedEvents) {
				$Email = $QPUser.email.Replace("*", "").Replace("?", "").Replace("[", "").Replace("]", "")
				$RelatedEvents = $OrgITGEvents | Where-Object { $_.source -like $QPUser.samAccountName -or $_.source -like $QPUser.userPrincipalName -or $_.source -like $Email } 	
			}

			if ($RelatedEvents -and ($RelatedEvents | Measure-Object).Count -gt 0 -and $QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId)) {
				# Clear ITG match from cache
				$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId).PSObject.Properties.Remove('ITG')
			}
		}

		if ($QPUser.integrations.itglue.active -ne $false) {
			# Check the cache to find an ITG match
			if ($QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId) -and $QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId).ITG) {
				$ITGPasswordMatch = $QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId).ITG
			}
		} else {
			# There is no ITG match, remove match from cache
			if ($QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId)) {
				$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId).PSObject.Properties.Remove('ITG')
			}
		}

		# No cached match found (or it needs updating), get the match from QP
		if (!$ITGPasswordMatch -and $QPUser.integrations.itglue.active -ne $false) {
			$Headers = @{
				Integration = "itglue"
			}
			try {
				$Response = Invoke-WebRequest "$($QP_BaseURI)integrations/accounts/$($QPUser.qpId)?customer_id=$($Organization.QP.id)" -WebSession $QPWebSession -Headers $Headers
				$ITGPasswordMatch = $Response.Content | ConvertFrom-Json
				$ITGPasswordMatch = $ITGPasswordMatch.id
			} catch {
				if ($_.Exception.Response.StatusCode.Value__ -eq 404) {
					$QPUser.integrations.itglue.active = $false
				}
			}
			Start-Sleep -Milliseconds 500

			if ($ITGPasswordMatch) {
				# Update cache
				$QPPasswordMatch_Cache.customers.($Organization.QP.id) | Add-Member -MemberType NoteProperty -Name $QPUser.qpId -Value $null -ErrorAction Ignore
				if (!$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId)) {
					$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId) = [PSCustomObject]@{}
				}
				$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId) | Add-Member -MemberType NoteProperty -Name ITG -Value $null -ErrorAction Ignore
				$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId) | Add-Member -MemberType NoteProperty -Name AT -Value $null -ErrorAction Ignore
				$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId).ITG = $ITGPasswordMatch
			}
		}

		if ($ITGPasswordMatch) {
			if (!$PasswordMatches.($QPUser.qpId)) {
				$PasswordMatches.($QPUser.qpId) = @{}
			}
			$PasswordMatches.($QPUser.qpId).ITG = $ITGPasswordMatch
		}
	}

	# Match QP users to Autotask passwords and update cached list
	foreach ($QPUser in $QP_EndUsers.users) {
		$ATPasswordMatch = $false

		# Check if this password was recently updated by checking the event log, if so, clear the match cache for this password
		if ($OrgATEvents) {
			$RelatedEvents = $OrgATEvents | Where-Object { $_.details -like "*Quickpass Account: $($QPUser.displayName)*" -or $_.details -like "*Datto Autotask Contact: $($QPUser.displayName)*" }

			if (!$RelatedEvents) {
				$Email = $QPUser.email.Replace("*", "").Replace("?", "").Replace("[", "").Replace("]", "")
				$RelatedEvents = $OrgATEvents | Where-Object { $_.source -like $QPUser.samAccountName -or $_.source -like $QPUser.userPrincipalName -or $_.source -like $Email }
			}

			if ($RelatedEvents -and ($RelatedEvents | Measure-Object).Count -gt 0 -and $QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId)) {
				# Clear Autotask match from cache
				$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId).PSObject.Properties.Remove('AT')
			}
		}

		if ($QPUser.integrations.dattoAutoTask.active -ne $false) {
			# Check the cache to find an Autotask match
			if ($QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId) -and $QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId).AT) {
				$ATPasswordMatch = $QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId).AT
			}
		} else {
			# There is no Autotask match, remove match from cache
			if ($QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId)) {
				$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId).PSObject.Properties.Remove('AT')
			}
		}

		# No cached match found (or it needs updating), get the match from QP
		if (!$ATPasswordMatch -and $QPUser.integrations.dattoAutoTask.active -ne $false) {
			$Headers = @{
				Integration = "dattoAutoTask"
			}
			try {
				$Response = Invoke-WebRequest "$($QP_BaseURI)integrations/accounts/$($QPUser.qpId)?customer_id=$($Organization.QP.id)" -WebSession $QPWebSession -Headers $Headers
				$ATPasswordMatch = $Response.Content | ConvertFrom-Json
				$ATPasswordMatch = $ATPasswordMatch.id
			} catch {
				if ($_.Exception.Response.StatusCode.Value__ -eq 404) {
					$QPUser.integrations.dattoAutoTask.active = $false
				}
			}
			Start-Sleep -Milliseconds 500

			if ($ATPasswordMatch) {
				# Update cache
				$QPPasswordMatch_Cache.customers.($Organization.QP.id) | Add-Member -MemberType NoteProperty -Name $QPUser.qpId -Value $null -ErrorAction Ignore
				if (!$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId)) {
					$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId) = [PSCustomObject]@{}
				}
				$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId) | Add-Member -MemberType NoteProperty -Name ITG -Value $null -ErrorAction Ignore
				$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId) | Add-Member -MemberType NoteProperty -Name AT -Value $null -ErrorAction Ignore
				$QPPasswordMatch_Cache.customers.($Organization.QP.id).($QPUser.qpId).AT = $ATPasswordMatch
			}
		}

		if ($ATPasswordMatch) {
			if (!$PasswordMatches.($QPUser.qpId)) {
				$PasswordMatches.($QPUser.qpId) = @{}
			}
			$PasswordMatches.($QPUser.qpId).AT = $ATPasswordMatch
		}
	}

	# Update/fix phone numbers and emails
	foreach ($QPUser in $QP_EndUsers.users) {
		$ATContact = $false
		$PasswordMatch = $PasswordMatches.($QPUser.qpId)
		if ($PasswordMatch -and $PasswordMatch.AT) {
			$ATContact = $Autotask_Contacts | Where-Object { $_.id -eq $PasswordMatch.AT }
		}

		if ($QPUser.qpId -notin $QPPasswordMatch_Cache.firstTimeUpdates.($Organization.QP.id)) {
			$FirstTimeUpdateComplete = $true
			if (!$ATContact) {
				$FirstTimeUpdateComplete = $false
			}
			# Update phone number from Autotask
			# Even if QP already has a number, overwrite it, unless the user answered an onboarding email
			# Autotask generally has better data than AD which is where QP gets the phone number from by default
			if ($ATContact -and $ATContact.mobilePhone -and (!$QPUser.onboardingCompleted -or !$QPUser.phoneNumber)) {
				$NewPhone = $false
				$CurPhone = $ATContact.mobilePhone -replace '\D', ""
				if ($CurPhone.StartsWith("011")) {
					$NewPhone = $CurPhone.Substring(3)
				} elseif ($CurPhone.length -eq 10) {
					$NewPhone = "1$($CurPhone)"
				} elseif ($CurPhone.length -gt 10 -and !$QPUser.phoneNumber) {
					$NewPhone = $CurPhone
				}
				if ($NewPhone -and (!$QPUser.phoneNumber -or $NewPhone -ne $QPUser.phoneNumber)) {
					$NewPhone = "+$($NewPhone)"

					$Body = @{
						id = $QPUser.qpId
						phoneNumber = $NewPhone
					}
					try {
						$Response = Invoke-WebRequest "$($QP_BaseURI)user/changePhoneNumber" -WebSession $QPWebSession -Body $Body -Method 'POST'
						$Success = $Response.Content | ConvertFrom-Json
						Write-Output "Updated phone number: $($QPUser.displayName) to $($NewPhone)"
						if (!$Success -or !$Success.id) {
							$FirstTimeUpdateComplete = $false
						}
					} catch {
						$FirstTimeUpdateComplete = $false
					}
				}

			} elseif ($QPUser.phoneNumber -and $QPUser.phoneNumber -like "+*") {
				# No need to update the phone from Autotask, but lets ensure it has the correct country code
				if ($QPUser.phoneNumber.Length -ne 12 -or $QPUser.phoneNumber -notlike "+1*") {
					$NewPhone = $false
					if ($QPUser.phoneNumber.Length -eq 11) {
						$NewPhone = "+1$($QPUser.phoneNumber.Substring(1))"
					} elseif ($QPUser.phoneNumber.Length -lt 11 -and $QPUser.phoneNumber.Length -gt 0) {
						$NewPhone = "REMOVE"
					}

					if ($NewPhone) {
						if ($NewPhone -eq "REMOVE") {
							$NewPhone = ""
						}
						$Body = @{
							id = $QPUser.qpId
							phoneNumber = $NewPhone
						}
						try {
							$Response = Invoke-WebRequest "$($QP_BaseURI)user/changePhoneNumber" -WebSession $QPWebSession -Body $Body -Method 'POST'
							$Success = $Response.Content | ConvertFrom-Json
							if (!$NewPhone) {
								Write-Output "Removed bad phone number: $($QPUser.displayName)"
							} else {
								Write-Output "Updated broken phone number: $($QPUser.displayName) to $($NewPhone)"
							}
							if (!$Success -or !$Success.id) {
								$FirstTimeUpdateComplete = $false
							}
						} catch {
							$FirstTimeUpdateComplete = $false
						}
					}
				}
			}

			
			if ($FirstTimeUpdateComplete) {
				$QPPasswordMatch_Cache.firstTimeUpdates.($Organization.QP.id) += $QPUser.qpId
			}
		}

		if (!$QPUser.email -and $ATContact -and $ATContact.emailAddress) {
			# Update email address if missing in QP
			$NewEmail = $false
			if (IsValidEmail $ATContact.emailAddress) {
				$NewEmail = $ATContact.emailAddress.Trim()
			} elseif ($ATContact.emailAddress2 -and (IsValidEmail $ATContact.emailAddress2)) {
				$NewEmail = $ATContact.emailAddress2.Trim()
			} elseif ($ATContact.emailAddress3 -and (IsValidEmail $ATContact.emailAddress3)) {
				$NewEmail = $ATContact.emailAddress3.Trim()
			}

			$Body = @{
				id = $QPUser.qpId
				email = $NewEmail
			}
			try {
				$Response = Invoke-WebRequest "$($QP_BaseURI)user/changeEmail" -WebSession $QPWebSession -Body $Body -Method 'POST'
				Write-Output "Updated email: $($QPUser.displayName) to $($NewEmail)"
			} catch {
				if ($_.ErrorDetails.Message -like "Another user with the same email already exists") {
					$QPFixes.Add([PSCustomObject]@{
						Company = $Organization.QP.name
						id = $QPUser.qpId
						Name = $QPUser.displayName
						FixType = "Could not update email to '$($NewEmail)' because another user exists with that email."
					})
				}
			}
		}
	}
}

# Update QP to ITG matching cache
if ($QPPasswordMatch_Cache) {
	$QPPasswordMatch_Cache.lastUpdated = (Get-Date).ToString()
	$QPPasswordMatch_Cache | ConvertTo-Json -Depth 10 | Out-File -PSPath "./QPPasswordMatchingCache.json" -Force
}

# Send an email if any issues were found
if ($QPFixes -and $Email_APIKey -and $Email_APIKey.Key) {
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
							<th>Company/Contact ID</th>
							<th>Name</th>
							<th>Issue</th>
						</tr>'
	$QPFixes | ForEach-Object {
		$HTMLBody += '
						<tr>
							<td style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 5px; border: 1px solid #000000;">{0}</td>
							<td style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 5px; border: 1px solid #000000;">{1}</td>
							<td style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 5px; border: 1px solid #000000;">{2}</td>
							<td style="font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 5px; border: 1px solid #000000;">{3}</td>
						</tr>' -f $_.Company, $_.id, $_.Name, $_.FixType
	}
	$HTMLBody += '
						</tbody>
					</table>
					<div class="mobile_table_fallback" style="display: none;">
						Table version hidden. You can view a tabular version of the above data on a desktop.
					</div><br />'

	$HTMLEmail = $EmailTemplate -f `
					"Issues were found during the Quickpass cleanup that need to be manually fixed.", 
					"Quickpass Issues Found", 
					$HTMLBody, 
					"Please resolve these issues manually."

	$Path = "C:\Temp\QPCleanupFixes.csv"
	if (Test-Path $Path) {
        Remove-Item $Path
    }
	$QPFixes | Export-CSV -NoTypeInformation -Path $Path
	$QPFixes_Encoded = [System.Convert]::ToBase64String([IO.File]::ReadAllBytes($Path))

	$mailbody = @{
		"From" = $EmailFrom
		"To" = $EmailTo
		"Subject" = "Quickpass Cleanup - Manual Fixes Required"
		"HTMLContent" = $HTMLEmail
		"Attachments" = @(
			@{
				Base64Content = $QPFixes_Encoded
				Filename = "QPCleanupFixes.csv"
				ContentType = "text/csv"
			}
		)
	} | ConvertTo-Json -Depth 6

	$headers = @{
		'x-api-key' = $Email_APIKey.Key
	}
	
	Invoke-RestMethod -Method Post -Uri $Email_APIKey.Url -Body $mailbody -Headers $headers -ContentType application/json
}
