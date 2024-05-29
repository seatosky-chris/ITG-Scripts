####################################################################
$ITGAPIKey = @{
	Url = "https://api.itglue.com"
	Key = ""
}
$AutotaskAPIKey = @{
	Url = "https://webservices1.autotask.net/atservicesrest"
	Username = ""
	Key = ""
	IntegrationCode = ""
}
$LastUpdatedUpdater_APIURL = ""

$CyberQP_APIVendorID = $false # This is the API vendor ID for CyberQP/Quickpass. We use this to check if a contact was made by Quickpass. Set to $false to disable.

# See README for full instructions on setup.
# Find the nuget package location with "dotnet nuget locals global-packages -l" in a terminal
# Then navigate to the "libphonenumber-csharp" folder and find the latest PhoneNumbers.dll for a version of .Net that will work on this system
# Use the full path for this constant.
$phoneNumbersDLLPath = "C:\Users\Administrator\.nuget\packages\libphonenumber-csharp\8.12.34\lib\net46\PhoneNumbers.dll"
####################################################################

### This code is common for every company and can be ran before looping through multiple companies
$CurrentTLS = [System.Net.ServicePointManager]::SecurityProtocol
if ($CurrentTLS -notlike "*Tls12" -and $CurrentTLS -notlike "*Tls13") {
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	Write-Output "This device is using an old version of TLS. Temporarily changed to use TLS v1.2."
	Write-PSFMessage -Level Warning -Message "Temporarily changed TLS to TLS v1.2."
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
	$AutotaskConnected = $true
}

# Prepare the phone number formatter. This uses C# and Google's libphonenumber library.
Add-Type -Path $phoneNumbersDLLPath

Add-Type -TypeDefinition @"
using System;
using PhoneNumbers;
using System.Text.RegularExpressions;

public class PhoneNumberLookup 
{
    public static string FormatPhoneNumber(string phoneString)
    {
        PhoneNumberUtil phoneNumberUtil = PhoneNumbers.PhoneNumberUtil.GetInstance();

        string parsedPhoneString;
        if (phoneString.Trim().StartsWith("+") || phoneString.Trim().StartsWith("011")) 
        {
            parsedPhoneString = phoneString.Trim();
        } 
        else 
        {
            string extension = "";
        
            if (Regex.IsMatch(phoneString, "[a-zA-Z]")) 
            {
                Match letterMatches = Regex.Match(phoneString, "[a-zA-Z]");
                extension = " " + phoneString.Substring(letterMatches.Index);
                parsedPhoneString = phoneString.Substring(0, letterMatches.Index);
                parsedPhoneString = Regex.Replace(parsedPhoneString, "\\D", "");
            } 
            else 
            {
                parsedPhoneString = Regex.Replace(phoneString, "\\D", "");
            }
            
            if (parsedPhoneString.Length == 10) 
            {
                parsedPhoneString = "+1 " + parsedPhoneString;
            } 
            else if (parsedPhoneString.Length > 10) 
            {
                parsedPhoneString = "+" + parsedPhoneString;
            }

            parsedPhoneString += extension;
        }

        PhoneNumber phoneNumber = phoneNumberUtil.Parse(parsedPhoneString, "CA");

        if (phoneNumber.CountryCode == 1) 
        {
            return phoneNumberUtil.Format(phoneNumber, PhoneNumbers.PhoneNumberFormat.NATIONAL);
        } 
        else 
        {
            //return phoneNumberUtil.Format(phoneNumber, PhoneNumbers.PhoneNumberFormat.INTERNATIONAL);
            return phoneNumberUtil.FormatOutOfCountryCallingNumber(phoneNumber, "CA");
        }
    }
}
"@ -ReferencedAssemblies @($phoneNumbersDLLPath, "System.Text.RegularExpressions")

function FormatPhoneNumber($phoneString, $APIVendorID = $false) {
	# Quickpass phone number fix
	if ($CyberQP_APIVendorID -and $phoneString -like "+*") {
		if ($APIVendorID -and $APIVendorID -eq $CyberQP_APIVendorID -and $phoneString -notlike "+1*" -and $phoneString.length -lt 12) {
			$phoneString = "+1$($phoneString.Substring(1))"
		} elseif ($phoneString.length -eq 11) {
			$phoneString = "+1$($phoneString.Substring(1))"
		} elseif ($phoneString.length -lt 7) {
			$phoneString = $phoneString.Substring(1)
		}
	}

	if ($phoneString.StartsWith("011 ") -or $phoneString -match "^\(\d\d\d\) " -or $phoneString.Trim() -eq "email") {
        $formattedPhoneNumber = $phoneString.Trim();
    } else {
        try {
            $formattedPhoneNumber = [PhoneNumberLookup]::FormatPhoneNumber($phoneString);
        } catch {
            $formattedPhoneNumber = $phoneString;
        }
    }

	return $formattedPhoneNumber;
}

# Get all the possible companies and match them together
$AutotaskCompanies = Get-AutotaskAPIResource -resource Companies -SimpleSearch "isactive eq $true"
$ITGlueCompanies = Get-ITGlueOrganizations -page_size 1000

if ($ITGlueCompanies -and $ITGlueCompanies.data) {
	$ITGlueCompanies = $ITGlueCompanies.data | Where-Object { $_.attributes.'organization-status-name' -eq "Active" }
}

if (!$AutotaskCompanies -or !$ITGlueCompanies) {
	exit
}

$Companies = @()
foreach ($Company in $AutotaskCompanies) {
	$ITGCompany = $ITGlueCompanies | Where-Object { $_.attributes.name.Trim() -like $Company.companyName.Trim() }

	if ($ITGCompany) {
		$Companies += [PSCustomObject]@{
			AutotaskID = $Company.id
			AutotaskName = $Company.companyName
			ITGID = $ITGCompany.id
			ITGName = $ITGCompany.attributes.name
		}
	}
}

if (!$Companies) {
	exit
}

# Run through each company and get/compare contacts for cleanup
foreach ($Company in $Companies) {
	Write-Host "Auditing: $($Company.AutotaskName)" -ForegroundColor Green
	$AutotaskContacts = Get-AutotaskAPIResource -Resource Contacts -SimpleSearch "companyID eq $($Company.AutotaskID)"
	$ITGContacts = Get-ITGlueContacts -organization_id $Company.ITGID -page_size 1000
	$ITGLocations = Get-ITGlueLocations -org_id $Company.ITGID

	if ($ITGContacts.Error -or $ITGLocations.Error) {
		Write-Error "An error occurred trying to get the existing contact or locations from ITG. Exiting..."
		if ($ITGContacts.Error) {
			Write-Error $ITGContacts.Error
		} else {
			Write-Error $ITGLocations.Error
		}
		exit 1
	}

	if ($ITGContacts -and $ITGContacts.data) {
		$ITGContacts = $ITGContacts.data | Where-Object { $_.attributes.'psa-integration' -ne 'disabled' }
	}
	if ($ITGLocations -and $ITGLocations.data) {
		$ITGLocations = $ITGLocations.data
	}

	if (!$AutotaskContacts -or !$ITGContacts) {
		continue
	}

	# Match contacts from ITG to Autotask
	$Contacts = @()
	foreach ($Contact in $ITGContacts) {
		$PrimaryEmail = $Contact.attributes.'contact-emails' | Where-Object { $_.primary -eq "True" }
		if ($PrimaryEmail) {
			$PrimaryEmail = $PrimaryEmail.value.Trim()
		} else {
			$PrimaryEmail = ""
		}
		$AutotaskContact = $AutotaskContacts | Where-Object { 
			$_.firstName.Trim() -like $Contact.attributes.'first-name'.Trim() -and 
			$_.lastName.Trim() -like $Contact.attributes.'last-name'.Trim() -and
			$_.emailAddress.Trim() -like $PrimaryEmail
		}

		if (($AutotaskContact | Measure-Object).Count -eq 0) {
			$AutotaskContact = $AutotaskContacts | Where-Object { 
				$_.firstName.Trim() -like $Contact.attributes.'first-name'.Trim() -and 
				$_.lastName.Trim() -like $Contact.attributes.'last-name'.Trim()
			}
		}

		if (($AutotaskContact | Measure-Object).Count -gt 1) {
			$PhoneMatches = @($Contact.attributes.'contact-phones'.value)
			$PhoneMatches += ""
			$AutotaskContact = $AutotaskContact | Where-Object { $_.title.Trim() -eq ($Contact.attributes.title | Out-String).Trim() -and ((($_.mobilePhone -replace '\D', '') -in $PhoneMatches -and ($_.phone -replace '\D', '') -in $PhoneMatches) -or !$Contact.attributes.'contact-phones'.value) }
		}

		if (($AutotaskContact | Measure-Object).Count -gt 1 -and $ITGLocations -and $Contact.attributes.'location-id') {
			$Location = $ITGLocations | Where-Object { $_.id -eq $Contact.attributes.'location-id' }
			if ($Location.attributes.'psa-integration' -eq 'enabled') {
				$AutotaskContact = $AutotaskContact | Where-Object { $_.city -like $Location.attributes.city -and ($_.zipCode -replace '\W', '') -like ($Location.attributes.'postal-code' -replace '\W', '') }
				if (($AutotaskContact | Measure-Object).Count -gt 1) {
					$AutotaskContact = $AutotaskContact | Where-Object { $_.addressLine -like $Location.attributes.'address-1' }
				}
			}
		}

		if ($AutotaskContact) {
			$Contacts += [PSCustomObject]@{
				AutotaskIDs = @($AutotaskContact.id)
				AutotaskName = (($AutotaskContact | Select-Object -First 1).firstName + " " + ($AutotaskContact | Select-Object -First 1).lastName)
				ITGID = $Contact.id
				ITGName = $Contact.attributes.name
			}
		}
	}

	if (!$Contacts) {
		continue
	}

	# Loop through each matched contact and update data where applicable
	# This is for things where a match with IT Glue DOES matter
	# Targets: Deactivating autotask contacts that are terminated in ITG
	foreach ($ContactMatch in $Contacts) {
		$ITGContact = $ITGContacts | Where-Object { $_.id -eq $ContactMatch.ITGID }
		$AutotaskContactMatches = @($AutotaskContacts | Where-Object { $_.id -in $ContactMatch.AutotaskIDs })

		# Skip if there is nothing to change
		if ($ITGContact.attributes.'contact-type-name' -ne "Terminated") {
			continue
		}

		foreach ($AutotaskContact in $AutotaskContactMatches) {
			$ContactUpdate = 
			[PSCustomObject]@{
				id = $AutotaskContact.id
			}

			$ChangesMade = $false

			if ($AutotaskContact.isActive -gt 0) {
				$ChangesMade = $true
				$ContactUpdate | Add-Member -NotePropertyName isActive -NotePropertyValue 0
				$ContactUpdate.isActive = 0
			}

			if ($ChangesMade) {
				$Changes += $ContactUpdate
				Set-AutotaskAPIResource -Resource CompanyContactsChild -ID $Contact.id -body $ContactUpdate
			}
		}
	}

	# Loop through every Autotask contact and update where applicable
	# This is for things where a match with IT Glue DON'T matter
	# Targets: Email2At in title and Phone Number formatting
	foreach ($Contact in $AutotaskContacts) {
		if (!$Contact.phone -and !$Contact.mobilePhone -and !$Contact.alternatePhone -and $Contact.title -ne "Email2AT Contact") {
			continue
		}

		$ContactUpdate = 
		[PSCustomObject]@{
			id = $Contact.id
		}

		$ChangesMade = $false

		# Update phone number formatting and titles in Autotask
		if ($Contact.phone) {
			$FormattedPhone = FormatPhoneNumber -phoneString $Contact.phone -APIVendorID $Contact.apiVendorID

			if ($FormattedPhone -like "*ext.*") {
				$Extension = $FormattedPhone.Substring($FormattedPhone.IndexOf("ext.") + 5).Trim() 
				if (!$Contact.extension) {
					$ChangesMade = $true
					$ContactUpdate | Add-Member -NotePropertyName extension -NotePropertyValue ""
					$ContactUpdate.extension = $Extension
					$FormattedPhone = $FormattedPhone.Substring(0, $FormattedPhone.IndexOf("ext.")).Trim()
				} elseif ($Contact.extension.Trim() -like $Extension) {
					$FormattedPhone = $FormattedPhone.Substring(0, $FormattedPhone.IndexOf("ext.")).Trim()
				}
			}

			if ($Contact.phone -ne $FormattedPhone) {
				$ChangesMade = $true
				$ContactUpdate | Add-Member -NotePropertyName phone -NotePropertyValue ""
				$ContactUpdate.phone = $FormattedPhone
			}
		}

		if ($Contact.mobilePhone) {
			$FormattedPhone = FormatPhoneNumber -phoneString $Contact.mobilePhone -APIVendorID $Contact.apiVendorID
			if ($Contact.mobilePhone -ne $FormattedPhone) {
				$ChangesMade = $true
				$ContactUpdate | Add-Member -NotePropertyName mobilePhone -NotePropertyValue ""
				$ContactUpdate.mobilePhone = $FormattedPhone
			}
		}

		if ($Contact.alternatePhone) {
			$FormattedPhone = FormatPhoneNumber -phoneString $Contact.alternatePhone -APIVendorID $Contact.apiVendorID
			if ($Contact.alternatePhone -ne $FormattedPhone) {
				$ChangesMade = $true
				$ContactUpdate | Add-Member -NotePropertyName alternatePhone -NotePropertyValue ""
				$ContactUpdate.alternatePhone = $FormattedPhone
			}
		}

		if ($Contact.title -eq "Email2AT Contact") {
			$ChangesMade = $true
			$ContactUpdate | Add-Member -NotePropertyName title -NotePropertyValue ""
			$ContactUpdate.title = ""
		}

		if ($ChangesMade) {
			Set-AutotaskAPIResource -Resource CompanyContactsChild -ID $Contact.id -body $ContactUpdate
		}
	}

	# Update / Create the "Scripts - Last Run" ITG page which shows when the contact cleanup (and other scripts) last ran
	if ($LastUpdatedUpdater_APIURL -and $Company.ITGID) {
		$Headers = @{
			"x-api-key" = $ITGAPIKey.Key
		}
		$Body = @{
			"apiurl" = $ITGAPIKey.Url
			"itgOrgID" = $Company.ITGID
			"HostDevice" = $env:computername
			"contact-cleanup" = (Get-Date).ToString("yyyy-MM-dd")
		}
	
		$Params = @{
			Method = "Post"
			Uri = $LastUpdatedUpdater_APIURL
			Headers = $Headers
			Body = ($Body | ConvertTo-Json)
			ContentType = "application/json"
		}			
		Invoke-RestMethod @Params 
	}
}
