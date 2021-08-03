####################################################################
$APIKEy =  "<ITG API KEY>"
$APIEndpoint = "https://api.itglue.com"
$orgID = "<ITG ORG ID>"
$DevicePrefix =  "<DEVICE PREFIX>" # Generally the companies acronym
$FolderID = $false # the folder id to put the passwords into, $false if root
$OverwritePasswords = $false # If a password already exists in ITG, update it and overwrite the ITG password with the imported one (when $true)
$SelectFormPath = ".\Forms\SelectForm\SelectForm\MainWindow.xaml"

<# .SYNOPSIS
	These are the password types.
.PARAMETER Name
	Name is the label used for naming the password in ITG.
.PARAMETER SelectLabel
	SelectLabel is the label used for the select box here in this script.
.PARAMETER Category
	Category is the Password category associated with this type in ITG.
.PARAMETER Embedded
	If this password should be embedded to other assets, include this here. Otherwise false. Options: "Configuration", "Contact", $false
.PARAMETER Linked
	An optional parameter, can be used instead of the Embedded param. If used it will link to an asset (just like with embedded) but will create a general password and make the link as a related item. Options: "Configuration", "Contact", $false
.PARAMETER Notes
	An optional parameter, can be used to set the notes of a password to predefined text.
.PARAMETER Defaults
	An optional parameter that is a hashtable of default column selections. You can map the following keys to specific column labels: Name, Username, Password
#>
$ImportTypes = @(
	@{ "Name" = "Local Admin"; "SelectLabel" = "Local Admin"; "Category" = "Configurations - Local Admin (Workstation / Server)"; "Embedded" = "Configuration"; "Defaults" = @{ "Password" = "Local Admin"; "Name" = "Hostname" } }
	@{ "Name" = "BIOS"; "SelectLabel" = "BIOS"; "Category" = "Configurations - BIOS"; "Embedded" = "Configuration"; "Defaults" = @{ "Password" = "BIOS"; "Name" = "Hostname" } }
	@{ "Name" = "Local User"; "SelectLabel" = "Local User Account (non-admin)"; "Category" = "Configurations - Local User Account (Workstation)"; "Embedded" = "Configuration" }
	@{ "Name" = "Office Key"; "SelectLabel" = "Office Key"; "Category" = "Application / Software - Office Key"; "Embedded" = "Configuration"; "Notes" = "This is an old Office key and can be deleted if this computer now uses O365."; "Defaults" = @{ "Password" = "Office (key)"; "Name" = "Hostname" } }

	@{ "Name" = "AD"; "SelectLabel" = "AD"; "Category" = "Active Directory"; "Embedded" = $false; "Linked" = "Contact"; "Defaults" = @{ "Username" = "Username"; "Password" = "Password"; "Name" = "User" } }
	@{ "Name" = "O365"; "SelectLabel" = "O365"; "Category" = "Email Account / O365 User"; "Embedded" = $false; "Linked" = "Contact" }
	@{ "Name" = "AD & O365"; "SelectLabel" = "AD & O365"; "Category" = "Active Directory"; "Embedded" = $false; "Linked" = "Contact"; "Defaults" = @{ "Username" = "Username"; "Password" = "Password"; "Name" = "User" } }
	@{ "Name" = "Email"; "SelectLabel" = "Email (Generic, not O365)"; "Category" = "Email Account / O365 User"; "Embedded" = $false; "Linked" = "Contact" }
)

# Key comes from $ImportTypes.Name
$PresetUsernames = @{
	"Local Admin" = ".\Administrator"
	"BIOS" = ""
	"Office Key" = ""
}
####################################################################


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

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

# Gets a list of all password categories
$ITGPasswordCategories = Get-ITGluePasswordCategories
$PasswordCategories = @{}
$ITGPasswordCategories.data | ForEach-Object {
	$PasswordCategories[$_.id] = $_.attributes.name
}

function getPasswordCategoryID($Category) {
	($PasswordCategories.GetEnumerator() | Where-Object { $_.Value -like $Category }).Name
}

$PresetUsernameTypes = $PresetUsernames.Keys


# Loads a WPF form and returns the loaded form
function loadForm($Path) {
	$inputXML = Get-Content $Path -Raw
	$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
	[xml]$XAML = $inputXML
	$reader = (New-Object System.Xml.XmlNodeReader $XAML) 
	try {
		$Form = [Windows.Markup.XamlReader]::Load( $reader )
	} catch {
		Write-Warning $_.Exception
		throw
	}

	# this finds all of the possible variables in the form (btn, listbox, textbox) and maps them to powershell variables with "var_" appended to the objects name. e.g. var_btnSave
	$XAML.SelectNodes("//*[@Name]") | ForEach-Object {
		#"trying item $($_.Name)"
		try {
			Set-Variable -Name "var_$($_.Name)" -Value $Form.FindName($_.Name) -Scope 1 -ErrorAction Stop
		} catch {
			throw
		}
	}

	return $Form
}

foreach ($ImportType in $ImportTypes) {
	if (!$ImportType.ContainsKey("Linked")) {
		$ImportType.Linked = $false;
	}
	if (!$ImportType.ContainsKey("Embedded")) {
		$ImportType.Embedded = $false;
	}
}


# Select what type of import this is (stored in $ImportType)
$Form = loadForm -Path($SelectFormPath)
$Form.Title = "Select the import type"
$var_lblDescription.Content = "Select the type of import you are doing:"
$i = 0
foreach ($ImportType in $ImportTypes) { 
	$var_lstSelectOptions.Items.Insert($i, $ImportType.SelectLabel) | Out-Null
	$i++
}
$var_txtNotes.Text = "- Local Admin, Bios, Local User, and Office Key Account will link to configurations
- AD, O365, & Email will link to Contacts"
$var_btnSave.IsEnabled = $false

$ImportType = $null
$var_lstSelectOptions.Add_SelectionChanged({
	$var_btnSave.IsEnabled = $true
	$script:ImportType = $ImportTypes[$var_lstSelectOptions.SelectedIndex]
})

$var_btnSave.Add_Click({
	Write-Host "Import type '$($ImportType.Name)' selected." -ForegroundColor Yellow
	$Form.Close()
})

$Form.ShowDialog() | out-null

if (!$ImportType) {
	Write-Host "You must select an import type! Exiting." -ForegroundColor Red
	exit;
}

# File selector
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
	InitialDirectory = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
	Filter = 'HTML (*.html;*.htm)|*.html;*.htm'
}
$null = $FileBrowser.ShowDialog()
$HTML_File = $FileBrowser.FileName

# Import HTML file
$HTML = New-Object -ComObject "HTMLFile"
$HTML.IHTMLDocument2_write($(Get-Content $HTML_File -raw -Encoding UTF8))

$Content = ($HTML.all.tags("div") | Where-Object {$_.id -eq "main-content"}).innerHTML

# Get Table headers
$Matches = ($Content | Select-String -pattern '(<h[1-6].*?>(.+?)<\/h2>\s*)?<div class="?table-wrap"?>\s*<table ' -AllMatches).Matches
$i = 0
$TableNames = $Matches | ForEach-Object { if ($_.Groups[2].Value) { $_.Groups[2].Value } else { $i++; "No Name #$i" } }

# Get Tables
$Tables = $HTML.all.tags("table")

if (($Tables | Measure-Object).Count -gt 1) {
	# Present the user with the list of tables and let them choose the correct one for this import
	$Form = loadForm -Path(".\Forms\Password-Import_Table-Selector-GUI\Password-Import_Table-Selector-GUI\MainWindow.xaml")

	# update the listbox with the table names
	foreach ($TableName in $TableNames) {
		$var_lstTableSelection.Items.Add($TableName) | Out-Null
	}

	# on listbox select, show example
	$var_lstTableSelection.Add_SelectionChanged({
		$SelectedTable = $var_lstTableSelection.SelectedItem
		$index = $TableNames.IndexOf($SelectedTable)
		$script:Table = $Tables[$index]
		$var_webExample.NavigateToString("<table>" + $Table.innerHTML + "</table>") 
	})

	$var_btnSave.Add_Click({
		$Form.Close()
	})

	$Form.ShowDialog() | out-null
} else {
	$Table = $Tables[0]
}

# A table was selected (or only one exists), lets parse it
if ($Table) {
	$Headers = $Table.rows[0].cells | ForEach-Object { $_.innerText }

	# Choose the passwords column (stored in $PasswordsColumn)
	if ($ImportType.Defaults -and $ImportType.Defaults.Password -and ($Headers | Where-Object { $_ -like $ImportType.Defaults.Password })) {
		$PasswordsColumn = $Headers.IndexOf($ImportType.Defaults.Password);
	} else {
		$Form = loadForm -Path($SelectFormPath)
		$Form.Title = "Select the password column"
		$var_lblDescription.Content = "Select the column with the passwords to import:"
		$i = 0
		foreach ($Col in $Headers) { 
			$var_lstSelectOptions.Items.Insert($i, $Col) | Out-Null
			$i++
		}
		$var_grpNotes.Visibility = "Hidden"
		$var_btnSave.IsEnabled = $false

		$PasswordsColumn = $null
		$var_lstSelectOptions.Add_SelectionChanged({
			$var_btnSave.IsEnabled = $true
			$script:PasswordsColumn = $var_lstSelectOptions.SelectedIndex
		})

		$var_btnSave.Add_Click({
			$Form.Close()
		})

		$Form.ShowDialog() | out-null
	}

	if ($ImportType.Name -notin $PresetUsernameTypes) {
		# Choose the username column (stored in $UsernamesColumn)
		if ($ImportType.Defaults -and $ImportType.Defaults.Username -and ($Headers | Where-Object { $_ -like $ImportType.Defaults.Username })) {
			$UsernamesColumn = $Headers.IndexOf($ImportType.Defaults.Username);
		} else {
			$Form = loadForm -Path($SelectFormPath)
			$Form.Title = "Select the username column"
			$var_lblDescription.Content = "Select the column with the usernames to import:"
			$i = 0
			foreach ($Col in $Headers) { 
				$var_lstSelectOptions.Items.Insert($i, $Col) | Out-Null
				$i++
			}
			$var_txtNotes.Text = "To save no usernames, don't select anything in the list and just click save."

			$UsernamesColumn = $null
			$var_lstSelectOptions.Add_SelectionChanged({
				$script:UsernamesColumn = $var_lstSelectOptions.SelectedIndex
			})

			$var_btnSave.Add_Click({
				$Form.Close()
			})

			$Form.ShowDialog() | out-null
		}
	}

	# Choose the naming column (stored in $NamingColumn)
	if ($ImportType.Defaults -and $ImportType.Defaults.Name -and ($Headers | Where-Object { $_ -like $ImportType.Defaults.Name })) {
		$NamingColumn = $Headers.IndexOf($ImportType.Defaults.Name);
	} else {
		$Form = loadForm -Path($SelectFormPath)
		$Form.Title = "Select the naming column"
		$var_lblDescription.Content = "Select the column to be used for naming:"
		$i = 0
		foreach ($Col in $Headers) { 
			$var_lstSelectOptions.Items.Insert($i, $Col) | Out-Null
			$i++
		}

		if ($ImportType.Category -like "Configurations*" -or $ImportType.Embedded -like "Configuration" -or $ImportType.Linked -like "Configuration") {
			$var_txtNotes.Text = "This should be the configuration's name.
			If this value appears to be an asset tag, the script will auto append the company's prefix.
			Naming Example: $($ImportType.Name) - XXX-1111"
		} elseif ($ImportType.Embedded -like "Contact" -or $ImportType.Linked -like "Contact") {
			$var_lstSelectOptions.SelectionMode = "Multiple"
			$var_txtNotes.Text = "This should be the contact's full name.
			If the name is split between First and Last name columns, you may select both.
			Naming Example: $($ImportType.Name) - John Smith"
		} else {
			$var_txtNotes.Text = "Naming Example: $($ImportType.Name) - Name from Column"
		}
		$var_btnSave.IsEnabled = $false

		$NamingColumn = $null
		$var_lstSelectOptions.Add_SelectionChanged({
			$var_btnSave.IsEnabled = $true
			if ($var_lstSelectOptions.SelectedItems.Count -gt 1) {
				$script:NamingColumn = @()
				$var_lstSelectOptions.SelectedItems | ForEach-Object {
					$script:NamingColumn += $var_lstSelectOptions.Items.IndexOf($_)
				}
			} else {
				$script:NamingColumn = $var_lstSelectOptions.SelectedIndex
			}
		})

		$var_btnSave.Add_Click({
			$Form.Close()
		})

		$Form.ShowDialog() | out-null
	}

	if ($ImportType.Embedded -or $ImportType.Linked) {
		# Choose the column for matching to the embedded/linked asset (stored in $MatchingColumns)
		$Form = loadForm -Path($SelectFormPath)
		if ($ImportType.Embedded) {
			$Form.Title = "Select the embed matching column(s)"
			$var_lblDescription.Content = "Select the column(s) for matching this to the embedded asset:"
		} else {
			$Form.Title = "Select the linked matching column(s)"
			$var_lblDescription.Content = "Select the column(s) for matching this to the linked asset:"
		}
		$var_lstSelectOptions.SelectionMode = "Multiple"

		$i = 0
		foreach ($Col in $Headers) { 
			$var_lstSelectOptions.Items.Insert($i, $Col) | Out-Null
			$i++
		}
		if ($ImportType.Embedded -like "Configuration" -or $ImportType.Linked -like "Configuration") {
			$var_txtNotes.Text = "Choose columns that contain asset tags, hostnames, and/or serial numbers."
		} elseif ($ImportType.Embedded -like "Contact" -or $ImportType.Linked -like "Contact") {
			$var_txtNotes.Text = "Choose columns that contain contact's full names, first/last names, and/or email addresses."
		}
		$var_btnSave.IsEnabled = $false

		$MatchingColumns = $null
		$var_lstSelectOptions.Add_SelectionChanged({
			$var_btnSave.IsEnabled = $true
			$script:MatchingColumns = @()
			$var_lstSelectOptions.SelectedItems | ForEach-Object {
				$script:MatchingColumns += $var_lstSelectOptions.Items.IndexOf($_)
			}
		})

		$var_btnSave.Add_Click({
			$Form.Close()
		})

		$Form.ShowDialog() | out-null
	}

	# We now know which columns to use, lets loop through the table and save each password and its associated values in an array
	$Passwords_Parsed = @()
	$Length = $Table.rows.length 
	for ($i = 1; $i -lt $Length; $i++) {
		$Row = $Table.rows[$i]

		if ($Row.cells[0].innerText -eq $Headers[0] -and $Row.cells[1].innerText -eq $Headers[1]) {
			continue; # skip header row
		}

		$Password = $Row.cells[$PasswordsColumn].innerText
		if (!$Password) {
			continue;
		}

		if ($Password -match "\s") {
			$PasswordHTML = $Row.cells[$PasswordsColumn].innerHTML
			$PasswordHTML = $PasswordHTML -replace "<s>.+?<\/s>", '' # remove strikethrough text
			$PasswordHTML = $PasswordHTML -replace '<[^>]+>', '' # remove html
			$Password = [System.Web.HttpUtility]::HtmlDecode($PasswordHTML.Trim())
		}

		$Username = ""
		if ($UsernamesColumn) {
			$Username = $Row.cells[$UsernamesColumn].innerText
			$Username = $Username -replace "\s", ''
		} elseif ($ImportType.Name -in $PresetUsernameTypes) {
			$Username = $PresetUsernames[$ImportType.Name]
		}
		$Username = [System.Web.HttpUtility]::HtmlDecode($Username)

		if ($NamingColumn.Count -gt 1) {
			$Order = @()
			$NamingColumn | ForEach-Object { if ($Headers[$_] -like "*First*") { $Order += $_ } }
			$NamingColumn | ForEach-Object { if ($Headers[$_] -notlike "*First*" -and $Headers[$_] -notlike "*Last*") { $Order += $_ } }
			$NamingColumn | ForEach-Object { if ($Headers[$_] -like "*Last*") { $Order += $_ } }
			$Name = ($Order | ForEach-Object { $Row.Cells[$_].innerText }) -join ' '
		} else {
			$Name = $Row.cells[$NamingColumn].innerText
		}
		$Name = [System.Web.HttpUtility]::HtmlDecode($Name)

		if ($ImportType.Name -eq "Office Key") {
			$Matches = ($Password | Select-String -pattern '((Office .+?)|([0-9]{4} H&B)|(H&B [0-9]{4})): (([A-Za-z0-9]{5}-){4}[A-Za-z0-9]{5})').Matches
			if ($Matches -and $Matches.Groups -and $Matches.Groups[1] -and $Matches.Groups[5]) {
				$KeyType = $Matches.Groups[1].Value.Trim()
				$Name += " ($KeyType)"
				$Password = $Matches.Groups[5].Value.Trim()
			} else {
				$Matches = ($Password | Select-String -pattern '(([A-Za-z0-9]{5}-){4}[A-Za-z0-9]{5})').Matches
				if ($Matches -and $Matches.Groups -and $Matches.Groups[1]) {
					$Password = $Matches.Groups[1].Value.Trim()
				} else {
					continue;
				}
			}
		}

		$Matching = @{}
		$MatchingColumns | ForEach-Object { 
			$Type = $Headers[$_]
			if ($Type -like "Asset*") {
				$Type = "AssetTag"
			} elseif ($Type -like "Serial*" -or $Type -like "Service Tag") {
				$Type = "SerialNumber"
			} elseif ($Type -like "Name" -and $ImportType.Embedded -like "Configuration") {
				$Type = "Hostname"
			} elseif ($Type -like "User" -or $Type -like "Full Name") {
				$Type = "Name"
			} elseif ($Type -like "First*") {
				$Type = "FirstName"
			} elseif ($Type -like "Last*") {
				$Type = "LastName"
			} elseif ($Type -like "Email*") {
				$Type = "Email"
			}
			# TODO: Show a form to match any types not in the above list
			$Matching[$Type] = $Row.cells[$_].innerText
		}

		$Passwords_Parsed += @{
			Username = $Username
			Password = $Password.Trim()
			Name = $Name
			Matching = $Matching
		}
	}

	if ($ImportType.Embedded -like "Configuration" -or $ImportType.Linked -like "Configuration") {
		# Get full configurations list from ITG
		Write-Host "Downloading all ITG configurations"
		$FullConfigurationsList = (Get-ITGlueConfigurations -page_size 1000 -organization_id $OrgID).data
	} elseif ($ImportType.Embedded -like "Contact" -or $ImportType.Linked -like "Contact") {
		Write-Host "Downloading all ITG contacts"
		$FullContactList = (Get-ITGlueContacts -page_size 1000 -organization_id $OrgID).data
	}

	$CategoryID = getPasswordCategoryID -Category $ImportType.Category

	# Get full passwords list from ITG for the import type category (to avoid making duplicates)
	$FullPasswordList = (Get-ITGluePasswords -page_size 10000 -organization_id $OrgID -filter_password_category_id $CategoryID).data

	# We now have a nicely formatted list of passwords, lets query the ITG data for matches and start adding them
	$ITGCreatePasswords = @()
	$RelatedItems = @{}
	foreach ($PasswordInfo in $Passwords_Parsed) {
		if (!$PasswordInfo.Name -or !$PasswordInfo.Password -or ($ImportType.Embedded -and !$PasswordInfo.Matching)) {
			continue;
		}

		$PasswordName =  "{0} - {1}" -f $ImportType.Name, $PasswordInfo.Name

		$ExistingPasswords = $FullPasswordList | Where-Object { $_.attributes.name -like $PasswordName -or ($_.attributes.name -like "*"+$ImportType.Name+"*" -and $_.attributes.name -like "*"+$PasswordInfo.Name+"*") }

		if (($ExistingPasswords | Measure-Object).Count -gt 0) {
			Write-Host "Check on password, may need updating: $PasswordName" -ForegroundColor Yellow
			if ($OverwritePasswords) {
				# TODO: Update ITG
			} else {
				continue;
			}
		}

		$MatchingAsset = @()
		if (($ImportType.Embedded -eq "Configuration" -or $ImportType.Linked -eq "Configuration") -and $PasswordInfo.Matching) {
			if (!$MatchingAsset -and $PasswordInfo.Matching.SerialNumber) {
				$MatchingAsset = $FullConfigurationsList | Where-Object { $_.attributes.'serial-number' -like $PasswordInfo.Matching.SerialNumber }
			}
			if (!$MatchingAsset -and $PasswordInfo.Matching.AssetTag) {
				$MatchingAsset = $FullConfigurationsList | Where-Object { $_.attributes.'asset-tag' -like $PasswordInfo.Matching.AssetTag }
			}
			if (!$MatchingAsset -and $PasswordInfo.Matching.Hostname) {
				$MatchingAsset = $FullConfigurationsList | Where-Object { $_.attributes.'hostname' -like $PasswordInfo.Matching.Hostname -or $_.attributes.'name' -like $PasswordInfo.Matching.Hostname }
			}
			if (!$MatchingAsset -and $PasswordInfo.Matching.AssetTag) {
				$MatchingAsset = $FullConfigurationsList | Where-Object { $_.attributes.'hostname' -like ($DevicePrefix + '-' + $PasswordInfo.Matching.AssetTag) -or $_.attributes.'name' -like ($DevicePrefix + '-' + $PasswordInfo.Matching.AssetTag) }
			}
			if (!$MatchingAsset -and $PasswordInfo.Matching.AssetTag) {
				$MatchingAsset = $FullConfigurationsList | Where-Object { $_.attributes.hostname -like ("*-" + $PasswordInfo.Matching.AssetTag) -or $_.attributes.name -like ("*-" + $PasswordInfo.Matching.AssetTag) }
			}
		} elseif (($ImportType.Embedded -eq "Contact" -or $ImportType.Linked -eq "Contact") -and $PasswordInfo.Matching) {
			if (!$MatchingAsset -and $PasswordInfo.Matching.Name -and $PasswordInfo.Matching.Email) {
				$MatchingAsset = $FullContactList | Where-Object { $_.attributes.'name' -like $PasswordInfo.Matching.Name -and ($_.attributes.'contact-emails' | Where-Object { $_.primary -eq 'True' }).value -like $PasswordInfo.Matching.Email }
			}
			if (!$MatchingAsset -and $PasswordInfo.Matching.Name -and $PasswordInfo.Matching.Email) {
				$MatchingAsset = $FullContactList | Where-Object { $_.attributes.'name' -like $PasswordInfo.Matching.Name -and $_.attributes.'contact-emails'.value -like $PasswordInfo.Matching.Email }
			}
			if (!$MatchingAsset -and $PasswordInfo.Username) {
				$MatchingAsset = $FullContactList | Where-Object { $_.attributes.'notes' -match ".*(Username: " + $PasswordInfo.Username + "(\s|W|$)).*" }
			}
			if (!$MatchingAsset -and $PasswordInfo.Matching.Name) {
				$MatchingAsset = $FullContactList | Where-Object { $_.attributes.'name' -like $PasswordInfo.Matching.Name }
			}
			if (!$MatchingAsset -and $PasswordInfo.Matching.FirstName -and $PasswordInfo.Matching.LastName) {
				$MatchingAsset = $FullContactList | Where-Object { $_.attributes.'first-name' -like $PasswordInfo.Matching.FirstName -and $_.attributes.'last-name' -like $PasswordInfo.Matching.LastName }
			}
			if (!$MatchingAsset -and $PasswordInfo.Matching.Email) {
				$MatchingAsset = $FullContactList | Where-Object { ($_.attributes.'contact-emails' | Where-Object { $_.primary -eq 'True' }).value -like $PasswordInfo.Matching.Email }
			}
			if (!$MatchingAsset -and $PasswordInfo.Matching.Email) {
				$MatchingAsset = $FullContactList | Where-Object { $_.attributes.'contact-emails'.value -like $PasswordInfo.Matching.Email }
			}
		}


		$PasswordAssetBody = 
		@{
			type = 'passwords'
			attributes = @{
				"name" = $PasswordName
				"username" = $PasswordInfo.Username
				"password" = $PasswordInfo.Password
				"password-category-id" = $CategoryID
				"password-category-name" = $ImportType.Category
			}
		}

		if ($FolderID) {
			$PasswordAssetBody.attributes.'password-folder-id' = $FolderID
		}

		if ($ImportType.Notes) {
			$PasswordAssetBody.attributes.notes = $ImportType.Notes
		}

		if (($ImportType.Embedded -or $ImportType.Linked) -and !$MatchingAsset) {
			# If no matching asset, display a form to let the user choose the device
			$Form = loadForm -Path(".\Forms\ManualMatching\ManualMatching\MainWindow.xaml")

			$var_lblPasswordName.Content = $PasswordName

			$MatchingNotes = @()
			if (($ImportType.Embedded -eq "Configuration" -or $ImportType.Linked -eq "Configuration") -and $PasswordInfo.Matching) {
				if ($PasswordInfo.Matching.Hostname) {
					$MatchingNotes += "Hostname: " + $PasswordInfo.Matching.Hostname
				}
				if ($PasswordInfo.Matching.AssetTag) {
					$MatchingNotes += "Asset Tag: " + $PasswordInfo.Matching.AssetTag
				}
				if ($PasswordInfo.Matching.SerialNumber) {
					$MatchingNotes += "S/N: " + $PasswordInfo.Matching.SerialNumber
				}
			} elseif (($ImportType.Embedded -eq "Contact" -or $ImportType.Linked -eq "Contact") -and $PasswordInfo.Matching) {
				if ($PasswordInfo.Matching.Name) {
					$MatchingNotes += "Name: " + $PasswordInfo.Matching.Name
				} elseif ($PasswordInfo.Matching.FirstName -and $PasswordInfo.Matching.LastName) {
					$MatchingNotes += "Name: " + $PasswordInfo.Matching.FirstName + " " + $PasswordInfo.Matching.LastName
				}
				if ($PasswordInfo.Username) {
					$MatchingNotes += "Username: " + $PasswordInfo.Username
				}
				if ($PasswordInfo.Matching.Email) {
					$MatchingNotes += "Email: " + $PasswordInfo.Matching.Email
				}	
			}

			$var_lblMatchingNotes.Content = $MatchingNotes -join ", "

			function cmbItems($Items, $Filter = "") {
				$FilteredItems = $Items | Where-Object { $_ -like "*$Filter*" } | Sort-Object
				$var_cmbMatch.Items.Clear()
				foreach ($Item in $FilteredItems) {
					$var_cmbMatch.Items.Add($Item) | Out-Null
				}
			}

			# update the listbox with the configurations / contacts
			$Items = $null
			if ($ImportType.Embedded -like "Configuration" -or $ImportType.Linked -like "Configuration") {
				$Items = $FullConfigurationsList.attributes.'name'
			} elseif ($ImportType.Embedded -like "Contact" -or $ImportType.Linked -like "Contact") {
				$Items = $FullContactList.attributes.'name'
			}

			cmbItems -Items $Items

			$var_cmbMatch.Add_KeyUp({
				if ($_.Key -eq "Down" -or $_.Key -eq "Up") {
					$var_cmbMatch.IsDropDownOpen = $true
				} elseif ($_.Key -ne "Enter" -and $_.Key -ne "Tab" -and $_.Key -ne "Return") {
					$var_cmbMatch.IsDropDownOpen = $true
					cmbItems -Items $Items -Filter $var_cmbMatch.Text
				}
			})

			$var_cmbMatch.Add_SelectionChanged({
				$SelectedAsset = $var_cmbMatch.SelectedItem
				if ($ImportType.Embedded -eq "Configuration" -or $ImportType.Linked -eq "Configuration") {
					$script:MatchingAsset = $FullConfigurationsList | Where-Object { $_.attributes.'name' -eq $SelectedAsset }
				} elseif ($ImportType.Embedded -eq "Contact" -or $ImportType.Linked -eq "Contact") {
					$script:MatchingAsset = $FullContactList | Where-Object { $_.attributes.'name' -like $SelectedAsset }
				}
			})

			$var_btnIgnore.Add_Click({
				Write-Host "Password skipped! ($PasswordName)"
				$script:MatchingAsset = @()
				$Form.Close()
				continue;
			})

			$var_btnSave.Add_Click({
				$Form.Close()
			})

			$Form.ShowDialog() | out-null
		}

		# If multiple matches, narrow down
		if (($ImportType.Embedded -eq "Configuration" -or $ImportType.Linked -eq "Configuration") -and ($MatchingAsset | Measure-Object).Count -gt 1) {
			$MatchingAsset = $MatchingAsset | Where-Object { $_.attributes.archived -eq $false }
			if (($MatchingAsset | Measure-Object).Count -gt 1) {
				$MatchingAsset = $MatchingAsset | Sort-Object -Property id -Descending | Select-Object -First 1
			}
		} elseif (($ImportType.Embedded -eq "Contact" -or $ImportType.Linked -eq "Contact") -and ($MatchingAsset | Measure-Object).Count -gt 1) {
			$MatchingAsset = $MatchingAsset | Where-Object { $_.attributes.'contact-type-name' -like "Employee*" }
			if (($MatchingAsset | Measure-Object).Count -gt 1) {
				$MatchingAsset = $MatchingAsset | Sort-Object -Property id -Descending | Select-Object -First 1
			}
		}

		if ($ImportType.Embedded -and $MatchingAsset) {
			$PasswordAssetBody.attributes.'resource-id' = $MatchingAsset.id
			$PasswordAssetBody.attributes.'resource-name' = $MatchingAsset.attributes.name
			if ($ImportType.Embedded -like 'Contact') {
				$PasswordAssetBody.attributes.'resource-type' = 'Contact'
			} elseif ($ImportType.Embedded -like "Configuration") {
				$PasswordAssetBody.attributes.'resource-type' = 'Configuration'
			}
		}

		if ($ImportType.Linked -and $MatchingAsset) {
			$RelatedItemsBody = @{
				type = 'related_items'
				attributes = @{
					'destination-id' = $MatchingAsset.id
					'destination-name' = $MatchingAsset.attributes.name
				}
			}
			if ($ImportType.Linked -like 'Contact') {
				$RelatedItemsBody.attributes.'destination-type' = 'Contact'
				$RelatedItemsBody.attributes.'notes' = 'Password is for this contact.'
			} elseif ($ImportType.Linked -like "Configuration") {
				$RelatedItemsBody.attributes.'destination-type' = 'Configuration'
				$RelatedItemsBody.attributes.'notes' = 'Password is for this configuration.'
			}
			$RelatedItems[$PasswordAssetBody.attributes.name -replace '[\W]', ''] = $RelatedItemsBody
		}

		if ($PasswordInfo.Password -match "\s") {
			Write-Host "The password for '$PasswordName' contains white space and is likely 2 different passwords. Please review and fix this manually." -ForegroundColor Red
		}

		$ITGCreatePasswords += $PasswordAssetBody
	}
	
	if ($ITGCreatePasswords) {
		$ITGCreatePasswords.attributes | Select-Object -Property @{Name="name"; E={$_.name}}, @{Name="username"; E={$_.username}}, @{Name="password"; E={$_.password}}, @{Name = "matched resource name"; E={
			if ($_."resource-name") { $_."resource-name" } elseif ($RelatedItems[$_.name -replace '[\W]', '']) { $RelatedItems[$_.name -replace '[\W]', ''].attributes.'destination-name' }
		}} |
			Sort-Object -Property name | Out-GridView -Title "Passwords Collected. Please Review." -PassThru
		Write-Host "Uploading passwords..."
		if (($ITGCreatePasswords | Measure-Object).Count -gt 60) {
			# if more than 60 passwords, upload in batch of 50 ($Batch)
			$Batch = 50
			$Response = $null
			for ($i = 0; $i -lt [Math]::Ceiling(($ITGCreatePasswords | Measure-Object).Count / $Batch); $i++) {
				$ThisResponse = New-ITGluePasswords -organization_id $orgID -data ($ITGCreatePasswords | Select-Object -Skip ($i*$Batch) -First $Batch)
				if ($Response) {
					$Response.data += $ThisResponse.data
				} else {
					$Response = $ThisResponse
				}
			}
		} else {
			$Response = New-ITGluePasswords -organization_id $orgID -data $ITGCreatePasswords
		}
		Write-Host "Passwords uploaded!" -ForegroundColor Green

		if ($RelatedItems.count -gt 0) {
			Write-Host "Creating related items..."
			foreach ($Password in $Response.data) {
				$Name = $Password.attributes.Name -replace '[\W]', ''
				if ($RelatedItems[$Name]) {
					New-ITGlueRelatedItems -resource_type 'passwords' -resource_id $Password.id -data $RelatedItems[$Name] | Out-Null
				}
			}
			Write-Host "Related items created!" -ForegroundColor Green
		}
	}

} else {
	Write-Host "No table found for import. Please try again." -ForegroundColor Red
	Read-Host "Press ENTER to close..." 
}
Read-Host "Press ENTER to close..." 