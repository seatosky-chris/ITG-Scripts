####################################################################
$APIKEy =  "<ITG API KEY>"
$APIEndpoint = "https://api.itglue.com"
$orgID = "<ITG ORG ID>"
$DevicePrefix =  "<DEVICE PREFIX>" # Generally the companies acronym
$SelectFormPath = ".\Forms\SelectForm\SelectForm\MainWindow.xaml"

$ImportTypes = @(
	@{ "Name" = "Notes - Contacts"; "SelectLabel" = "Notes - Contacts"; "Embedded" = "Contact"; "Defaults" = @{ "Notes" = "Notes" } }
	@{ "Name" = "Notes - Configurations"; "SelectLabel" = "Notes - Configurations"; "Embedded" = "Configuration"; "Defaults" = @{ "Notes" = "Notes" } }
)
####################################################################


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Web

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

# Select what type of import this is (stored in $ImportType)
$Form = loadForm -Path($SelectFormPath)
$Form.Title = "Select the import type"
$var_lblDescription.Content = "Select the type of import you are doing:"
$i = 0
foreach ($ImportType in $ImportTypes) { 
	$var_lstSelectOptions.Items.Insert($i, $ImportType.SelectLabel) | Out-Null
	$i++
}
$var_txtNotes.Text = "Use 'Notes - Contacts' for uploading to contacts, and 'Notes - Configurations' for uploading to configurations."
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

	# Choose the notes column (stored in $NotesColumn)
	if ($ImportType.Defaults -and $ImportType.Defaults.Notes -and ($Headers | Where-Object { $_ -like $ImportType.Defaults.Notes })) {
		$NotesColumn = $Headers.IndexOf($ImportType.Defaults.Notes);
	} else {
		$Form = loadForm -Path($SelectFormPath)
		$Form.Title = "Select the notes column"
		$var_lblDescription.Content = "Select the column with the notes to import:"
		$i = 0
		foreach ($Col in $Headers) { 
			$var_lstSelectOptions.Items.Insert($i, $Col) | Out-Null
			$i++
		}
		$var_grpNotes.Visibility = "Hidden"
		$var_btnSave.IsEnabled = $false

		$NotesColumn = $null
		$var_lstSelectOptions.Add_SelectionChanged({
			$var_btnSave.IsEnabled = $true
			$script:NotesColumn = $var_lstSelectOptions.SelectedIndex
		})

		$var_btnSave.Add_Click({
			$Form.Close()
		})

		$Form.ShowDialog() | out-null
	}

	if ($ImportType.Embedded) {
		# Choose the column for matching to the embedded/linked asset (stored in $MatchingColumns)
		$Form = loadForm -Path($SelectFormPath)
		$Form.Title = "Select the embed matching column(s)"
		$var_lblDescription.Content = "Select the column(s) for matching this to the embedded asset:"
		$var_lstSelectOptions.SelectionMode = "Multiple"

		$i = 0
		foreach ($Col in $Headers) { 
			$var_lstSelectOptions.Items.Insert($i, $Col) | Out-Null
			$i++
		}
		if ($ImportType.Embedded -like "Configuration") {
			$var_txtNotes.Text = "Choose columns that contain asset tags, hostnames, and/or serial numbers."
		} elseif ($ImportType.Embedded -like "Contact") {
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

	# We now know which columns to use, lets loop through the table and save each note and its associated values in an array
	$Notes_Parsed = @()
	$Length = $Table.rows.length 
	for ($i = 1; $i -lt $Length; $i++) {
		$Row = $Table.rows[$i]

		if ($Row.cells[0].innerText -eq $Headers[0] -and $Row.cells[1].innerText -eq $Headers[1]) {
			continue; # skip header row
		}

		$Notes = $Row.cells[$NotesColumn].innerText
		if (!$Notes) {
			continue;
		}
		$Notes = [System.Web.HttpUtility]::HtmlDecode($Notes.Trim())
		
		$Matching = @{}
		$MatchingColumns | ForEach-Object { 
			$Type = $Headers[$_]
			if ($Type -like "Asset*") {
				$Type = "AssetTag"
			} elseif ($Type -like "Serial*" -or $Type -like "Service Tag") {
				$Type = "SerialNumber"
			} elseif ($Type -like "Name" -and $ImportType.Embedded -like "Configuration") {
				$Type = "Hostname"
			} elseif ($Type -like "User" -or $Type -like "Full Name" -or $Type -like "Account Name") {
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

		$Notes_Parsed += @{
			Notes = $Notes
			Matching = $Matching
		}
	}

	if ($ImportType.Embedded -like "Configuration") {
		# Get full configurations list from ITG
		Write-Host "Downloading all ITG configurations"
		$FullConfigurationsList = (Get-ITGlueConfigurations -page_size 1000 -organization_id $OrgID).data
	} elseif ($ImportType.Embedded -like "Contact") {
		Write-Host "Downloading all ITG contacts"
		$FullContactList = (Get-ITGlueContacts -page_size 1000 -organization_id $OrgID).data
	}

	# We now have a nicely formatted list of notes, lets query the ITG data for matches and start adding them
	$ITGCreateNotes = @()
	foreach ($NotesInfo in $Notes_Parsed) {
		if (!$NotesInfo.Notes -or ($ImportType.Embedded -and !$NotesInfo.Matching)) {
			continue;
		}

		$MatchingAsset = @()
		if ($ImportType.Embedded -eq "Configuration" -and $NotesInfo.Matching) {
			if (!$MatchingAsset -and $NotesInfo.Matching.SerialNumber) {
				$MatchingAsset = $FullConfigurationsList | Where-Object { $_.attributes.'serial-number' -like $NotesInfo.Matching.SerialNumber }
			}
			if (!$MatchingAsset -and $NotesInfo.Matching.AssetTag) {
				$MatchingAsset = $FullConfigurationsList | Where-Object { $_.attributes.'asset-tag' -like $NotesInfo.Matching.AssetTag }
			}
			if (!$MatchingAsset -and $NotesInfo.Matching.Hostname) {
				$MatchingAsset = $FullConfigurationsList | Where-Object { $_.attributes.'hostname' -like $NotesInfo.Matching.Hostname -or $_.attributes.'name' -like $NotesInfo.Matching.Hostname }
			}
			if (!$MatchingAsset -and $NotesInfo.Matching.AssetTag) {
				$MatchingAsset = $FullConfigurationsList | Where-Object { $_.attributes.'hostname' -like ($DevicePrefix + '-' + $NotesInfo.Matching.AssetTag) -or $_.attributes.'name' -like ($DevicePrefix + '-' + $NotesInfo.Matching.AssetTag) }
			}
			if (!$MatchingAsset -and $NotesInfo.Matching.AssetTag) {
				$MatchingAsset = $FullConfigurationsList | Where-Object { $_.attributes.hostname -like ("*-" + $NotesInfo.Matching.AssetTag) -or $_.attributes.name -like ("*-" + $NotesInfo.Matching.AssetTag) }
			}
		} elseif ($ImportType.Embedded -eq "Contact" -and $NotesInfo.Matching) {
			if (!$MatchingAsset -and $NotesInfo.Matching.Name -and $NotesInfo.Matching.Email) {
				$MatchingAsset = $FullContactList | Where-Object { $_.attributes.'name' -like $NotesInfo.Matching.Name -and ($_.attributes.'contact-emails' | Where-Object { $_.primary -eq 'True' }).value -like $NotesInfo.Matching.Email }
			}
			if (!$MatchingAsset -and $NotesInfo.Matching.Name -and $NotesInfo.Matching.Email) {
				$MatchingAsset = $FullContactList | Where-Object { $_.attributes.'name' -like $NotesInfo.Matching.Name -and $_.attributes.'contact-emails'.value -like $NotesInfo.Matching.Email }
			}
			if (!$MatchingAsset -and $NotesInfo.Matching.Name) {
				$MatchingAsset = $FullContactList | Where-Object { $_.attributes.'name' -like $NotesInfo.Matching.Name }
			}
			if (!$MatchingAsset -and $NotesInfo.Matching.FirstName -and $NotesInfo.Matching.LastName) {
				$MatchingAsset = $FullContactList | Where-Object { $_.attributes.'first-name' -like $NotesInfo.Matching.FirstName -and $_.attributes.'last-name' -like $NotesInfo.Matching.LastName }
			}
			if (!$MatchingAsset -and $NotesInfo.Matching.Email) {
				$MatchingAsset = $FullContactList | Where-Object { ($_.attributes.'contact-emails' | Where-Object { $_.primary -eq 'True' }).value -like $NotesInfo.Matching.Email }
			}
			if (!$MatchingAsset -and $NotesInfo.Matching.Email) {
				$MatchingAsset = $FullContactList | Where-Object { $_.attributes.'contact-emails'.value -like $NotesInfo.Matching.Email }
			}
		}


		if ($ImportType.Embedded -eq "Contact") {
			$AssetBody = 
			@{
				type = 'contacts'
				attributes = @{
					"notes" = $NotesInfo.Notes
				}
			}
		} elseif ($ImportType.Embedded -eq "Configuration") {
			$AssetBody = 
			@{
				type = 'configurations'
				attributes = @{
					"notes" = $NotesInfo.Notes
				}
			}
		}

		if ($ImportType.Embedded -and !$MatchingAsset) {
			# If no matching asset, display a form to let the user choose the device
			$Form = loadForm -Path(".\Forms\ManualMatching\ManualMatching\MainWindow.xaml")

			$var_lblPasswordName.Content = $PasswordName

			$MatchingNotes = @()
			if ($ImportType.Embedded -eq "Configuration" -and $NotesInfo.Matching) {
				if ($NotesInfo.Matching.Hostname) {
					$MatchingNotes += "Hostname: " + $NotesInfo.Matching.Hostname
				}
				if ($NotesInfo.Matching.AssetTag) {
					$MatchingNotes += "Asset Tag: " + $NotesInfo.Matching.AssetTag
				}
				if ($NotesInfo.Matching.SerialNumber) {
					$MatchingNotes += "S/N: " + $NotesInfo.Matching.SerialNumber
				}
			} elseif ($ImportType.Embedded -eq "Contact" -and $NotesInfo.Matching) {
				if ($NotesInfo.Matching.Name) {
					$MatchingNotes += "Name: " + $NotesInfo.Matching.Name
				} elseif ($NotesInfo.Matching.FirstName -and $NotesInfo.Matching.LastName) {
					$MatchingNotes += "Name: " + $NotesInfo.Matching.FirstName + " " + $NotesInfo.Matching.LastName
				}
				if ($NotesInfo.Matching.Email) {
					$MatchingNotes += "Email: " + $NotesInfo.Matching.Email
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
			if ($ImportType.Embedded -like "Configuration") {
				$Items = $FullConfigurationsList.attributes.'name'
			} elseif ($ImportType.Embedded -like "Contact") {
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
				if ($ImportType.Embedded -eq "Configuration") {
					$script:MatchingAsset = $FullConfigurationsList | Where-Object { $_.attributes.'name' -eq $SelectedAsset }
				} elseif ($ImportType.Embedded -eq "Contact") {
					$script:MatchingAsset = $FullContactList | Where-Object { $_.attributes.'name' -like $SelectedAsset }
				}
			})

			$var_btnIgnore.Add_Click({
				Write-Host "Notes skipped! ($($NotesInfo.Notes))"
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
		if ($ImportType.Embedded -eq "Configuration" -and ($MatchingAsset | Measure-Object).Count -gt 1) {
			$MatchingAsset = $MatchingAsset | Where-Object { $_.attributes.archived -eq $false }
			if (($MatchingAsset | Measure-Object).Count -gt 1) {
				$MatchingAsset = $MatchingAsset | Sort-Object -Property id -Descending | Select-Object -First 1
			}
		} elseif ($ImportType.Embedded -eq "Contact" -and ($MatchingAsset | Measure-Object).Count -gt 1) {
			$MatchingAsset = $MatchingAsset | Where-Object { $_.attributes.'contact-type-name' -like "Employee*" }
			if (($MatchingAsset | Measure-Object).Count -gt 1) {
				$MatchingAsset = $MatchingAsset | Sort-Object -Property id -Descending | Select-Object -First 1
			}
		}

		if ($MatchingAsset) {
			$AssetBody.attributes.id = $MatchingAsset.id
			if ($MatchingAsset.attributes.notes) {
				if ($MatchingAsset.attributes.notes -like "*$($NotesInfo.Notes)*") {
					continue;
				}
				$AssetBody.attributes.notes = "$($MatchingAsset.attributes.notes) <br> $($NotesInfo.Notes)"
			}
		} else {
			continue;
		}

		$ITGCreateNotes += $AssetBody
	}

	if ($ITGCreateNotes) {
		$ITGCreateNotes.attributes | Select-Object -Property @{Name="id"; E={$_.id}}, @{Name="asset"; E={
			$id = $_.id
			if ($ImportType.Embedded -eq "Configuration") {
				($FullConfigurationsList | Where-Object { $_.id -eq $id }).attributes.name
			} elseif ($ImportType.Embedded -eq "Contact") {
				($FullContactList | Where-Object { $_.id -eq $id }).attributes.name
			}
		}}, @{Name="notes"; E={$_.notes}} |
		 Sort-Object -Property id | Out-GridView -Title "Notes Collected. Please Review." -PassThru
		Write-Host "Uploading notes..."
		if (($ITGCreateNotes | Measure-Object).Count -gt 60) {
			# if more than 60 notes, upload in batch of 50 ($Batch)
			$Batch = 50
			$Response = $null
			for ($i = 0; $i -lt [Math]::Ceiling(($ITGCreateNotes | Measure-Object).Count / $Batch); $i++) {
				if ($ImportType.Embedded -eq "Contact") {
					$ThisResponse = Set-ITGlueContacts -organization_id $orgID -data ($ITGCreateNotes | Select-Object -Skip ($i*$Batch) -First $Batch)
				} elseif ($ImportType.Embedded -eq "Configuration") {
					$ThisResponse = Set-ITGlueConfigurations -organization_id $orgID -data ($ITGCreateNotes | Select-Object -Skip ($i*$Batch) -First $Batch)
				}
				if ($Response) {
					$Response.data += $ThisResponse.data
				} else {
					$Response = $ThisResponse
				}
			}
		} else {
			if ($ImportType.Embedded -eq "Contact") {
				$Response = Set-ITGlueContacts -organization_id $orgID -data $ITGCreateNotes
			} elseif ($ImportType.Embedded -eq "Configuration") {
				$Response = Set-ITGlueConfigurations -organization_id $orgID -data $ITGCreateNotes
			}
		}
		Write-Host "Notes uploaded!" -ForegroundColor Green
	}

} else {
	Write-Host "No table found for import. Please try again." -ForegroundColor Red
	Read-Host "Press ENTER to close..." 
}
Read-Host "Press ENTER to close..." 