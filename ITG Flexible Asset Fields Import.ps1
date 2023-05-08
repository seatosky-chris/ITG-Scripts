###
# File: \ITG Flexible Asset Fields Import.ps1
# Project: Scripts
# Created Date: Monday, May 8th 2023, 3:03:48 pm
# Author: Chris Jantzen
# -----
# Last Modified: Mon May 08 2023
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

#####################################################################
$APIKEy =  "APIKEYHERE"
$APIEndpoint = "https://api.itglue.com"
$ExportDir = "C:\Temp\ITG Flexible Assets"
#####################################################################
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

Write-Host "Ensure backup directory exists" -ForegroundColor Green
if (!(Test-Path $ExportDir)) { 
	Write-Host "Backup directory does not exist, please update configuration." -ForegroundColor Red
	exit 1
}

# Ensure they are using the latest TLS version
$CurrentTLS = [System.Net.ServicePointManager]::SecurityProtocol
if ($CurrentTLS -notlike "*Tls12" -and $CurrentTLS -notlike "*Tls13") {
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	Write-Host "This device is using an old version of TLS. Temporarily changed to use TLS v1.2."
}

# Connect to IT Glue
If (Get-Module -ListAvailable -Name "ITGlueAPI") {Import-module ITGlueAPI -Force} Else { install-module ITGlueAPI -Force; import-module ITGlueAPI -Force}
Add-ITGlueBaseURI -base_uri $APIEndpoint
Add-ITGlueAPIKey $APIKEy
Write-Host "Configured the ITGlue API"

Write-Host "Getting Existing Flexible Assets" -ForegroundColor Green
$FlexAssetTypes = (Get-ITGlueFlexibleAssetTypes -page_size 1000).data

# Select backup json file
Write-Host "Choose the json file(s) to import:"
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = $ExportDir; Filter = 'Json files (*.json)|*.json'; MultiSelect = $true; }
$null = $FileBrowser.ShowDialog()

foreach ($FileName in $FileBrowser.FileNames) {

	if ($FileName) {
		$FlexAssetJson = Get-Content $FileName | Out-String | ConvertFrom-Json

		if ($FlexAssetJson -and $FlexAssetJson.name) {
			Write-Host "Loaded flexible asset: $($FlexAssetJson.name)" -ForegroundColor Green
		} else {
			Write-Host "Could not load the flexible asset json file. Exiting..." -ForegroundColor Red
			exit 1
		}
	} else {
		Write-Host "Could not get the file. Exiting..." -ForegroundColor Red
		exit 1
	}

	$NewAssetName = $FlexAssetJson.name
	if ($FlexAssetTypes.attributes.name -contains $FlexAssetJson.name) {
		$i = 2
		while ($i -le 100) {
			if ($FlexAssetTypes.attributes.name -notcontains ($FlexAssetJson.name + " " + $i)) {
				break
			}
			$i++
		}
		$NewAssetName = ($FlexAssetJson.name + " " + $i)
		$CreateDuplicateFlexAsset = [System.Windows.MessageBox]::Show("A flex asset with the name '$($FlexAssetJson.name)' already exists. Would you like to continue importing it? (The new asset will be named: $($NewAssetName))", 'Create Duplicate Flex Asset?', 'YesNo')

		if ($CreateDuplicateFlexAsset -ne "Yes") {
			Write-Host "Skipping asset creation for: $($FlexAssetJson.name)" -ForegroundColor Yellow
			continue
		}
	}

	# Reformat json backup for import
	$ImportJson = @{
		type = "flexible_asset_types"
		attributes = @{
			name = $NewAssetName
			description = $FlexAssetJson.description
			icon = $FlexAssetJson.icon
			"show-in-menu" = $true
		}
		relationships = @{
			"flexible-asset-fields" = @{
				data = @()
			}
		}
	}

	foreach ($FlexAsset_FieldSet in $FlexAssetJson.fields) {
		$ParsedFlexAssetAttributes = @{
			type = "flexible_asset_fields"
			attributes = $FlexAsset_FieldSet
		}

		$ParsedFlexAssetAttributes.attributes.PSObject.Properties.Remove("created-at")
		$ParsedFlexAssetAttributes.attributes.PSObject.Properties.Remove("updated-at")
		$ParsedFlexAssetAttributes.attributes.PSObject.Properties.Remove("flexible-asset-type-id")

		$ImportJson.relationships.'flexible-asset-fields'.data += $ParsedFlexAssetAttributes
	}

	# Create new flexible asset
	$Result = New-ITGlueFlexibleAssetTypes -data $ImportJson

	if ($Result -and $Result.data -and $Result.data.id) {
		Write-Host "Created new flex asset: $($NewAssetName)  (ID: $($Result.data.id))" -ForegroundColor Green
	} else {
		Write-Host "Could not correctly create the new flex asset: $($NewAssetName)" -ForegroundColor Yellow
	}
}
