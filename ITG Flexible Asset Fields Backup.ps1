###
# File: \ITG Flexible Asset Fields Backup.ps1
# Project: Scripts
# Created Date: Thursday, March 9th 2023, 1:10:11 pm
# Author: Chris Jantzen
# -----
# Last Modified: Thu Mar 09 2023
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
$BackupsToKeep = 10
#####################################################################

Write-Host "Creating backup directory" -ForegroundColor Green
if (!(Test-Path $ExportDir)) { new-item $ExportDir -ItemType Directory }
$DateTimestamp = $(get-date -f yyyy-MM-dd_HH-mm-ss)
$FullExportDir = "$($ExportDir)\$($DateTimestamp)"
New-Item $FullExportDir -ItemType Directory

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

Function Remove-InvalidFileNameChars {
	param(
	  [Parameter(Mandatory=$true,
		Position=0,
		ValueFromPipeline=$true,
		ValueFromPipelineByPropertyName=$true)]
	  [String]$Name
	)
  
	$invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
	$re = "[{0}]" -f [RegEx]::Escape($invalidChars)
	return ($Name -replace $re)
}

Write-Host "Getting Flexible Assets" -ForegroundColor Green
$i = 0
$FlexAssetTypes = (Get-ITGlueFlexibleAssetTypes -page_size 1000).data
foreach ($FlexAsset in $FlexAssetTypes) {
	$FlexAssetFields = Get-ITGlueFlexibleAssetFields -flexible_asset_type_id $FlexAsset.id
	if ($FlexAssetFields -and $FlexAssetFields.data) {
		$i++
		$FlexAssetFields = $FlexAssetFields.data

		$FlexAssetObject = $FlexAsset.attributes
		$FlexAssetObject | Add-Member -NotePropertyName fields -NotePropertyValue $null
		$FlexAssetObject.fields = $FlexAssetFields.attributes

		$Filename = (Remove-InvalidFileNameChars $FlexAsset.attributes.name).ToLower() -replace ' ', '_'
		$FlexAssetObject | ConvertTo-Json -Depth 8 | Out-File -FilePath "$($FullExportDir)\$($Filename).json"
		Write-Host "Backed up flexible asset: $($FlexAsset.attributes.name) to $Filename"
	}
}

Write-Host "Backed up $i flexible assets."

# Cleanup old folders
$ExistingFolders = Get-ChildItem -Path $ExportDir -Directory
$BackupFoldersToKeep = $ExistingFolders.Name | Sort-Object -Descending | Select-Object -First $BackupsToKeep

foreach ($Folder in $ExistingFolders) {
	if ($Folder -notin $BackupFoldersToKeep) {
		Remove-Item -LiteralPath "$($ExportDir)\$($Folder)" -Force -Recurse
	}
}
