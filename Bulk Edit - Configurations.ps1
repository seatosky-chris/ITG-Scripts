#####################################################################
$APIKEy =  "<ITG API KEY>"
$APIEndpoint = "https://api.itglue.com"
$orgID = "ORGIDHERE"
$NewGateway = "192.1.1.254"
#####################################################################
If(Get-Module -ListAvailable -Name "ITGlueAPI") {Import-module ITGlueAPI} Else { install-module ITGlueAPI -Force; import-module ITGlueAPI}
#Settings IT-Glue logon information
Add-ITGlueBaseURI -base_uri $APIEndpoint
Add-ITGlueAPIKey $APIKEy

$ConfigList = Get-ITGlueConfigurations -page_size "1000" -organization_id $OrgID
$i = 1
while ($ConfigList.links.next) {
	$i++
	$Configurations_Next = Get-ITGlueConfigurations -page_size "1000" -page_number $i -organization_id $OrgID
    if (!$Configurations_Next -or $Configurations_Next.Error) {
		# We got an error querying configurations, wait and try again
		Start-Sleep -Seconds 2
		$Configurations_Next = Get-ITGlueConfigurations -page_size "1000" -page_number $i -organization_id $OrgID

		if (!$Configurations_Next -or $Configurations_Next.Error) {
			Write-Error "An error occurred trying to get the existing configurations from ITG. Exiting..."
			Write-Error $Configurations_Next.Error
			exit 1
		}
	}
	$ConfigList.data += $Configurations_Next.data
	$ConfigList.links = $Configurations_Next.links
}

if (!$ConfigList) {
	Write-Warning "There was an issue getting the Configurations from ITG. Exiting..."
	exit 1
}

$ConfigList = $ConfigList.data.attributes | Out-GridView -PassThru

foreach($Config in $ConfigList){
    $ConfigID = ($config.'resource-url' -split "/")[-1]
    $UpdatedConfig = 
        @{
            type = 'Configurations'
            attributes = @{
                        "default-gateway" = $NewGateway
            }
        }
    Set-ITGlueConfigurations -id $ConfigID -data $UpdatedConfig
}