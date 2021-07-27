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

$ConfigList = (Get-ITGlueConfigurations -page_size 1000 -organization_id $OrgID).data.attributes | Out-GridView -PassThru
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