###############################################################################
#Connect-MgGraph
#Connect to your Azure Active Directory with "Application Adminstrator" or "Global Administrator" Role
###############################################################################


param (
    [Parameter(Mandatory = $true, ParameterSetName = "CreateAzureApplication")]
    [Switch]$CreateAzureApplication,

    [Parameter(Mandatory = $true, ParameterSetName = "DeleteAzureApplication")]
    [Switch]$DeleteAzureApplication
)

If ($CreateAzureApplication)
{
	#Connect-MgGraph
	Disconnect-MgGraph -ErrorAction SilentlyContinue
	Connect-MgGraph -Scopes "Application.Read.All","Application.ReadWrite.All","User.Read.All","AppRoleAssignment.ReadWrite.All"
	$TenantId = (Get-MgContext).TenantId

	#Create AAD Application
	$AppName =  "RoomStatisticsGraph"
	$App = New-MgApplication -DisplayName $AppName
	$AppID = $App.AppId
	$APPObjectID = $App.Id

	#Create a Self Signed Certificate
	$Subject = $AppName
	$NotAfter = (Get-Date).AddMonths(+24)
	$Cert = New-SelfSignedCertificate -Subject $Subject -CertStoreLocation "Cert:\CurrentUser\My" -KeySpec Signature -NotAfter $Notafter -KeyExportPolicy Exportable
	#$ThumbPrint = $Cert.ThumbPrint

	<#
	#View Certificates in the Current User Certificate Store
	Get-ChildItem -Path cert:\CurrentUser\my\$ThumbPrint | Format-Table

	#Export Certificate as Base64 (PEM Format)
	$CurrentLocation = (Get-Location).path
	$Base64 = [convert]::tobase64string((get-item cert:\currentuser\my\$ThumbPrint).RawData)
	$Base64Block = $Base64 |
	ForEach-Object {
		$line = $_

		for ($i = 0; $i -lt $Base64.Length; $i += 64)
		{
			$length = [Math]::Min(64, $line.Length - $i)
			$line.SubString($i, $length)
		}
	}
	$base64Block2 = $Base64Block | Out-String

	$Value = "-----BEGIN CERTIFICATE-----`r`n"
	$Value += "$Base64Block2"
	$Value += "-----END CERTIFICATE-----"
	$Value
	Set-Content -Path "$CurrentLocation\$Subject-BASE64.cer" -Value $Value
	#>

	#Add Certificate to AzureAD App
	$keyCreds = @{ 
		Type = "AsymmetricX509Cert";
		Usage = "Verify";
		key = $cert.RawData
	}
	
	try {
		Update-MgApplication -ApplicationId $APPObjectID  -KeyCredentials $keyCreds
	} catch {
		Write-Error $Error[0]
	}

	#Add Application Permissions
	#Calendars.Read	Application	798ee544-9d2d-430c-a058-570e29e34338
	#Calendars.ReadBasic.All	Application	8ba4a692-bc31-4128-9094-475872af8a53
	$params = @{
		RequiredResourceAccess = @(
			@{
				ResourceAppId = "00000003-0000-0000-c000-000000000000"
				ResourceAccess = @(
					@{
						Id = "798ee544-9d2d-430c-a058-570e29e34338"
						Type = "Role"
					}
				)
			}
		)
	}
	Update-MgApplication -ApplicationId $APPObjectID -BodyParameter $params

	#Grant Permission
	#Get-AzureADServicePrincipalOAuth2PermissionGrant
	<#
	$ServicePrincipal = Get-MgServicePrincipal -Filter "DisplayName eq 'RoomStatisticsGraph'"
	$params = @{
		PrincipalId = "bb6c9bfd-7426-40ea-9292-0013f75a7e85"
		ResourceId = "00000003-0000-0000-c000-000000000000"
		AppRoleId = "798ee544-9d2d-430c-a058-570e29e34338"
	}
	New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipalId -BodyParameter $params
	#>
	
	#Grants interactive Consent
	$URL = "https://login.microsoftonline.com/$TenantId/adminconsent?client_id=$APPId"
	Start-Process $URL


	#Add Redirect URI
	$RedirectURI = @()
	$RedirectURI += "https://login.microsoftonline.com/common/oauth2/nativeclient"
	$RedirectURI += "msal" + $AppId + "://auth"

	$params = @{
		RedirectUris = @($RedirectURI)
	}
	Update-MgApplication -ApplicationId $APPObjectID -IsFallbackPublicClient -PublicClient $params
}