###############################################################################
# Creates the AzureAD App with Certificate Authentication and Redirect
# Requires Azure Active Directory Role "Application Adminstrator" or "Global Administrator" 
# 
# Creates the Exchange Online Service Principal and RBAC Permissions
# Requires "Exchange Administrator" to create the RBAC Permission
#
# 2023.05.01 - Initial Version - @AndresBohren
###############################################################################
#Require the following Modules
#Requires -Modules ExchangeOnlineManagement
#Requires -Modules Microsoft.Graph.Authentication
#Requires -Modules Microsoft.Graph.Application


param (
    [Parameter(Mandatory = $true, ParameterSetName = "CreateAzureApplication")]
    [Switch]$CreateAzureApplication,

    [Parameter(Mandatory = $true, ParameterSetName = "DeleteAzureApplication")]
    [Switch]$DeleteAzureApplication
)

#Variables
$AppName =  "RoomStatisticsGraph"
$ManagementScopeName = "AllRooms"

#Exchange Connection
Write-Verbose "Check Exchange Connection"
$ExConnection = Get-ConnectionInformation
If ($Null -eq $ExConnection)
{
	Write-Host "Connect to Exchange Online"
	Connect-ExchangeOnline -ShowBanner:$false
} else {
	Write-Host "Connection to Exchange Online already Exists"
}

#Connect-MgGraph
Disconnect-MgGraph -ErrorAction SilentlyContinue
Connect-MgGraph -Scopes "Application.Read.All","Application.ReadWrite.All","User.Read.All","AppRoleAssignment.ReadWrite.All"
$TenantId = (Get-MgContext).TenantId

###############################################################################
# Create Azure AD Application
###############################################################################
If ($CreateAzureApplication)
{

	#Check if Application already Exists	
	$AADApps = Get-MgApplication -ConsistencyLevel eventual -Count appCount -Search "DisplayName:$AppName"
	If ($AADApps.Count -eq 0)
	{

		##Create AAD Application
		$App = New-MgApplication -DisplayName $AppName
		$AppID = $App.AppId
		$APPObjectID = $App.Id


		##Check Certificate
		$Cert = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object {$_.Subject -eq "CN=$AppName"}
		If ($Null -ne $Cert)
		{
			#Certificate Exists already
			$CertExpiryDate = $cert.NotAfter
			If (($CertExpiryDate).AddDays(30) -lt (get-date))
			{
				#Certificate will soon expire
				Write-Host "WARNING: Certificate expiration: $CertExpiryDate" -ForegroundColor Yellow
			} else {
				Write-Verbose "Certificate already Exists"
			}
		} else {
			#Create a Self Signed Certificate
			Write-Verbose "Creating a Self Signed Certificate"
			$Subject = $AppName
			$NotAfter = (Get-Date).AddMonths(+24)
			$Cert = New-SelfSignedCertificate -Subject $Subject -CertStoreLocation "Cert:\CurrentUser\My" -KeySpec Signature -NotAfter $Notafter -KeyExportPolicy Exportable
		}


		##Add Certificate to AzureAD App
		Write-Verbose "Add Certificate to AzureAD App"
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


		##Add Redirect URI
		Write-Verbose "Add Redirect URI to AzureAD App"
		$RedirectURI = @()
		$RedirectURI += "https://login.microsoftonline.com/common/oauth2/nativeclient"
		$RedirectURI += "msal" + $AppId + "://auth"

		$params = @{
			RedirectUris = @($RedirectURI)
		}
		Update-MgApplication -ApplicationId $APPObjectID -IsFallbackPublicClient -PublicClient $params

		#Eventually Wait

		##Check Serviceprincipal | Enterprise App
		Write-Verbose "Check ServicePrincipal"	
		$ServicePrincipalDetails = Get-MgServicePrincipal -ConsistencyLevel eventual -Search "DisplayName:$AppName"
		If ($Null -eq $ServicePrincipalDetails)
		{
			#Service Principal | Enterprise App does not exist
			$ServicePrincipalID=@{
				"AppId" = "$AppID"
				}
				$ServicePrincipalDetails = New-MgServicePrincipal -BodyParameter $ServicePrincipalId 
			  	$ServicePrincipalDetails | Format-List id, DisplayName, AppId, SignInAudience
		} else {
			#Check Exchange Online Service Principal
			$ExServicePrincipal = Get-ServicePrincipal | Where-Object {$_.AppId -eq "$AppID"} -ErrorAction SilentlyContinue
			If ($Null -eq $ExServicePrincipal)
			{	
				#Create Exchange Online Service Principal
				Write-Verbose "Create Exchange Online Service Principal"
				$ExServicePrincipal = New-ServicePrincipal -AppId $ServicePrincipalDetails.AppId -ServiceId $ServicePrincipalDetails.Id -DisplayName "EXO Serviceprincipal $($ServicePrincipalDetails.Displayname)"
			}
		}

		## Check Exchange Online RBAC Management Scope
		Write-Verbose "Check Exchange Online RBAC Management Scope"		
		$EXOManagementScope = Get-ManagementScope -Identity "$ManagementScopeName" -ErrorAction SilentlyContinue
		If ($Null -eq $EXOManagementScope)
		{
			#Create Exchange Online RBAC Management Scope
			Write-Verbose "Create Exchange Online RBAC Management Scope"
			$EXOManagementScope = New-ManagementScope -Name "$ManagementScopeName" -RecipientRestrictionFilter "RecipientTypeDetails -eq 'RoomMailbox'"
		}
		
		## Check Exchange Online Management Role Assignment
		Write-Verbose "Check Exchange Online Management Role Assignment"
		$ServiceId = $ExServicePrincipal.ServiceId
		$MRA = Get-ManagementRoleAssignment  | Where-Object {$_.App -eq $ServiceId}
		If ($Null -eq $MRA)
		{
			#Create Exchange Online Management Role Assignment
			Write-Verbose "Create Exchange Online Management Role Assignment"
			$MRA = New-ManagementRoleAssignment -App $ServiceId -Role "Application Calendars.Read" -CustomResourceScope "$ManagementScopeName"
		} 
		
		
		#DEBUG
		#Write-Host "Exchange Online Management Role Assignment"
		#$MRA | Format-List

		Write-Host "Store this Information" -ForegroundColor Green
		Write-Host "AppID: $AppID"
		Write-Host "CertificateThumbprint: $CertificateThumbprint"
		Write-Host "TenantID: $TenantID"

	} else {
		## Application already Exist
		Write-host "Application already Exists"

		$AppID = $AADApps.AppId
		$TenantID = $AADApps.PublisherDomain

		Write-Host "Store this Information" -ForegroundColor Green
		Write-Host "AppID: $AppID"
		Write-Host "CertificateThumbprint: $CertificateThumbprint"
		Write-Host "TenantID: $TenantID"
	}

}

###############################################################################
# Delete Azure AD Application
###############################################################################
If ($DeleteAzureApplication)
{
	$AADApps = Get-MgApplication -ConsistencyLevel eventual -Count appCount -Search "DisplayName:$AppName"
	If ($AADApps.Count -eq 1)
	{
		$AppID = $AADApps.AppId
		$TenantID = $AADApps.PublisherDomain

		Write-Host "Deleting the Application"

		#Remove Echange Online Management Role
		## Check Exchange Online Management Role Assignment
		Write-Verbose "Check Exchange Online Management Role Assignment"
		$ExServicePrincipal = Get-ServicePrincipal | Where-Object {$_.AppId -eq "$AppID"} -ErrorAction SilentlyContinue
		$ServiceId = $ExServicePrincipal.ServiceId
		$MRA = Get-ManagementRoleAssignment  | Where-Object {$_.App -eq $ServiceId}
		If ($MRA)
		{
			#Remove Exchange Online Management Role Assignment
			Write-Verbose "Remove Exchange Online Management Role Assignment"
			$MRA | Remove-ManagementRoleAssignment
		} 
		## Check Exchange Online RBAC Management Scope
		Write-Verbose "Check Exchange Online RBAC Management Scope"		
		$EXOManagementScope = Get-ManagementScope -Identity "$ManagementScopeName" -ErrorAction SilentlyContinue
		If ($EXOManagementScope)
		{
			#Remove Exchange Online RBAC Management Scope
			Write-Verbose "Remove Exchange Online RBAC Management Scope"
			$EXOManagementScope | Remove-ManagementScope
		}
		
		#Remove AzureAD ServicePrincipal / Enterprise App
		##Check Serviceprincipal | Enterprise App
		Write-Verbose "Check ServicePrincipal"	
		$ServicePrincipalDetails = Get-MgServicePrincipal -ConsistencyLevel eventual -Search "DisplayName:$AppName"
		If ($ServicePrincipalDetails)
		{
			#Service Principal | Enterprise App does exist
			$MgServicePrincipalId = $ServicePrincipalDetails.Id
			$ServicePrincipalDetails | Remove-MgServicePrincipal -ServicePrincipalId $MgServicePrincipalId
		} else {
			#Check Exchange Online Service Principal
			$ExServicePrincipal = Get-ServicePrincipal | Where-Object {$_.AppId -eq "$AppID"} -ErrorAction SilentlyContinue
			If ($ExServicePrincipal)
			{	
				#Remove Exchange Online Service Principal
				Write-Verbose "Remove Exchange Online Service Principal"
				$ExServicePrincipal | Remove-ServicePrincipal -Identity $ServicePrincipalDetails.AppId
			}
		}

		#Remove AzureAD Application
		Write-Verbose "Remove Enterprise Application"
		$AADApps | Remove-MgApplication -ApplicationId $AppID
	}
}