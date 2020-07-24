##############################################################################
# Get-RoomStatistics.ps1
# Created a combination of the two Room Statistics Scripts to fulfill my Needs
# http://stackoverflow.com/questions/10921563/extract-the-report-of-room-calendar-from-exchage-server-using-powershell-scripti
# https://gallery.technet.microsoft.com/office/Exchange-Meeting-Room-2aab769a
#
# Version 1.0 - 21.07.2020 Andres Bohren - Initial Version
#
# Prerequisits
# - Exchange Online Powershell V2 (Module ExchangeOnlineManagement)
# - EWS Managed API 2.2
# - Account with Exchange Administrator Role
# - EWS Account with Impersonation Permission
##############################################################################
<#
.SYNOPSIS
Gather statistics regarding meeting room usage

.DESCRIPTION
This script uses Exchange Web Services to connect to one or more meeting rooms and gather statistics regarding their usage between to specific dates.
The Output will saved to a CSV File

IMPORTANT:
  - Maximum of 1000 meetings per room are returned 
  (Limitation of CalendarView) https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.calendarview.-ctor?view=exchange-ews-api#Microsoft_Exchange_WebServices_Data_CalendarView__ctor_System_DateTime_System_DateTime_System_Int32_
 
.EXAMPLE
 .\Get-RoomStatistics.ps1 -Startdate "01/01/2020" -EndDate "12/31/2020" 

#>


param (
    [Parameter(Mandatory=$true)][DateTime]$StartDate,
    [Parameter(Mandatory=$true)][DateTime]$EndDate
	
)


###############################################################################
#EWS Connection
###############################################################################
[string]$Mailbox = "ewservice@domain.tld"
[string]$Username = "ewservice@domain.tld"
[string]$Password = "YourSecurePassword"
[string]$EWSURL = "https://outlook.office365.com/EWS/Exchange.asmx"

# Load EWS Managed API DLL  and Connect to Exchange
[string]$EwsApiDll = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
Import-Module -Name $EwsApiDll

## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
#[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true} 

#Connect to Exchange
$EWService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1 ) 
If ($EWSURL -eq "") 
	{		
		Write-Host ("Using Autodiscover")
		$EWService.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain)
		$EWService.AutodiscoverUrl($Mailbox, {$True}) 
	} else {
		$EWService.Url = $EWSURL
		$EWService.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain)
		Write-Host ("Using EWS URL")
	}
 

###############################################################################
#Exchange Online V2 Powershell Module
###############################################################################
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$IsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
#EXO V2 Module - REST Based Module
If ((Get-Module ExchangeOnlineManagement -ListAvailable) -eq $null)
{
	Write-Host "ExchangeOnlineManagement V2 Module not found. Try to install..."
	If ($IsAdmin -eq $false)
	{
		Write-Host "WARNING: PS must be running <As Administrator> to install the Module" -ForegroundColor Red
	} else {
		Install-Module ExchangeOnlineManagement -Confirm:$false
	}
} else {
	Write-Host "Loading Module: ExchangeOnlineManagement V2"
	Import-Module ExchangeOnlineManagement
}

###############################################################################
#Connect to Exchange Online
###############################################################################
If (!(Get-PSSession | Where {(($_.ComputerName -like "outlook.office365.com") -AND ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.State -ne "Open"))})) {
	Write-Host "Connect to Exchange Online" -ForegroundColor green
    Connect-ExchangeOnline
	#Connect-ExchangeOnline -CertificateThumbprint "cf1dcd32e78b6ccc9a89be93b31a98a30fe7f760" -AppID "f38d26a7-740e-425f-aef5-2da3f3d595db" -Organization "icewolfch.onmicrosoft.com"
	#$CertPassword = ConvertTo-SecureString "DemoPassword!" -AsPlainText -Force
	#Connect-ExchangeOnline -CertificateFilePath "E:\Scripting\ExOPowershell.pfx" -CertificatePassword $CertPassword -AppID "f38d26a7-740e-425f-aef5-2da3f3d595db" -Organization "icewolfch.onmicrosoft.com"
}Else {
    Write-Host "Connection to Exchange Online already exists" -ForegroundColor yellow
}


###############################################################################
#Get Room Mailboxes
###############################################################################
$Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails RoomMailbox 

#DEBUG
$MailboxCount = $Mailboxes | measure
Write-Host "DEBUG: Mailboxes found: $($MailboxCount.Count)"


$rptcollection = @()
$obj = @{}
$start = Get-Date 
$i = 0
foreach($MBX in $Mailboxes)
{
	$i +=1
	$DisplayName = $MBX.DisplayName
	$PrimarySMTPAddress = $MBX.PrimarySMTPAddress
	
	Write-Host "Working on: $DisplayName"
	Write-Host "Working on: $PrimarySMTPAddress"
	$Duration = (New-TimeSpan -Start ($start) -End (Get-Date)).totalseconds

	$WorkingDays = ($MBX | Get-mailboxCalendarConfiguration -WarningAction SilentlyContinue).WorkDays.ToString() 
	$WorkingHoursStartTime = ($MBX |    Get-mailboxCalendarConfiguration -WarningAction SilentlyContinue).WorkingHoursStartTime 
	$WorkingHoursEndTime = ($MBX | Get-mailboxCalendarConfiguration -WarningAction SilentlyContinue).WorkingHoursEndTime 

	if($WorkingDays -eq "Weekdays"){$WorkingDays = "Monday,Tuesday,Wednesday,Thursday,Friday"}
	if($WorkingDays -eq "AllDays"){$WorkingDays = "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday"}
	if($WorkingDays -eq "WeekEndDays"){$WorkingDays = "Saturday,Sunday"}

	## Optional section for Exchange Impersonation  
	$EWService.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $PrimarySMTPAddress) 

	$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$PrimarySMTPAddress)   
	$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWService,$folderid)

	if($Calendar.TotalCount -gt 0)
	{
		$offset = 0
		$MeetingCount= 0
		$RecurringMeetingCount = 0
		
		#Create Hashtable
		$topOrganizers = @{}
		$topAttendees = @{}
		
		do { 
			$cvCalendarview = new-object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,1000)
			$cvCalendarview.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
			
			$frCalendarResult = $Calendar.FindAppointments($cvCalendarview)
			
			#DEBUG
			$ResultCount = $frCalendarResult | measure
			Write-Host "DEBUG ResultCount: $($ResultCount.Count)"
			
			$inPolicy = New-TimeSpan
			$OutOfPolicy = New-TimeSpan
			$TotalDuration = New-timespan
			$BookableTime = New-TimeSpan
			
			foreach ($apApointment in $frCalendarResult.Items)
			{
						
				$MeetingCount +=1
				$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
				$apApointment.load($psPropset)
			
				If ($apApointment.IsRecurring) 
				{
					$RecurringMeetingCount = $RecurringMeetingCount +1
				}

			
				# Top Organizers
				If ($apApointment.Organizer -and $topOrganizers.ContainsKey($apApointment.Organizer.Address)) 
				{
					$topOrganizers.Set_Item($apApointment.Organizer.Address, $topOrganizers.Get_Item($apApointment.Organizer.Address) + 1)
				} Else {
					$topOrganizers.Add($apApointment.Organizer.Address, 1)
				}
		
				# Top Required Attendees
				ForEach ($attendant in $apApointment.RequiredAttendees) 
				{
					If (!$attendant.Address) {Continue}
					If ($topAttendees.ContainsKey($attendant.Address)) 
					{
						$topAttendees.Set_Item($attendant.Address, $topAttendees.Get_Item($attendant.Address) + 1)
					} Else {
						$topAttendees.Add($attendant.Address, 1)
					}
				}
			

				 if($apApointment.IsAllDayEvent -eq $false)
				 {

					if($apApointment.Duration)
					 {
						if($WorkingDays.split(",") -contains ($apApointment.start).dayofweek)
						{
							$TotalDuration = $TotalDuration.add((new-timespan -End $apApointment.End.tolongTimeString() -start $apApointment.start.tolongTimeString()))

							#Only count to inPolicy if within the workinghours time
							if($apApointment.start.tolongTimeString() -lt $WorkingHoursStartTime)
							{   
								$tStart = $WorkingHoursStartTime.ToString()
							}   
							else
							{
								$tStart = $apApointment.start.ToLongTimeString()
							}

							if($apApointment.End.tolongTimeString() -gt $WorkingHoursEndTime)
							{   
								$tEnd = $WorkingHoursEndTime.ToString()
							}   
							else
							{
								$tEnd = $apApointment.End.ToLongTimeString()

							}

							$Duration = New-TimeSpan -Start $tStart -End $tEnd
							$inPolicy = $inPolicy.add($Duration)
						}
					}
				 }
			}
		
			$offset = $offset + 1000
		} while($Items.MoreAvailable)

		#Calculate to total hours of bookable time between the 2 dates
		for ($d=$Startdate;$d -le $Enddate;$d=$d.AddDays(1))
		{
		  if ($WorkingDays.split(",") -contains $d.DayOfWeek) 
		  {
			$BookableTime += $WorkingHoursEndTime - $WorkingHoursStartTime

		  }
		}

		#Save result....
		$rptobj = "" | Select StartDate,EndDate,RoomEmail,DisplayName,Meetings,RecurringMeetings,inPolicy,Out-Of-Policy,TotalDuration,BookableTime,BookedPercentage,TopOrganizers,TopAttandees
		$rptobj.StartDate = $StartDate
		$rptobj.EndDate = $EndDate
		$rptobj.RoomEmail = $MBX.PrimarySMTPAddress
		$rptobj.DisplayName = $MBX.DisplayName
		$rptobj.Meetings = $MeetingCount
		$rptobj.RecurringMeetings = $RecurringMeetingCount
		$rptobj.inPolicy =  '{0:f2}' -f ($inPolicy.TotalHours)
		$rptobj."Out-Of-Policy" =  '{0:f2}' -f (($TotalDuration - $inPolicy).TotalHours)
		$rptobj.TotalDuration =  '{0:f2}' -f ($TotalDuration.TotalHours)
		$rptobj.BookableTime =  '{0:f2}' -f ($BookableTime.TotalHours)
		$rptobj.BookedPercentage =  '{0:f2}' -f (($inPolicy.TotalHours / $BookableTime.TotalHours) * 100)
		$rptobj.TopOrganizers = [String] ($topOrganizers.GetEnumerator() | Sort -Property Value -Descending | Select -First 10 | % {"$($_.Key) ($($_.Value)),"})
		$rptobj.TopAttandees =  [String] ($topAttendees.GetEnumerator() | Sort -Property Value -Descending | Select -First 10 | % {"$($_.Key) ($($_.Value)),"})
		$rptcollection += $rptobj

	} 
}

$rptcollection

###############################################################################
# Export Results
###############################################################################
$Filename = "MeetingRoomStats_$((Get-Date).ToString('yyyyMMdd')).csv"
Write-Host "Export CSV as $Filename"
$rptcollection | Export-Csv $Filename -Encoding UTF8 -NoTypeInformation -delimiter ";"


