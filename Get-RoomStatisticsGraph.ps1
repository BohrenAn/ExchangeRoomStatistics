###############################################################################
# Get Exchange Room Statistics with MgGraph
# Version 0.1 - 28.02.2023 Andres Bohren - Initial Version
###############################################################################
# Prerequisits
# - Exchange Online Powershell V3 (Module ExchangeOnlineManagement)
# - Account with Exchange Administrator Role / Exchange Recipient
# - Graph.Calendar Permissions for the Room Mailboxes
##############################################################################
<#
.SYNOPSIS
Gather statistics regarding meeting room usage

.DESCRIPTION
This script uses the Exchange Online Management PowerShell Module and Microsoft Graph to connect to one or more Meeting Rooms and gather statistics regarding their usage between to specific dates.
The Output will saved to a CSV File
 
.EXAMPLE
 .\Get-RoomStatisticsGraph.ps1 -Startdate "01/01/2020" -EndDate "12/31/2020" [-Mailboxes <ArrayOfEmailAddresses>] 
#>

param (
    [Parameter(Mandatory=$true)][DateTime]$StartDate,
    [Parameter(Mandatory=$true)][DateTime]$EndDate,
	[Parameter(Mandatory=$false)][Array]$Mailboxes
	
)

###############################################################################
# Variables
###############################################################################
$TenantId = "icewolfch.onmicrosoft.com"
$Scope = "https://graph.microsoft.com/.default" 
$AppID = "c1a5903b-cd73-48fe-ac1f-e71bde968412" #DelegatedMail
$CertificateThumbprint = "07EFF3918F47995EB53B91848F69B5C0E78622FD" #O365Powershell3.cer

###############################################################################
# Connect to Exchange Online
###############################################################################
$ConnInfo  = Get-ConnectionInformation
If ($Null -eq $ConnInfo) {
	Write-Host "Connect to Exchange Online" -ForegroundColor green
    Connect-ExchangeOnline
	#Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbprint -AppID $AppID -Organization $TenantId	
}Else {
    Write-Host "Connection to Exchange Online already exists" -ForegroundColor yellow
}


###############################################################################
# Connect to Microsoft Graph
###############################################################################
Connect-MgGraph -Scopes Calendars.Read
#Connect-MgGraph -AppId $AppID -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId

$ScriptStart = Get-Date 

###############################################################################
# Getting Room Mailboxes
###############################################################################
Write-Host "Getting Room Mailboxes"
$Rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited

#Loop through the Room Mailboxes
$i = 0
Foreach ($MBX in $Rooms)
{
	#Do for each Room
	$i = $i + 1
	$DisplayName = $MBX.DisplayName
	$PrimarySMTPAddress = $MBX.PrimarySMTPAddress

	Write-Host "Working on: $DisplayName [$i]"
	Write-Host "Working on: $PrimarySMTPAddress [$i]"
	#$Duration = (New-TimeSpan -Start ($start) -End (Get-Date)).totalseconds

	$WorkingDays = ($MBX | Get-MailboxCalendarConfiguration -WarningAction SilentlyContinue).WorkDays.ToString() 
	$WorkingHoursStartTime = ($MBX | Get-mailboxCalendarConfiguration -WarningAction SilentlyContinue).WorkingHoursStartTime 
	$WorkingHoursEndTime = ($MBX | Get-mailboxCalendarConfiguration -WarningAction SilentlyContinue).WorkingHoursEndTime 

	if($WorkingDays -eq "Weekdays"){$WorkingDays = "Monday,Tuesday,Wednesday,Thursday,Friday"}
	if($WorkingDays -eq "AllDays"){$WorkingDays = "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday"}
	if($WorkingDays -eq "WeekEndDays"){$WorkingDays = "Saturday,Sunday"}

	#Variables for Calendar
	#$offset = 0
	$MeetingCount = 0
	$OnlineMeetingCount = 0
	$RecurringMeetingCount = 0
	$AllDayMeetingCount = 0
	$topOrganizers = @{}
	$topAttendees = @{}
	$rptcollection = @()

	# Example for One Mailbox
	$StartDate = "2023-01-01T00:00"
	$EndDate = "2023-12-13T23:59"
	$PrimarySMTPAddress = "postmaster@icewolf.ch"
	$CalendarItems = Get-MgUserEvent -UserId $PrimarySMTPAddress -Filter "start/dateTime ge '$startdate' and end/dateTime lt '$Enddate'"
	#$CalendarItems = Get-MgUserEvent -UserId $PrimarySMTPAddress -Filter "start/dateTime ge '2023-01-01T00:00' and end/dateTime lt '2023-12-31T23:59'"
	#https://graph.microsoft.com/v1.0/users/a.bohren@icewolf.ch/calendar/events?start/dateTime ge '2023-01-01T00:00' and end/dateTime lt '2023-12-31T23:59'

	If ($Null -ne $CalendarItems)
	{
		#Calendar Items found		
		Write-Verbose "CalendarItems found: $($CalendarItems.Count)"

		#$InPolicy = New-TimeSpan
		#$OutOfPolicy = New-TimeSpan
		$TotalDuration = New-timespan
		$BookableTime = New-TimeSpan

		#Loop through the Calendar Items
		foreach ($Appointment in $CalendarItems)
		{
			#Increase Meeting Count
			$MeetingCount = $MeetingCount + 1
			
			#Recurring Meeting
			If ($Appointment.Type -eq "seriesMaster")
			{
				$RecurringMeetingCount = $RecurringMeetingCount +1
			}

			# Top Organizers
			If ($Appointment.Organizer.EmailAddress.Address -and $topOrganizers.ContainsKey($Appointment.Organizer.EmailAddress.Address)) 
			{
				$topOrganizers.Set_Item($Appointment.Organizer.EmailAddress.Address, $topOrganizers.Get_Item($Appointment.Organizer.EmailAddress.Address) + 1)
			} Else {
				$topOrganizers.Add($Appointment.Organizer.EmailAddress.Address, 1)
			}
	
			# Top Required Attendees
			ForEach ($Attendees in $Appointment.Attendees.EmailAddress) 
			{
				Foreach ($Attendee in $Attendees)
				{					
					If ($topAttendees.ContainsKey($Attendee.Address)) 
					{
						$topAttendees.Set_Item($Attendee.Address, $topAttendees.Get_Item($attendant.Address) + 1)
					} Else {
						$topAttendees.Add($Attendee.Address, 1)
					}
				}
			}

			#OnlineMeeting
			If ($Appointment.IsOnlineMeeting -eq $true)
			{
				$OnlineMeetingCount = $OnlineMeetingCount + 1
			}
			
			#All Day Event
			if($Appointment.IsAllDay -eq $true)
			{
				#All Day Event
				$Duration = New-TimeSpan -Start $WorkingHoursStartTime -End $WorkingHoursEndTime
				$AllDayMeetingCount = $AllDayMeetingCount + 1
			} else {

				#Not an All Day Event
				[DateTime]$AppointmentStart = $Appointment.Start.DateTime
				[DateTime]$AppointmentEnd = $Appointment.End.DateTime

				if($Appointment.Duration)
				{
					if($WorkingDays.split(",") -contains ($apApointment.start).dayofweek)
					{
						$TotalDuration = $TotalDuration.add((new-timespan -End $apApointment.End.tolongTimeString() -start $apApointment.start.tolongTimeString()))

						#Only count to inPolicy if within the workinghours time
						if($apApointment.start.tolongTimeString() -lt $WorkingHoursStartTime)
						{   
							$tStart = $WorkingHoursStartTime.ToString()
						} else {
						   $tStart = $apApointment.start.ToLongTimeString()
						}

						if($apApointment.End.tolongTimeString() -gt $WorkingHoursEndTime)
						{   
							$tEnd = $WorkingHoursEndTime.ToString()
						} else {
							$tEnd = $apApointment.End.ToLongTimeString()
						}

					   $Duration = New-TimeSpan -Start $tStart -End $tEnd
					   $inPolicy = $inPolicy.add($Duration)
				   }
			   }
			} 

			#Calculate to total hours of bookable time between the 2 dates
			for ($d=$Startdate;$d -le $Enddate;$d=$d.AddDays(1))
			{
				if ($WorkingDays.split(",") -contains $d.DayOfWeek) 
				{
					$BookableTime += $WorkingHoursEndTime - $WorkingHoursStartTime
				}
			}

			#Save result....
			#ReportStartDate;ReportEndDate;RoomEmailaddress;RoomDisplayName;MeetingCount;OnlineMeetingCount;RecurringMeetingCount;AllDayMeetingCount;
			#TotalMinutes;TotalAttandees;AvgEventsPerDay;AvgAttendees;MostActiveMeetingOrganizer
			$rptobj = "" | Select-Object ReportStartDate,ReportEndDate,RoomEmail,DisplayName,WorkingHoursStartTime,WorkingHoursEndTime,MeetingCount,OnlineMeetingCount,RecurringMeetingCount,AllDayMeetingCount,TotalDuration,BookableTime,BookedPercentage,TopOrganizers,TopAttandees
			$rptobj.ReportStartDate = $StartDate
			$rptobj.ReportEndDate = $EndDate
			$rptobj.RoomEmail = $MBX.PrimarySMTPAddress
			$rptobj.DisplayName = $MBX.DisplayName
			#$rptobj.WorkingDays = $WorkingDays
			$rptobj.WorkingHoursStartTime = $WorkingHoursStartTime
			$rptobj.WorkingHoursEndTime = $WorkingHoursEndTime
			$rptobj.MeetingCount = $MeetingCount
			$rptobj.OnlineMeetingCount = $OnlineMeetingCount
			$rptobj.RecurringMeetingCount = $RecurringMeetingCount
			$rptobj.AllDayMeetingCount = $AllDayMeetingCount			
			$rptobj.inPolicy =  '{0:f2}' -f ($inPolicy.TotalHours)
			$rptobj."Out-Of-Policy" =  '{0:f2}' -f (($TotalDuration - $inPolicy).TotalHours)
			$rptobj.TotalDuration =  '{0:f2}' -f ($TotalDuration.TotalHours)
			$rptobj.BookableTime =  '{0:f2}' -f ($BookableTime.TotalHours)
			$rptobj.BookedPercentage =  '{0:f2}' -f (($inPolicy.TotalHours / $BookableTime.TotalHours) * 100)
			$rptobj.TopOrganizers = [String] ($topOrganizers.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10 | ForEach-Object {"$($_.Key) ($($_.Value)),"})
			$rptobj.TopAttandees =  [String] ($topAttendees.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10 | ForEach-Object {"$($_.Key) ($($_.Value)),"})
			$rptcollection += $rptobj
			}
		}
	}

	$rptcollection



$ScriptEnd = Get-Date 
$ScriptDuration = New-TimeSpan $ScriptStart -End $ScriptEnd


###############################################################################
# Export Results
###############################################################################
$Filename = "MeetingRoomStats_$((Get-Date).ToString('yyyyMMdd')).csv"
Write-Host "Export CSV as $Filename"
$rptcollection | Export-Csv $Filename -Encoding UTF8 -NoTypeInformation -delimiter ";"