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
	[Parameter(Mandatory=$false)][String]$Mailbox
	
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
    Connect-ExchangeOnline ShowBanner:$false
	#Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbprint -AppID $AppID -Organization $TenantId	
}Else {
    Write-Host "Connection to Exchange Online already exists" -ForegroundColor yellow
}


###############################################################################
# Connect to Microsoft Graph
###############################################################################
#Connect-MgGraph -Scopes Calendars.Read
Connect-MgGraph -AppId $AppID -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId

$ScriptStart = Get-Date 

###############################################################################
# Getting Room Mailboxes
###############################################################################
If ($Null -eq $Mailbox)
{
	Write-Host "Getting Room Mailboxes"
	$Mailboxes = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited
} else {
	$Mailboxes = Get-Mailbox -Identity $Mailbox
}

#Loop through the Room Mailboxes
$i = 0
Foreach ($MBX in $Mailboxes)
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

	if($WorkingDays -eq "Weekdays"){$WorkingDaysArray = "Monday,Tuesday,Wednesday,Thursday,Friday"}
	if($WorkingDays -eq "AllDays"){$WorkingDaysArray = "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday"}
	if($WorkingDays -eq "WeekEndDays"){$WorkingDaysArray = "Saturday,Sunday"}

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
	#$StartDate = "2023-01-01T00:00"
	#$EndDate = "2023-12-13T23:59"
	#$PrimarySMTPAddress = "postmaster@icewolf.ch"
	$CalendarItems = Get-MgUserEvent -UserId $PrimarySMTPAddress -Filter "start/dateTime ge '$StartDate' and end/dateTime lt '$EndDate'"
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
			#Subject
			Write-Verbose "Subject: $Appointment.Subject"

			#Increase Meeting Count
			$MeetingCount = $MeetingCount + 1
			
			#Recurring Meeting
			If ($Appointment.Type -eq "seriesMaster")
			{
				$RecurringMeetingCount = $RecurringMeetingCount +1
				Write-Verbose "Recurring Meeting"
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
				Write-Verbose "IsOnlineMeeting"
			}
			
			#All Day Event
			if($Appointment.IsAllDay -eq $true)
			{
				#All Day Event
				$TotalDuration = $TotalDuration.add((New-Timespan -Start $WorkingHoursStartTime -End $WorkingHoursEndTime))
				$AllDayMeetingCount = $AllDayMeetingCount + 1
				Write-Verbose "IsAllDay"
			} else {

				#Not an All Day Event
				[DateTime]$AppointmentStart = $Appointment.Start.DateTime
				[DateTime]$AppointmentEnd = $Appointment.End.DateTime

				#Only Count if Start and End is in WorkingDays
				if($WorkingDaysArray.split(",") -contains $AppointmentStart.dayofweek -and $WorkingDaysArray.split(",") -contains $AppointmentEnd.dayofweek)
				{
					$TotalDuration = $TotalDuration.add((new-timespan -Start $AppointmentStart -End $AppointmentEnd))					
				}
			}
			Write-Verbose "TotalDurationHours: $TotalDuration.Hours"
		}

		#Calculate to total hours of bookable time between the 2 dates
		for ($d=$Startdate;$d -le $Enddate;$d=$d.AddDays(1))
		{
			if ($WorkingDaysArray.split(",") -contains $d.DayOfWeek) 
			{
				$BookableTime += $WorkingHoursEndTime - $WorkingHoursStartTime
			}
		}
		Write-Verbose "BookableTime: $BookableTime.Hours"
		

		#Save result....
		#ReportStartDate;ReportEndDate;RoomEmailaddress;RoomDisplayName;MeetingCount;OnlineMeetingCount;RecurringMeetingCount;AllDayMeetingCount;TotalMinutes;TotalAttandees;AvgEventsPerDay;AvgAttendees;MostActiveMeetingOrganizer
		$rptobj = "" | Select-Object ReportStartDate,ReportEndDate,RoomEmail,DisplayName,WorkingDays,WorkingHoursStartTime,WorkingHoursEndTime,MeetingCount,OnlineMeetingCount,RecurringMeetingCount,AllDayMeetingCount,TotalDuration,BookableTime,BookedPercentage,TopOrganizers,TopAttandees
		$rptobj.ReportStartDate = $StartDate
		$rptobj.ReportEndDate = $EndDate
		$rptobj.RoomEmail = $MBX.PrimarySMTPAddress
		$rptobj.DisplayName = $MBX.DisplayName
		$rptobj.WorkingDays = $WorkingDays
		$rptobj.WorkingHoursStartTime = $WorkingHoursStartTime
		$rptobj.WorkingHoursEndTime = $WorkingHoursEndTime
		$rptobj.MeetingCount = $MeetingCount
		$rptobj.OnlineMeetingCount = $OnlineMeetingCount
		$rptobj.RecurringMeetingCount = $RecurringMeetingCount
		$rptobj.AllDayMeetingCount = $AllDayMeetingCount			
		#$rptobj.inPolicy =  '{0:f2}' -f ($inPolicy.TotalHours)
		#$rptobj."Out-Of-Policy" =  '{0:f2}' -f (($TotalDuration - $inPolicy).TotalHours)
		$rptobj.TotalDuration =  '{0:f2}' -f ($TotalDuration.TotalHours)
		$rptobj.BookableTime =  '{0:f2}' -f ($BookableTime.TotalHours)
		$rptobj.BookedPercentage =  '{0:f2}' -f (($TotalDuration.TotalHours / $BookableTime.TotalHours) * 100)
		$rptobj.TopOrganizers = [String] ($topOrganizers.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10 | ForEach-Object {"$($_.Key) ($($_.Value)),"})
		$rptobj.TopAttandees =  [String] ($topAttendees.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10 | ForEach-Object {"$($_.Key) ($($_.Value)),"})
		$rptcollection += $rptobj
		}
	}
	

	$rptcollection

#Script Run Time
$ScriptEnd = Get-Date 
$ScriptDuration = New-TimeSpan $ScriptStart -End $ScriptEnd
Write-Host "ScriptDuration: $ScriptDuration.TotalMinutes"

###############################################################################
# Export Results
###############################################################################
$Filename = "MeetingRoomStats_$((Get-Date).ToString('yyyyMMdd')).csv"
Write-Host "Export CSV as $Filename"
$rptcollection | Export-Csv $Filename -Encoding UTF8 -NoTypeInformation -delimiter ";"