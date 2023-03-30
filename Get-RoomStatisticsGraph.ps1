###############################################################################
# Get Exchange Room Statistics with MgGraph
# Version 0.1 - 28.02.2023 Andres Bohren - Initial Version
###############################################################################
# Prerequisits
# - Exchange Online Powershell V3 (Module ExchangeOnlineManagement)
# - Account with Exchange Administrator Role / Exchange Recipient
# - Graph Calendars.Read Permissions for the Room Mailboxes
##############################################################################
<#
.SYNOPSIS
Gather statistics regarding meeting room usage

.DESCRIPTION
This script uses the Exchange Online Management PowerShell Module and Microsoft Graph to connect to one or more Meeting Rooms and gather statistics regarding their usage between to specific dates.
The Output will saved to a CSV File

.PARAMETER StartDate
The Start Date for the Report

.PARAMETER StartDate
The End Date for the Report

.PARAMETER Mailbox
A specifix Mailbox to run the Report against

.PARAMETER AppID
An AzureAD App that has the "Calendars.Read" Permission. Requires also the "CertificateThumbprint" and "TenantId" Parameter.

.PARAMETER CertificateThumbprint
The Certificate used for Authenticate against the AzureAD App specified with the AppID Parameter.

.PARAMETER TenantId
The TenantId of the Tenant where the Azure AD App is registered.
Examples: 
-TenantID <tenant.onmicrosoft.com>
-TenantId <GUID of the Tenant>

.EXAMPLE
 .\Get-RoomStatisticsGraph.ps1 -Startdate "01/01/2023" -EndDate "12/31/2023" [-Mailbox <ArrayOfEmailAddresses>] 

 .\Get-RoomStatisticsGraph.ps1 -Startdate "01/01/2023" -EndDate "12/31/2023" [-Mailbox <ArrayOfEmailAddresses>] [-AppID <AppID>] [-CertificateThumbprint <CertificateThumbprint>] [-TenantId <TenantId>]
#>

param (
	[Parameter(Mandatory=$true)][DateTime]$StartDate,
	[Parameter(Mandatory=$true)][DateTime]$EndDate,
	[Parameter(Mandatory=$false)][String]$Mailbox,
	[Parameter(Mandatory=$false)][String]$AppID,
	[Parameter(Mandatory=$false)][String]$CertificateThumbprint,
	[Parameter(Mandatory=$false)][String]$TenantId
	
)

###############################################################################
# Variables
###############################################################################
#$TenantId = "icewolfch.onmicrosoft.com"
#$AppID = "c1a5903b-cd73-48fe-ac1f-e71bde968412" #DelegatedMail
#$CertificateThumbprint = "07EFF3918F47995EB53B91848F69B5C0E78622FD" #O365Powershell3.cer

###############################################################################
# Connect to Microsoft Graph
###############################################################################
If ($Null -eq $AppID)
{
	#Delegated Authentication
	Connect-MgGraph -Scopes Calendars.Read
} else {
	#App Authentication with Certificate
	Connect-MgGraph -AppId $AppID -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId
}

$ScriptStart = Get-Date 

###############################################################################
# Connect to Exchange Online
###############################################################################
$ConnInfo  = Get-ConnectionInformation
If ($Null -eq $ConnInfo) {
	Write-Host "Connect to Exchange Online" -ForegroundColor green
	Connect-ExchangeOnline ShowBanner:$false
}Else {
	Write-Host "Connection to Exchange Online already exists" -ForegroundColor yellow
}

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

	$WorkingDays = ($MBX | Get-MailboxCalendarConfiguration -WarningAction SilentlyContinue).WorkDays.ToString() 
	$WorkingHoursStartTime = ($MBX | Get-mailboxCalendarConfiguration -WarningAction SilentlyContinue).WorkingHoursStartTime 
	$WorkingHoursEndTime = ($MBX | Get-mailboxCalendarConfiguration -WarningAction SilentlyContinue).WorkingHoursEndTime 

	if($WorkingDays -eq "Weekdays"){$WorkingDaysArray = "Monday,Tuesday,Wednesday,Thursday,Friday"}
	if($WorkingDays -eq "AllDays"){$WorkingDaysArray = "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday"}
	if($WorkingDays -eq "WeekEndDays"){$WorkingDaysArray = "Saturday,Sunday"}

	#Variables for Calendar
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
		
		#Save result
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