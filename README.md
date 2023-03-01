# ExchangeRoomStatisticsGraph
Version 0.1 - 28.03.2022 Andres Bohren - Initial Version (Not Completed)

Prerequisits
- Exchange Online Powershell V3 (Module ExchangeOnlineManagement)
- Account with Exchange Administrator or Exchange Recipient Admin Role
- Graph Permissions for Calendar.Read

Get-RoomStatistics.ps1 -Startdate "01/01/2022" -EndDate "12/31/2022" [-Mailboxes <ArrayOfEmailAddresses>]

# ExchangeRoomStatistics
Version 1.0 - 21.07.2020 Andres Bohren - Initial Version

Prerequisits
- Exchange Online Powershell V2 (Module ExchangeOnlineManagement)
- EWS Managed API 2.2
- Account with Exchange Administrator Role
- EWS Account with Impersonation Permission

Get-RoomStatistics.ps1 -Startdate "01/01/2020" -EndDate "12/31/2020"