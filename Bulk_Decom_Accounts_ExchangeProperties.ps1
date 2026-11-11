Connect-ExchangeOnline
Remove-CalendarEvents -Identity "alisonhowe@awmcharis.com" -CancelOrganizedMeetings -QueryWindowInDays 365
Remove-CalendarEvents -Identity "amandaparker@awmcharis.com" -CancelOrganizedMeetings -QueryWindowInDays 365
Remove-CalendarEvents -Identity "dakotakott@awmcharis.com" -CancelOrganizedMeetings -QueryWindowInDays 365
Remove-CalendarEvents -Identity "hosannaoettle@awmcharis.com" -CancelOrganizedMeetings -QueryWindowInDays 365
Remove-CalendarEvents -Identity "joshuawilliams@awmcharis.com" -CancelOrganizedMeetings -QueryWindowInDays 365
Remove-CalendarEvents -Identity "lukethomas@awmcharis.com" -CancelOrganizedMeetings -QueryWindowInDays 365
Remove-CalendarEvents -Identity "ravenjewell@awmcharis.com" -CancelOrganizedMeetings -QueryWindowInDays 365
Remove-CalendarEvents -Identity "annadueck@awmcharis.com" -CancelOrganizedMeetings -QueryWindowInDays 365
Remove-CalendarEvents -Identity "carolineolson@awmcharis.com" -CancelOrganizedMeetings -QueryWindowInDays 365
Remove-CalendarEvents -Identity "charlesbishop@awmcharis.com" -CancelOrganizedMeetings -QueryWindowInDays 365
Get-InboxRule -Mailbox alisonhowe@awmcharis.com| Remove-InboxRule -Confirm:$false
Get-InboxRule -Mailbox amandaparker@awmcharis.com| Remove-InboxRule -Confirm:$false
Get-InboxRule -Mailbox annadueck@awmcharis.com| Remove-InboxRule -Confirm:$false
Get-InboxRule -Mailbox carolineolson@awmcharis.com| Remove-InboxRule -Confirm:$false
Get-InboxRule -Mailbox charlesbishop@awmcharis.com| Remove-InboxRule -Confirm:$false
Get-InboxRule -Mailbox dakotakott@awmcharis.com| Remove-InboxRule -Confirm:$false
Get-InboxRule -Mailbox hosannaoettle@awmcharis.com| Remove-InboxRule -Confirm:$false
Get-InboxRule -Mailbox joshuawilliams@awmcharis.com| Remove-InboxRule -Confirm:$false
Get-InboxRule -Mailbox lukethomas@awmcharis.com| Remove-InboxRule -Confirm:$false
Get-InboxRule -Mailbox ravenjewell@awmcharis.com| Remove-InboxRule -Confirm:$false
