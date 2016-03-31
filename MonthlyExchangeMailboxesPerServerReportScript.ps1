add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010 -erroraction silentlyContinue
Set-ADServerSettings -ViewEntireForest $True
###################################################################################################################################################################################################################################### 
# Exchange 2010 and 2007 Month End Report for number of mailboxes per server. 
# Author: Tiens van Zyl
# Email: tiens.vanzyl@gmail.com
# Date 30 March 2016
# Updated by Tiens van Zyl (Date)
# Updates: Version 1
# This script exports and e-mail's an HTML document containing a table showing the number of mailboxes per server
# 1. The script exports an HTML file with the date that the script is run appended i.e. EndofMonthNumberOfMailboxesPerServer 2016-03-30.html
# 2. The HTML document shows a table containing the server name and the number of mailboxes on that server. 
# 3. Change the <H2></H2> element to another heading if you'd like. This is printed above the table in the HTML file.
# 4. Enter your mailbox server/s name in place of "ServerName". Use a wilcard if you have more than one server you need to query i.e. mailbox0* if your servers names are mailbox01, mailbox02 etc.
# 5. Set the path to where you'd like to export the HTML file. Currently set to export to C:\Exchange_AutomatedScripts\MonthlyReports\EndOfMonthMailboxesPerServerReport\EndofMonthNumberOfMailboxesPerServer.html 
######################################################################################################################################################################################################################################

$a = "<style>"
$a = $a + "BODY{background-color:white;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:grey}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:white}"
$a = $a + "</style>"

$file = "C:\Exchange_AutomatedScripts\MonthlyReports\EndOfMonthMailboxesPerServerReport\EndofMonthNumberOfMailboxesPerServer $(get-date -f yyyy-MM-dd).html"

Get-MailboxServer ServerName | Get-Mailbox -ResultSize unlimited | Group-Object -Property:ServerName | Select-Object name,count | ConvertTo-HTML -head $a -body "<H2>Number of Mailboxes per Server (2007 and 2010)</H2>" | Out-File "$file"

$smtpServer = "YourSMTPRelayServer"

$att = new-object Net.Mail.Attachment($file)

$msg = new-object Net.Mail.MailMessage

$smtp = new-object Net.Mail.SmtpClient($smtpServer)

$msg.From = "Sender SMTP Address"

$msg.To.Add("Recipient@domain.com, Recipient2@domain.com")

$msg.Subject = "Your Email Subject Here"

$msg.Body = "What Ever Text you want in the Body of the e-mail sent"

$msg.Attachments.Add($att)

$smtp.Send($msg)

$att.Dispose()





