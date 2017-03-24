Param([string]$Computername = $env:COMPUTERNAME)
 
$head = @"
<Title>Stale Computers</Title>
<style>
Body {
font-family: "Tahoma", "Arial", "Helvetica", sans-serif;
background-color:#FFFFFF;
}

table {
border-collapse:collapse;
width:60%
}

td {
font-size:12pt;
border:1px #00008c solid;
padding:5px 5px 5px 5px;
}


th {
font-size:24pt;
text-align:left;
padding-top:5px;
padding-bottom:4px;
background-color:#00008c;
color:#ffffff;
}

name tr{
color:#000000;
background-color:#00008c;
}
</style>
"@
 
#convert output to html as a string
import-module activedirectory 
$domain = "DOMAIN.LOCAL" 
$DaysInactive = 90 
$time = (Get-Date).Adddays(-($DaysInactive))
$computers = Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -Properties LastLogonTimeStamp 
 
# Output hostname and lastLogonTimestamp into CSV
$html = $computers| select-object Name,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}}  |  # export-csv OLD_Computer.csv -notypeinformation
ConvertTo-Html -Head $head -precontent "<h2>Stale Computers</h2>" -PostContent "<h6> report run $(Get-Date)</h6>" |
Out-String
 
#send as mail body
$paramHash = @{
 To = "TO@DOMAIN.COM"
 from = "FROM@DOMAIN.COM"
 BodyAsHtml = $True
 Body = $html
 SmtpServer = "ironport.mybofi.local" 
 #Port = 587
 Subject = "Inactive Computer Accounts (90 days)"
}
 
Send-MailMessage @paramHash
