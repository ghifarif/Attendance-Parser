$ifile = "D:\$VENDOR\Attendant log"
$ofile = "D:\$VENDOR\report\weekly-report.csv"
$From = "attendance-report@localhost"
$To = "ghifari@company.com"
$Cc = "receiver1@company.com", "receiver2@company.com"
$Subject = "[Report] Weekly Attendance"
$Body = "Dear Customer, Please find attendance report on attachment"
$SMTPServer = "$SMTPIP"
$SMTPPort = "$SMTPPORT"

New-Item -path $ofile -ItemType File
Get-ChildItem $ifile -Recurse | Where-Object {$_.CreationTime -lt (Get-Date).AddDays(-2) -and $_.CreationTime -gt (Get-Date).AddDays(-7) } | ForEach-Object {
    $_ | Get-Content | Out-File $ofile -Append
    Start-Sleep -s 5
}
Send-MailMessage -From $From -to $To -Cc $Cc -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -Attachments $ofile –DeliveryNotificationOption OnSuccess
Start-Sleep -s 5
Remove-Item -path $ofile