$EmailFrom = "notifications@somedomain.com"
$EmailTo = "nguyenlehoang2911@gmail.com" 
$Subject = "Notification from XYZ" 
$Body = "this is a notification from XYZ Notifications.." 
$SMTPServer = "smtp.gmail.com" 
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587) 
$SMTPClient.EnableSsl = $true 
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential("XXX", "XXXX"); 
$SMTPClient.Send($EmailFrom, $EmailTo, $Subject, $Body)