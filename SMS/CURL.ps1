$Curl = "curl.exe"
$AccountId = 'xxxxxxxxxxxxxxxx'
$AuthenToken = 'xxxxxxxxxxxx'
$BaseLink = 'https://api.twilio.com/2010-04-01/Accounts/'
$From = 'xxxxxxxxx'
$To = 'xxxxxxxxxxxx'
$Body = 'xxxxxxxxxxxxxxxxx'
$Proxy = 'xxxxxxxxxxxxxxxx'

#GET
#try{
#    $Command = 'C:\curl\src\curl.exe -x {0} -G {1} \ -u {2}:{3}' -f $Proxy, $BaseLink, $AccountId, $AuthenToken
#    Write-host $Command
#    Invoke-Expression -Command $Command
#}
#catch{
#    $ErrorMessage = $_.Exception.Message
#    Write-host $ErrorMessage
#}

#POST
try{
    $url = ($BaseLink + $AccountId + "/Messages.json")
    $Command = '{0} -x {1} -X POST {2} \ --data-urlencode To="{3}" \ --data-urlencode From="{4}"  \ --data-urlencode Body="{5}"  \ -u {6}:{7}' -f 
                $Curl,
                $Proxy, 
                $url, 
                $To,
                $From,
                $Body,
                $AccountId, 
                $AuthenToken
    Write-host $Command
    Invoke-Expression -Command $Command
}
catch{
    $ErrorMessage = $_.Exception.Message
    Write-host $ErrorMessage
}
#Read-host "Complete"