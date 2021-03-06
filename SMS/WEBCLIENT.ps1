function ConvertToBody($body){
    $temp = ($body.GetEnumerator() | % { "$($_.Key)=$($_.Value)" }) -join '&'
    $result = $temp.Replace("+","%2B").Replace(" ","%20")
    Return $result
}
$AccountId = 'xxxxxxxxxxxxxxxxxxxxxxxxx'
$AuthenToken = 'xxxxxxxxxxxxxxxxxxxxxxx'
$url = 'https://api.twilio.com/2010-04-01/Accounts/{0}/Messages.json' -f $AccountId
$params = @{
    To = "xxxxxxxxxxxx" ;
    From = "xxxxxxxxxxx";
    Body = "xxxxxxxxxxxxxxxx";
}

#2. setup basic authenication
$credentials = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($AccountId + ":" + $AuthenToken));
write-host $credentials
$authen = ("Basic {0}" -f $credentials)
write-host $authen


#format body data
$para = ConvertToBody($params)
Write-host $para
$postbytes = [System.Text.Encoding]::ASCII.GetBytes($para);

$webclient = New-Object System.Net.WebClient
$webclient.Headers.Add("Authorization", $authen);
$webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");


$resp = $webClient.UploadData($url,'POST',$postbytes )