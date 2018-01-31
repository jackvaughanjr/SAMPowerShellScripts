Function Convert-FromUnixdate ($UnixDate)
{
   [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').`
   AddSeconds($UnixDate))
}
 
Function Get-LocalTime($UTCTime)
{
$strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
$LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
Return $LocalTime
}
 
function Convert-DateString ([String]$Date, [String[]]$Format)
{
   $result = New-Object DateTime
  
   $convertible = [DateTime]::TryParseExact(
      $Date,
      $Format,
      [System.Globalization.CultureInfo]::InvariantCulture,
      [System.Globalization.DateTimeStyles]::None,
      [ref]$result)
  
   if ($convertible) { $result }
}
 
###### CONNECTION INFORMATION FROM OFFICE365 APPLICATION REGISTERED AT https://apps.dev.microsoft.com/#/appList
Add-Type -AssemblyName System.Web
$client_id = "<<CLIENTID>>"
$client_secret = "<<CLIENTSECRET>>"
$redirectUrl = "<<REDIRECTURL>>"
 
###### GET NEW ACCESS TOKEN USING REFRESH TOKEN
$loginUrl = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=" +
            [System.Web.HttpUtility]::UrlEncode($redirectUrl) +
            "&client_id=$client_id" +
            "&prompt=login"
 
###### REFRESH TOKEN FILE PATH
$refreshTokenPath = "<<PATHTOTOKEN>>"
 
###### VERIFY REFRESH TOKEN EXISTS AND GET NEW ACCESS TOKEN
if (Test-Path -Path $refreshTokenPath) {
    $refresh_token = Get-Content -Path $refreshTokenPath
 
    $refreshPostRequest =
        "grant_type=refresh_token" + "&" +
        "redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($redirectUrl) + "&" +
        "client_id=$client_id" + "&" +
        "client_secret=" + [System.Web.HttpUtility]::UrlEncode("$client_secret") + "&" +
        "refresh_token=" + [System.Web.HttpUtility]::UrlEncode("$refresh_token") + "&" +
        "resource=" + [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/")
 
    $refreshAuthorization =
        Invoke-RestMethod   -Method Post `
                            -ContentType application/x-www-form-urlencoded `
                            -Uri 'https://login.microsoftonline.com/common/oauth2/token' `
                            -Body $refreshPostRequest
     
    ### SAVE NEW REFRESH TOKEN TO FILE FOR NEXT RUN
    Out-File -InputObject $($refreshAuthorization.refresh_token) -FilePath $refreshTokenPath -Force
}
 
###### IF NO REFRESH TOKEN, ACCESS TOKEN MUST BE GENERATED MANUALLY
else {
    Send-MailMessage -From '<<FROM>>' `
                    -To '<<TO>>' `
                    -Cc '<<CC>>' `
                    -Subject '<<SUBJECT>>' `
                    -Body 'The refresh token has expired.  Log into the server and run the new token script.' `
                    -SmtpServer '<<SMTPSERVER>>'
}
 
$access_token = $refreshAuthorization.access_token
$refresh_token = $refreshAuthorization.refresh_token
 
 
 
###### COLLECT ALL MESSAGES WITH SUBJECT "Time to review the plan"
$mail =
    Invoke-RestMethod   -Headers @{Authorization =("Bearer "+ $access_token)} `
                        -Uri 'https://outlook.office.com/api/v2.0/Users/<<EMAILADDRESS>>/messages?$select=Sender,Subject,ReceivedDateTime&$filter=Subject eq ''Time to review the plan''&$orderby=Subject,ReceivedDateTime desc' `
                        -Method Get
 
$messages = $mail.value
 
 
###### WINDOW OF TIME TO CHECK EMAILS AGAINST (CURRENTLY THE PREVIOUS DAY, 2000 - 0000)
$timeAllowedStart = (Get-Date -Hour 20 -Minute 0 -Second 0).AddDays(-1)
$timeAllowedEnd   = (Get-Date -Hour 00 -Minute 0 -Second 0)
 
###### CHECK MOST RECENT MESSAGE AGAINST THE ALLOWABLE TIME WINDOW - EMAIL IF FOUND, ALERT IF NOT FOUND
$message = $messages[0]
$timeReceived = Get-Date -Date $($message.ReceivedDateTime)
if (($timeReceived -gt $timeAllowedStart) -and ($timeReceived -lt $timeAllowedEnd)) {
    $msgReceived = '<span class="up">Received - No Action Required</span>'
         
    Send-MailMessage -From '<<FROM>>' `
                        -To '<<TO>>' `
                        -Cc '<<CC>>' `
                        -Subject '<<SUBJECT>>' `
                        -Body "Email with received date of $timeReceived found. No alert fired." `
                        -SmtpServer '<<SMTPSERVER>>'
}
else {
    $msgReceived = '<span class="critical">Not Received - Escalate!</span>'
}
 
###### HTML OUTPUT FOR SOLARWINDS APPLICATION CHECK
$currentTime = Get-Date
 
$htmlHead = @"
    <title><<HTML TITLE>></title>
    <style>
    body {font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;}
    .critical {color: #ff0000; font-weight: bold}
    .up {color: #087a00; font-weight: bold}
    </style>
"@
 
$htmlBody = @"
    <h1>
    Jack Cooper Email Verification
    </h1>
    <p>
    Email status: $msgReceived<br />
    Most recent email date: $timeReceived<br />
    Last run time: $currentTime
    </p>
    <<EMAIL BODY HERE>>
"@
 
ConvertTo-Html  -Head $htmlHead `
                -Body $htmlBody |
                Out-File <<PATH FOR HTML>> -Force
