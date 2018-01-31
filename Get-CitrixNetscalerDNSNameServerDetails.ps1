Import-Module NetScalerConfiguration

###SEPARATE USERNAME AND PASSWORD
 Function SecureStringToString($value)
{
    [System.IntPtr] $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($value);
    try
    {
        [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr);
    }
    finally
    {
        [System.Runtime.InteropServices.Marshal]::FreeBSTR($bstr);
    }
}

$nsCreds = Get-Credential -credential ${CREDENTIAL}
[string] $nsUsername = $nsCreds.Username
[string] $nsPassword = SecureStringToString $nsCreds.Password

###SET UP BASIC CONNECTION 
Set-NSMgmtProtocol -Protocol http
$nsAddress = '${IP}'
$nsSession = Connect-NSAppliance -NSAddress $nsAddress -NSUserName $nsUsername -NSPassword $nsPassword

$nsResponse = Invoke-NSNitroRestApi -NSSession $nsSession `
                                    -OperationMethod GET `
                                    -ResourceType 'dnsnameserver' 

$swOutputMsg = @()
$swOutputStat = 0

### NOTIFY IF NO NAME SERVERS ARE CONFIGURED
if (!$($nsResponse.dnsnameserver)) {
    Write-Host "Message.NameServers: No Name Server entries were found!"
    Write-Host "Statistic.NameServers: 1"
    exit 0
}

### VERIFY ALL NAME SERVER ENTRIES ARE "UP"
foreach ($nameserver in $nsResponse.dnsnameserver) {
    $nameserverAddress = $nameserver.ip
    $nameserverState = $nameserver.state
    $nameserverStatus = $nameserver.nameserverstate

    if ($nameserverState -ne 'ENABLED') {
        $swOutputMsg += "Name Server $nameserverAddress is set to $nameserverState "
    }
    elseif ($nameserver.nameserverstate -eq 'UP') {
        $swOutputMsg += "$nameserverAddress - $nameserverStatus"
    }
    else {
        $swOutputStat++
        $swOutputMsg += "Name Server $nameserverAddress is $nameserverStatus "
    }
}

$message = $swOutputMsg -join '<br/>'

### CLOSE THE NITRO API SESSION
Disconnect-NSAppliance -NSSession $nsSession

### OUTPUT TO SOLARWINDS
Write-Host "Message.NameServers: $message"
Write-Host "Statistic.NameServers: $swOutputStat"
