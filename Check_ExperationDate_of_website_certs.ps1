clear
$daysUntilExpire = 80
$timeoutMs = 10000
$websites = @(
"https://gaadt.net",
"https://another_site_to_monitor.com/",
"https://yet_another_site_to_monitor.com/")
# certificate validation true / false
[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
foreach ($site in $websites)
{
Write-Host Check $site -f Green
# create webrequest
$req = [Net.HttpWebRequest]::Create($site)
$req.Timeout = $timeoutMs
try {$req.GetResponse() |Out-Null} catch {Write-Host URL check error $site`: $_ -f Yellow}
$expDate = $req.ServicePoint.Certificate.GetExpirationDateString()
$certExpDate = [datetime]::ParseExact($expDate, “dd/MM/yyyy HH:mm:ss”, $null)
[int]$certExpiresIn = ($certExpDate - $(get-date)).Days
$certName = $req.ServicePoint.Certificate.GetName()
$certThumbprint = $req.ServicePoint.Certificate.GetCertHashString()
$certEffectiveDate = $req.ServicePoint.Certificate.GetEffectiveDateString()
$certIssuer = $req.ServicePoint.Certificate.GetIssuerName()
if ($certExpiresIn -gt $daysUntilExpire)
    {Write-Host The $site certificate expires in $certExpiresIn days [$certExpDate] -f Green}
else
    {
    $message= "The $site certificate expires in $certExpiresIn days"
    $messagetitle= "Renew certificate"
    Write-Host $message [$certExpDate]. Details:`n`nCert name: $certName`Cert thumbprint: $certThumbprint`nCert effective date: $certEffectiveDate`nCert issuer: $certIssuer -f Yellow
    }
write-host "________________" `n

