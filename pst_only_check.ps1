#Prüft die PST Archiv Ordner und versendet eine Email, wenn andere Dateien abliegen.

[string[]] $WRONG_FILES = @()
$WRONG_FILES = Get-ChildItem -Path ('\\[Server_FQDN]\Ordnerpfad\') -Recurse | where {! $_.PSIsContainer} | where {$_.extension -ne ".pst"} | select FullName


ForEach ($Eintrag in $WRONG_FILES)
{
$messagelist = $messagelist + $Eintrag + "`r`n" + "`r`n"
}

$mailbody=@"
Folgende Dateien wurden beim Scan erfasst:


$messagelist

Bitte kontaktieren Sie die Benutzer, damit diese die Daten entfernen. 
Das Laufwerk ist nur zum Ablegen von Mail Archiven gedacht!
"@

If ($WRONG_FILES){
 Send-MailMessage -to "[AdminMail.example.com]" -From "Petze@example.com" -Subject "Falsche Dateien auf PST Archiv Laufwerk gefunden!" -Body "$mailbody" -smtpServer "[IP_of_SMTP_Server]"
  	}
 


