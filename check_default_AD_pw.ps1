#Check UserAccounts for Bad Passwords...

$MyDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$OutFileName = "dict_att_Userlist.log"


clear

Function Test-ADAuthentication {
    param($username,$password)
    (new-object directoryservices.directoryentry "",$username,$password).psbase.name -ne $null
}




$var_PW	= Read-Host "Eingabe zu pr�fendes Kennwort"


$var_Logfileausgabe = "Liste der gef�hrdeten Benutzern mit Kennwort " +$var_PW +" :"
$var_Logfileausgabe>$MyDir\$OutFileName


Get-ADUser -Filter "*" | ForEach-Object {
											$var_user = $_.SamAccountName
											$var_Result = Test-ADAuthentication $var_User $var_PW

											if ($var_Result -eq "true") 
												{
													Write-Host "Der Benutzername - " $var_user " - ist gef�hrdet! " -foregroundcolor red
													$var_user >> $MyDir\$OutFileName
												}

											Else 
												{
													Write-Host "Der Benutzername - " $var_user " - besitzt nicht das gepr�fte Kennwort (" $var_PW "). " -foregroundcolor green
												}
										}
