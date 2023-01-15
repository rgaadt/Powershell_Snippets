#Check for local Certs

﻿# PS3.0
Get-ChildItem -Cert: -Recurse -ExpiringInDays 30

#PS2.0
Get-ChildItem -Path Cert: -Recurse | where { $_.notafter -le (get-date).AddDays(30) -AND $_.notafter -gt (get-date)} | select thumbprint, subject
