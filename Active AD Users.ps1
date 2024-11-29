#how to check active users

Get-ADUser -Filter { enabled -eq $true} -Properties Samaccountname,Whencreated | Select Samaccountname,Whencreated
