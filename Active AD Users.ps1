#how to check active users

$activeuser = Get-ADUser -Filter { enabled -eq $true} -Properties name,Samaccountname,Whencreated | Select Samaccountname,Whencreated,name
$activeuser.name.count
$activeuser.Samaccountname