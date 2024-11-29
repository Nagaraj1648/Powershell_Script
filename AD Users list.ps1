# How to check AD Users list

$ADUsers = Get-ADUser -Filter * -Properties Whencreated | select samaccountname,whencreated
($ADUsers).count
$ADUsers | select samaccountname