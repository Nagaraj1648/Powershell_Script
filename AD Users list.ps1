# How to check AD Users list

Get-ADUser -Filter * -Properties Whencreated | select samaccountname,whencreated