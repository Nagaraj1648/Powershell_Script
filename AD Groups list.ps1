# How to check AD group list

Get-ADGroup -Filter * -Properties Whencreated | select samaccountname,whencreated