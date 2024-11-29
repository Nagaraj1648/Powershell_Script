# How to check AD computer list

Get-ADComputer -Filter * -Properties Whencreated | select samaccountname,whencreated