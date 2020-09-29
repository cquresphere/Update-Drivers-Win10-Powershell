# Update-Drivers-Win10-Powershell
Powershell Script to Update Drivers in Windows 10

Base for this project was script from StackExchange with Update. The problem was it works only on Windows 7. 
I thought it would be easy to adjust but it turns out it requires a lot of testing. 

This script: 

1. starts with checking if terminal is runing with administrator priviledge.
If it is not it propmts notification about requirement. 

2. Then it installs required PowerShell modules.
