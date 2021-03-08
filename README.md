# Update-Drivers-Win10-Powershell
Powershell Script to Update Drivers in Windows 10

Base for this project was script from StackExchange. The problem was it works only on Windows 7. 
I thought it would be easy to adjust but it turns out it requires a lot of testing.

So after testing and experimenting I get this script. 

# This script: 

1. Starts with checking if terminal is runing with administrator priviledge.
If it is not it propmts notification about requirement. 

2. Install and Import Windows Powershell Update Module

3. Add ServiceID for Windows Update

4. Start Service wuauserv 

5. Search and list all missing Drivers

6. Download the Drivers from Microsoft

7. Check if the Drivers are all downloaded and trigger the Installation

8. Pause and give the service time to update

9. Stop Windows Update Service

# Roadmap / Plans

1. Next stage is to add logging feature to be able to see all details about preupdate and post update drivers. [✔]
2. Another feature I would like to add is measurements to count time spent on downloading and installing. [✔]
3. I think also it would be nice to have good error hadling [partialy done]

# References:
[1] https://superuser.com/questions/1243011/how-to-automatically-update-all-devices-in-device-manager

[2] https://rzander.azurewebsites.net/script-to-install-or-update-drivers-directly-from-microsoft-catalog/

[3] https://www.tenforums.com/tutorials/76207-update-upgrade-windows-10-using-powershell.html
