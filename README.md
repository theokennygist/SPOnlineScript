### Author: Kenny.soetjipto
### SharePoint Developer
### RXP Services
* kenny.soetjipto@rxpservices.com 

### Description
This repository contains scripts to automate SharePoint online administration

### Quick summary ###
It is used to automate SP online administration and tasks are made into functions based at each file. 

Therefore, each function can be invoked using for or foreach loop
* Version: 0.1

### Setup ###

* Summary of set up
It needs to have Powershell online
You must be assigned as "global administrator" role on the SharePoint Online site on which you are running the Windows PowerShell cmdlet.
Add ISAPI folder which contain DLLs needed
Change include file at the top at the right folder

* Configuration
- Put ISAPI folders in the C:\program files\common files\Microsoft Shared\Web Server Extension\15\
- If DLL files are already there, then you do not need to put that ISAPI Folder in there. I provide it in case the user does not have it
- Change the $gUrl with your own SP Online URL site collection
- Change the $adminUrl with your SP Online admin page url
- Change $username with your own username
- Create PassFile.ps1 and put this code: $password = "your_password"
- Change $files values with path of the ISAPI folder in InitialConnectionScript.ps1
