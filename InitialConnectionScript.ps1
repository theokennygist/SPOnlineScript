#-------------------------------------------------------------------------------------------------------------------------------------- 
# Author   : Kenny Soetjipto
# Job Title: SharePoint Developer
# Company  : RXPServices
#-------------------------------------------------------------------------------------------------------------------------------------- 

#-------------------------------------------------------------------------------------------------------------------------------------- 
# Description: 
# This file contains initial global variable declaration
# It has connection function which connect to SP Online
# It has open connection which is needed to execute any -spo cmdlet
#-------------------------------------------------------------------------------------------------------------------------------------- 

#--------------------------------------------------------------------------------------------------------------------------------------
# Functions:
# CONNECTSPONLINE, OPENCONNECTION
#--------------------------------------------------------------------------------------------------------------------------------------
if ((Get-Module Microsoft.Online.SharePoint.PowerShell).Count -eq 0) 
{
    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
}

Set-ExecutionPolicy bypass -Force 

#include Password file into this script. This file path is outside the project folder
. "C:\projects\PassFile.ps1"

#array to hold DLL files path
[string []] $files = @("C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll",
            "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll",
            "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll",            
            "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.DocumentManagement.dll" 
            );

#iterate the DLL and check if the files are exist
foreach($file in $files)
{
    Get-ChildItem $file
    
    if(Test-Path $file)
    {
        Add-Type -Path $file
    }
    else
    {
        write-host "file is not exist. Download the DLL files and put into the C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI folder"
    }   
}

#--------------------------------------------------------------------------------------------------------------------------------------
#Global variables declaration
#--------------------------------------------------------------------------------------------------------------------------------------
$gUrl = "https://rxpservices.sharepoint.com/sites/teamtest" #CHANGE THIS URL
$secPass = ConvertTo-SecureString $password -AsPlainText -Force
$username = "kenny@rxpservices.onmicrosoft.com"; #CHANGE THIS USERNAME

#--------------------------------------------------------------------------------------------------------------------------------------
#Sharepoint Object declaration
#--------------------------------------------------------------------------------------------------------------------------------------
$cContext = New-Object Microsoft.SharePoint.Client.ClientContext($gUrl);
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $secPass)
$gQuery = New-Object Microsoft.SharePoint.Client.CamlQuery;


#--------------------------------------------------------------------------------------------------------------------------------------
#CONNECTSPONLINE
#function to add [Text] SiteColumn
#it gives DisplayName and Name with same value(optional)
#It puts in Custom Column group(optional)
#--------------------------------------------------------------------------------------------------------------------------------------
function connectSPOnline()
{

    if($username -eq $null -or $gUrl -eq $null)
    {
        write-host -backgroundColor DarkRed "Either username or password or URL is empty";
    }
    else
    { 
        $cContext.Credentials = $credentials
        
        if($cContext -ne $null)
        {            
            write-host -BackgroundColor Green "You are connected to SP Online";            
        }
        else
        {
            write-host -BackgroundColor red "Connection issue";
        } 
    }    
}
#--------------------------------------------------------------------------------------------------------------------------------------
#instantiate Web and Lists object
#invoke connection to make sure that is connected. Otherwise, it will give error
#--------------------------------------------------------------------------------------------------------------------------------------
connectSPOnline

$gWeb = $cContext.Web;
$gList = $cContext.Web.Lists;
$gField = $cContext.Web.Fields;

$cContext.load($gWeb)
$cContext.ExecuteQuery()
$cContext.load($gList)
$cContext.ExecuteQuery()
$cContext.Load($gField)
$cContext.ExecuteQuery()

#--------------------------------------------------------------------------------------------------------------------------------------
#OPENCONNECTION
#This function is used when SP Admin need to run -SPO cmdlet such as create site Collections
#--------------------------------------------------------------------------------------------------------------------------------------
function openConnection()
{
    $cred = get-credential $username
    $adminUrl = "https://rxpservices-admin.sharepoint.com"        
    Connect-SPOService -Url $adminUrl -Credential $cred
}