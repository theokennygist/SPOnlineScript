. "C:\projects\SPOnline\InitialConnectionScript.ps1";
#-------------------------------------------------------------------------------------------------------------------------------------- 
# Author   : Kenny Soetjipto
# Job Title: SharePoint Developer
# Company  : RXPServices
#-------------------------------------------------------------------------------------------------------------------------------------- 

#-------------------------------------------------------------------------------------------------------------------------------------- 
# Description: 
# file to create common site collections in SP online
# function to create sites with default quota of 500MB
# url is taken from the title
# common template name:
# team: STS#0
# Project: PROJECTSITE#0
# Publishing site: CMSPUBLISHING#0
# url for templates to lookup through links below: 
# http://blogs.technet.com/b/araviraj/archive/2008/06/18/sharepoint-templates-types.aspx
# http://www.funwithsharepoint.com/sharepoint-2013-site-templates-codes-for-powershell/
#----------------------------------------------------------------------------------------------- 

#--------------------------------------------------------------------------------------------------------------------------------------
# Functions:
# TEAM, PROJECT, TESTSITE, CREATESITE
#--------------------------------------------------------------------------------------------------------------------------------------

#--------------------------------------------------------------------------------------------------------------------------------------
#TEAM
#function to create a site with template of TEAM
#Provide sitename and owner values as parameters
#--------------------------------------------------------------------------------------------------------------------------------------
function teamSite([string]$siteName, [string]$owner)
{
    $templateType = "STS#0"
    $owner = "kenny@rxpservices.onmicrosoft.com";

    createSite $siteName $templateType $owner
}

#--------------------------------------------------------------------------------------------------------------------------------------
# PROJECT
#function to create a site with template of PROJECT
#Provide sitename and owner values as parameters
#--------------------------------------------------------------------------------------------------------------------------------------
function projectSite([string]$siteName, [string]$owner)
{    
    $templateType = "PROJECTSITE#0";
    #$owner = "kenny@rxpservices.onmicrosoft.com";
    
    createSite $siteName $templateType $owner     
}

#--------------------------------------------------------------------------------------------------------------------------------------
# TESTSITE
#Is used to test whether the site is exist in SP Online
#Is invoked in createSite function
#if exists return true else return false
#--------------------------------------------------------------------------------------------------------------------------------------
function testSite([string]$siteName)
{
    $siteName = "https://rxpservices.sharepoint.com/sites/teamtest"    
    $sites = Get-SPOSite
    $siteNames = $sites | select -ExpandProperty Url
    $exists = ($siteNames -contains $siteName)
    $exists    
    return $exists
}
#--------------------------------------------------------------------------------------------------------------------------------------
# CREATESITE
#A function which actally create the site in SP Online
#--------------------------------------------------------------------------------------------------------------------------------------
function createSite([string]$siteName, [string]$template, [string]$owner)
{
    openConnection
    
    #$siteName = "KennyProject"
    #$template = "PROJECTSITE#0"
    $stQuota = 500   
    $sitesList = Get-SPOSite
    $rootUrl = $sitesList[0];
    $rootUrlValue = $rootUrl | select -ExpandProperty url

    $url = $rootUrlValue+ "sites/" +$siteName  

    if(!(testSite($url)))
    { 
        New-SPOSite -Url $url -title $siteName -Owner $owner -Template $template -StorageQuota $stQuota
    }
    else
    {
        write-host "Site is already exist in SP online" -ForegroundColor Red -BackgroundColor Yellow
    }
}