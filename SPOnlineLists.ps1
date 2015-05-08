. "C:\projects\SPOnline\InitialConnectionScript.ps1";
#-------------------------------------------------------------------------------------------------------------------------------------- 
# Author   : Kenny Soetjipto
# Job Title: SharePoint Developer
# Company  : RXPServices
#-------------------------------------------------------------------------------------------------------------------------------------- 

#--------------------------------------------------------------------------------------------------------------------------------------
# Description:
# This file contains all Lists operational
#--------------------------------------------------------------------------------------------------------------------------------------

#--------------------------------------------------------------------------------------------------------------------------------------
# Functions in this file:
# DOCUMENTLIST, PICTURELIST, CREATELIST, GETALLSPLIST, GETSPLIST, TESTLIST
#--------------------------------------------------------------------------------------------------------------------------------------

#--------------------------------------------------------------------------------------------------------------------------------------
# DOCUMENTLIST
# Function to create Document list
#--------------------------------------------------------------------------------------------------------------------------------------
function documentList([string]$listTitle)
{
    $templateType = "DocumentLibrary"
    createList $listTitle $templateType
}

#--------------------------------------------------------------------------------------------------------------------------------------
# PICTURELIST
# Function to create Picture List
#--------------------------------------------------------------------------------------------------------------------------------------
function pictureList([string]$listTitle)
{
    $templateType = "PictureLibrary"
    createList $listTitle $templateType
}

#--------------------------------------------------------------------------------------------------------------------------------------
# CREATELIST
# Function which create any type of Lists that are specified above such as Picture or Document
#--------------------------------------------------------------------------------------------------------------------------------------
function createList([string]$listTitle, [Microsoft.SharePoint.Client.ListTemplateType]$templateType = "genericList", [Microsoft.SharePoint.Client.QuickLaunchOptions]$quickLaunch = "DefaultValue")
{
    #$listTitle = "DocumentTest"
    
    $listResult = getSPList $listTitle
    
    #if the list is not exist, then create new one
    if($listResult.Title -eq $null)
    {
        $listCreationInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
        $listCreationInfo.TemplateType = $templateType
        $listCreationInfo.Title = $listTitle
        $listCreationInfo.QuickLaunchOption = $quickLaunch

        $list = $gWeb.Lists.Add($listCreationInfo)

        $cContext.ExecuteQuery()
        write-host "List $listTitle is created successfully" -ForegroundColor Black -BackgroundColor Cyan
    }
    else
    {
        write-host "The list with that name already exist" -BackgroundColor Red
    }
    
}

#--------------------------------------------------------------------------------------------------------------------------------------
# GETALLSPLIST
# get all SPLists as parameter specified
#--------------------------------------------------------------------------------------------------------------------------------------
function getAllSPList()
{
    $listColl = $gList;
    $listColl | select -Property Title
    return $listColl
}

#--------------------------------------------------------------------------------------------------------------------------------------
# GETSPLIST
# function to get specific List with return List type
#--------------------------------------------------------------------------------------------------------------------------------------
function getSPList($listNameParam)
{
    #$listNameParam = "Documents" 
    $listResult = $gList | where {$_.Title -eq $listNameParam}          
    return $listResult;
}

#--------------------------------------------------------------------------------------------------------------------------------------
#TESTLIST
#Is used to test whether the list is exist or not
#--------------------------------------------------------------------------------------------------------------------------------------
function testList([string]$listName)
{
    #$listName = "Documents"
    $lists = $gList | select -ExpandProperty Title
    $exists = ($lists -contains $listName)
    return $exists
}


