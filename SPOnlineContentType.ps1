. "C:\projects\SPOnline\InitialConnectionScript.ps1";

#-------------------------------------------------------------------------------------------------------------------------------------- 
# Author   : Kenny Soetjipto
# Job Title: SharePoint Developer
# Company  : RXPServices
#-------------------------------------------------------------------------------------------------------------------------------------- 

#--------------------------------------------------------------------------------------------------------------------------------------
# Description
# A file which contains all content types operational
# It contains global variables: gCTypes, ctypeID
#--------------------------------------------------------------------------------------------------------------------------------------

#--------------------------------------------------------------------------------------------------------------------------------------
# Functions:
# ENABLECT, GETALLCONTENTTYPES, GETALLSPLIST, CREATECONTENTTYPE, ADDCONTENTTYPETOLIST, TESTCONTENTTYPE, TESTLIST
#--------------------------------------------------------------------------------------------------------------------------------------

#--------------------------------------------------------------------------------------------------------------------------------------
# GLOBAL VARIABLE:
# gCTypes, ctypeID
#--------------------------------------------------------------------------------------------------------------------------------------

#--------------------------------------------------------------------------------------------------------------------------------------
#ENABLECT
#is used to enable content type for specified a list in parameter
#--------------------------------------------------------------------------------------------------------------------------------------
function enableCT($listName)
{
    #$listName = "Kenny"
    $docList = $gList.GetByTitle($listName)
    
    #if list is exist
    if(testList($docList))
    {
        $docList.ContentTypesEnabled = $true
        $docList.Update()
        $cContext.ExecuteQuery()        
    }
    else
    {
        Write-Host "The list is not exist" -BackgroundColor Yellow -ForegroundColor Red
    }
}

#--------------------------------------------------------------------------------------------------------------------------------------
#GETALLCONTENTTYPES
#is used to get all content types Name and ID
#$ctypeID has name and id and GLOBAL
#$gCTypes has contentTypes object and GLOBAL
#--------------------------------------------------------------------------------------------------------------------------------------
function getAllContentType()
{
    $Global:gCTypes = $gWeb.ContentTypes
    $cContext.Load($gCTypes)
    $cContext.ExecuteQuery()
    $Global:ctypeID = $gCTypes | select Name, Id
    $ctypeID  
}

#--------------------------------------------------------------------------------------------------------------------------------------
#GETALLSPLIST
#get SPList as parameter specified
#--------------------------------------------------------------------------------------------------------------------------------------
function getAllSPList()
{
    $listColl = $gList;
    $listColl | select -ExpandProperty Title
    return $listColl
}

#--------------------------------------------------------------------------------------------------------------------------------------
#CREATECONTENTTYPE
#is used to create new content type
#--------------------------------------------------------------------------------------------------------------------------------------
function createContentType([string] $name, [string] $parentCTName, [string] $description)
{
    #$parentCTName = "Document"
    #$description = "Test Desc"
    #$name = "KennyTest2"
    $group = "Custom Content Types"    

    #Test Parent Content Type. If parent content type is exist
    if(testContentType($parentCTName))
    {
        $ctsName = $ctypeID.GetEnumerator() | Where-Object {$_.Name -eq $parentCTName}
        $parentID = $ctsName.Id.ToString()

        $parentType = $gCTypes.GetById($parentID)
        $cContext.Load($parentType)
        $cContext.ExecuteQuery()
        
        if(!(testContentType($name)))
        {
            $newCT = New-Object Microsoft.SharePoint.Client.ContentTypeCreationInformation
            $newCT.Description = $description
            $newCT.Name = $name
            $newCT.Group = $group
            $newCT.ParentContentType = $parentType
        
            $ctReturn = $gCTypes.Add($newCT)
            $cContext.Load($ctReturn)
            $cContext.ExecuteQuery()
            Write-Host "Content type $name is created" -BackgroundColor Cyan -ForegroundColor Black 
        }
        else
        {
            write-host "The content type is exist" -BackgroundColor red
        }        
       
    }
    else #if Parent is not exist
    {
        write-host "Parent Content Type is not exist" -BackgroundColor red
    }    
}

#--------------------------------------------------------------------------------------------------------------------------------------
#ADDCONTENTTYPETOLIST
#Function to add contentType into List
#to Test, it's better to invoke CREATECONTENTTYPE function and use that content type in this function
#--------------------------------------------------------------------------------------------------------------------------------------
function addContentTypeToList([string]$listTitle, [string]$ctName)
{
    $ctName = "KennyTest2"
    $listTitle = "Documents"

    if(testContentType($ctName))
    {        
        if(testList($listTitle))
        {                                
            $ctResultName = $ctypeID.GetEnumerator() | Where-Object {$_.Name -eq $ctName}                       
            $ctResultID = $ctResultName.Id.ToString()

            $ctResult = $gWeb.ContentTypes.GetById($ctResultID)
            $cContext.Load($ctResult)
            $cContext.ExecuteQuery()
    
            $listResult = $gList.GetByTitle($listTitle)
            $listCTResults = $listResult.contentTypes
            $cContext.Load($listResult)
            $cContext.ExecuteQuery()

            $ctReturn = $listCTResults.AddExistingContentType($ctResult)
            $cContext.Load($ctReturn)
            $cContext.ExecuteQuery()
            Write-Host "content type $ctResultName.Name added to $listResult.Title" -BackgroundColor Cyan -ForegroundColor black
        }
        else
        {
            write-host "List is not exist. Create new List or choose existing List" -BackgroundColor Red
        }
        
    }
    else
    {
        write-host "Content Type is not exist. Create Content type first" -BackgroundColor Red
    }
    
}

#--------------------------------------------------------------------------------------------------------------------------------------
#TESTCONTENTTYPE
#Is used to test whether content type is exist or not
#--------------------------------------------------------------------------------------------------------------------------------------
function testContentType($ctName)
{
    #$ctName = "Document"
    $cts = $gWeb.ContentTypes | select -ExpandProperty Name
    $exists = ($cts -contains $ctName)
    return $exists
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
