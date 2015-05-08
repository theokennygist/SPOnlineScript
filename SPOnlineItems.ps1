. "C:\projects\SPOnline\SPOnlineLists.ps1"

#-------------------------------------------------------------------------------------------------------------------------------------- 
# Author   : Kenny Soetjipto
# Job Title: SharePoint Developer
# Company  : RXPServices
#--------------------------------------------------------------------------------------------------------------------------------------

#--------------------------------------------------------------------------------------------------------------------------------------
# Description:
# A file which contains Items operational
#--------------------------------------------------------------------------------------------------------------------------------------

#--------------------------------------------------------------------------------------------------------------------------------------
#GETSPLISTITEMS
#Is used to retrieve items in the column specified in columnName parameter
#Show 100 items. If need more or less, change it in CreateAllItemsQuery function
#print all items on the screen
#--------------------------------------------------------------------------------------------------------------------------------------

#--------------------------------------------------------------------------------------------------------------------------------------
# Functions:
# GETSPLISTITEMS, SHOWALLFOLDERS, GETALLFOLDERS, GETFOLDER, CREATEFOLDER, ADDFILE, ADDITEM, TESTLIST
#--------------------------------------------------------------------------------------------------------------------------------------


#--------------------------------------------------------------------------------------------------------------------------------------
# GETSPLISTITEMS
# function to show all folders in the list which is specified in parameter
# print all folders name
#--------------------------------------------------------------------------------------------------------------------------------------
function getSPListItems([string]$listName, [string] $columnName)
{
    $listName = "AnnouncementTest"
    $listObj = getSPList $listName
    $columnName = "Title"
    $query = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery(100, $columnName)
    $items = $listObj.GetItems($query)
    $cContext.Load($items)
    $cContext.ExecuteQuery()

    foreach($i in $items)
    {
        write-host $i["Title"]
    }
    
}

#--------------------------------------------------------------------------------------------------------------------------------------
# SHOWALLFOLDERS
# function to show all folders in the list which is specified in parameter
# print all folders name
#--------------------------------------------------------------------------------------------------------------------------------------
function showAllFolders($listName)
{
   $listName = "Documents"    
   $listColl = $gList.GetByTitle($listName)   

   $folders = $listColl.RootFolder.Folders
   $cContext.Load($folders)
   $cContext.ExecuteQuery()  
   
   $folders | select Name 
  
}
#--------------------------------------------------------------------------------------------------------------------------------------
# GETALLFOLDERS
# It will retrieve all folders object in list specified in parameter
# return Folders object
#--------------------------------------------------------------------------------------------------------------------------------------
function getAllFolders($listName)
{
   $listName = "Documents"    
   $listColl = $gList.GetByTitle($listName)   

   $folders = $listColl.RootFolder.Folders
   $cContext.Load($folders)
   $cContext.ExecuteQuery()
   return $folders  
}

#--------------------------------------------------------------------------------------------------------------------------------------
# GETFOLDER
# It will retrieve folder specified in folderName parameter
# return 1 folder result otherwise it will print error
#-------------------------------------------------------------------------------------------------------------------------------------- 
function getFolder($listName, $folderName)
{
    $listName = "Documents"
    $folderName = "FolderTest1"
    $listObj = $gList.GetByTitle($listName)

    $folders = $listObj.RootFolder.Folders
    $cContext.Load($folders)
    $cContext.ExecuteQuery()
    $folderResult = $folders | Where-Object{$_.Name -eq $folderName}

    if($folderResult -eq $null)
    {
        write-host "Folder is not exist in $listName List"
    }   
    
    return $folderResult
}

#--------------------------------------------------------------------------------------------------------------------------------------
# CREATEFOLDER
# Creates folder with name specified from the folderName parameter
# It will print out successful message if folder is created in the list
#--------------------------------------------------------------------------------------------------------------------------------------
function createFolder($listName, $folderName)
{
    #$listName = "Documents";
    #$folderName = "Test20";
    $listColl = $gList.GetByTitle($listName);
    $cContext.Load($listColl.RootFolder.Folders)
    $cContext.ExecuteQuery()

    $cContext.Load($listColl.RootFolder.Folders.Add($folderName))
    $listColl.Update()
    $cContext.ExecuteQuery()
    write-host "$folderName is created in the $listName" -BackgroundColor Cyan -ForegroundColor Black
}

#--------------------------------------------------------------------------------------------------------------------------------------
# ADDFILE
# Uploads file from local to SharePoint
# It will print successful message if the file uploaded to the SharePoint
#--------------------------------------------------------------------------------------------------------------------------------------  
function addFile([string]$listName, [string]$localFile)
{
    $listName = "Documents"

    if(testList $listName)
    {
        $localFile = Get-ChildItem "C:\projects\READ THIS FIRST123.docx";
        
        if($localFile.Exists)
        {
            $list = $gWeb.Lists.GetByTitle($listName)       
            $rootFolder = $list.RootFolder.ServerRelativeUrl            
            $localFileName = $localFile.Name
            $fileUrl = $rootFolder+ '/' + $localFileName
            [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($cContext, $fileUrl, $localFile.OpenRead(), $true);
        }
        else
        {
            write-host "File is not exist locally" -BackgroundColor Red
        }       
       
    }
    else
    {
        write-host "$listName List is not exist. Create the List" -BackgroundColor Red -ForegroundColor Black
    }
}

#--------------------------------------------------------------------------------------------------------------------------------------
# ADDITEM
# function to addItems into List specified in the paramater
#--------------------------------------------------------------------------------------------------------------------------------------
function addItem($listName, $columnName, $value)
{
    $listName = "AnnouncementTest"
    $columnName = "Title"
    $value = "Awesome Test2"

    $list = $gList.GetByTitle($listName)
    $cContext.Load($list)
    $cContext.ExecuteQuery()
    
    if(testList $listName)
    {
        $newItem = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $listItem = $list.AddItem($newItem)
        $listItem.set_item($columnName, $value);
        $listItem.update()
        $cContext.Load($listItem)
        $cContext.ExecuteQuery()
        Write-Host "Item is inserted into $listName List" -BackgroundColor Cyan
    }
    else
    {
        write-host "List is not exist. Create list first" -BackgroundColor Red
    }
    
}

#--------------------------------------------------------------------------------------------------------------------------------------
# TESTLIST
# Is used to test whether the list is exist or not
#--------------------------------------------------------------------------------------------------------------------------------------
function testList([string]$listName)
{
    #$listName = "Documents"    
    $lists = $gList | select -ExpandProperty Title
    $exists = ($lists -contains $listName)
    return $exists
}



