. "C:\projects\SPOnline\SPOnlineScript.ps1"

#----------------------------------------------------------------------------- 
# Author   : Kenny Soetjipto
# Job Title: SharePoint Developer
# Company  : RXPServices
#----------------------------------------------------------------------------- 

#----------------------------------------------------------------------------- 
# Description: 
# This file contains any Site Columns operational
# It does not contain all columns. It contains common site columns
# If needed, copy and paste any functions below and change the FieldType value
#----------------------------------------------------------------------------- 

#--------------------------------------------------------------------------------------------------------------------------------------
# Functions available:
# TEXTSITECOLUMN, METADATASITECOLUMN, NUMBERSITECOLUMN, CHOICESITECOLUMN, CURRENCYSITECOLUMN
# DATETIMESITECOLUMN, USERSITECOLUMN, INSERTSITECOLUMN, SHOWALLCOLUMNS
#--------------------------------------------------------------------------------------------------------------------------------------

#--------------------------------------------------------------------------------------------------------------------------------------
# TEXTSITECOLUMN
# function to add [Text] SiteColumn
# it gives DisplayName and Name with same value(optional)
# It puts in Custom Column group(optional)
#--------------------------------------------------------------------------------------------------------------------------------------
function textSiteColumn([string]$columnName)
{

    #$columnName="KennyTitle3" I used this for test
    #$Name="KennyTitle3" I used this for test
    
    $Group='Custom Columns' #optional. Can be changed 
    $xmlField = "<Field Type='Text' DisplayName='$columnName' Name='$columnName' Required='False' MaxLength='255' Group='$Group'/>"
    
    #INVOKE insertSiteColumn function
    insertSiteColumn $columnName $xmlField
    
}

#--------------------------------------------------------------------------------------------------------------------------------------
# METADATASITECOLUMN
# function to add [Text] SiteColumn
# it gives DisplayName and Name with same value(optional)
# It puts in Custom Column group(optional)
#--------------------------------------------------------------------------------------------------------------------------------------
function metadataSiteColumn([string] $columnName)
{
    $Group='Custom Columns'
    $columnName = "KennyMetadata"
    $xmlField = "<Field Type='TaxonomyFieldType' DisplayName='$columnName' Name='$columnName' Required='False' MaxLength='255' Group='$Group'/>"
    insertSiteColumn $columnName $xmlField
}
#--------------------------------------------------------------------------------------------------------------------------------------
# NUMBERSITECOLUMN
# function to add [Number] SiteColumn
# it gives DisplayName and Name with same value(optional)
# It puts in Custom Column group(optional)
#--------------------------------------------------------------------------------------------------------------------------------------
function numberSiteColumn([string] $columnName)
{
    $Group='Custom Columns'
    $xmlField = "<Field Type='Number' DisplayName='$columnName' Name='$columnName' required='FALSE' Group='$Group' />"

    insertSiteColumn $columnName $xmlField
}
#--------------------------------------------------------------------------------------------------------------------------------------
# CHOICESITECOLUMN
# function to add [Choice] SiteColumn
# it gives DisplayName and Name with same value(optional)
# It puts in Custom Column group(optional)
# to use this:
# provide columnName 
# provide values for choices with ; for each choice
# example: choiceSiteColumn choiceColumn "choice1;choice2;choice3"
#--------------------------------------------------------------------------------------------------------------------------------------
function choiceSiteColumn([string] $columnName, [string]$values)
{
    $Group = "Custom Columns"
    $options = ""
    $valueArray = $values.Split(";")
    
    foreach($val in $valueArray)
    {
        $options = $options+ "<CHOICE>$val</CHOICE>"
    }

    $xmlField = "<Field Type='Choice' DisplayName='$columnName' Name='$columnName'  
        required='FALSE' Group='$Group'><CHOICES>$options</CHOICES> </Field>"

    insertSiteColumn $columnName $xmlField
}
#--------------------------------------------------------------------------------------------------------------------------------------
# CURRENCYSITECOLUMN
# function to add [Currency] SiteColumn
# it gives DisplayName and Name with same value(optional)
# It puts in Custom Column group(optional)
#--------------------------------------------------------------------------------------------------------------------------------------
function currencySiteColumn([string] $columnName)
{
    $Group = "Custom Columns"
    $xmlField = "<Field Type='Currency' DisplayName='$columnName' Name='$columnName' required='FALSE' Group='$Group' />"
    insertSiteColumn $columnName $xmlField
}
#--------------------------------------------------------------------------------------------------------------------------------------
# DATETIMESITECOLUMN
# function to add [DateTime] SiteColumn
# it gives DisplayName and Name with same value(optional)
# It puts in Custom Column group(optional)
#--------------------------------------------------------------------------------------------------------------------------------------
function dateTimeSiteColumn([string] $columnName)
{

    $Group = "Custom Columns"
    $xmlField = "<Field Type='DateTime' DisplayName='$columnName' Name='$columnName' required='FALSE' Group='$Group' />"
    insertSiteColumn $columnName $xmlField
}

#--------------------------------------------------------------------------------------------------------------------------------------
# USERSITECOLUMN
# function to add [UserSite] SiteColumn
# it gives DisplayName and Name with same value(optional)
# It puts in Custom Column group(optional)
#--------------------------------------------------------------------------------------------------------------------------------------
function userSiteColumn([string] $columnName, [string]$multi, [int] $selectionMode)
{
    $Group = "Custom Columns"
    $xmlField = "<Field Type='UserMulti' DisplayName='$columnName' Name='$columnName' StaticName='$fieldName' 
    UserSelectionScope='0' UserSelectionMode='$selectionMode' Sortable='FALSE' Required='FALSE' 
    Mult='$multi' Group='$Group'/>"
    insertSiteColumn $columnName $xmlField
}

#--------------------------------------------------------------------------------------------------------------------------------------
# TESTSITECOLUMN
# function to test whether site column exist or not before insert into SP Online
# it is used in the insertSiteColumn
#--------------------------------------------------------------------------------------------------------------------------------------
function testSiteColumn([string] $columnName)
{	
    $columns = $gWeb.Fields
    $cContext.Load($columns)
    $cContext.ExecuteQuery()
    $columnNames = $columns | select -ExpandProperty Title
    $exists = ($columnNames -contains $columnName)
    return $exists
}

#--------------------------------------------------------------------------------------------------------------------------------------
# INSERTSITECOLUMN
# function which insert site column into SP ONLINE
# this function is invoked in other type of site column functions
#--------------------------------------------------------------------------------------------------------------------------------------
function insertSiteColumn([string] $columnName, [string]$xmlField)
{
    if(!(testSiteColumn $columnName))
    {
        $columnOption = [Microsoft.SharePoint.Client.AddFieldOptions]::AddToNoContentType
        $siteColumn = $gField.AddFieldAsXml($xmlField, $true, $columnOption);
        $cContext.ExecuteQuery()

        write-host "Site Column $columnName is inserted into the site" -ForegroundColor black -BackgroundColor Green
    }
    else
    {
        write-host "Site column $columnName already exist" -ForegroundColor Red -BackgroundColor Yellow
    }
}
#--------------------------------------------------------------------------------------------------------------------------------------
# SHOWALLCOLUMNS
# get all columns from the list specified in parameter
#--------------------------------------------------------------------------------------------------------------------------------------
function showAllColumns($listName)
{
    $listName = "AnnouncementTest"
    $list = $gWeb.Lists.getByTitle($listName)
    $columns = $list.Fields
    $cContext.Load($columns)
    $cContext.ExecuteQuery()    

    foreach($item in $columns)
    {
	    write-host $item.Title
    }
}
