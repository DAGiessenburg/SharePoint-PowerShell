<#
.Synopsis
    Get Files and Folders amount
.DESCRIPTION
    Get amount of files and (sub)folders in a folder from SharePoint Online Library
.EXAMPLE
    Get-FilesAndFolderAmount -SiteUrl "https://[tenant].sharepoint.com/sites/[site]" -FolderURL "/[Library]/[Folder]" 
.INPUTS
    $SiteUrl               SharePoint Online Site collection url
    $FolderURL             Path form what folder the files and subfolders need to be counted
.OUTPUTS
    Amount of files
    Amount of folders
.NOTES
    Created on:   	2024-05-22
    Created by:   	Daniel Giessenburg
    Filename:     	SPO-PS-GetFilesAndFolderAmount.ps1
    Organization: 	GSSNBRG
    Version:        2024-05-22
.COMPONENT
    Project: Github\SharePoint-PowerShell
    Github: https://github.com/DAGiessenburg/SharePoint-PowerShell
#>


Function Get-FilesAndFolderAmount(){
    param([string]$SiteUrl,[string]$FolderURL)

    #Get-FilesAndFolderAmount
    Write-Host "Get-FilesAndFolderAmount" -ForegroundColor Yellow
            
    #Required parameters
    if(-not($SiteUrl)) { Throw "You must supply a value for SiteUrl" }
    Write-Host "- SiteUrl: "$SiteUrl  -ForegroundColor DarkYellow

    #Connect to SharePoint site
    Connect-PnPOnline -Url $SiteUrl -Interactive
    $connection = Get-PnPConnection

    #Get files form Folder
    $itemsFiles = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderURL -Recursive -ItemType File -Connection $connection 
    $itemsFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderURL -Recursive -ItemType Folder -Connection $connection 

    write-host $FolderURL -ForegroundColor Yellow
    write-host "Amount of Files: "$itemsFiles.count -ForegroundColor DarkYellow
    write-host "Amount of Folders: "$itemsFolders.count -ForegroundColor DarkYellow
}

#Get-FilesAndFolderAmount -SiteUrl "https://[tenant].sharepoint.com/sites/[site]" -FolderURL "/[Library]/[Folder]" 