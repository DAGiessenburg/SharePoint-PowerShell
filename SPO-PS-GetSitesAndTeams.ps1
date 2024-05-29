<#
.SYNOPSIS
    Get SharePoint site collections & MS Teams environments
.DESCRIPTION
    Get a list of all the SharePoint site collections and MS Teams enviroments in the tenant
.EXAMPLE
    Get-FilesAndFolderAmount -AdminUrl "https://[tenant]-admin.sharepoint.com/"
.INPUTS
    $AdminUrl              SharePoint Online Admin Site collection url
.OUTPUTS
    List with SharePoint site collections and MS Teams enviroments
.NOTES
    Created on:   	2024-05-29
    Created by:   	Daniel Giessenburg
    Filename:     	SPO-PS-GetSitesAndTeams.ps1
    Organization: 	GSSNBRG
    Version:        2024-05-29
.REQUIRMENTS
    PowerShell 5.1 or 7.2
    PowerShell Module: PnP PowerShell | https://github.com/pnp/powershell
.COMPONENT
    Project: Github\SharePoint-PowerShell
    Github: https://github.com/DAGiessenburg/SharePoint-PowerShell
#>

Function Get-SitesAndTeams(){
    param([string]$AdminUrl,[string]$FilterResults,[string]$CsvExportPath)

    #Get-FilesAndFolderAmount
    Write-Host " Get-SitesAndTeams" -ForegroundColor Yellow
            
    #Required parameters
    if(-not($AdminUrl)) { Throw "You must supply a value for AdminUrl" }
    Write-Host "- AdminUrl: "$AdminUrl  -ForegroundColor DarkGreen
    Write-Host "- FilterResults: "$FilterResults  -ForegroundColor DarkYellow
    Write-Host "- CsvExportPath: "$CsvExportPath  -ForegroundColor DarkYellow

    #Connect to SharePoint site
    Connect-PnPOnline -Url $AdminUrl -Interactive
    $connection = Get-PnPConnection

    #If filter gave with function
    if($FilterResults){
        $sites = Get-PnPTenantSite -Detailed -Filter $FilterResults -Connection $connection | Select-Object -Property Title,Url,Template,Owner,IsHubSite,HubSiteId,GroupId,IsTeamsConnected,IsTeamsChannelConnected,LocaleId,LastContentModifiedDate,SharingCapability
    }else{
        #$sites = Get-PnPTenantSite -Detailed -Connection $connection | Select-Object -Property *
        $sites = Get-PnPTenantSite -Detailed -Connection $connection | Select-Object -Property Title,Url,Template,Owner,IsHubSite,HubSiteId,GroupId,IsTeamsConnected,IsTeamsChannelConnected,LocaleId,LastContentModifiedDate,SharingCapability
    }

    #If there are sites to process
    if($sites){
        $Data = @() #Data object to store information
        $i = 0; #Counter for progressbar
        $amountOfSites= $sites.count #amount of total sites to process

        #Process every site
        foreach($site in $sites){
            $i++
            #Start Progressbar / update bar every site
            Write-Progress -Activity "Processing sites:" -Status "Processing site $i/$amountOfSites" -PercentComplete ($i/$amountOfSites*100)

            #SiteType
            [string]$siteType = ""
            switch ($siteType) {
                "GROUP#0" { $siteType = "MS Teams" }
                "TEAMCHANNEL#1" { $siteType = "MS Teams channel" }
                "STS#3" { $siteType = "SharePoint Team site (no M365 Group)" }
                "SITEPAGEPUBLISHING#0" { $siteType = "SharePoint Communication site" }
                "RedirectSite#0" { $siteType = "SharePoint Redirect Site" }
                Default { $siteType = "Other" }
            }

            #Connected with Hubsite
            [string]$connectedHubSite = $site.HubSiteId
            [string]$connectedHubSiteName = ""
            If($connectedHubSite -ne "00000000-0000-0000-0000-000000000000"){
                $getHubName = Get-PnPTenantSite -Identity $connectedHubSite -Connection $connection | Select-Object -Property Title
                $connectedHubSiteName = $getHubName.Title
                $connectedHubSite = "Yes"
            }else{
                $connectedHubSite = "No"
            }

            #Connected with Group
            [string]$connectedGroup = $site.GroupId
            [string]$connectedGroupName = ""
            If($connectedGroup -ne "00000000-0000-0000-0000-000000000000"){
                $getGroupName = Get-PnPMicrosoft365Group -Identity $connectedGroup -Connection $connection | Select-Object -Property DisplayName
                $connectedGroupName = $getGroupName.DisplayName
                $connectedGroup = "Yes"
            }else{
                $connectedGroup = "No"
            }

            #Site Data into a object
            $siteData = [pscustomobject]@{
                "Site Name" = $site.Title
                "Site Url" = $site.Url
                "Type" = $siteType
                "Template Code" = $site.Template
                "LocaleId" = $site.LocaleId
                "Owner email" = $site.Owner
                "Site is Hubsite" = $site.IsHubSite
                "Connected with Hubsite" = $connectedHubSite
                "Connected with Hubsite name" = $connectedHubSiteName
                "Connected with Group" = $connectedGroup
                "Connected with Group name" = $connectedGroupName
                "Connected with MS Team" = $site.IsTeamsConnected
                "Connected with MS Team Channel" = $site.IsTeamsChannelConnected
                "Last Content Modified Date" = $site.LastContentModifiedDate
                "Sharing" = $site.SharingCapability
            }

            #Write-host $siteData           
            $Data += $siteData
        }
        
        #End progressbar
        Write-Progress -Activity "Processing sites:" -Completed

        #Show Data as Table in terminal
        $Data | Format-Table -AutoSize

        #Export Data to CSV file
        if($CsvExportPath){
            $Data | Export-Csv -Path $CsvExportPath -Delimiter ';' -NoTypeInformation
        }
    }
    
}

#Get-SitesAndTeams -AdminUrl "https://[tenant]-admin.sharepoint.com/" #-FilterResults "Url -like '/Project*'" -CsvExportPath "C:\temp\Export-SitesAndTeams.csv"
#Get-SitesAndTeams -AdminUrl "https://[tenant]-admin.sharepoint.com/" #-CsvExportPath "C:\temp\Export-SitesAndTeams.csv"