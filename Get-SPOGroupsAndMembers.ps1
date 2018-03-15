# Connect to SharePoint Online
$cred = Get-Credential
$adminSite = "https://amerikleen-admin.sharepoint.com/"
$rootSite =  "https://amerikleen.sharepoint.com/"
Connect-SPOService -Url $adminSite -Credential $cred

# Get all groups on the root site
$groups = (Get-SPOSiteGroup -Site $rootSite | Sort-Object LoginName)

# Get all members
foreach ($group in $groups){
    Get-SPOUser -Site $rootSite -Group $group.LoginName | Select-Object `
        @{Name="GroupName";Expression={$group.LoginName}}, `
        @{Name="User";Expression={$_.DisplayName}}, `
        @{Name="LoginName";Expression={$_.LoginName}}
}

# Disconnect from SharePoint Online
Disconnect-SPOService
